"""Local Qwen3-TTS inference worker executed inside venv_qwen_tts."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import traceback
from pathlib import Path


def _reserve_ipc_stdout():
    """Redirect fd 1 (stdout) to stderr and force UTF-8 streams for clean JSON IPC."""
    ipc_fd = os.dup(1)
    os.dup2(2, 1)
    sys.stdout = sys.stderr
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    return os.fdopen(ipc_fd, "w", encoding="utf-8", buffering=1)


def _load_model(model_dir: Path):
    import torch
    from qwen_tts import Qwen3TTSModel

    use_cuda = torch.cuda.is_available()
    kwargs = {
        "device_map": "cuda:0" if use_cuda else "cpu",
        "dtype": torch.bfloat16 if use_cuda else torch.float32,
        "attn_implementation": "sdpa",
    }
    return Qwen3TTSModel.from_pretrained(str(model_dir), **kwargs), use_cuda


def _sanitize_text(text: object) -> str:
    """Coerce input to clean printable UTF-8 str without lone surrogates or control chars."""
    if not isinstance(text, str):
        text = str(text)
    # Strip invisible/control characters that can break tokenizers or TTS alignment
    cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ufeff\u200b-\u200d]", "", text)
    return cleaned.encode("utf-8", errors="ignore").decode("utf-8").strip()


def _estimate_max_new_tokens(text: str) -> int:
    """Bound codec steps to ~2x conservative speaking length so missed EOS cannot hang."""
    return max(96, min(1200, len(text.strip()) * 6 + 64))


def _split_text_chunks(text: str, max_chars: int = 80) -> list[str]:
    """Split long slide scripts on Chinese/English sentence boundaries."""
    cleaned = re.sub(r"\s+", " ", _sanitize_text(text))
    if not cleaned:
        return []
    if len(cleaned) <= max_chars:
        return [cleaned]
    parts = re.split(r"(?<=[。！？；!?;\n])", cleaned)
    chunks: list[str] = []
    buf = ""
    for part in parts:
        piece = part.strip()
        if not piece:
            continue
        if buf and len(buf) + len(piece) > max_chars:
            chunks.append(buf)
            buf = piece
        else:
            buf = f"{buf}{piece}" if not buf else f"{buf} {piece}" if piece[0].isascii() else f"{buf}{piece}"
    if buf:
        chunks.append(buf)
    return [c for c in (chunks or [cleaned]) if isinstance(c, str) and c]


def _synthesize_one(
    model,
    *,
    text: str,
    speaker: str,
    language: str,
    instruct: str | None,
    out_path: Path,
    max_new_tokens: int | None = None,
) -> dict:
    import numpy as np
    import soundfile as sf

    clean_text = _sanitize_text(text)
    if not clean_text:
        raise ValueError("请输入数字人台词。")

    chunks = _split_text_chunks(clean_text)
    if not chunks:
        raise ValueError("请输入有效数字人台词。")
    audio_parts: list[np.ndarray] = []
    sample_rate = 24000

    for idx, chunk in enumerate(chunks):
        token_cap = int(max_new_tokens) if max_new_tokens else _estimate_max_new_tokens(chunk)
        kwargs = {
            "text": chunk,
            "language": language or "Chinese",
            "speaker": speaker or "serena",
            "max_new_tokens": token_cap,
            "temperature": 0.75,
            "top_p": 0.9,
            "top_k": 40,
            "repetition_penalty": 1.08,
        }
        if instruct and instruct.strip():
            kwargs["instruct"] = instruct.strip()

        wavs, sr = model.generate_custom_voice(**kwargs)
        if not wavs:
            raise RuntimeError("Qwen3-TTS 未返回音频数据。")
        sample_rate = int(sr)
        audio_parts.append(np.asarray(wavs[0], dtype=np.float32))
        if idx < len(chunks) - 1:
            audio_parts.append(np.zeros(int(sample_rate * 0.12), dtype=np.float32))

    # Pad 0.30s leading / 0.12s trailing silence so OS/Bluetooth audio wake-up never clips first word
    lead_pad = np.zeros(int(sample_rate * 0.30), dtype=np.float32)
    tail_pad = np.zeros(int(sample_rate * 0.12), dtype=np.float32)
    combined = np.concatenate([lead_pad, *audio_parts, tail_pad])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    sf.write(str(out_path), combined, sample_rate, subtype="PCM_16")
    dur = len(combined) / float(sample_rate)
    return {
        "ok": True,
        "out_path": str(out_path),
        "sr": sample_rate,
        "duration_sec": round(dur, 4),
    }


def _serve_loop(model, ipc_out, use_cuda: bool = True) -> int:
    ipc_out.write(json.dumps({"ready": True, "cuda": use_cuda}, ensure_ascii=False) + "\n")
    ipc_out.flush()
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
            if req.get("cmd") == "quit":
                ipc_out.write(json.dumps({"ok": True, "bye": True}, ensure_ascii=False) + "\n")
                ipc_out.flush()
                break
            res = _synthesize_one(
                model,
                text=req["text"],
                speaker=req.get("speaker", "serena"),
                language=req.get("language", "Chinese"),
                instruct=req.get("instruct"),
                out_path=Path(req["out_path"]),
                max_new_tokens=req.get("max_new_tokens"),
            )
        except Exception as exc:  # noqa: BLE001
            tb_text = traceback.format_exc()
            try:
                log_path = Path(__file__).resolve().parents[1] / ".cache" / "qwen3_worker_error.log"
                log_path.parent.mkdir(parents=True, exist_ok=True)
                log_path.write_text(
                    f"req_keys={sorted(req.keys()) if isinstance(req, dict) else type(req)}\n"
                    f"text_type={type(req.get('text')).__name__ if isinstance(req, dict) else None}\n"
                    f"{tb_text}",
                    encoding="utf-8",
                )
            except Exception:  # noqa: BLE001
                pass
            tb_lines = [ln.strip() for ln in tb_text.strip().splitlines() if ln.strip()]
            where = tb_lines[-2] if len(tb_lines) >= 2 else ""
            res = {"ok": False, "error": f"{type(exc).__name__}: {exc}" + (f" ({where})" if where else "")}
        ipc_out.write(json.dumps(res, ensure_ascii=False) + "\n")
        ipc_out.flush()
    return 0


def main() -> int:
    ipc_out = _reserve_ipc_stdout()
    parser = argparse.ArgumentParser(description="Qwen3-TTS local worker")
    parser.add_argument("--model-dir", required=True)
    parser.add_argument("--serve", action="store_true")
    parser.add_argument("--text", default="")
    parser.add_argument("--speaker", default="serena")
    parser.add_argument("--language", default="Chinese")
    parser.add_argument("--instruct", default="")
    parser.add_argument("--out", default="")
    args = parser.parse_args()

    model_dir = Path(args.model_dir)
    if not model_dir.is_dir():
        ipc_out.write(json.dumps({"ok": False, "error": f"找不到模型目录: {model_dir}"}, ensure_ascii=False) + "\n")
        return 1

    model, use_cuda = _load_model(model_dir)
    if args.serve:
        return _serve_loop(model, ipc_out, use_cuda=use_cuda)

    if not args.out:
        ipc_out.write(json.dumps({"ok": False, "error": "缺少 --out 输出路径"}, ensure_ascii=False) + "\n")
        return 1

    res = _synthesize_one(
        model,
        text=args.text,
        speaker=args.speaker,
        language=args.language,
        instruct=args.instruct or None,
        out_path=Path(args.out),
    )
    ipc_out.write(json.dumps(res, ensure_ascii=False) + "\n")
    ipc_out.flush()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
