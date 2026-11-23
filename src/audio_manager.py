"""Stage 5/6 Audio cache, local WAV/MP3 import, and Windows audition helper."""

from __future__ import annotations

import hashlib
import io
import json
import shutil
import subprocess
import wave
from pathlib import Path

try:
    import winsound
except ImportError:
    winsound = None

ROOT = Path(__file__).resolve().parents[1]
BUNDLED_FFMPEG = ROOT / "bin" / "ffmpeg.exe"
DEFAULT_FFMPEG = Path("D:/ffmpeg/bin/ffmpeg.exe")
SUPPORTED_AUDIO_EXTS = {".wav", ".mp3", ".m4a", ".ogg", ".flac", ".aac", ".wma", ".mp4", ".mov", ".mkv"}


def find_ffmpeg() -> str | None:
    if BUNDLED_FFMPEG.is_file():
        return str(BUNDLED_FFMPEG)
    if DEFAULT_FFMPEG.is_file():
        return str(DEFAULT_FFMPEG)
    found = shutil.which("ffmpeg")
    if found:
        return found
    try:
        import imageio_ffmpeg
        exe = imageio_ffmpeg.get_ffmpeg_exe()
        if exe and Path(exe).is_file():
            return exe
    except Exception:  # noqa: BLE001
        pass
    return None


def validate_wav_bytes(data: bytes) -> float:
    """Validate non-empty WAV payload and return duration in seconds."""
    try:
        with wave.open(io.BytesIO(data), "rb") as reader:
            frames = reader.getnframes()
            rate = reader.getframerate()
            if frames <= 0 or rate <= 0:
                raise ValueError("音频文件为空。")
            return frames / float(rate)
    except (wave.Error, EOFError) as exc:
        raise ValueError("音频不是有效的 WAV 格式。") from exc


def cache_key(
    slide: int,
    script: str,
    voice: str,
    style: str,
    pace: str,
    model: str,
    engine: str = "gemini",
) -> str:
    data = {
        "engine": engine,
        "slide": int(slide),
        "script": script.strip(),
        "voice": voice,
        "style": style.strip(),
        "pace": pace.strip(),
        "model": model,
    }
    if engine == "qwen3":
        data["qwen3_pad_v"] = 2
    payload = json.dumps(data, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


class PageAudioCache:
    """Per-project, per-page WAV cache so editing one page keeps other pages valid."""

    def __init__(self, project_dir: Path):
        self.project_dir = Path(project_dir)
        self.audio_dir = self.project_dir / "audio"
        self.audio_dir.mkdir(parents=True, exist_ok=True)
        self.index_path = self.audio_dir / "cache_index.json"
        self.index: dict[str, dict] = self._load_index()

    def _load_index(self) -> dict[str, dict]:
        if self.index_path.is_file():
            try:
                return json.loads(self.index_path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                return {}
        return {}

    def _save_index(self) -> None:
        self.index_path.write_text(json.dumps(self.index, ensure_ascii=False, indent=2), encoding="utf-8")

    def get_cached_tts(
        self, slide: int, script: str, voice: str, style: str, pace: str, model: str, engine: str = "gemini",
    ) -> Path | None:
        entry = self.index.get(str(slide))
        if not entry or entry.get("source") != "tts":
            return None
        expected = cache_key(slide, script, voice, style, pace, model, engine=engine)
        if entry.get("hash") != expected:
            return None
        wav_path = self.audio_dir / entry.get("file", "")
        if wav_path.is_file() and wav_path.stat().st_size > 44:
            return wav_path
        return None

    def put_tts(
        self, slide: int, script: str, voice: str, style: str, pace: str, model: str, wav_bytes: bytes, engine: str = "gemini",
    ) -> Path:
        duration = validate_wav_bytes(wav_bytes)
        key = cache_key(slide, script, voice, style, pace, model, engine=engine)
        filename = f"slide_{int(slide):03d}_{key}.wav"
        wav_path = self.audio_dir / filename
        wav_path.write_bytes(wav_bytes)
        self.index[str(slide)] = {
            "source": "tts",
            "hash": key,
            "file": filename,
            "duration_sec": round(duration, 3),
        }
        self._save_index()
        return wav_path

    def import_local_audio(self, slide: int, source_file: Path) -> Path:
        src = Path(source_file)
        if not src.is_file():
            raise FileNotFoundError(f"找不到音频文件: {src}")
        ext = src.suffix.lower()
        if ext not in SUPPORTED_AUDIO_EXTS:
            raise ValueError(f"不支持的音频格式: {ext}")
        raw_hash = hashlib.sha256(src.read_bytes()).hexdigest()[:16]
        filename = f"slide_{int(slide):03d}_local_{raw_hash}.wav"
        dest = self.audio_dir / filename
        if ext == ".wav":
            wav_bytes = src.read_bytes()
            duration = validate_wav_bytes(wav_bytes)
            dest.write_bytes(wav_bytes)
        else:
            ffmpeg = find_ffmpeg()
            if not ffmpeg:
                raise RuntimeError("未找到 FFmpeg，无法转换非 WAV 音频，请提供 WAV 文件。")
            cmd = [
                ffmpeg, "-y", "-i", str(src),
                "-ac", "1", "-ar", "24000", "-sample_fmt", "s16",
                str(dest),
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, check=False)
            if res.returncode != 0 or not dest.is_file():
                raise RuntimeError(f"音频转码失败: {res.stderr.strip()}")
            duration = validate_wav_bytes(dest.read_bytes())
        self.index[str(slide)] = {
            "source": "local",
            "hash": raw_hash,
            "original_name": src.name,
            "file": filename,
            "duration_sec": round(duration, 3),
        }
        self._save_index()
        return dest


def play_wav_async(wav_path: Path) -> None:
    if winsound is None:
        raise RuntimeError("当前平台不支持 winsound 试听。")
    path = Path(wav_path)
    if not path.is_file():
        raise FileNotFoundError(f"音频文件不存在: {path}")
    winsound.PlaySound(str(path), winsound.SND_FILENAME | winsound.SND_ASYNC)


def stop_playback() -> None:
    if winsound is not None:
        winsound.PlaySound(None, winsound.SND_PURGE)
