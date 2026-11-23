"""Generate three test WAV files using a key read from standard input."""

import json
import sys
from pathlib import Path

from gemini_tts import MODELS, PREFERRED_VOICES, synthesize


def main():
    output = Path(sys.argv[1]).resolve()
    output.mkdir(parents=True, exist_ok=True)
    api_key = sys.stdin.readline().strip()
    if not api_key:
        raise SystemExit("未读取到 API Key。")
    script = "同学们，请观察页面中的评价方式，想一想还需要哪些证据？"
    results = []
    for voice in PREFERRED_VOICES:
        path = output / f"{voice}.wav"
        if path.exists():
            raise SystemExit(f"文件已存在，请选择空目录：{path}")
        try:
            wav = synthesize(api_key, script, voice, "自然、清晰", "适中", MODELS[0])
            path.write_bytes(wav)
            results.append({"voice": voice, "model": MODELS[0], "status": "成功", "bytes": len(wav)})
        except Exception as error:
            results.append({"voice": voice, "model": MODELS[0], "status": "失败", "error": str(error)})
    (output / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
