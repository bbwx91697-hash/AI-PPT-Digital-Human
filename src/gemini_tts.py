"""Gemini 3.8 TTS single-turn REST adapter."""

import base64
import io
import json
import urllib.error
import urllib.request
import wave

MODELS = ("gemini-3.8-flash-tts", "gemini-3.8-flash-lite-tts")
PREFERRED_VOICES = ("Leda", "Achird", "Aoede")
ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/interactions"


def synthesize(api_key, text, voice, style, pace, model=MODELS[0]):
    if model not in MODELS:
        raise ValueError("不支持的模型")
    if not api_key or not text.strip() or not voice:
        raise ValueError("请输入 API Key、台词并选择音色")
    payload = {
        "model": model,
        "input": [{
            "type": "user_input",
            "content": [{
                "type": "text",
                "text": text,
                "annotations": [{"type": "speech_metadata", "style": f"普通话，{style}，语速{pace}"}],
            }],
        }],
        "response_format": {"type": "audio"},
        "generation_config": {"speech_config": [{"voice": voice}]},
    }
    request = urllib.request.Request(
        ENDPOINT,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            result = json.loads(response.read())
    except urllib.error.HTTPError as error:
        body = error.read(65536).decode("utf-8", errors="replace")
        body = body.replace(api_key, "[已隐藏密钥]")
        raise RuntimeError(f"语音生成失败，HTTP {error.code}。服务端诊断：{body}") from None
    except urllib.error.URLError as error:
        raise RuntimeError("语音服务连接失败，请检查网络。") from None
    for step in reversed(result.get("steps", [])):
        for part in step.get("content", []):
            if part.get("type") == "audio" and part.get("data"):
                try:
                    audio = base64.b64decode(part["data"], validate=True)
                    with wave.open(io.BytesIO(audio), "rb") as wav:
                        if wav.getnframes() <= 0:
                            raise ValueError("语音服务返回了空音频")
                    return audio
                except (ValueError, wave.Error, EOFError):
                    raise ValueError("语音服务返回的音频格式无效") from None
    raise ValueError("语音服务没有返回音频")
