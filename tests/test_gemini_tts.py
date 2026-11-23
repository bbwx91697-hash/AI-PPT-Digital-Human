import base64
import importlib.util
import io
import urllib.error
import unittest
import wave
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location(
    "gemini_tts", Path(__file__).resolve().parents[1] / "src" / "gemini_tts.py"
)
tts = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tts)


class Response:
    def __init__(self, data):
        self.data = data

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def read(self):
        return self.data


class GeminiTests(unittest.TestCase):
    def test_transcript_is_verbatim_and_response_is_wav(self):
        import json

        buffer = io.BytesIO()
        with wave.open(buffer, "wb") as output:
            output.setparams((1, 2, 24000, 0, "NONE", "not compressed"))
            output.writeframes(b"\0\0" * 100)
        body = json.dumps({"steps": [{"content": [{"type": "audio", "data": base64.b64encode(buffer.getvalue()).decode()}]}]}).encode()
        with patch.object(tts.urllib.request, "urlopen", return_value=Response(body)) as call:
            wav = tts.synthesize("secret", "同学们好", "Kore", "温和", "较慢")
        request = call.call_args.args[0]
        payload = json.loads(request.data)
        part = payload["input"][0]["content"][0]
        self.assertEqual(part["text"], "同学们好")
        self.assertIn("较慢", part["annotations"][0]["style"])
        self.assertEqual(wav, buffer.getvalue())

    def test_empty_audio_is_rejected(self):
        with patch.object(tts.urllib.request, "urlopen", return_value=Response(b'{"steps": []}')):
            with self.assertRaisesRegex(ValueError, "音频"):
                tts.synthesize("secret", "你好", "Kore", "", "正常")

    def test_service_diagnostics_are_visible_and_key_is_redacted(self):
        error = urllib.error.HTTPError("https://example.test", 429, "rate limit", {}, io.BytesIO(
            b'{"error":{"status":"RESOURCE_EXHAUSTED","message":"quota exhausted secret"}}'
        ))
        with patch.object(tts.urllib.request, "urlopen", side_effect=error):
            with self.assertRaises(RuntimeError) as caught:
                tts.synthesize("secret", "你好", "Leda", "", "正常")
        self.assertIn("RESOURCE_EXHAUSTED", str(caught.exception))
        self.assertIn("quota exhausted", str(caught.exception))
        self.assertNotIn("secret", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
