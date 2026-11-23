import io
import subprocess
import sys
import tempfile
import unittest
import wave
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import audio_manager
import tts_service


def _wav_bytes(seconds: float = 1.2, rate: int = 24000) -> bytes:
    buf = io.BytesIO()
    with wave.open(buf, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(b"\x01\x00" * int(rate * seconds))
    return buf.getvalue()


class Stage5AudioTests(unittest.TestCase):
    def test_dpapi_key_roundtrip_without_plaintext(self):
        with tempfile.TemporaryDirectory() as temp:
            key_path = Path(temp) / "gemini_key.dpapi"
            secret = "test-secret-key-12345"
            tts_service.save_api_key(secret, key_path)
            raw = key_path.read_text(encoding="utf-8", errors="ignore")
            self.assertNotIn(secret, raw)
            self.assertEqual(tts_service.load_api_key(key_path), secret)

    def test_per_page_cache_invalidation_only_affects_edited_slide(self):
        with tempfile.TemporaryDirectory() as temp:
            cache = audio_manager.PageAudioCache(Path(temp))
            p7 = cache.put_tts(7, "第7页问题", "Leda", "自然", "适中", "gemini-3.8-flash-tts", _wav_bytes(1.0))
            p8 = cache.put_tts(8, "第8页问题", "Aoede", "自然", "适中", "gemini-3.8-flash-tts", _wav_bytes(1.5))
            self.assertEqual(cache.get_cached_tts(7, "第7页问题", "Leda", "自然", "适中", "gemini-3.8-flash-tts"), p7)
            self.assertEqual(cache.get_cached_tts(8, "第8页问题", "Aoede", "自然", "适中", "gemini-3.8-flash-tts"), p8)
            # Editing slide 7 misses slide 7 cache while slide 8 stays cached
            self.assertIsNone(cache.get_cached_tts(7, "修改后的第7页问题", "Leda", "自然", "适中", "gemini-3.8-flash-tts"))
            self.assertEqual(cache.get_cached_tts(8, "第8页问题", "Aoede", "自然", "适中", "gemini-3.8-flash-tts"), p8)

    def test_local_wav_and_mp3_import(self):
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            cache = audio_manager.PageAudioCache(tmp / "proj")
            src_wav = tmp / "teacher.wav"
            src_wav.write_bytes(_wav_bytes(1.4))
            imported = cache.import_local_audio(3, src_wav)
            self.assertTrue(imported.is_file())
            self.assertAlmostEqual(audio_manager.validate_wav_bytes(imported.read_bytes()), 1.4, places=2)

            ffmpeg = audio_manager.find_ffmpeg()
            if ffmpeg:
                src_mp3 = tmp / "teacher.mp3"
                subprocess.run([ffmpeg, "-y", "-i", str(src_wav), str(src_mp3)], capture_output=True, check=True)
                imported_mp3 = cache.import_local_audio(4, src_mp3)
                self.assertTrue(imported_mp3.is_file())
                self.assertGreater(audio_manager.validate_wav_bytes(imported_mp3.read_bytes()), 1.0)

    def test_edge_tts_and_windows_offline_tts_generate_valid_wav(self):
        edge_wav = tts_service.synthesize_speech(
            "同学们，请思考为什么过程评价需要贯穿课前、课中和课后。",
            engine="edge",
            voice="zh-CN-XiaoxiaoNeural",
            pace="适中",
        )
        self.assertGreater(audio_manager.validate_wav_bytes(edge_wav), 1.5)

        win_wav = tts_service.synthesize_speech(
            "测试离线语音。",
            engine="windows",
            voice="Microsoft Huihui Desktop",
            pace="适中",
        )
        self.assertGreater(audio_manager.validate_wav_bytes(win_wav), 0.3)

    def test_cache_isolates_different_engines(self):
        with tempfile.TemporaryDirectory() as temp:
            cache = audio_manager.PageAudioCache(Path(temp))
            cache.put_tts(1, "你好", "zh-CN-XiaoxiaoNeural", "", "适中", "", _wav_bytes(1.0), engine="edge")
            self.assertIsNotNone(cache.get_cached_tts(1, "你好", "zh-CN-XiaoxiaoNeural", "", "适中", "", engine="edge"))
            self.assertIsNone(cache.get_cached_tts(1, "你好", "zh-CN-XiaoxiaoNeural", "", "适中", "", engine="windows"))
            self.assertIsNone(cache.get_cached_tts(1, "你好", "zh-CN-XiaoxiaoNeural", "", "适中", "", engine="qwen3"))

    def test_qwen3_local_tts_warm_worker_generates_valid_wav(self):
        import qwen3_worker

        self.assertTrue(tts_service.is_qwen3_available())
        self.assertEqual(tts_service.default_engine(), "qwen3")
        self.assertEqual(tts_service._normalize_qwen3_voice("vivian · Vivian · 明亮女声"), "vivian")
        self.assertLessEqual(qwen3_worker._estimate_max_new_tokens("各位老师好！"), 150)
        self.assertEqual(qwen3_worker._sanitize_text("\ufeff测试\x00\u200b台词\x0b！"), "测试台词！")
        chunks = qwen3_worker._split_text_chunks("第一句介绍背景。第二句说明转向器壳体结构。" * 4, max_chars=50)
        self.assertGreater(len(chunks), 1)
        try:
            wav1 = tts_service.synthesize_speech(
                "各位老师好，今天我汇报的题目是转向器壳体轴承压装机设计。",
                engine="qwen3",
                voice="serena",
                style="温和、亲切、清晰",
                pace="适中",
            )
            self.assertGreater(audio_manager.validate_wav_bytes(wav1), 1.5)

            wav2 = tts_service.synthesize_speech(
                "请同学们回答以下问题。",
                engine="qwen3",
                voice="vivian",
                style="干练、清晰",
                pace="偏快",
            )
            self.assertGreater(audio_manager.validate_wav_bytes(wav2), 1.0)
            import struct
            with wave.open(io.BytesIO(wav2), "rb") as w:
                lead_samples = struct.unpack("<" + "h" * int(w.getframerate() * 0.30), w.readframes(int(w.getframerate() * 0.30)))
            self.assertLess(max(abs(x) for x in lead_samples), 150)
        finally:
            tts_service.shutdown_qwen3_worker()

    def test_qwen3_worker_manager_timeout_kills_hung_subprocess(self):
        mgr = tts_service._Qwen3WorkerManager()
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            fake_worker = tmp / "hung_worker.py"
            fake_worker.write_text(
                "import json, sys, time\n"
                "print(json.dumps({'ready': True}), flush=True)\n"
                "for _ in sys.stdin:\n"
                "    time.sleep(10)\n",
                encoding="utf-8",
            )
            fake_model = tmp / "model"
            fake_model.mkdir()
            (fake_model / "model.safetensors").write_bytes(b"0")
            try:
                with self.assertRaises(TimeoutError):
                    mgr.request(
                        {"text": "测试超时"},
                        python_path=Path(sys.executable),
                        model_dir=fake_model,
                        worker_script=fake_worker,
                        timeout_sec=0.8,
                    )
                self.assertIsNone(mgr._proc)
            finally:
                mgr.shutdown()


if __name__ == "__main__":
    unittest.main()
