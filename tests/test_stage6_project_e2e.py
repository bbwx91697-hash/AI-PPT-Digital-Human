import hashlib
import py_compile
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import audio_manager
import openxml_merge
import ppt_inspector
import project_store
import timing_builder

P = f"{{{openxml_merge.P}}}"


class Stage6ProjectTests(unittest.TestCase):
    def test_app_py_compiles_and_imports(self):
        py_compile.compile(str(ROOT / "src/app.py"), doraise=True)
        import app  # noqa: F401

    def test_inspect_detects_animated_and_clean_slides(self):
        info = ppt_inspector.inspect_pptx(ROOT / "input/说课.pptx")
        self.assertEqual(info.slide_count, 16)
        self.assertTrue(info.slides[0].has_existing_animations)
        self.assertFalse(info.slides[6].has_existing_animations)

    def test_extract_idle_frame_from_default_gif(self):
        with tempfile.TemporaryDirectory() as temp:
            idle_png = Path(temp) / "idle.png"
            project_store.extract_idle_frame(ROOT / "assets/avatar/default_avatar.gif", idle_png)
            self.assertEqual(idle_png.read_bytes(), (ROOT / "assets/avatar/default_avatar_idle.png").read_bytes())

    def test_e2e_export_two_decks_with_real_gemini_wavs_and_custom_positions(self):
        leda = ROOT / "generated/audio/gemini_voice_test/Leda.wav"
        achird = ROOT / "generated/audio/gemini_voice_test/Achird.wav"
        aoede = ROOT / "generated/audio/gemini_voice_test/Aoede.wav"
        self.assertTrue(leda.is_file() and achird.is_file() and aoede.is_file())

        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            # Deck 1: 说课.pptx (auto mode, pages 5[animated], 7, 8 with distinct positions & WAVs)
            pdir1, data1, info1 = project_store.create_or_load_project(ROOT / "input/说课.pptx", projects_root=tmp / "projects")
            data1["mode"] = "auto"
            configs1 = {
                5: (leda, 0.03, 0.60, 0.16, 0.35),
                7: (achird, 0.81, 0.60, 0.16, 0.35),
                8: (aoede, 0.81, 0.08, 0.16, 0.35),
            }
            for slide, (wav_p, l, t, w, h) in configs1.items():
                cfg = data1["pages"][str(slide)]
                cfg["enabled"] = True
                cfg["audio_path"] = str(wav_p)
                cfg["left_pct"], cfg["top_pct"], cfg["width_pct"], cfg["height_pct"] = l, t, w, h
            out1 = tmp / "说课_Stage6_自动实测.pptx"
            project_store.export_presentation(pdir1, data1, info1, out1)
            self.assertTrue(out1.is_file())

            with ZipFile(ROOT / "input/说课.pptx") as old, ZipFile(out1) as new:
                self.assertIsNone(new.testzip())
                for s in range(1, 17):
                    if s not in configs1:
                        self.assertEqual(old.read(f"ppt/slides/slide{s}.xml"), new.read(f"ppt/slides/slide{s}.xml"))
                # Verify slide 5 kept its original mainSeq intact
                old5 = ET.fromstring(old.read("ppt/slides/slide5.xml"))
                new5 = ET.fromstring(new.read("ppt/slides/slide5.xml"))
                old_main5 = next(n for n in old5.iter(P + "cTn") if n.get("nodeType") == "mainSeq")
                new_main5 = next(n for n in new5.iter(P + "cTn") if n.get("nodeType") == "mainSeq")
                self.assertEqual(
                    ET.canonicalize(ET.tostring(old_main5, encoding="unicode")),
                    ET.canonicalize(ET.tostring(new_main5, encoding="unicode")),
                )
                # Verify real Gemini WAV durations are embedded in timing XML
                for s, (wav_p, *_rest) in configs1.items():
                    expected_ms = str(timing_builder.wav_duration_ms(wav_p.read_bytes()))
                    slide_xml = new.read(f"ppt/slides/slide{s}.xml").decode("utf-8")
                    self.assertIn(f'dur="{expected_ms}"', slide_xml)

            # Deck 2: second-source.pptx (manual mode, pages 1 & 2)
            pdir2, data2, info2 = project_store.create_or_load_project(ROOT / "tests/fixtures/second-source.pptx", projects_root=tmp / "projects")
            data2["mode"] = "manual"
            data2["pages"]["1"]["enabled"] = True
            data2["pages"]["1"]["audio_path"] = str(leda)
            data2["pages"]["2"]["enabled"] = True
            data2["pages"]["2"]["audio_path"] = str(aoede)
            out2 = tmp / "second_Stage6_手动实测.pptx"
            project_store.export_presentation(pdir2, data2, info2, out2)
            self.assertTrue(out2.is_file())
            with ZipFile(out2) as new2:
                self.assertIsNone(new2.testzip())
                gifs = [n for n in new2.namelist() if n.endswith(".gif")]
                wavs = [n for n in new2.namelist() if n.endswith(".wav")]
                self.assertEqual(len(gifs), 1)
                self.assertEqual(len(wavs), 2)
                self.assertEqual(
                    hashlib.sha256(new2.read(gifs[0])).digest(),
                    hashlib.sha256((ROOT / "assets/avatar/default_avatar.gif").read_bytes()).digest(),
                )


if __name__ == "__main__":
    unittest.main()
