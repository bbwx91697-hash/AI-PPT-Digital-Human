import io
import tempfile
import unittest
import wave
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

import sys
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import openxml_merge
import timing_builder

P = f"{{{timing_builder.P}}}"


def _make_wav(duration_sec: float, rate: int = 24000) -> bytes:
    buf = io.BytesIO()
    with wave.open(buf, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(rate)
        w.writeframes(b"\x00\x00" * int(rate * duration_sec))
    return buf.getvalue()


class Stage4TimingTests(unittest.TestCase):
    def test_wav_duration_ms_reads_actual_length(self):
        wav = _make_wav(3.45)
        self.assertEqual(timing_builder.wav_duration_ms(wav), 3450)

    def test_generated_auto_and_manual_files_have_mutual_exclusion(self):
        auto_path = ROOT / "output/第四轮_自动状态修复.pptx"
        manual_path = ROOT / "output/第四轮_手动状态修复.pptx"
        self.assertTrue(auto_path.exists())
        self.assertTrue(manual_path.exists())

        with ZipFile(auto_path) as z:
            for page in (7, 8):
                root = ET.fromstring(z.read(f"ppt/slides/slide{page}.xml"))
                spids = {p.get("name"): p.get("id") for p in root.iter(P + "cNvPr")}
                av, idle = spids["AI Teaching Avatar"], spids["AI Teaching Idle"]
                tm_root = next(n for n in root.iter(P + "cTn") if n.get("nodeType") == "tmRoot")
                auto_par = tm_root.find(f"./{P}childTnLst/{P}par")
                self.assertIsNotNone(auto_par)
                # At t=0: Avatar visible, Idle hidden; at dur_ms: Idle visible, Avatar hidden
                sets = [(s.find(f".//{P}spTgt").get("spid"), s.find(f".//{P}strVal").get("val"))
                        for s in auto_par.iter(P + "set")]
                self.assertEqual(sets, [(av, "visible"), (idle, "hidden"), (idle, "visible"), (av, "hidden")])
                # Also retains interactive click fallbacks on both Avatar and Idle
                inter = [n for n in root.iter(P + "cTn") if n.get("nodeType") == "interactiveSeq"]
                self.assertEqual(len(inter), 2)

        with ZipFile(manual_path) as z:
            for page in (1, 2):
                root = ET.fromstring(z.read(f"ppt/slides/slide{page}.xml"))
                spids = {p.get("name"): p.get("id") for p in root.iter(P + "cNvPr")}
                av, idle = spids["AI Teaching Avatar"], spids["AI Teaching Idle"]
                inter = [n for n in root.iter(P + "cTn") if n.get("nodeType") == "interactiveSeq"]
                self.assertEqual(len(inter), 2)
                for seq in inter:
                    sets = [(s.find(f".//{P}spTgt").get("spid"), s.find(f".//{P}strVal").get("val"))
                            for s in seq.iter(P + "set")]
                    self.assertEqual(sets, [(av, "visible"), (idle, "hidden"), (idle, "visible"), (av, "hidden")])

    def test_generated_files_preserve_original_objects_and_unchanged_slides(self):
        cases = [
            (ROOT / "input/说课.pptx", ROOT / "output/第四轮_自动状态修复.pptx", {7, 8}),
            (ROOT / "tests/fixtures/second-source.pptx", ROOT / "output/第四轮_手动状态修复.pptx", {1, 2}),
        ]
        for source, target, target_pages in cases:
            with ZipFile(source) as old, ZipFile(target) as new:
                self.assertIsNone(new.testzip())
                paths = openxml_merge.slide_paths(old)
                for idx, path in enumerate(paths, 1):
                    if idx not in target_pages:
                        self.assertEqual(old.read(path), new.read(path), f"{target.name} slide {idx}")
                        continue
                    old_tree = [
                        ET.canonicalize(ET.tostring(el, encoding="unicode"))
                        for el in ET.fromstring(old.read(path)).find(f".//{P}spTree")
                    ]
                    new_tree = [
                        ET.canonicalize(ET.tostring(el, encoding="unicode"))
                        for el in ET.fromstring(new.read(path)).find(f".//{P}spTree")
                        if not any(p.get("name") in openxml_merge.MANAGED_NAMES for p in el.iter(P + "cNvPr"))
                    ]
                    self.assertEqual(old_tree, new_tree, f"{target.name} slide {idx} original shapes")

    def test_preserve_existing_animations_on_animated_slide(self):
        from xml.dom import minidom
        source = ROOT / "input/说课.pptx"
        with ZipFile(source) as old:
            doc = minidom.parseString(old.read("ppt/slides/slide1.xml"))
        old_timing = doc.getElementsByTagNameNS(timing_builder.P, "timing")[0]
        old_main_xml = old_timing.getElementsByTagNameNS(timing_builder.P, "cTn")
        old_main = next(n for n in old_main_xml if n.getAttribute("nodeType") == "mainSeq")
        old_main_canon = ET.canonicalize(f'<w xmlns:p="{timing_builder.P}">{old_main.toxml()}</w>')

        merged = timing_builder.build_timing_element(
            doc, mode="manual", avatar_spid="9001", audio_spid="9002", idle_spid="9003",
            dur_ms=4200, existing_timing=old_timing,
        )
        new_main = next(
            n for n in merged.getElementsByTagNameNS(timing_builder.P, "cTn")
            if n.getAttribute("nodeType") == "mainSeq"
        )
        self.assertEqual(old_main_canon, ET.canonicalize(f'<w xmlns:p="{timing_builder.P}">{new_main.toxml()}</w>'))
        inter = [
            n for n in merged.getElementsByTagNameNS(timing_builder.P, "cTn")
            if n.getAttribute("nodeType") == "interactiveSeq"
        ]
        self.assertEqual(len(inter), 2)


if __name__ == "__main__":
    unittest.main()
