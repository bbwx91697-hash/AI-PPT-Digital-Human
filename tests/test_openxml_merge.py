import importlib.util
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("openxml_merge", ROOT / "src/openxml_merge.py")
merge = importlib.util.module_from_spec(spec)
spec.loader.exec_module(merge)


class MergeTests(unittest.TestCase):
    def test_original_parts_are_preserved_and_media_embedded(self):
        source = ROOT / "tests/fixtures/second-source.pptx"
        donor = ROOT / "tests/fixtures/round2-manual.pptx"
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "result.pptx"
            merge.merge(source, donor, output, [1, 2])
            with ZipFile(source) as old, ZipFile(output) as new:
                unchanged = set(old.namelist()) - {
                    "[Content_Types].xml", "ppt/slides/slide1.xml", "ppt/slides/slide2.xml",
                    "ppt/slides/_rels/slide1.xml.rels", "ppt/slides/_rels/slide2.xml.rels",
                }
                for name in unchanged:
                    self.assertEqual(old.read(name), new.read(name), name)
                for index in (1, 2):
                    slide = new.read(f"ppt/slides/slide{index}.xml")
                    self.assertIn(b"AI Teaching Avatar", slide)
                    self.assertIn(b"AI Teaching Audio", slide)
                self.assertTrue(any(x.endswith(".gif") for x in new.namelist()))
                self.assertEqual(sum(x.endswith(".wav") for x in new.namelist()), 2)

    def test_replacement_has_one_avatar_and_one_audio(self):
        source = ROOT / "tests/fixtures/second-source.pptx"
        donor = ROOT / "tests/fixtures/round2-manual.pptx"
        with tempfile.TemporaryDirectory() as temp:
            first = Path(temp) / "first.pptx"
            second = Path(temp) / "second.pptx"
            merge.merge(source, donor, first, [1, 2])
            merge.merge(first, donor, second, [1, 2])
            with ZipFile(second) as archive:
                for index in (1, 2):
                    doc = ET.fromstring(archive.read(f"ppt/slides/slide{index}.xml"))
                    props = list(doc.iter(f"{{{merge.P}}}cNvPr"))
                    self.assertEqual(sum(x.get("name") == "AI Teaching Avatar" for x in props), 1)
                    self.assertEqual(sum(x.get("name") == "AI Teaching Audio" for x in props), 1)

    def test_original_animation_is_never_overwritten(self):
        source = ROOT / "input/说课.pptx"
        donor = ROOT / "tests/fixtures/round2-auto.pptx"
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "result.pptx"
            with self.assertRaisesRegex(ValueError, "existing animations"):
                merge.merge(source, donor, output, [1])
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
