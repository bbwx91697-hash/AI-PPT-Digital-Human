import importlib.util
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("restart_experiment", ROOT / "src/restart_experiment.py")
experiment = importlib.util.module_from_spec(spec)
spec.loader.exec_module(experiment)


class RestartTest(unittest.TestCase):
    def test_only_target_timing_is_changed(self):
        source = ROOT / "output/第三轮_说课_固定人物_自动验收.pptx"
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "probe.pptx"
            experiment.create_candidate(source, output, [7, 8])
            with ZipFile(source) as before, ZipFile(output) as after:
                for name in before.namelist():
                    if name not in {"ppt/slides/slide7.xml", "ppt/slides/slide8.xml"}:
                        self.assertEqual(before.read(name), after.read(name), name)
                for page in (7, 8):
                    doc = ET.fromstring(after.read(f"ppt/slides/slide{page}.xml"))
                    nodes = list(doc.iter(f"{{{experiment.P}}}cTn"))
                    self.assertTrue(any(x.get("nodeType") == "mainSeq" and x.get("restart") == "always" for x in nodes))


if __name__ == "__main__":
    unittest.main()
