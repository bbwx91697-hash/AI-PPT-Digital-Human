import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


class TimingTests(unittest.TestCase):
    def test_manual_experiment_keeps_two_separate_click_sequences(self):
        path = ROOT / "output/第四轮_手动状态实验_v2.pptx"
        with ZipFile(path) as deck:
            for page in (1, 2):
                root = ET.fromstring(deck.read(f"ppt/slides/slide{page}.xml"))
                sequences = [node for node in root.iter(P + "cTn")
                             if node.get("nodeType") == "interactiveSeq"]
                self.assertEqual(len(sequences), 2)
                effects = [[node.get("presetClass") for node in sequence.iter(P + "cTn")
                            if node.get("presetClass")] for sequence in sequences]
                self.assertEqual(effects, [["mediacall", "entr"], ["mediacall", "exit", "entr"]])


if __name__ == "__main__":
    unittest.main()
