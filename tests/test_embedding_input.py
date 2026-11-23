import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location(
    "embedding_input", Path(__file__).resolve().parents[1] / "src" / "embedding_input.py"
)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class InputTests(unittest.TestCase):
    def test_rejects_duplicate_pages(self):
        with self.assertRaisesRegex(ValueError, "重复"):
            module.validate_pages([{"slide": 7}, {"slide": 7}], 16)

    def test_rejects_page_outside_deck(self):
        with self.assertRaisesRegex(ValueError, "页码"):
            module.validate_pages([{"slide": 17}], 16)

    def test_accepts_distinct_pages(self):
        module.validate_pages([{"slide": 7}, {"slide": 8}], 16)


if __name__ == "__main__":
    unittest.main()
