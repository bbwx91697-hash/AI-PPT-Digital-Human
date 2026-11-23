import unittest
from pathlib import Path

from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[1]


class IdleImageTests(unittest.TestCase):
    def test_idle_matches_original_gif_first_frame(self):
        with Image.open(ROOT / "assets/avatar/default_avatar.gif") as gif:
            gif.seek(0)
            first = gif.convert("RGBA")
        with Image.open(ROOT / "assets/avatar/default_avatar_idle.png") as idle:
            self.assertEqual(first.size, idle.size)
            self.assertIsNone(ImageChops.difference(first, idle.convert("RGBA")).getbbox())
            self.assertEqual(idle.convert("RGBA").getpixel((10, 100))[3], 0)


if __name__ == "__main__":
    unittest.main()
