import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

import app
import layout_geometry
import openxml_merge
import ppt_inspector
import project_store

P = f"{{{openxml_merge.P}}}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


class LayoutAndDragTests(unittest.TestCase):
    def test_clamp_and_corner_resize_geometry(self):
        l, t, w, h = layout_geometry.clamp_position(0.42, 0.31, 0.16, 0.35)
        self.assertEqual((l, t, w, h), (0.42, 0.31, 0.16, 0.35))

        # Clamps out-of-bounds drag inside slide
        cl, ct, cw, ch = layout_geometry.clamp_position(0.95, -0.1, 0.16, 0.35)
        self.assertAlmostEqual(cl, 0.84, places=3)
        self.assertEqual(ct, 0.0)

        # Resize SE handle with locked aspect ratio
        nl, nt, nw, nh = layout_geometry.resize_from_handle(
            "se", (0.40, 0.30, 0.16, 0.35), 0.08, 0.05, lock_aspect=True, aspect_ratio_hw=2.0
        )
        self.assertEqual((nl, nt), (0.40, 0.30))
        self.assertAlmostEqual(nw, 0.24, places=3)
        self.assertAlmostEqual(nh, 0.48, places=3)

        # Resize NW handle shrinks from top-left while keeping bottom-right anchored
        nl2, nt2, nw2, nh2 = layout_geometry.resize_from_handle(
            "nw", (0.40, 0.30, 0.20, 0.40), 0.05, 0.10, lock_aspect=False
        )
        self.assertAlmostEqual(nl2 + nw2, 0.60, places=3)
        self.assertAlmostEqual(nt2 + nh2, 0.70, places=3)

    def test_custom_center_position_exports_exact_emu_coordinates(self):
        wav = ROOT / "generated/audio/gemini_voice_test/Leda.wav"
        self.assertTrue(wav.is_file())
        with tempfile.TemporaryDirectory() as temp:
            tmp = Path(temp)
            pdir, data, info = project_store.create_or_load_project(
                ROOT / "tests/fixtures/second-source.pptx", projects_root=tmp / "projects"
            )
            custom_box = (0.37, 0.26, 0.22, 0.44)
            cfg = data["pages"]["1"]
            cfg["enabled"] = True
            cfg["audio_path"] = str(wav)
            cfg["left_pct"], cfg["top_pct"], cfg["width_pct"], cfg["height_pct"] = custom_box
            out = tmp / "custom_pos.pptx"
            project_store.export_presentation(pdir, data, info, out)

            with ZipFile(out) as z:
                pres = ET.fromstring(z.read("ppt/presentation.xml"))
                sld_sz = pres.find(f".//{P}sldSz")
                cx, cy = int(sld_sz.get("cx")), int(sld_sz.get("cy"))
                slide1 = ET.fromstring(z.read("ppt/slides/slide1.xml"))
                for target_name in ("AI Teaching Avatar", "AI Teaching Idle"):
                    pic = next(
                        el for el in slide1.find(f".//{P}spTree")
                        if el.find(f".//{P}cNvPr") is not None and el.find(f".//{P}cNvPr").get("name") == target_name
                    )
                    xfrm = pic.find(f".//{A}xfrm")
                    off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
                    self.assertAlmostEqual(int(off.get("x")) / cx, custom_box[0], places=2)
                    self.assertAlmostEqual(int(off.get("y")) / cy, custom_box[1], places=2)
                    self.assertAlmostEqual(int(ext.get("cx")) / cx, custom_box[2], places=2)
                    self.assertAlmostEqual(int(ext.get("cy")) / cy, custom_box[3], places=2)

    def test_gui_drag_resize_and_apply_all(self):
        pdir, data, info = project_store.create_or_load_project(ROOT / "input/说课.pptx")
        thumbs = ppt_inspector.export_thumbnails(Path(data["source_copy"]), pdir / "thumbs")
        gui = app.DigitalHumanApp()
        try:
            gui.proj_dir, gui.proj_data, gui.deck_info, gui.thumb_paths = pdir, data, info, thumbs
            gui._populate_tree()
            gui.tree.selection_set("7")
            gui.on_select_slide()
            gui.update()

            # Simulate dragging avatar center leftward/upward onto mid-slide
            x0, y0, x1, y1 = gui._box_canvas_coords()
            cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
            _ox, _oy, rw, rh = gui._slide_rect
            gui._on_canvas_press(SimpleNamespace(x=cx, y=cy))
            self.assertEqual(gui._drag_state["mode"], "move")
            gui._on_canvas_drag(SimpleNamespace(x=cx - rw * 0.35, y=cy - rh * 0.25))
            gui._on_canvas_release(SimpleNamespace())

            moved_l = gui.proj_data["pages"]["7"]["left_pct"]
            moved_t = gui.proj_data["pages"]["7"]["top_pct"]
            self.assertLess(moved_l, 0.60)
            self.assertLess(moved_t, 0.45)

            # Simulate resizing via bottom-right corner handle
            rx0, ry0, rx1, ry1 = gui._box_canvas_coords()
            gui._on_canvas_press(SimpleNamespace(x=rx1, y=ry1))
            self.assertEqual(gui._drag_state["mode"], "se")
            gui._on_canvas_drag(SimpleNamespace(x=rx1 + rw * 0.05, y=ry1 + rh * 0.08))
            gui._on_canvas_release(SimpleNamespace())
            self.assertGreater(gui.proj_data["pages"]["7"]["width_pct"], 0.17)
        finally:
            gui.destroy()


if __name__ == "__main__":
    unittest.main()
