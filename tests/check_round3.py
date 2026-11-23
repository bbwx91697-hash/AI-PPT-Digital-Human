"""Audit preserved slide parts, original objects, and embedded media."""

import hashlib
import runpy
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[1]
MODULE = runpy.run_path(str(ROOT / "src/openxml_merge.py"))
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
FIXTURES = ROOT / "tests/fixtures"
AVATAR = ROOT / "assets/avatar/default_avatar.gif"


def original_objects(root):
    tree = root.find(".//" + P + "spTree")
    return [
        ET.canonicalize(ET.tostring(item, encoding="unicode"))
        for item in tree
        if not any(p.get("name") in MODULE["NAMES"] for p in item.iter(P + "cNvPr"))
    ]


def check(source, target, pages, audio_files):
    with ZipFile(source) as old, ZipFile(target) as new:
        assert new.testzip() is None
        paths = MODULE["slide_paths"](old)
        assert len(paths) == len(MODULE["slide_paths"](new))
        for name in old.namelist():
            if name.startswith("ppt/media/"):
                assert old.read(name) == new.read(name), name
        for i, path in enumerate(paths, 1):
            rel_path = MODULE["relationship_path"](path)
            if i not in pages:
                assert old.read(path) == new.read(path), path
                if rel_path in old.namelist():
                    assert old.read(rel_path) == new.read(rel_path), rel_path
                continue
            before = ET.fromstring(old.read(path))
            after = ET.fromstring(new.read(path))
            assert original_objects(before) == original_objects(after), i
            names = {}
            for shape in after.find(".//" + P + "spTree"):
                props = list(shape.iter(P + "cNvPr"))
                if props and props[0].get("name") in MODULE["NAMES"]:
                    names[props[0].get("name")] = shape
            assert set(names) == MODULE["NAMES"]
            relationships = {r.get("Id"): r for r in ET.fromstring(new.read(rel_path))}
            for name, wanted in (("AI Teaching Avatar", AVATAR), ("AI Teaching Audio", audio_files[pages.index(i)])):
                payloads = []
                for item in names[name].iter():
                    for key, value in item.attrib.items():
                        if key.startswith(R) and value in relationships:
                            relation = relationships[value]
                            assert relation.get("TargetMode") != "External"
                            part = MODULE["posixpath"].normpath(MODULE["posixpath"].join(
                                MODULE["posixpath"].dirname(path), relation.get("Target")))
                            payloads.append(new.read(part))
                assert any(hashlib.sha256(data).digest() == hashlib.sha256(wanted.read_bytes()).digest()
                           for data in payloads), name
        print(f"{target.name}: {len(paths)} slides, original objects and media preserved; page media embedded")


if __name__ == "__main__":
    check(ROOT / "input/说课.pptx", ROOT / "output/第三轮_说课_固定人物_自动验收.pptx",
          [7, 8], [FIXTURES / "tone-7.wav", FIXTURES / "tone-8.wav"])
    check(FIXTURES / "second-source.pptx", ROOT / "output/第三轮_两页_固定人物_手动验收.pptx",
          [1, 2], [FIXTURES / "tone-7.wav", FIXTURES / "tone-8.wav"])
    differences = []
    for i in list(range(1, 7)) + list(range(9, 17)):
        old = Image.open(ROOT / f"generated/previews/original/slide-{i:02}.png").convert("RGB")
        new = Image.open(ROOT / f"generated/previews/round3-final-main/slide-{i:02}.png").convert("RGB")
        if ImageChops.difference(old, new).getbbox():
            differences.append(i)
    assert not differences, differences
    print("14 unchanged original slides: exact PowerPoint export pixel match")
