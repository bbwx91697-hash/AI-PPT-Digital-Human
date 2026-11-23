"""Check the source, output media package, ordering, and original renders."""

import hashlib
import json
import posixpath
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "input/说课.pptx"
SHA256 = "deee90bf0ddf088b96c92d010e815423f54bfdc49d682be6a242dd10dcfdf481"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def slide_paths(archive):
    presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
    rels = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    targets = {rel.get("Id"): rel.get("Target") for rel in rels}
    return [
        posixpath.normpath(posixpath.join("ppt", targets[item.get(R + "id")]))
        for item in presentation.find(P + "sldIdLst")
    ]


def slide_signature(archive, path):
    page = ET.fromstring(archive.read(path))
    return [
        (shape.tag, next((p.get("name") for p in shape.iter(P + "cNvPr")), ""),
         tuple("".join(node.text or "" for node in part.iter() if node.tag.endswith("}t"))
               for part in shape.iter() if part.tag.endswith("}p")))
        for shape in page.find(".//" + P + "spTree")
    ]


def check(mode, external_source):
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest() == SHA256
    assert hashlib.sha256(external_source.read_bytes()).hexdigest() == SHA256
    suffix = "" if mode == "auto" else "_手动"
    target = ROOT / f"output/说课_数字人提问样片{suffix}.pptx"
    original_renders = ROOT / "generated/previews/original"
    final_renders = ROOT / "generated/previews" / mode
    with ZipFile(SOURCE) as source, ZipFile(target) as result:
        assert result.testzip() is None
        registration = {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"}
        for part in set(source.namelist()) - registration:
            assert result.read(part) == source.read(part), part
        before, after = slide_paths(source), slide_paths(result)
        assert len(before) == 16 and len(after) == 17
        for index, old_path in enumerate(before, 1):
            new_index = index if index < 7 else index + 1
            assert slide_signature(source, old_path) == slide_signature(result, after[new_index - 1]), index
            with Image.open(original_renders / f"slide-{index:02}.png") as old:
                with Image.open(final_renders / f"slide-{new_index:02}.png") as new:
                    assert ImageChops.difference(old.convert("RGB"), new.convert("RGB")).getbbox() is None, index
        folder, name = posixpath.split(after[6])
        rels = ET.fromstring(result.read(f"{folder}/_rels/{name}.rels"))
        media_rels = [r for r in rels if r.get("Type").rsplit("/", 1)[-1] in {"media", "video"}]
        assert media_rels
        media = (ROOT / "generated/videos/slide07-composite.mp4").read_bytes()
        for rel in media_rels:
            assert rel.get("TargetMode") != "External"
            assert result.read(posixpath.normpath(posixpath.join(folder, rel.get("Target")))) == media
    print(f"{mode}: 16 original slide object signatures and rendered pixels match; MP4 embedded on slide 7")


if __name__ == "__main__":
    for name in ("auto", "manual"):
        check(name, Path(sys.argv[1]))
    questions = json.loads((ROOT / "reports/question_candidates.json").read_text(encoding="utf-8"))
    assert len(questions) == 3 and all(item["review_status"] == "待审核" for item in questions)
    print("Three candidates require teacher review.")
