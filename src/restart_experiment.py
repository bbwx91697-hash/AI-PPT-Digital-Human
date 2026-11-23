"""Generate a PowerPoint restart-timing experiment on an existing test deck."""

import runpy
import sys
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
HELPERS = runpy.run_path(str(Path(__file__).with_name("openxml_merge.py")))


def create_candidate(source, output, pages):
    source, output = Path(source), Path(output)
    if source.resolve() == output.resolve() or output.exists():
        raise ValueError("Output must be a new file")
    with ZipFile(source) as before:
        paths = HELPERS["slide_paths"](before)
        updates = {}
        for page in pages:
            path = paths[page - 1]
            doc = minidom.parseString(before.read(path))
            timing = doc.getElementsByTagNameNS(P, "timing")
            if len(timing) != 1:
                raise ValueError(f"Page {page} has no timing")
            for node in timing[0].getElementsByTagNameNS(P, "cTn"):
                if node.getAttribute("nodeType") == "mainSeq" or node.getAttribute("presetClass") == "mediacall":
                    node.setAttribute("restart", "always")
            updates[path] = doc.toxml(encoding="utf-8")
        with ZipFile(output, "x") as after:
            for info in before.infolist():
                after.writestr(info, updates.get(info.filename, before.read(info.filename)))


if __name__ == "__main__":
    create_candidate(sys.argv[1], sys.argv[2], [int(x) for x in sys.argv[3:]])
