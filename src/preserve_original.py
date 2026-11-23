"""Import a PowerPoint-authored sample slide while preserving original PPTX parts."""

import posixpath
import sys
from pathlib import Path
from xml.dom import minidom
from zipfile import ZIP_DEFLATED, ZipFile

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PRES = "ppt/presentation.xml"
PRES_RELS = "ppt/_rels/presentation.xml.rels"
TYPES = "[Content_Types].xml"


def parse(archive, part):
    return minidom.parseString(archive.read(part))


def append(parent, name, attributes):
    doc = parent.ownerDocument
    element = doc.createElementNS(parent.namespaceURI, name)
    for key, value in attributes.items():
        element.setAttribute(key, str(value))
    return parent.appendChild(element)


def slides(doc):
    nodes = doc.getElementsByTagNameNS(P, "sldIdLst")
    return list(nodes[0].getElementsByTagNameNS(P, "sldId"))


def rels(doc):
    return list(doc.getElementsByTagNameNS(PR, "Relationship"))


def add_sample(source_path, authored_path, output_path):
    with ZipFile(source_path) as source, ZipFile(authored_path) as authored:
        presentation = parse(source, PRES)
        source_slides = slides(presentation)
        donor_slides = slides(parse(authored, PRES))
        if (len(source_slides), len(donor_slides)) != (16, 17):
            raise ValueError("Expected 16 original slides and 17 authored slides")
        donor_id = donor_slides[6].getAttributeNS(R, "id")
        donor_rels = rels(parse(authored, PRES_RELS))
        donor = next(r for r in donor_rels if r.getAttribute("Id") == donor_id)
        donor_path = posixpath.join("ppt", donor.getAttribute("Target"))
        folder, name = posixpath.split(donor_path)
        new_rels = parse(authored, f"{folder}/_rels/{name}.rels")
        new_parts = {"ppt/slides/slide17.xml": authored.read(donor_path)}

        for relationship in rels(new_rels):
            relation_type = relationship.getAttribute("Type").rsplit("/", 1)[-1]
            target = posixpath.normpath(posixpath.join(folder, relationship.getAttribute("Target")))
            if relation_type == "slideLayout" and target in source.namelist():
                continue
            if relationship.getAttribute("TargetMode") == "External":
                raise ValueError("Sample slide must use embedded media")
            if relation_type in {"media", "video"} and target.endswith(".mp4"):
                filename = "avatar-sample.mp4"
            elif relation_type == "image" and target.endswith(".png"):
                filename = "avatar-sample-poster.png"
            else:
                raise ValueError(f"Unsupported sample relationship: {relation_type}, {target}")
            new_parts[f"ppt/media/{filename}"] = authored.read(target)
            relationship.setAttribute("Target", "../media/" + filename)

        presentation_rels = parse(source, PRES_RELS)
        relationship_id = "rIdAvatarSample"
        append(presentation_rels.documentElement, "Relationship", {
            "Id": relationship_id, "Type": R + "/slide", "Target": "slides/slide17.xml"
        })
        node = presentation.createElementNS(P, "p:sldId")
        node.setAttribute("id", str(max(int(s.getAttribute("id")) for s in source_slides) + 1))
        node.setAttributeNS(R, "r:id", relationship_id)
        source_slides[6].parentNode.insertBefore(node, source_slides[6])

        types = parse(source, TYPES)
        for part, media_type in {
            "ppt/slides/slide17.xml": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            "ppt/media/avatar-sample.mp4": "video/mp4",
            "ppt/media/avatar-sample-poster.png": "image/png",
        }.items():
            append(types.documentElement, "Override", {"PartName": "/" + part, "ContentType": media_type})
        new_parts["ppt/slides/_rels/slide17.xml.rels"] = new_rels.toxml(encoding="utf-8")
        overlap = set(new_parts).intersection(source.namelist())
        if overlap:
            raise ValueError(f"Sample parts already in original: {overlap}")
        updates = {
            PRES: presentation.toxml(encoding="utf-8"),
            PRES_RELS: presentation_rels.toxml(encoding="utf-8"),
            TYPES: types.toxml(encoding="utf-8"),
        }
        pending = output_path.with_suffix(".pending")
        with ZipFile(pending, "w", compression=ZIP_DEFLATED) as result:
            for part in source.namelist():
                result.writestr(part, updates[part] if part in updates else source.read(part))
            for part, payload in new_parts.items():
                result.writestr(part, payload)
        pending.replace(output_path)


if __name__ == "__main__":
    add_sample(*map(Path, sys.argv[1:]))
