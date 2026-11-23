"""Transplant COM-authored avatar/audio objects onto original PPTX slides."""

import argparse
import hashlib
import json
import posixpath
import sys
from pathlib import Path
from xml.dom import minidom
from zipfile import ZIP_DEFLATED, ZipFile

try:
    from timing_builder import build_timing_element, wav_duration_ms
except ImportError:
    from src.timing_builder import build_timing_element, wav_duration_ms

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NAMES = {"AI Teaching Avatar", "AI Teaching Audio"}
IDLE_NAME = "AI Teaching Idle"
MANAGED_NAMES = NAMES | {IDLE_NAME}


def slide_paths(archive):
    presentation = minidom.parseString(archive.read("ppt/presentation.xml"))
    relationships = minidom.parseString(archive.read("ppt/_rels/presentation.xml.rels"))
    targets = {node.getAttribute("Id"): node.getAttribute("Target")
               for node in relationships.getElementsByTagNameNS(PR, "Relationship")}
    return [posixpath.normpath(posixpath.join("ppt", targets[node.getAttributeNS(R, "id")]))
            for node in presentation.getElementsByTagNameNS(P, "sldId")]


def relationship_path(slide):
    folder, name = posixpath.split(slide)
    return f"{folder}/_rels/{name}.rels"


def named_shapes(doc):
    result = {}
    tree = doc.getElementsByTagNameNS(P, "spTree")[0]
    for node in tree.childNodes:
        if node.nodeType != node.ELEMENT_NODE:
            continue
        properties = node.getElementsByTagNameNS(P, "cNvPr")
        if properties and properties[0].getAttribute("name") in MANAGED_NAMES:
            name = properties[0].getAttribute("name")
            if name in result:
                raise ValueError(f"Duplicate managed object: {name}")
            result[name] = node
    return result


def merge(source_path, donor_path, output_path, pages, *, mode=None, preserve_animations=False):
    source_path, donor_path, output_path = map(Path, (source_path, donor_path, output_path))
    if output_path.resolve() in (source_path.resolve(), donor_path.resolve()) or output_path.exists():
        raise ValueError("Output must be a new path distinct from input files")
    with ZipFile(source_path) as source, ZipFile(donor_path) as donor:
        original_paths, donor_paths = slide_paths(source), slide_paths(donor)
        if len(original_paths) != len(donor_paths):
            raise ValueError("Slide count mismatch")
        if not pages or len(set(pages)) != len(pages) or any(not 1 <= p <= len(original_paths) for p in pages):
            raise ValueError("Invalid or duplicated page number")
        updates = {}
        media = {}
        types = minidom.parseString(source.read("[Content_Types].xml"))
        existing_ext = {node.getAttribute("Extension").lower()
                        for node in types.getElementsByTagNameNS(CT, "Default")}
        for ext, mime in (("gif", "image/gif"), ("wav", "audio/wav"), ("png", "image/png")):
            if ext not in existing_ext:
                node = types.createElementNS(CT, "Default")
                node.setAttribute("Extension", ext)
                node.setAttribute("ContentType", mime)
                types.documentElement.appendChild(node)
        updates["[Content_Types].xml"] = types.toxml(encoding="utf-8")
        occupied = set(source.namelist())
        for page in pages:
            original_slide, donor_slide = original_paths[page - 1], donor_paths[page - 1]
            original_doc = minidom.parseString(source.read(original_slide))
            donor_doc = minidom.parseString(donor.read(donor_slide))
            old = named_shapes(original_doc)
            old_managed_ids = {node.getElementsByTagNameNS(P, "cNvPr")[0].getAttribute("id") for node in old.values()}
            original_timing = original_doc.getElementsByTagNameNS(P, "timing")
            preserved_timing = None
            if original_timing:
                refs = {node.getAttribute("spid") for node in original_timing[0].getElementsByTagNameNS(P, "spTgt")}
                is_managed_only = bool(old) and bool(refs) and refs.issubset(old_managed_ids)
                if not is_managed_only:
                    if not preserve_animations:
                        raise ValueError(f"Slide {page} contains existing animations; preserving them is not yet supported")
                    preserved_timing = original_timing[0]
                original_timing[0].parentNode.removeChild(original_timing[0])
            for node in old.values():
                node.parentNode.removeChild(node)
            new_shapes = named_shapes(donor_doc)
            if not NAMES.issubset(new_shapes):
                raise ValueError(f"Slide {page} does not contain both donor objects")
            for i in range(donor_doc.documentElement.attributes.length):
                attr = donor_doc.documentElement.attributes.item(i)
                if attr.name.startswith("xmlns:") and not original_doc.documentElement.hasAttribute(attr.name):
                    original_doc.documentElement.setAttribute(attr.name, attr.value)
            rel_path = relationship_path(original_slide)
            donor_rel_path = relationship_path(donor_slide)
            if rel_path in source.namelist():
                original_rels = minidom.parseString(source.read(rel_path))
            else:
                original_rels = minidom.parseString(f'<Relationships xmlns="{PR}"/>')
            donor_rels = minidom.parseString(donor.read(donor_rel_path))
            rel_index = {r.getAttribute("Id"): r for r in donor_rels.getElementsByTagNameNS(PR, "Relationship")}
            used_rel_ids = {r.getAttribute("Id") for r in original_rels.getElementsByTagNameNS(PR, "Relationship")}
            used_shape_ids = {int(n.getAttribute("id")) for n in original_doc.getElementsByTagNameNS(P, "cNvPr")}
            shape_map, rel_map = {}, {}
            new_id_by_name = {}
            audio_dur_ms = 8000
            tree = original_doc.getElementsByTagNameNS(P, "spTree")[0]
            shape_order = ["AI Teaching Avatar", "AI Teaching Audio"]
            if IDLE_NAME in new_shapes:
                shape_order.append(IDLE_NAME)
            for name in shape_order:
                shape = original_doc.importNode(new_shapes[name], True)
                props = shape.getElementsByTagNameNS(P, "cNvPr")[0]
                old_id = props.getAttribute("id")
                new_id = max(used_shape_ids, default=0) + 1
                used_shape_ids.add(new_id)
                shape_map[old_id] = str(new_id)
                new_id_by_name[name] = str(new_id)
                props.setAttribute("id", str(new_id))
                for element in [shape] + list(shape.getElementsByTagName("*")):
                    for i in range(element.attributes.length):
                        attr = element.attributes.item(i)
                        if attr.namespaceURI != R or not attr.value:
                            continue
                        old_rel = attr.value
                        if old_rel not in rel_map:
                            donor_rel = rel_index[old_rel]
                            if donor_rel.getAttribute("TargetMode") == "External":
                                raise ValueError("External media relationship is not supported")
                            target = posixpath.normpath(posixpath.join(posixpath.dirname(donor_slide), donor_rel.getAttribute("Target")))
                            if not target.startswith("ppt/media/"):
                                raise ValueError(f"Unexpected donor relationship: {target}")
                            payload = donor.read(target)
                            ext = posixpath.splitext(target)[1].lower()
                            if ext not in (".gif", ".wav", ".png"):
                                raise ValueError(f"Unsupported media type: {ext}")
                            if ext == ".wav" and name == "AI Teaching Audio":
                                audio_dur_ms = wav_duration_ms(payload)
                            media_name = "ppt/media/ai_teaching_" + hashlib.sha256(payload).hexdigest() + ext
                            if media_name in occupied and source.read(media_name) != payload:
                                raise ValueError("Media collision")
                            media[media_name] = payload
                            rel_id = 1
                            while f"rId{rel_id}" in used_rel_ids:
                                rel_id += 1
                            rel_map[old_rel] = f"rId{rel_id}"
                            used_rel_ids.add(rel_map[old_rel])
                            rel = original_rels.createElementNS(PR, "Relationship")
                            rel.setAttribute("Id", rel_map[old_rel])
                            rel.setAttribute("Type", donor_rel.getAttribute("Type"))
                            rel.setAttribute("Target", posixpath.relpath(media_name, posixpath.dirname(original_slide)))
                            original_rels.documentElement.appendChild(rel)
                        attr.value = rel_map[old_rel]
                tree.appendChild(shape)

            if mode is not None:
                timing = build_timing_element(
                    original_doc,
                    mode=mode,
                    avatar_spid=new_id_by_name["AI Teaching Avatar"],
                    audio_spid=new_id_by_name["AI Teaching Audio"],
                    idle_spid=new_id_by_name.get(IDLE_NAME),
                    dur_ms=audio_dur_ms,
                    existing_timing=preserved_timing,
                    old_managed_ids=old_managed_ids,
                )
            else:
                timing_nodes = donor_doc.getElementsByTagNameNS(P, "timing")
                if len(timing_nodes) != 1:
                    raise ValueError(f"Slide {page} has no donor timing")
                timing = original_doc.importNode(timing_nodes[0], True)
                for target in timing.getElementsByTagNameNS(P, "spTgt"):
                    old_id = target.getAttribute("spid")
                    if old_id not in shape_map:
                        raise ValueError(f"Slide {page} donor timing includes unrelated shapes")
                    target.setAttribute("spid", shape_map[old_id])

            extension = next((node for node in original_doc.documentElement.childNodes
                              if node.nodeType == node.ELEMENT_NODE and node.namespaceURI == P
                              and node.localName == "extLst"), None)
            original_doc.documentElement.insertBefore(timing, extension)
            updates[original_slide] = original_doc.toxml(encoding="utf-8")
            updates[rel_path] = original_rels.toxml(encoding="utf-8")
        updates.update(media)
        with ZipFile(output_path, "x", compression=ZIP_DEFLATED) as output:
            for info in source.infolist():
                output.writestr(info, updates.pop(info.filename, source.read(info.filename)))
            for name, data in updates.items():
                output.writestr(name, data)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Merge AI Teaching media into original PPTX.")
    parser.add_argument("source")
    parser.add_argument("donor")
    parser.add_argument("output")
    parser.add_argument("page_file")
    parser.add_argument("--mode", choices=("auto", "manual"), default=None)
    parser.add_argument("--preserve-animations", action="store_true")
    args = parser.parse_args()
    pages = json.loads(Path(args.page_file).read_text(encoding="utf-8-sig"))
    merge(
        args.source,
        args.donor,
        args.output,
        [int(page["slide"]) for page in pages],
        mode=args.mode,
        preserve_animations=args.preserve_animations,
    )
