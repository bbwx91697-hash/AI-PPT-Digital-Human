"""Build clean PowerPoint Open XML <p:timing> trees with GIF/Idle mutual exclusion."""

from __future__ import annotations

import io
import wave
from xml.dom import minidom

P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def wav_duration_ms(wav_bytes: bytes, fallback_ms: int = 8000) -> int:
    """Return WAV duration in milliseconds (minimum 500 ms)."""
    try:
        with wave.open(io.BytesIO(wav_bytes), "rb") as reader:
            rate = reader.getframerate()
            frames = reader.getnframes()
            if rate > 0 and frames > 0:
                return max(500, int(round(frames * 1000.0 / rate)))
    except (wave.Error, EOFError, ValueError):
        pass
    return fallback_ms


class _IdGen:
    def __init__(self, start: int = 1):
        self.value = start

    def next(self) -> int:
        current = self.value
        self.value += 1
        return current


def _visibility_par(ids: _IdGen, spid: str, visible: bool, preset_class: str, node_type: str) -> str:
    val = "visible" if visible else "hidden"
    return (
        f'<p:par><p:cTn id="{ids.next()}" presetID="1" presetClass="{preset_class}" '
        f'presetSubtype="0" fill="hold" nodeType="{node_type}">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst><p:set><p:cBhvr>'
        f'<p:cTn id="{ids.next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
        f'<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        f'</p:cBhvr><p:to><p:strVal val="{val}"/></p:to></p:set></p:childTnLst>'
        f'</p:cTn></p:par>'
    )


def _play_effects_group(
    ids: _IdGen,
    audio_spid: str,
    avatar_spid: str,
    idle_spid: str | None,
    dur_ms: int,
    first_node_type: str,
    start_cond_xml: str,
) -> str:
    """Build one complete speaking -> idle transition group."""
    start_pars = [
        f'<p:par><p:cTn id="{ids.next()}" presetID="1" presetClass="mediacall" '
        f'presetSubtype="0" fill="hold" restart="always" nodeType="{first_node_type}">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>'
        f'<p:cTn id="{ids.next()}" dur="{dur_ms}" fill="hold"/>'
        f'<p:tgtEl><p:spTgt spid="{audio_spid}"/></p:tgtEl>'
        f'</p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>'
    ]
    if idle_spid:
        start_pars.append(_visibility_par(ids, avatar_spid, True, "entr", "withEffect"))
        start_pars.append(_visibility_par(ids, idle_spid, False, "exit", "withEffect"))

    end_block = ""
    if idle_spid:
        end_pars = [
            _visibility_par(ids, idle_spid, True, "entr", "afterEffect"),
            _visibility_par(ids, avatar_spid, False, "exit", "withEffect"),
        ]
        end_block = (
            f'<p:par><p:cTn id="{ids.next()}" fill="hold">'
            f'<p:stCondLst><p:cond delay="{dur_ms}"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(end_pars)}</p:childTnLst>'
            f'</p:cTn></p:par>'
        )

    return (
        f'<p:par><p:cTn id="{ids.next()}" fill="hold">'
        f'<p:stCondLst>{start_cond_xml}</p:stCondLst>'
        f'<p:childTnLst>'
        f'<p:par><p:cTn id="{ids.next()}" fill="hold">'
        f'<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
        f'<p:childTnLst>{"".join(start_pars)}</p:childTnLst>'
        f'</p:cTn></p:par>'
        f'{end_block}'
        f'</p:childTnLst></p:cTn></p:par>'
    )


def _audio_media_node(ids: _IdGen, audio_spid: str) -> str:
    return (
        f'<p:audio><p:cMediaNode vol="80000">'
        f'<p:cTn id="{ids.next()}" fill="remove" display="0">'
        f'<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
        f'<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>'
        f'</p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{audio_spid}"/></p:tgtEl>'
        f'</p:cMediaNode></p:audio>'
    )


def _interactive_seq(
    ids: _IdGen,
    trigger_spid: str,
    audio_spid: str,
    avatar_spid: str,
    idle_spid: str | None,
    dur_ms: int,
) -> str:
    seq_id = ids.next()
    body = _play_effects_group(
        ids,
        audio_spid=audio_spid,
        avatar_spid=avatar_spid,
        idle_spid=idle_spid,
        dur_ms=dur_ms,
        first_node_type="clickEffect",
        start_cond_xml='<p:cond delay="0"/>',
    )
    return (
        f'<p:seq concurrent="1" nextAc="seek">'
        f'<p:cTn id="{seq_id}" restart="always" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">'
        f'<p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{trigger_spid}"/></p:tgtEl></p:cond></p:stCondLst>'
        f'<p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync>'
        f'<p:childTnLst>{body}</p:childTnLst>'
        f'</p:cTn>'
        f'<p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{trigger_spid}"/></p:tgtEl></p:cond></p:nextCondLst>'
        f'</p:seq>'
    )


def build_timing_element(
    doc: minidom.Document,
    mode: str,
    avatar_spid: str,
    audio_spid: str,
    idle_spid: str | None,
    dur_ms: int,
    existing_timing: minidom.Element | None = None,
    old_managed_ids: set[str] | None = None,
) -> minidom.Element:
    """Build or merge a <p:timing> element for the target slide."""
    if mode not in ("auto", "manual"):
        raise ValueError(f"Unsupported mode: {mode}")

    if existing_timing is not None:
        return _merge_existing_timing(
            doc, existing_timing, old_managed_ids or set(),
            mode, avatar_spid, audio_spid, idle_spid, dur_ms,
        )

    ids = _IdGen(1)
    root_id = ids.next()
    children: list[str] = []

    if mode == "auto":
        start_cond = f'<p:cond delay="0"/><p:cond evt="onBegin" delay="0"><p:tn val="{root_id}"/></p:cond>'
        auto_group = _play_effects_group(
            ids,
            audio_spid=audio_spid,
            avatar_spid=avatar_spid,
            idle_spid=idle_spid,
            dur_ms=dur_ms,
            first_node_type="withEffect",
            start_cond_xml=start_cond,
        )
        children.append(auto_group)

    children.append(_audio_media_node(ids, audio_spid))
    children.append(_interactive_seq(ids, avatar_spid, audio_spid, avatar_spid, idle_spid, dur_ms))
    if idle_spid:
        children.append(_interactive_seq(ids, idle_spid, audio_spid, avatar_spid, idle_spid, dur_ms))

    xml_str = (
        f'<p:timing xmlns:p="{P}"><p:tnLst><p:par>'
        f'<p:cTn id="{root_id}" dur="indefinite" restart="never" nodeType="tmRoot">'
        f'<p:childTnLst>{"".join(children)}</p:childTnLst>'
        f'</p:cTn></p:par></p:tnLst></p:timing>'
    )
    return doc.importNode(minidom.parseString(xml_str).documentElement, True)


def _node_targets_only(node: minidom.Element, managed_ids: set[str]) -> bool:
    if not managed_ids:
        return False
    spids = {t.getAttribute("spid") for t in node.getElementsByTagNameNS(P, "spTgt") if t.getAttribute("spid")}
    return bool(spids) and spids.issubset(managed_ids)


def _merge_existing_timing(
    doc: minidom.Document,
    existing_timing: minidom.Element,
    old_managed_ids: set[str],
    mode: str,
    avatar_spid: str,
    audio_spid: str,
    idle_spid: str | None,
    dur_ms: int,
) -> minidom.Element:
    timing = doc.importNode(existing_timing, True)
    tm_root = next(
        (n for n in timing.getElementsByTagNameNS(P, "cTn") if n.getAttribute("nodeType") == "tmRoot"),
        None,
    )
    if tm_root is None:
        raise ValueError("Existing slide timing lacks tmRoot")
    root_children = next(
        (c for c in tm_root.childNodes if c.nodeType == c.ELEMENT_NODE and c.localName == "childTnLst"),
        None,
    )
    if root_children is None:
        root_children = doc.createElementNS(P, "p:childTnLst")
        tm_root.appendChild(root_children)

    # Strip previous AI Teaching nodes before re-inserting updated ones.
    for child in list(root_children.childNodes):
        if child.nodeType != child.ELEMENT_NODE:
            continue
        if child.localName in ("audio", "seq", "par") and _node_targets_only(child, old_managed_ids):
            if child.localName == "seq":
                ctn = next((n for n in child.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName == "cTn"), None)
                if ctn and ctn.getAttribute("nodeType") == "mainSeq":
                    continue
            root_children.removeChild(child)

    main_seq_ctn = next(
        (n for n in timing.getElementsByTagNameNS(P, "cTn") if n.getAttribute("nodeType") == "mainSeq"),
        None,
    )
    main_seq_list = None
    if main_seq_ctn is not None:
        main_seq_list = next(
            (c for c in main_seq_ctn.childNodes if c.nodeType == c.ELEMENT_NODE and c.localName == "childTnLst"),
            None,
        )
        if main_seq_list is not None:
            for item in list(main_seq_list.childNodes):
                if item.nodeType == item.ELEMENT_NODE and _node_targets_only(item, old_managed_ids):
                    main_seq_list.removeChild(item)

    max_id = max(
        (int(n.getAttribute("id")) for n in timing.getElementsByTagNameNS(P, "cTn") if n.getAttribute("id").isdigit()),
        default=1,
    )
    ids = _IdGen(max_id + 1)

    def _import_fragment(frag_xml: str) -> minidom.Element:
        wrapped = f'<p:wrap xmlns:p="{P}">{frag_xml}</p:wrap>'
        parsed = minidom.parseString(wrapped).documentElement
        return doc.importNode(parsed.firstChild, True)

    if mode == "auto":
        root_id = tm_root.getAttribute("id") or "1"
        start_cond = f'<p:cond delay="0"/><p:cond evt="onBegin" delay="0"><p:tn val="{root_id}"/></p:cond>'
        auto_group = _play_effects_group(
            ids, audio_spid, avatar_spid, idle_spid, dur_ms, "withEffect", start_cond,
        )
        root_children.appendChild(_import_fragment(auto_group))

    root_children.appendChild(_import_fragment(_audio_media_node(ids, audio_spid)))
    root_children.appendChild(_import_fragment(_interactive_seq(ids, avatar_spid, audio_spid, avatar_spid, idle_spid, dur_ms)))
    if idle_spid:
        root_children.appendChild(_import_fragment(_interactive_seq(ids, idle_spid, audio_spid, avatar_spid, idle_spid, dur_ms)))

    return timing
