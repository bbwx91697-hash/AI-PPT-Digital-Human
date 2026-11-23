"""Inspect PPTX metadata, text previews, existing animations, and export slide thumbnails."""

from __future__ import annotations

import hashlib
import subprocess
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

try:
    import openxml_merge
except ImportError:
    from src import openxml_merge

P = f"{{{openxml_merge.P}}}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
EMU_PER_POINT = 12700.0


@dataclass
class SlideSummary:
    slide: int
    title: str
    preview_text: str
    has_existing_animations: bool
    avatar_box: tuple[float, float, float, float] | None


@dataclass
class DeckInfo:
    slide_count: int
    width_pt: float
    height_pt: float
    slides: list[SlideSummary]
    sha256: str


def inspect_pptx(pptx_path: Path) -> DeckInfo:
    path = Path(pptx_path)
    raw = path.read_bytes()
    digest = hashlib.sha256(raw).hexdigest()
    with ZipFile(path) as archive:
        pres = ET.fromstring(archive.read("ppt/presentation.xml"))
        sld_sz = pres.find(f".//{P}sldSz")
        cx = int(sld_sz.get("cx", "12192000")) if sld_sz is not None else 12192000
        cy = int(sld_sz.get("cy", "6858000")) if sld_sz is not None else 6858000
        width_pt = cx / EMU_PER_POINT
        height_pt = cy / EMU_PER_POINT

        slides: list[SlideSummary] = []
        for idx, rel_slide_path in enumerate(openxml_merge.slide_paths(archive), start=1):
            root = ET.fromstring(archive.read(rel_slide_path))
            texts: list[str] = []
            managed_ids: set[str] = set()
            avatar_box = None
            sp_tree = root.find(f".//{P}spTree")
            if sp_tree is not None:
                for child in sp_tree:
                    cnv = child.find(f".//{P}cNvPr")
                    name = cnv.get("name", "") if cnv is not None else ""
                    if name in openxml_merge.MANAGED_NAMES:
                        if cnv is not None and cnv.get("id"):
                            managed_ids.add(cnv.get("id"))
                        if name == "AI Teaching Avatar":
                            xfrm = child.find(f".//{A}xfrm")
                            if xfrm is not None:
                                off = xfrm.find(f"{A}off")
                                ext = xfrm.find(f"{A}ext")
                                if off is not None and ext is not None and cx > 0 and cy > 0:
                                    avatar_box = (
                                        round(int(off.get("x", "0")) / cx, 4),
                                        round(int(off.get("y", "0")) / cy, 4),
                                        round(int(ext.get("cx", "0")) / cx, 4),
                                        round(int(ext.get("cy", "0")) / cy, 4),
                                    )
                        continue
                    for t_node in child.iter(f"{A}t"):
                        if t_node.text and t_node.text.strip():
                            texts.append(t_node.text.strip())

            title = texts[0] if texts else f"第 {idx} 页"
            preview = " ".join(texts)[:140]
            timing = root.find(f".//{P}timing")
            has_anim = False
            if timing is not None:
                refs = {n.get("spid") for n in timing.iter(f"{P}spTgt") if n.get("spid")}
                has_anim = not (managed_ids and refs and refs.issubset(managed_ids))

            slides.append(
                SlideSummary(
                    slide=idx,
                    title=title[:60],
                    preview_text=preview,
                    has_existing_animations=has_anim,
                    avatar_box=avatar_box,
                )
            )
        return DeckInfo(
            slide_count=len(slides),
            width_pt=width_pt,
            height_pt=height_pt,
            slides=slides,
            sha256=digest,
        )


def export_thumbnails(pptx_path: Path, thumb_dir: Path, *, width: int = 960, height: int = 540) -> list[Path]:
    """Export all slide thumbnails to PNG via PowerPoint COM, cached by source SHA-256."""
    src = Path(pptx_path).resolve()
    out_dir = Path(thumb_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256(src.read_bytes()).hexdigest()
    marker = out_dir / ".source_sha256"

    info = inspect_pptx(src)
    expected = [out_dir / f"slide_{i:03d}.png" for i in range(1, info.slide_count + 1)]
    if marker.is_file() and marker.read_text(encoding="utf-8").strip() == digest:
        if all(p.is_file() for p in expected):
            return expected

    ps = (
        "$ErrorActionPreference = 'Stop'; "
        "$app = New-Object -ComObject PowerPoint.Application; "
        f"$deck = $app.Presentations.Open('{src}', -1, 0, 0); "
        "try { "
        "for ($i = 1; $i -le $deck.Slides.Count; $i++) { "
        "$name = ('slide_{0:D3}.png' -f $i); "
        f"$dest = Join-Path '{out_dir}' $name; "
        f"$deck.Slides.Item($i).Export($dest, 'PNG', {int(width)}, {int(height)}); "
        "} "
        "} finally { $deck.Close() }"
    )
    res = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=False,
    )
    if res.returncode != 0:
        raise RuntimeError(f"导出幻灯片缩略图失败: {(res.stderr or res.stdout).strip()}")
    marker.write_text(digest, encoding="utf-8")
    return expected
