"""Project configuration persistence, global avatar management, and PPTX export."""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from pathlib import Path
from PIL import Image

try:
    import embedding_input
    import ppt_inspector
    import tts_service
except ImportError:
    from src import embedding_input, ppt_inspector, tts_service

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_GIF = ROOT / "assets" / "avatar" / "default_avatar.gif"
DEFAULT_IDLE = ROOT / "assets" / "avatar" / "default_avatar_idle.png"
PROJECTS_ROOT = ROOT / "projects"
EMBED_SCRIPT = ROOT / "src" / "embed_media.ps1"

DEFAULT_BOX = (0.81, 0.60, 0.16, 0.35)


def extract_idle_frame(gif_path: Path, output_png: Path) -> Path:
    """Extract frame 0 of an animated GIF as an RGBA PNG for static idle state."""
    out = Path(output_png)
    out.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(gif_path) as img:
        img.seek(0)
        frame = img.convert("RGBA")
        frame.save(out, format="PNG")
    return out


def _safe_slug(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\s]+', "_", name.strip())
    return cleaned.strip("_") or "project"


def default_page_config(slide: int, avatar_box: tuple[float, float, float, float] | None = None) -> dict:
    left, top, width, height = avatar_box or DEFAULT_BOX
    engine = tts_service.default_engine()
    default_voice = (
        tts_service.QWEN3_VOICE_CODES[0]
        if engine == "qwen3"
        else tts_service.EDGE_VOICE_CODES[0]
    )
    return {
        "slide": int(slide),
        "enabled": False,
        "script": "",
        "engine": engine,
        "voice": default_voice,
        "style": "温和、亲切、清晰",
        "pace": "适中",
        "model": "gemini-3.8-flash-tts",
        "left_pct": left,
        "top_pct": top,
        "width_pct": width,
        "height_pct": height,
        "audio_path": "",
    }


def create_or_load_project(source_pptx: Path, projects_root: Path = PROJECTS_ROOT) -> tuple[Path, dict, ppt_inspector.DeckInfo]:
    """Import a copy of source_pptx into a project folder (never mutates original)."""
    src = Path(source_pptx).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"找不到 PPTX 文件: {src}")
    info = ppt_inspector.inspect_pptx(src)
    slug = f"{_safe_slug(src.stem)}_{info.sha256[:8]}"
    proj_dir = Path(projects_root) / slug
    proj_dir.mkdir(parents=True, exist_ok=True)

    source_copy = proj_dir / "source.pptx"
    if not source_copy.is_file():
        shutil.copy2(src, source_copy)

    config_file = proj_dir / "project.json"
    if config_file.is_file():
        data = json.loads(config_file.read_text(encoding="utf-8"))
    else:
        data = {
            "name": src.stem,
            "original_source_name": src.name,
            "source_copy": str(source_copy),
            "mode": "auto",
            "avatar_gif": str(DEFAULT_GIF),
            "avatar_idle": str(DEFAULT_IDLE),
            "pages": {},
        }

    pages = data.setdefault("pages", {})
    for s in info.slides:
        key = str(s.slide)
        if key not in pages:
            pages[key] = default_page_config(s.slide, s.avatar_box)
    save_project(proj_dir, data)
    return proj_dir, data, info


def save_project(proj_dir: Path, data: dict) -> None:
    path = Path(proj_dir) / "project.json"
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def build_pages_payload(data: dict, info: ppt_inspector.DeckInfo) -> list[dict]:
    """Build validated pages.json payload with PowerPoint point coordinates."""
    payload: list[dict] = []
    for key in sorted(data.get("pages", {}), key=lambda k: int(k)):
        cfg = data["pages"][key]
        if not cfg.get("enabled"):
            continue
        slide_num = int(cfg["slide"])
        audio_raw = cfg.get("audio_path", "").strip()
        if not audio_raw or not Path(audio_raw).is_file():
            raise ValueError(f"第 {slide_num} 页已勾选启用，但尚未生成或导入音频。")
        left_pct = float(cfg.get("left_pct", DEFAULT_BOX[0]))
        top_pct = float(cfg.get("top_pct", DEFAULT_BOX[1]))
        width_pct = float(cfg.get("width_pct", DEFAULT_BOX[2]))
        height_pct = float(cfg.get("height_pct", DEFAULT_BOX[3]))
        if (
            left_pct < 0 or top_pct < 0 or width_pct <= 0 or height_pct <= 0
            or left_pct + width_pct > 1.001 or top_pct + height_pct > 1.001
        ):
            raise ValueError(f"第 {slide_num} 页数字人位置超出幻灯片范围。")
        # Clamp tiny floating-point edge rounding inside bounds.
        left_pt = max(0.0, round(info.width_pt * left_pct, 2))
        top_pt = max(0.0, round(info.height_pt * top_pct, 2))
        w_pt = min(info.width_pt - left_pt, round(info.width_pt * width_pct, 2))
        h_pt = min(info.height_pt - top_pt, round(info.height_pt * height_pct, 2))
        payload.append({
            "slide": slide_num,
            "audio": str(Path(audio_raw).resolve()),
            "left": left_pt,
            "top": top_pt,
            "width": w_pt,
            "height": h_pt,
        })
    if not payload:
        raise ValueError("请至少勾选并配置一页需要嵌入数字人的幻灯片。")
    embedding_input.validate_pages(payload, info.slide_count)
    return payload


def export_presentation(proj_dir: Path, data: dict, info: ppt_inspector.DeckInfo, destination: Path) -> Path:
    dest = Path(destination).resolve()
    source_copy = Path(data["source_copy"]).resolve()
    if dest == source_copy:
        raise ValueError("输出文件不能覆盖项目源文件。")
    if dest.exists():
        dest.unlink()
    payload = build_pages_payload(data, info)
    pages_json = Path(proj_dir) / "export_pages.json"
    pages_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    avatar_gif = Path(data.get("avatar_gif") or DEFAULT_GIF).resolve()
    avatar_idle = Path(data.get("avatar_idle") or DEFAULT_IDLE).resolve()
    if not avatar_gif.is_file():
        raise FileNotFoundError(f"找不到数字人 GIF: {avatar_gif}")
    if not avatar_idle.is_file():
        avatar_idle = extract_idle_frame(avatar_gif, Path(proj_dir) / "avatar_idle.png")

    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(EMBED_SCRIPT),
        "-Source",
        str(source_copy),
        "-Destination",
        str(dest),
        "-Avatar",
        str(avatar_gif),
        "-IdleAvatar",
        str(avatar_idle),
        "-PagesJson",
        str(pages_json),
        "-Mode",
        data.get("mode", "auto"),
        "-PythonExe",
        sys.executable,
        "-ExperimentalIdle",
        "-PreserveAnimations",
    ]
    res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", check=False)
    if res.returncode != 0 or not dest.is_file():
        raise RuntimeError(f"嵌入 PPTX 失败: {(res.stderr or res.stdout).strip()}")
    return dest
