"""Pure normalized (0..1) layout geometry helpers for interactive avatar positioning."""

from __future__ import annotations

MIN_WIDTH_PCT = 0.05
MIN_HEIGHT_PCT = 0.08


def clamp_position(left: float, top: float, width: float, height: float) -> tuple[float, float, float, float]:
    """Clamp box inside [0, 1] x [0, 1] while preserving width and height."""
    w = max(MIN_WIDTH_PCT, min(1.0, float(width)))
    h = max(MIN_HEIGHT_PCT, min(1.0, float(height)))
    l = max(0.0, min(1.0 - w, float(left)))
    t = max(0.0, min(1.0 - h, float(top)))
    return round(l, 4), round(t, 4), round(w, 4), round(h, 4)


def aspect_height_for_width(
    width_pct: float,
    slide_width_pt: float,
    slide_height_pt: float,
    img_width_px: int,
    img_height_px: int,
) -> float:
    """Compute height_pct matching image natural aspect ratio on the slide."""
    if slide_width_pt <= 0 or slide_height_pt <= 0 or img_width_px <= 0 or img_height_px <= 0:
        return width_pct
    w_pt = width_pct * slide_width_pt
    h_pt = w_pt * (img_height_px / float(img_width_px))
    return h_pt / slide_height_pt


def resize_from_handle(
    handle: str,
    orig_box: tuple[float, float, float, float],
    dx_pct: float,
    dy_pct: float,
    *,
    lock_aspect: bool = True,
    aspect_ratio_hw: float | None = None,
) -> tuple[float, float, float, float]:
    """Resize normalized (left, top, width, height) from a corner handle ('nw','ne','sw','se')."""
    ol, ot, ow, oh = orig_box
    oright = ol + ow
    obottom = ot + oh

    if handle not in ("nw", "ne", "sw", "se"):
        return clamp_position(ol + dx_pct, ot + dy_pct, ow, oh)

    if "w" in handle:
        new_left = max(0.0, min(oright - MIN_WIDTH_PCT, ol + dx_pct))
        new_w = oright - new_left
    else:
        new_left = ol
        new_w = max(MIN_WIDTH_PCT, min(1.0 - ol, ow + dx_pct))

    if lock_aspect and aspect_ratio_hw and aspect_ratio_hw > 0:
        new_h = new_w * aspect_ratio_hw
        if "n" in handle:
            new_top = obottom - new_h
            if new_top < 0.0:
                new_top = 0.0
                new_h = obottom
                new_w = new_h / aspect_ratio_hw
                new_left = oright - new_w if "w" in handle else ol
        else:
            new_top = ot
            if new_top + new_h > 1.0:
                new_h = 1.0 - new_top
                new_w = new_h / aspect_ratio_hw
                new_left = oright - new_w if "w" in handle else ol
        return clamp_position(new_left, new_top, new_w, new_h)

    if "n" in handle:
        new_top = max(0.0, min(obottom - MIN_HEIGHT_PCT, ot + dy_pct))
        new_h = obottom - new_top
    else:
        new_top = ot
        new_h = max(MIN_HEIGHT_PCT, min(1.0 - ot, oh + dy_pct))

    return clamp_position(new_left, new_top, new_w, new_h)
