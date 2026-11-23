"""Render an anonymous cutout and composite it over PowerPoint's slide 7 export."""

import argparse
import math
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets/avatar_placeholder"
VIDEOS = ROOT / "generated/videos"
PREVIEWS = ROOT / "generated/previews"
FONT = "C:/Windows/Fonts/msyh.ttc"
FPS = 30
SECONDS = 6
SIZE = (400, 500)
ORIGIN = (1472, 440)


def silhouette():
    scale = 3
    img = Image.new("RGBA", (SIZE[0] * scale, SIZE[1] * scale))
    draw = ImageDraw.Draw(img)

    def outline(start, curves, color):
        points = [start]
        for control1, control2, end in curves:
            first = points[-1]
            for step in range(1, 19):
                t = step / 18
                s = 1 - t
                points.append(tuple(
                    s**3 * first[i] + 3 * s**2 * t * control1[i]
                    + 3 * s * t**2 * control2[i] + t**3 * end[i]
                    for i in range(2)
                ))
        draw.polygon([(round(x * scale), round(y * scale)) for x, y in points], fill=color)

    outline((71, 437), [
        ((79, 369), (99, 313), (131, 301)),
        ((154, 288), (174, 284), (174, 261)),
        ((174, 251), (174, 242), (174, 234)),
        ((196, 234), (220, 235), (228, 235)),
        ((224, 280), (235, 285), (268, 299)),
        ((308, 314), (327, 364), (331, 437)),
        ((289, 454), (118, 454), (71, 437)),
    ], "#233e58")
    outline((199, 95), [
        ((149, 95), (123, 135), (124, 178)),
        ((121, 201), (129, 228), (142, 245)),
        ((158, 272), (180, 281), (200, 282)),
        ((223, 281), (247, 266), (261, 239)),
        ((276, 214), (280, 177), (270, 146)),
        ((260, 115), (234, 95), (199, 95)),
    ], "#526f87")
    outline((171, 286), [
        ((184, 301), (191, 308), (201, 315)),
        ((213, 307), (224, 295), (232, 286)),
        ((241, 289), (247, 293), (252, 297)),
        ((236, 317), (219, 334), (201, 344)),
        ((185, 333), (165, 317), (151, 297)),
        ((157, 293), (164, 289), (171, 286)),
    ], "#7698ac")
    return img.resize(SIZE, Image.Resampling.LANCZOS)


def draw_label(page):
    draw = ImageDraw.Draw(page)
    navy = "#233e58"
    x = 1488
    draw.rounded_rectangle((x, 398, x + 50, 405), 4, fill="#007f90")
    draw.text((x, 420), "数字人占位测试", font=ImageFont.truetype(FONT, 33), fill=navy)
    draw.text((x, 474), "静音样片 · 声音待接入", font=ImageFont.truetype(FONT, 23), fill="#4b6177")
    draw.text((x, 920), "点击播放  ·  Alt+P 重播", font=ImageFont.truetype(FONT, 22), fill="#4b6177")


def render(ffmpeg):
    avatar = silhouette()
    avatar.save(ASSETS / "silhouette.png")
    clip = VIDEOS / "avatar-alpha.webm"
    encode = [
        ffmpeg, "-hide_banner", "-loglevel", "error", "-y", "-f", "rawvideo",
        "-pix_fmt", "rgba", "-s", "400x500", "-r", str(FPS), "-i", "pipe:0",
        "-an", "-c:v", "libvpx-vp9", "-pix_fmt", "yuva420p", "-auto-alt-ref", "0",
        "-crf", "24", "-b:v", "0", str(clip),
    ]
    with subprocess.Popen(encode, stdin=subprocess.PIPE) as ffmpeg_input:
        for index in range(FPS * SECONDS):
            y = round(6 * math.sin(2 * math.pi * index / FPS / 3))
            frame = Image.new("RGBA", SIZE)
            frame.paste(avatar, (0, y))
            ffmpeg_input.stdin.write(frame.tobytes())
        ffmpeg_input.stdin.close()
        if ffmpeg_input.wait():
            raise RuntimeError("VP9 encoding failed")

    background = PREVIEWS / "original/slide-07.png"
    page = Image.open(background).convert("RGB")
    if page.size != (1920, 1080):
        raise ValueError("Expected a 1920x1080 PowerPoint slide 7 export")
    draw_label(page)
    labeled = PREVIEWS / "slide07-before-overlay.png"
    page.save(labeled)
    composite = VIDEOS / "slide07-composite.mp4"
    subprocess.run([
        ffmpeg, "-hide_banner", "-loglevel", "error", "-y", "-loop", "1",
        "-framerate", str(FPS), "-i", str(labeled), "-c:v", "libvpx-vp9", "-i", str(clip),
        "-filter_complex", f"[0:v][1:v]overlay={ORIGIN[0]}:{ORIGIN[1]}:shortest=1,format=yuv420p",
        "-an", "-c:v", "libx264", "-crf", "16", "-movflags", "+faststart", str(composite),
    ], check=True)
    subprocess.run([
        ffmpeg, "-hide_banner", "-loglevel", "error", "-y", "-i", str(composite),
        "-frames:v", "1", str(PREVIEWS / "slide07-after-overlay.png"),
    ], check=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--ffmpeg", default="ffmpeg")
    render(parser.parse_args().ffmpeg)
