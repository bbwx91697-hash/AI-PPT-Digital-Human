"""Verify the silent 1080p composite and sample its placement at 10 fps."""

import json
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parents[1]
CLIP = ROOT / "generated/videos/slide07-composite.mp4"
BACKGROUND = ROOT / "generated/previews/slide07-before-overlay.png"
ORIGINAL = ROOT / "generated/previews/original/slide-07.png"


def verify(ffmpeg, ffprobe):
    probe = subprocess.check_output([
        ffprobe, "-v", "error", "-show_streams", "-show_format", "-of", "json", str(CLIP)
    ])
    info = json.loads(probe)
    stream, = info["streams"]
    width, height = stream["width"], stream["height"]
    assert (width, height) == (1920, 1080)
    assert stream["codec_name"] == "h264" and stream["avg_frame_rate"] == "30/1"
    assert float(info["format"]["duration"]) == 6
    reference = Image.open(BACKGROUND).convert("RGB")
    original = Image.open(ORIGINAL).convert("RGB")
    assert ImageChops.difference(reference, original).getbbox()[0] >= 1450
    command = [
        ffmpeg, "-hide_banner", "-loglevel", "error", "-i", str(CLIP),
        "-vf", "fps=10", "-f", "rawvideo", "-pix_fmt", "rgb24", "pipe:1",
    ]
    sample_size = width * height * 3
    bounds = []
    with subprocess.Popen(command, stdout=subprocess.PIPE) as frames:
        while data := frames.stdout.read(sample_size):
            frame = Image.frombytes("RGB", (width, height), data)
            diff = ImageChops.difference(frame, reference).convert("L")
            above_threshold = diff.point(lambda pixel: 255 if pixel >= 55 else 0)
            box = above_threshold.getbbox()
            assert box and 1472 <= box[0] < box[2] <= 1872
            assert 510 <= box[1] < box[3] <= 915
            bounds.append(box)
        assert frames.wait() == 0
    assert len(bounds) == 60 and len(set(bounds)) > 1
    print(f"6.000 s, H.264, 1920x1080, 30 fps, no audio stream. Inspected {len(bounds)} samples at 10 fps.")
    print(f"Visible motion spans y={min(b[1] for b in bounds)}–{max(b[3] for b in bounds)} px.")


if __name__ == "__main__":
    verify(*sys.argv[1:])
