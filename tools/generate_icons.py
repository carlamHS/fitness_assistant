#!/usr/bin/env python3
"""
Generate app icons for NG Fitness Assistant.

Outputs:
- assets/icons/icon-1024.png
- assets/icons/icon-512.png
- assets/icons/icon-192.png
- assets/icons/icon-180.png
- assets/icons/favicon-32.png
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "assets" / "icons"

BG_COLOR = "#0A7A4D"
ACCENT_COLOR = "#D26E1D"
MAIN_COLOR = "#FFFFFF"


def rounded_box(draw: ImageDraw.ImageDraw, bbox: tuple[float, float, float, float], radius: float, fill: str) -> None:
    x0, y0, x1, y1 = bbox
    left = int(round(min(x0, x1)))
    top = int(round(min(y0, y1)))
    right = int(round(max(x0, x1)))
    bottom = int(round(max(y0, y1)))
    width = max(1, right - left)
    height = max(1, bottom - top)
    safe_radius = int(round(min(radius, width / 2, height / 2)))
    draw.rounded_rectangle((left, top, right, bottom), radius=safe_radius, fill=fill)


def draw_dumbbell(draw: ImageDraw.ImageDraw, size: int) -> None:
    cx = size * 0.5
    cy = size * 0.54

    plate_w = size * 0.095
    plate_h = size * 0.27
    gap = size * 0.16
    corner = size * 0.03

    left_x2 = cx - gap
    left_x1 = left_x2 - plate_w
    right_x1 = cx + gap
    right_x2 = right_x1 + plate_w

    y1 = cy - plate_h / 2
    y2 = cy + plate_h / 2

    rounded_box(draw, (left_x1, y1, left_x2, y2), corner, MAIN_COLOR)
    rounded_box(draw, (right_x1, y1, right_x2, y2), corner, MAIN_COLOR)

    # Bar
    bar_h = size * 0.09
    bar_y1 = cy - bar_h / 2
    bar_y2 = cy + bar_h / 2
    rounded_box(draw, (left_x2 - size * 0.02, bar_y1, right_x1 + size * 0.02, bar_y2), bar_h / 2, MAIN_COLOR)


def draw_protein_drop(draw: ImageDraw.ImageDraw, size: int) -> None:
    # Teardrop accent at top-right
    cx = size * 0.76
    cy = size * 0.31
    r = size * 0.10

    draw.polygon(
        [
            (cx, cy - r * 1.6),
            (cx - r * 0.78, cy - r * 0.1),
            (cx + r * 0.78, cy - r * 0.1),
        ],
        fill=ACCENT_COLOR,
    )
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=ACCENT_COLOR)

    # Small highlight
    hr = r * 0.22
    draw.ellipse((cx - r * 0.45 - hr, cy - r * 0.38 - hr, cx - r * 0.45 + hr, cy - r * 0.38 + hr), fill="#F8D9BF")


def build_master(size: int = 1024) -> Image.Image:
    img = Image.new("RGBA", (size, size), BG_COLOR)
    draw = ImageDraw.Draw(img)
    draw_dumbbell(draw, size)
    draw_protein_drop(draw, size)
    return img


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    master = build_master(1024)

    outputs = {
        "icon-1024.png": 1024,
        "icon-512.png": 512,
        "icon-192.png": 192,
        "icon-180.png": 180,
        "favicon-32.png": 32,
    }

    for name, size in outputs.items():
        target = OUT_DIR / name
        icon = master.resize((size, size), Image.Resampling.LANCZOS)
        icon.save(target, format="PNG", optimize=True)
        print(f"Wrote {target}")


if __name__ == "__main__":
    main()
