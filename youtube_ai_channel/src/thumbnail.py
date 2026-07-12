"""Miniatura simple del video (útil para catálogo; en Shorts es secundaria)."""
from __future__ import annotations

import textwrap

from PIL import Image, ImageDraw, ImageFont


def _font(size: int):
    for path in (
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "DejaVuSans-Bold.ttf",
    ):
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def make_thumbnail(cfg: dict, topic: str, out_path: str) -> str:
    W = cfg["video"]["ancho"]
    H = cfg["video"]["alto"]
    a = tuple(cfg["video"]["fondo"]["color_a"])
    b = tuple(cfg["video"]["fondo"]["color_b"])
    img = Image.new("RGB", (W, H))
    px = img.load()
    for y in range(H):
        t = y / max(1, H - 1)
        px_row = (
            int(a[0] + (b[0] - a[0]) * t),
            int(a[1] + (b[1] - a[1]) * t),
            int(a[2] + (b[2] - a[2]) * t),
        )
        for x in range(W):
            px[x, y] = px_row

    draw = ImageDraw.Draw(img)
    font = _font(96)
    wrapped = textwrap.fill(topic.upper(), width=16)
    # Texto centrado con borde negro para legibilidad.
    bbox = draw.multiline_textbbox((0, 0), wrapped, font=font, align="center")
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x, y = (W - tw) / 2, (H - th) / 2
    for dx in (-3, 0, 3):
        for dy in (-3, 0, 3):
            draw.multiline_text((x + dx, y + dy), wrapped, font=font, fill="black", align="center")
    draw.multiline_text((x, y), wrapped, font=font, fill="white", align="center")
    img.save(out_path)
    return out_path
