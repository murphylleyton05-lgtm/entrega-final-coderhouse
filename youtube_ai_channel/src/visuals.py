"""Fondo vertical del Short: gradiente animado (offline) o b-roll de Pexels."""
from __future__ import annotations

import os
import urllib.request

from PIL import Image

from . import ffutil


def _gradient_png(cfg: dict, out_png: str) -> str:
    W = cfg["video"]["ancho"]
    H = cfg["video"]["alto"]
    a = tuple(cfg["video"]["fondo"]["color_a"])
    b = tuple(cfg["video"]["fondo"]["color_b"])
    img = Image.new("RGB", (W, H))
    px = img.load()
    for y in range(H):
        t = y / max(1, H - 1)
        r = int(a[0] + (b[0] - a[0]) * t)
        g = int(a[1] + (b[1] - a[1]) * t)
        bl = int(a[2] + (b[2] - a[2]) * t)
        for x in range(W):
            px[x, y] = (r, g, bl)
    img.save(out_png)
    return out_png


def _gradient_video(cfg: dict, duration: float, workdir: str, out_mp4: str) -> str:
    W = cfg["video"]["ancho"]
    H = cfg["video"]["alto"]
    fps = cfg["video"]["fps"]
    png = _gradient_png(cfg, os.path.join(workdir, "fondo.png"))
    ken = cfg["video"]["fondo"].get("ken_burns", True)
    frames = max(1, int(duration * fps))
    if ken:
        # Zoom lento (Ken Burns) para dar movimiento a una imagen fija.
        vf = (
            f"scale={W*2}:{H*2},"
            f"zoompan=z='min(zoom+0.0008,1.2)':d={frames}:"
            f"x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':s={W}x{H}:fps={fps},"
            f"format=yuv420p"
        )
    else:
        vf = f"scale={W}:{H},format=yuv420p"
    ffutil.run([
        "-loop", "1", "-i", png,
        "-t", f"{duration:.2f}",
        "-vf", vf,
        "-r", str(fps),
        "-c:v", "libx264", "-pix_fmt", "yuv420p",
        out_mp4,
    ])
    return out_mp4


def _pexels_video(cfg: dict, query: str, duration: float, workdir: str, out_mp4: str) -> str | None:
    key = os.environ.get("PEXELS_API_KEY")
    if not key:
        return None
    try:
        url = (
            "https://api.pexels.com/videos/search"
            f"?query={urllib.parse.quote(query)}&orientation=portrait&per_page=5"
        )
        req = urllib.request.Request(url, headers={"Authorization": key})
        import json

        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.load(resp)
        videos = data.get("videos", [])
        if not videos:
            return None
        # Elegimos el primer archivo mp4 vertical de buena resolución.
        files = sorted(
            videos[0]["video_files"],
            key=lambda f: (f.get("height", 0)),
            reverse=True,
        )
        src = files[0]["link"]
        raw = os.path.join(workdir, "broll.mp4")
        urllib.request.urlretrieve(src, raw)
        W = cfg["video"]["ancho"]
        H = cfg["video"]["alto"]
        fps = cfg["video"]["fps"]
        # Recorte a 9:16, loop hasta cubrir la duración y oscurecido para leer los subs.
        ffutil.run([
            "-stream_loop", "-1", "-i", raw,
            "-t", f"{duration:.2f}",
            "-vf",
            f"scale={W}:{H}:force_original_aspect_ratio=increase,"
            f"crop={W}:{H},eq=brightness=-0.12,format=yuv420p",
            "-r", str(fps),
            "-an", "-c:v", "libx264", "-pix_fmt", "yuv420p",
            out_mp4,
        ])
        return out_mp4
    except Exception as exc:
        print(f"  [VISUAL] Pexels no disponible ({exc}). Uso gradiente.")
        return None


def make_background(cfg: dict, query: str, duration: float, workdir: str, out_mp4: str) -> str:
    """Genera el fondo del video para `duration` segundos."""
    tipo = cfg["video"]["fondo"].get("tipo", "gradiente")
    if tipo == "pexels":
        res = _pexels_video(cfg, query, duration, workdir, out_mp4)
        if res:
            return res
    return _gradient_video(cfg, duration, workdir, out_mp4)
