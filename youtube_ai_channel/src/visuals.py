"""Fondo vertical del Short.

Modos (config -> video.fondo.tipo):
  - "gameplay":  video en loop tipo satisfying/gameplay desde assets/backgrounds/
                 (recortado a 9:16, con inicio aleatorio y oscurecido). EL RECOMENDADO.
  - "pexels":    b-roll vertical descargado de Pexels (requiere PEXELS_API_KEY).
  - "gradiente": gradiente animado (Ken Burns). Solo como último recurso.

Si el modo es "gameplay" pero no hay clips, cae en un fondo DINÁMICO animado
(no plano) para que el video siga teniendo movimiento.
"""
from __future__ import annotations

import json
import os
import random
import urllib.parse
import urllib.request
from pathlib import Path

from PIL import Image

from . import ffutil

_VIDEO_EXTS = (".mp4", ".mov", ".mkv", ".webm", ".m4v", ".avi")


# --------------------------------------------------------------------------- #
# Modo GAMEPLAY: video en loop desde una carpeta de clips
# --------------------------------------------------------------------------- #
def _find_clips(cfg: dict) -> list[str]:
    carpeta = cfg["video"]["fondo"].get("carpeta_videos", "assets/backgrounds")
    base = Path(cfg["_root"]) / carpeta
    if not base.exists():
        return []
    return [str(p) for p in base.iterdir() if p.suffix.lower() in _VIDEO_EXTS]


def _postproc_filters(cfg: dict) -> str:
    """Filtros comunes: cubrir 9:16, (opcional) recortar marca de agua, oscurecer.

    Para sacar la marca de agua (típicamente abajo a la derecha) hacemos un zoom
    y sesgamos el recorte hacia arriba-izquierda, empujando esa esquina fuera de
    cuadro. Se controla con video.fondo.recortar_marca / zoom_marca.
    """
    W = cfg["video"]["ancho"]
    H = cfg["video"]["alto"]
    fondo = cfg["video"]["fondo"]
    brillo = fondo.get("oscurecer", -0.12)
    sat = fondo.get("saturacion", 1.1)

    if fondo.get("recortar_marca", False):
        z = float(fondo.get("zoom_marca", 1.22))
        w2, h2 = int(W * z), int(H * z)
        # Sesgo del recorte: 0 = pegado arriba-izquierda (saca abajo-derecha).
        xf = float(fondo.get("marca_sesgo_x", 0.12))
        yf = float(fondo.get("marca_sesgo_y", 0.0))
        scale_crop = (
            f"scale={w2}:{h2}:force_original_aspect_ratio=increase,"
            f"crop={W}:{H}:'(in_w-out_w)*{xf}':'(in_h-out_h)*{yf}'"
        )
    else:
        scale_crop = (
            f"scale={W}:{H}:force_original_aspect_ratio=increase,crop={W}:{H}"
        )

    return f"{scale_crop},eq=brightness={brillo}:saturation={sat},format=yuv420p"


def _gameplay_video(cfg: dict, duration: float, out_mp4: str) -> str | None:
    clips = _find_clips(cfg)
    if not clips:
        return None
    src = random.choice(clips)
    fps = cfg["video"]["fps"]
    src_dur = ffutil.media_duration(src)

    args: list[str] = []
    if src_dur > duration + 0.5:
        # Punto de inicio aleatorio para que no se repita el mismo segmento.
        start = round(random.uniform(0, src_dur - duration - 0.2), 2)
        args += ["-ss", f"{start}", "-i", src]
    else:
        # Clip corto: lo ponemos en loop infinito y cortamos por duración.
        args += ["-stream_loop", "-1", "-i", src]

    ffutil.run([
        *args,
        "-t", f"{duration:.2f}",
        "-vf", _postproc_filters(cfg),
        "-r", str(fps),
        "-an", "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p",
        out_mp4,
    ])
    print(f"  [VISUAL] Fondo gameplay: {Path(src).name}")
    return out_mp4


# --------------------------------------------------------------------------- #
# Modo PEXELS: b-roll relacionado al tema
# --------------------------------------------------------------------------- #
def _pexels_video(cfg: dict, query: str, duration: float, workdir: str, out_mp4: str) -> str | None:
    key = os.environ.get("PEXELS_API_KEY")
    if not key:
        return None
    try:
        url = (
            "https://api.pexels.com/videos/search?"
            + urllib.parse.urlencode({"query": query, "orientation": "portrait", "per_page": 5})
        )
        req = urllib.request.Request(url, headers={"Authorization": key})
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.load(resp)
        videos = data.get("videos", [])
        if not videos:
            return None
        files = sorted(videos[0]["video_files"], key=lambda f: f.get("height", 0), reverse=True)
        raw = os.path.join(workdir, "broll.mp4")
        urllib.request.urlretrieve(files[0]["link"], raw)
        fps = cfg["video"]["fps"]
        ffutil.run([
            "-stream_loop", "-1", "-i", raw,
            "-t", f"{duration:.2f}",
            "-vf", _postproc_filters(cfg),
            "-r", str(fps),
            "-an", "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p",
            out_mp4,
        ])
        print(f"  [VISUAL] Fondo Pexels para: {query}")
        return out_mp4
    except Exception as exc:
        print(f"  [VISUAL] Pexels no disponible ({exc}).")
        return None


# --------------------------------------------------------------------------- #
# Fondo DINÁMICO animado (sin nada externo) — reemplaza al gradiente plano
# --------------------------------------------------------------------------- #
def _dynamic_video(cfg: dict, duration: float, out_mp4: str) -> str:
    """Fondo hipnótico generado con ffmpeg (mandelbrot con zoom). No es plano."""
    W = cfg["video"]["ancho"]
    H = cfg["video"]["alto"]
    fps = cfg["video"]["fps"]
    brillo = cfg["video"]["fondo"].get("oscurecer", -0.15)
    ffutil.run([
        "-f", "lavfi",
        "-i", f"mandelbrot=size={W}x{H}:rate={fps}:maxiter=200",
        "-t", f"{duration:.2f}",
        "-vf", f"eq=brightness={brillo}:saturation=1.2,format=yuv420p",
        "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p",
        out_mp4,
    ])
    print("  [VISUAL] Fondo dinámico animado (agregá clips a assets/backgrounds/ "
          "para el estilo gameplay).")
    return out_mp4


# --------------------------------------------------------------------------- #
# Modo GRADIENTE (último recurso)
# --------------------------------------------------------------------------- #
def _gradient_video(cfg: dict, duration: float, workdir: str, out_mp4: str) -> str:
    W = cfg["video"]["ancho"]
    H = cfg["video"]["alto"]
    fps = cfg["video"]["fps"]
    a = tuple(cfg["video"]["fondo"].get("color_a", [18, 22, 48]))
    b = tuple(cfg["video"]["fondo"].get("color_b", [90, 30, 90]))
    png = os.path.join(workdir, "fondo.png")
    img = Image.new("RGB", (W, H))
    px = img.load()
    for y in range(H):
        t = y / max(1, H - 1)
        row = (int(a[0] + (b[0]-a[0])*t), int(a[1] + (b[1]-a[1])*t), int(a[2] + (b[2]-a[2])*t))
        for x in range(W):
            px[x, y] = row
    img.save(png)
    frames = max(1, int(duration * fps))
    vf = (
        f"scale={W*2}:{H*2},zoompan=z='min(zoom+0.0008,1.2)':d={frames}:"
        f"x='iw/2-(iw/zoom/2)':y='ih/2-(ih/zoom/2)':s={W}x{H}:fps={fps},format=yuv420p"
    )
    ffutil.run([
        "-loop", "1", "-i", png, "-t", f"{duration:.2f}",
        "-vf", vf, "-r", str(fps),
        "-c:v", "libx264", "-pix_fmt", "yuv420p", out_mp4,
    ])
    return out_mp4


# --------------------------------------------------------------------------- #
# Punto de entrada
# --------------------------------------------------------------------------- #
def make_background(cfg: dict, query: str, duration: float, workdir: str, out_mp4: str) -> str:
    tipo = cfg["video"]["fondo"].get("tipo", "gameplay")

    if tipo == "gameplay":
        res = _gameplay_video(cfg, duration, out_mp4)
        if res:
            return res
        # Sin clips: probamos Pexels y si no, fondo dinámico (nunca plano).
        pex_query = cfg["video"]["fondo"].get("pexels_query", "satisfying")
        res = _pexels_video(cfg, pex_query, duration, workdir, out_mp4)
        return res or _dynamic_video(cfg, duration, out_mp4)

    if tipo == "pexels":
        res = _pexels_video(cfg, query, duration, workdir, out_mp4)
        return res or _dynamic_video(cfg, duration, out_mp4)

    return _gradient_video(cfg, duration, workdir, out_mp4)
