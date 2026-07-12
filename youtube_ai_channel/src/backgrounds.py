"""Gestión de clips de fondo: cortar un gameplay largo en segmentos distintos.

Así cada Short usa un pedazo diferente y no se ven todos iguales. Combinado con
el inicio aleatorio del pipeline, da mucha variedad a partir de un solo video.
"""
from __future__ import annotations

import math
from pathlib import Path

from . import ffutil

_VIDEO_EXTS = (".mp4", ".mov", ".mkv", ".webm", ".m4v", ".avi")


def backgrounds_dir(cfg: dict) -> Path:
    carpeta = cfg["video"]["fondo"].get("carpeta_videos", "assets/backgrounds")
    d = Path(cfg["_root"]) / carpeta
    d.mkdir(parents=True, exist_ok=True)
    return d


def list_clips(cfg: dict) -> list[Path]:
    return [p for p in backgrounds_dir(cfg).iterdir() if p.suffix.lower() in _VIDEO_EXTS]


def split_source(
    cfg: dict,
    src: str,
    seg_len: float = 15.0,
    skip_start: float = 0.0,
    skip_end: float = 0.0,
    prefijo: str = "seg",
) -> list[str]:
    """Corta `src` en segmentos de `seg_len` seg. Los guarda en assets/backgrounds/.

    `skip_start` / `skip_end` descartan segundos del principio y del final (útil
    para saltar el intro/outro de TikTok, que no sirven como fondo).
    Borra segmentos previos con el mismo prefijo, re-codifica sin audio y a
    formato uniforme. Devuelve las rutas creadas.
    """
    src_path = Path(src)
    if not src_path.exists():
        raise FileNotFoundError(f"No existe el video: {src}")

    total = ffutil.media_duration(str(src_path))
    if total <= 0:
        raise RuntimeError("No pude leer la duración del video.")

    usable_start = max(0.0, skip_start)
    usable_end = max(usable_start, total - max(0.0, skip_end))
    usable = usable_end - usable_start
    if usable < 3:
        raise RuntimeError("Después de saltar intro/outro no queda video suficiente.")

    out_dir = backgrounds_dir(cfg)
    # Limpiamos segmentos previos del mismo prefijo para no dejar basura vieja.
    for old in out_dir.glob(f"{prefijo}_*.mp4"):
        old.unlink()

    n = max(1, math.ceil(usable / seg_len))
    creados: list[str] = []
    print(
        f"Cortando '{src_path.name}' (usable {usable:.1f}s de {total:.1f}s) "
        f"en ~{n} segmentos de {seg_len:.0f}s..."
    )

    for i in range(n):
        start = usable_start + i * seg_len
        length = min(seg_len, usable_end - start)
        if length < 3:  # descartamos colas muy cortas
            continue
        out = out_dir / f"{prefijo}_{i + 1:03d}.mp4"
        ffutil.run([
            "-ss", f"{start:.2f}", "-i", str(src_path),
            "-t", f"{length:.2f}",
            "-an",  # sin audio: la voz la pone el pipeline
            "-c:v", "libx264", "-preset", "veryfast", "-pix_fmt", "yuv420p",
            str(out),
        ])
        creados.append(str(out))
        print(f"  ✓ {out.name}  ({length:.1f}s)")

    print(f"Listo: {len(creados)} segmentos en {out_dir}")
    return creados
