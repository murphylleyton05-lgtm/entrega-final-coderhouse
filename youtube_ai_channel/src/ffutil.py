"""Utilidades de ffmpeg: localizar el binario, correr comandos y medir duración."""
from __future__ import annotations

import re
import shutil
import subprocess


def ffmpeg_bin() -> str:
    """Devuelve la ruta a ffmpeg. Prefiere el binario estático de imageio-ffmpeg."""
    try:
        import imageio_ffmpeg

        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        found = shutil.which("ffmpeg")
        if found:
            return found
        raise RuntimeError(
            "No se encontró ffmpeg. Instalá dependencias: pip install imageio-ffmpeg"
        )


def run(args: list[str]) -> subprocess.CompletedProcess:
    """Corre ffmpeg con los argumentos dados (sin incluir el binario)."""
    cmd = [ffmpeg_bin(), "-y", "-hide_banner", "-loglevel", "error", *args]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"ffmpeg falló:\n{' '.join(cmd)}\n{proc.stderr}")
    return proc


def media_duration(path: str) -> float:
    """Duración en segundos de un archivo de audio/video (parseando la salida de ffmpeg)."""
    proc = subprocess.run(
        [ffmpeg_bin(), "-hide_banner", "-i", path],
        capture_output=True,
        text=True,
    )
    # ffmpeg imprime la duración en stderr aunque "falle" por no tener salida.
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+\.\d+)", proc.stderr)
    if not m:
        return 0.0
    h, mn, s = int(m.group(1)), int(m.group(2)), float(m.group(3))
    return h * 3600 + mn * 60 + s
