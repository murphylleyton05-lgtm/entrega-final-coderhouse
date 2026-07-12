"""Ensamblado final: fondo + voz + subtítulos animados -> Short mp4 vertical."""
from __future__ import annotations

from . import ffutil


def assemble(cfg: dict, bg_video: str, audio: str, ass_file: str, out_mp4: str) -> str:
    fps = cfg["video"]["fps"]
    # El filtro `subtitles` (libass) quema los .ass sobre el video.
    # `-shortest` corta al terminar la voz (el fondo se genera un poco más largo).
    ass_escaped = ass_file.replace("\\", "/").replace(":", r"\:")
    ffutil.run([
        "-i", bg_video,
        "-i", audio,
        "-vf", f"subtitles='{ass_escaped}'",
        "-map", "0:v:0", "-map", "1:a:0",
        "-c:v", "libx264", "-preset", "medium", "-pix_fmt", "yuv420p",
        "-r", str(fps),
        "-c:a", "aac", "-b:a", "192k",
        "-shortest",
        "-movflags", "+faststart",
        out_mp4,
    ])
    return out_mp4
