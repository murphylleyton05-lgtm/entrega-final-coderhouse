"""Genera subtítulos .ass con resaltado palabra por palabra (estilo Shorts virales).

Usa karaoke de libass (\\kf): las palabras aparecen en blanco y se "encienden"
en amarillo a medida que la voz las pronuncia. Súper efectivo para retención.
"""
from __future__ import annotations

from pathlib import Path


def _fmt_time(seconds: float) -> str:
    """Formato de tiempo de ASS: H:MM:SS.cc (centésimas)."""
    if seconds < 0:
        seconds = 0
    cs = int(round(seconds * 100))
    h, cs = divmod(cs, 360000)
    m, cs = divmod(cs, 6000)
    s, cs = divmod(cs, 100)
    return f"{h}:{m:02d}:{s:02d}.{cs:02d}"


def _chunk(words, size):
    for i in range(0, len(words), size):
        yield words[i : i + size]


def build_ass(cfg: dict, word_timings, out_path: str) -> str:
    sub = cfg["video"]["subtitulos"]
    W = cfg["video"]["ancho"]
    H = cfg["video"]["alto"]
    per_line = int(sub.get("palabras_por_linea", 4))

    style = (
        f"Style: Main,{sub['fuente']},{sub['tamano']},"
        f"{sub['color_resaltado']},{sub['color_texto']},"  # Primary(sung)=resaltado, Secondary=texto
        f"{sub['color_borde']},&H64000000,"                 # Outline, Back
        f"1,0,0,0,100,100,0,0,1,{sub['grosor_borde']},2,5,60,60,120,1"
        # bold=1 ... Alignment=5 (centro), márgenes L/R/V
    )

    header = f"""[Script Info]
ScriptType: v4.00+
PlayResX: {W}
PlayResY: {H}
WrapStyle: 2
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
{style}

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""

    lines = []
    for group in _chunk(list(word_timings), per_line):
        if not group:
            continue
        start = group[0][1]
        end = group[-1][2]
        # Construye el texto con karaoke: \kf<duración en centésimas> por palabra.
        parts = []
        for idx, (word, w_start, w_end) in enumerate(group):
            nxt = group[idx + 1][1] if idx + 1 < len(group) else end
            k_cs = max(1, int(round((nxt - w_start) * 100)))
            clean = word.replace("{", "").replace("}", "").strip()
            parts.append(f"{{\\kf{k_cs}}}{clean} ")
        text = "".join(parts).strip()
        lines.append(
            f"Dialogue: 0,{_fmt_time(start)},{_fmt_time(end)},Main,,0,0,0,,{text}"
        )

    Path(out_path).write_text(header + "\n".join(lines) + "\n", encoding="utf-8")
    return out_path
