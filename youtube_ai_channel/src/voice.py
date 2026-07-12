"""Voz en off con edge-tts (gratis) y fallback offline (silencio con tiempos).

Devuelve la ruta del audio y una lista de tiempos por palabra
[(texto, inicio_seg, fin_seg)] que usan los subtítulos para sincronizarse.
"""
from __future__ import annotations

import asyncio

try:
    import edge_tts
except Exception:  # pragma: no cover
    edge_tts = None

from . import ffutil


async def _edge_synth(text: str, voice: str, rate: str, pitch: str, out_path: str):
    comm = edge_tts.Communicate(text, voice, rate=rate, pitch=pitch)
    audio = bytearray()
    words: list[tuple[str, float, float]] = []
    async for chunk in comm.stream():
        if chunk["type"] == "audio":
            audio.extend(chunk["data"])
        elif chunk["type"] == "WordBoundary":
            start = chunk["offset"] / 1e7
            dur = chunk["duration"] / 1e7
            words.append((chunk["text"], start, start + dur))
    with open(out_path, "wb") as fh:
        fh.write(bytes(audio))
    return words


def _estimate_timings(text: str) -> tuple[list[tuple[str, float, float]], float]:
    """Reparte los tiempos de forma uniforme (~2.6 palabras/seg) para el fallback."""
    tokens = [w for w in text.split() if w]
    wps = 2.6
    per = 1.0 / wps
    words, t = [], 0.0
    for w in tokens:
        words.append((w, t, t + per))
        t += per
    return words, t + 0.6  # colita de silencio


def synthesize(cfg: dict, text: str, out_path: str) -> tuple[str, list[tuple[str, float, float]]]:
    voz = cfg.get("voz", {})
    voice = voz.get("voz", "es-MX-JorgeNeural")
    rate = voz.get("velocidad", "+0%")
    pitch = voz.get("tono", "+0Hz")

    if edge_tts is not None:
        try:
            words = asyncio.run(_edge_synth(text, voice, rate, pitch, out_path))
            if words:
                return out_path, words
        except Exception as exc:
            print(f"  [VOZ] edge-tts no disponible ({exc}). Genero pista de respaldo.")

    # Fallback: audio silencioso de la duración estimada + tiempos calculados.
    words, dur = _estimate_timings(text)
    ffutil.run([
        "-f", "lavfi",
        "-i", f"anullsrc=r=24000:cl=mono",
        "-t", f"{dur:.2f}",
        "-c:a", "libmp3lame", "-q:a", "9",
        out_path,
    ])
    print("  [VOZ] AVISO: se usó audio de respaldo SIN voz (no había red para edge-tts).")
    return out_path, words


async def list_voices(lang_prefix: str = "es") -> list[str]:
    if edge_tts is None:
        return []
    voices = await edge_tts.list_voices()
    return sorted(
        v["ShortName"] for v in voices if v["ShortName"].startswith(lang_prefix)
    )
