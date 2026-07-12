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

from . import ffutil, tts_espeak


def _timings_from_duration(text: str, duration: float) -> list[tuple[str, float, float]]:
    """Reparte los tiempos de palabra proporcionalmente a una duración REAL de audio."""
    tokens = [w for w in text.split() if w]
    if not tokens:
        return []
    weights = [max(1, len(w)) for w in tokens]
    total = sum(weights)
    body = max(0.1, duration - 0.15)  # colita de silencio
    words, t = [], 0.0
    for w, wt in zip(tokens, weights):
        d = body * wt / total
        words.append((w, t, t + d))
        t += d
    return words


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

    # 1) edge-tts (voz neuronal, la mejor) — requiere internet.
    if edge_tts is not None:
        try:
            words = asyncio.run(_edge_synth(text, voice, rate, pitch, out_path))
            if words:
                return out_path, words
        except Exception as exc:
            print(f"  [VOZ] edge-tts sin conexión ({exc}). Uso voz offline (espeak).")

    # 2) espeak-ng (voz offline robótica, empaquetada en pip) — SIEMPRE audible.
    if tts_espeak.available():
        try:
            import os

            wav = os.path.splitext(out_path)[0] + ".wav"
            espeak_voice = voz.get("voz_espeak", "es-419")
            espeak_speed = voz.get("espeak_velocidad", 165)
            dur = tts_espeak.synth(text, wav, voice=espeak_voice, speed=espeak_speed)
            # Pasamos a mp3 para uniformidad.
            ffutil.run(["-i", wav, "-c:a", "libmp3lame", "-q:a", "4", out_path])
            os.remove(wav)
            print("  [VOZ] Voz offline (espeak-ng).")
            return out_path, _timings_from_duration(text, dur)
        except Exception as exc:
            print(f"  [VOZ] espeak falló ({exc}). Uso audio silencioso.")

    # 3) Último recurso: audio silencioso (para no romper el armado).
    words, dur = _estimate_timings(text)
    ffutil.run([
        "-f", "lavfi", "-i", "anullsrc=r=24000:cl=mono", "-t", f"{dur:.2f}",
        "-c:a", "libmp3lame", "-q:a", "9", out_path,
    ])
    print("  [VOZ] AVISO: audio SIN voz (no había edge-tts ni espeak).")
    return out_path, words


async def list_voices(lang_prefix: str = "es") -> list[str]:
    if edge_tts is None:
        return []
    voices = await edge_tts.list_voices()
    return sorted(
        v["ShortName"] for v in voices if v["ShortName"].startswith(lang_prefix)
    )
