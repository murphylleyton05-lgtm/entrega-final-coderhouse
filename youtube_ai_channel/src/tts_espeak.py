"""Voz offline con espeak-ng (empaquetado en pip, sin red ni descargas).

Es una voz robótica, pero 100% offline y sin claves. Sirve de respaldo cuando
edge-tts no tiene conexión (o el entorno bloquea el servicio online).
"""
from __future__ import annotations

import ctypes
import struct
import wave

_ready = False
_lib = None
_sr = 22050


def available() -> bool:
    try:
        import espeakng_loader  # noqa: F401
        return True
    except Exception:
        return False


def _init():
    global _ready, _lib, _sr
    if _ready:
        return
    import os

    import espeakng_loader

    _lib = ctypes.CDLL(espeakng_loader.get_library_path())
    path_arg = os.path.dirname(espeakng_loader.get_data_path()).encode()
    _lib.espeak_Initialize.restype = ctypes.c_int
    # 1 = AUDIO_OUTPUT_RETRIEVAL (nos devuelve el PCM por callback)
    _sr = _lib.espeak_Initialize(1, 0, path_arg, 0)
    _ready = True


def synth(text: str, out_wav: str, voice: str = "es-419", speed: int | None = None) -> float:
    """Sintetiza `text` a un WAV mono 16-bit. Devuelve la duración en segundos."""
    _init()
    samples: list[int] = []

    CB = ctypes.CFUNCTYPE(ctypes.c_int, ctypes.POINTER(ctypes.c_short), ctypes.c_int, ctypes.c_void_p)

    def _cb(wav, n, events):
        if n > 0:
            samples.extend(wav[i] for i in range(n))
        return 0

    cb = CB(_cb)  # mantener referencia viva durante la síntesis
    _lib.espeak_SetSynthCallback(cb)

    # Voz: probamos es-419 (latino) y caemos a es (España).
    if _lib.espeak_SetVoiceByName(voice.encode()) != 0:
        _lib.espeak_SetVoiceByName(b"es")

    if speed:
        # espeakRATE = 1  (80-450, default 175)
        _lib.espeak_SetParameter(1, int(speed), 0)

    data = text.encode("utf-8")
    # flags=1 => espeakCHARS_UTF8
    _lib.espeak_Synth(data, len(data) + 1, 0, 0, 0, 1, None, None)
    _lib.espeak_Synchronize()

    with wave.open(out_wav, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(_sr)
        if samples:
            w.writeframes(struct.pack("<%dh" % len(samples), *samples))
    return len(samples) / float(_sr)
