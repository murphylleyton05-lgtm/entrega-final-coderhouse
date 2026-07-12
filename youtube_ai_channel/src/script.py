"""Generación del guion hablado de un Short a partir de un tema."""
from __future__ import annotations

import re

from . import llm

# Estructura de un Short que retiene: gancho (0-3s) -> desarrollo -> remate/CTA.
_SYSTEM = (
    "Sos guionista de YouTube Shorts en español latino. Escribís narraciones "
    "para voz en off de 30 a 45 segundos (unas 90-120 palabras). Estructura: "
    "1) un gancho de una frase que genere curiosidad inmediata, "
    "2) el dato explicado de forma simple y asombrosa, "
    "3) un cierre corto con una micro-llamada a seguir el canal. "
    "Lenguaje hablado, frases cortas, sin emojis, sin acotaciones de escena. "
    "Solo el texto que se va a narrar."
)


def _fallback(topic: str) -> str:
    return (
        f"¿Sabías esto sobre {topic.lower()}? Prepará la cabeza porque es más "
        f"raro de lo que parece. {topic}. Y lo mejor es que la ciencia lo "
        f"explica de una forma que casi nadie conoce. La próxima vez que lo "
        f"pienses, te vas a acordar de este dato. Seguime para un dato nuevo "
        f"cada día."
    )


def write_script(cfg: dict, topic: str) -> dict:
    """Devuelve {'topic', 'narracion'} listo para convertir a voz."""
    user = (
        f"Tema del video: {topic}\n"
        f"Escribí solo la narración final, sin título ni encabezados."
    )
    raw = llm.complete(cfg, _SYSTEM, user, max_tokens=cfg["ia"]["max_tokens"])
    narracion = (raw or "").strip()
    if not narracion or len(narracion) < 40:
        narracion = _fallback(topic)
    # Limpieza: sacamos comillas envolventes y espacios raros.
    narracion = re.sub(r"\s+", " ", narracion).strip().strip('"').strip()
    return {"topic": topic, "narracion": narracion}
