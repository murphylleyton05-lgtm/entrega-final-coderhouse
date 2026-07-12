"""Metadata SEO del video: título, descripción y tags (con IA o plantilla)."""
from __future__ import annotations

from . import llm

_SYSTEM = (
    "Sos experto en SEO de YouTube Shorts en español. Devolvés SOLO un objeto "
    "JSON con las claves: titulo (max 90 caracteres, con gancho), "
    "descripcion (2-3 frases + saltos de línea con hashtags), "
    "tags (lista de 10-15 palabras clave sin #)."
)


def _fallback(cfg: dict, topic: str) -> dict:
    hashtags = cfg["youtube"].get("hashtags_fijos", ["#shorts"])
    return {
        "titulo": f"{topic} 🤯 #shorts",
        "descripcion": (
            f"{topic}.\n\nUn dato curioso nuevo cada día. Seguí el canal "
            f"{cfg['canal']['nombre']} para más.\n\n" + " ".join(hashtags)
        ),
        "tags": [
            "curiosidades", "datos curiosos", "sabias que", "shorts",
            "aprender", "ciencia", "datos", "cultura general", "curioso",
            "increible",
        ],
    }


def build_metadata(cfg: dict, topic: str, narracion: str) -> dict:
    user = (
        f"Canal: {cfg['canal']['nombre']}\n"
        f"Tema del video: {topic}\n"
        f"Narración: {narracion}\n"
        f"Generá el JSON de metadata."
    )
    raw = llm.complete(cfg, _SYSTEM, user, max_tokens=700)
    data = llm.extract_json(raw) if raw else None
    if not isinstance(data, dict) or "titulo" not in data:
        data = _fallback(cfg, topic)

    # Aseguramos hashtags fijos en la descripción y límite de título.
    hashtags = cfg["youtube"].get("hashtags_fijos", [])
    data["titulo"] = str(data.get("titulo", topic))[:100]
    desc = str(data.get("descripcion", ""))
    for h in hashtags:
        if h not in desc:
            desc += (" " if not desc.endswith("\n") else "") + h
    data["descripcion"] = desc
    tags = data.get("tags") or []
    data["tags"] = [str(t).lstrip("#") for t in tags][:15]
    return data
