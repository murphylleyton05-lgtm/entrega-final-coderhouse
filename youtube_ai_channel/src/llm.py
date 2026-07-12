"""Cliente de Claude para guiones y metadata, con degradado elegante a offline.

Si no hay `ANTHROPIC_API_KEY` o el SDK no está instalado, `available()` devuelve
False y los módulos que lo usan recurren a plantillas locales. Así el pipeline
completo corre de punta a punta incluso sin conexión ni claves.
"""
from __future__ import annotations

import json
import os
import re

try:
    import anthropic
except Exception:  # pragma: no cover - el SDK es opcional
    anthropic = None


def available() -> bool:
    return anthropic is not None and bool(os.environ.get("ANTHROPIC_API_KEY"))


def complete(system: str, user: str, model: str, max_tokens: int = 1200) -> str | None:
    """Una llamada a Claude. Devuelve el texto o None si no está disponible/falla."""
    if not available():
        return None
    try:
        client = anthropic.Anthropic()
        resp = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        return "".join(b.text for b in resp.content if b.type == "text").strip()
    except Exception as exc:  # red, cuota, etc. -> fallback
        print(f"  [IA] aviso: la llamada a Claude falló ({exc}). Uso plantilla offline.")
        return None


def extract_json(text: str):
    """Extrae el primer objeto/array JSON de una respuesta (tolerante a texto extra)."""
    if not text:
        return None
    # Bloque ```json ... ```
    fence = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
    candidate = fence.group(1) if fence else text
    # Primer { ... } o [ ... ] balanceado (búsqueda simple)
    start = min(
        [i for i in (candidate.find("{"), candidate.find("[")) if i != -1],
        default=-1,
    )
    if start == -1:
        return None
    for end in range(len(candidate), start, -1):
        try:
            return json.loads(candidate[start:end])
        except json.JSONDecodeError:
            continue
    return None
