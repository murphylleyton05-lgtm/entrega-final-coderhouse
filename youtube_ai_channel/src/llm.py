"""Motor de IA para guiones y metadata. Soporta Gemini (gratis) o Claude.

Elegí el proveedor en config.yaml -> ia.proveedor ("gemini" | "claude").
Si no hay clave o falla, los módulos que lo usan recurren a plantillas offline.
"""
from __future__ import annotations

import json
import os
import re
import urllib.request

try:
    import anthropic
except Exception:  # pragma: no cover
    anthropic = None


def _gemini(system: str, user: str, model: str, key: str, max_tokens: int) -> str | None:
    url = (
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
    )
    body = {
        "system_instruction": {"parts": [{"text": system}]},
        "contents": [{"parts": [{"text": user}]}],
        "generationConfig": {"maxOutputTokens": max_tokens, "temperature": 0.9},
    }
    req = urllib.request.Request(
        url, data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json"}, method="POST",
    )
    with urllib.request.urlopen(req, timeout=45) as resp:
        data = json.load(resp)
    cand = data.get("candidates", [])
    if not cand:
        return None
    parts = cand[0].get("content", {}).get("parts", [])
    return "".join(p.get("text", "") for p in parts).strip() or None


def _claude(system: str, user: str, model: str, max_tokens: int) -> str | None:
    if anthropic is None or not os.environ.get("ANTHROPIC_API_KEY"):
        return None
    client = anthropic.Anthropic()
    resp = client.messages.create(
        model=model, max_tokens=max_tokens, system=system,
        messages=[{"role": "user", "content": user}],
    )
    return "".join(b.text for b in resp.content if b.type == "text").strip() or None


def available(cfg: dict) -> bool:
    prov = cfg.get("ia", {}).get("proveedor", "gemini")
    if prov == "gemini":
        return bool(os.environ.get("GEMINI_API_KEY"))
    return anthropic is not None and bool(os.environ.get("ANTHROPIC_API_KEY"))


def complete(cfg: dict, system: str, user: str, max_tokens: int = 1200) -> str | None:
    """Una llamada al proveedor configurado. Devuelve texto o None si no está/falla."""
    ia = cfg.get("ia", {})
    prov = ia.get("proveedor", "gemini")
    try:
        if prov == "gemini":
            key = os.environ.get("GEMINI_API_KEY")
            if key:
                return _gemini(system, user, ia.get("modelo_gemini", "gemini-2.0-flash"),
                               key, max_tokens)
        else:
            return _claude(system, user, ia.get("modelo", "claude-opus-4-8"), max_tokens)
    except Exception as exc:
        print(f"  [IA] aviso: {prov} falló ({str(exc)[:150]}). Uso plantilla offline.")
    return None


def extract_json(text: str):
    """Extrae el primer objeto/array JSON de una respuesta (tolerante a texto extra)."""
    if not text:
        return None
    fence = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
    candidate = fence.group(1) if fence else text
    start = min([i for i in (candidate.find("{"), candidate.find("[")) if i != -1], default=-1)
    if start == -1:
        return None
    for end in range(len(candidate), start, -1):
        try:
            return json.loads(candidate[start:end])
        except json.JSONDecodeError:
            continue
    return None
