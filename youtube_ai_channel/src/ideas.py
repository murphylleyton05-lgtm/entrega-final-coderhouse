"""Generación de temas (ideas de video) para el nicho configurado."""
from __future__ import annotations

import random

from . import llm

# Banco de temas offline por si no hay IA disponible. Suficiente para arrancar
# y para demos; con la clave de Claude se generan temas frescos e ilimitados.
_FALLBACK_CURIOSIDADES = [
    "Por qué los pulpos tienen tres corazones y sangre azul",
    "El único metal líquido a temperatura ambiente además del mercurio",
    "Por qué la miel nunca se echa a perder",
    "El país que tiene más husos horarios del mundo",
    "Por qué los flamencos son rosados y no nacen así",
    "Cuánto tiempo tarda la luz del Sol en llegar a la Tierra",
    "El animal que puede sobrevivir en el espacio: el tardígrado",
    "Por qué el corazón de la ballena azul es del tamaño de un auto",
    "El lugar más seco de la Tierra donde nunca llovió",
    "Por qué los bananos son bayas pero las frutillas no",
    "Cuántas veces late tu corazón en toda tu vida",
    "El sonido que viaja más rápido en el agua que en el aire",
    "Por qué Saturno flotaría en una pileta gigante",
    "El idioma más hablado del mundo que no es el que crees",
    "Por qué no podés respirar y tragar al mismo tiempo",
    "La estrella tan densa que una cucharada pesa mil millones de toneladas",
    "Por qué los gatos ronronean incluso cuando les duele algo",
    "El desierto que alguna vez fue un océano lleno de ballenas",
    "Cuánto oro hay disuelto en todos los océanos del planeta",
    "Por qué el rayo es más caliente que la superficie del Sol",
]


def generate_topics(cfg: dict, n: int = 5, used: set[str] | None = None) -> list[str]:
    """Devuelve `n` ideas de video, evitando las ya usadas en `used`."""
    used = used or set()
    canal = cfg.get("canal", {})
    system = (
        "Sos un editor de contenido viral para YouTube Shorts en español. "
        "Generás títulos-tema breves, concretos y con gancho, de un dato curioso "
        "verificable cada uno. Nada de clickbait falso."
    )
    user = (
        f"Nicho del canal: {canal.get('descripcion_nicho', 'curiosidades')}\n"
        f"Dame {n} ideas NUEVAS de video (una línea cada una, sin numerar, sin comillas).\n"
        f"Evitá repetir estos temas ya usados:\n- " + "\n- ".join(sorted(used)[:40])
    )
    raw = llm.complete(system, user, cfg["ia"]["modelo"], max_tokens=600)
    if raw:
        ideas = [
            line.strip(" -•\t").strip()
            for line in raw.splitlines()
            if line.strip(" -•\t").strip()
        ]
        ideas = [i for i in ideas if i not in used]
        if ideas:
            return ideas[:n]

    # Fallback offline: elegimos del banco lo que no se haya usado.
    pool = [t for t in _FALLBACK_CURIOSIDADES if t not in used]
    random.shuffle(pool)
    return pool[:n] if pool else random.sample(_FALLBACK_CURIOSIDADES, min(n, len(_FALLBACK_CURIOSIDADES)))
