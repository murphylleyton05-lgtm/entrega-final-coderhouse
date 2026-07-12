"""Carga de configuración (config.yaml + variables de entorno / .env)."""
from __future__ import annotations

import os
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent


def _load_dotenv(path: Path) -> None:
    """Carga un .env simple sin depender de python-dotenv."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key, val = key.strip(), val.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = val


def load_config(path: str | Path | None = None) -> dict:
    """Devuelve el dict de config.yaml y carga el .env si existe."""
    _load_dotenv(ROOT / ".env")
    cfg_path = Path(path) if path else ROOT / "config.yaml"
    with open(cfg_path, "r", encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh)
    # Rutas útiles resueltas
    cfg["_root"] = str(ROOT)
    cfg["_salida"] = str(ROOT / cfg.get("salida", {}).get("carpeta", "output"))
    return cfg


def env(name: str, default: str | None = None) -> str | None:
    return os.environ.get(name, default)
