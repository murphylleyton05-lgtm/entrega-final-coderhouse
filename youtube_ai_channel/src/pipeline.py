"""Orquestador: de un tema a un Short terminado (mp4 + metadata + miniatura)."""
from __future__ import annotations

import json
import re
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path

from . import assemble, captions, ffutil, ideas, metadata, script, thumbnail, visuals, voice

_USED_FILE = "temas_usados.txt"


def _slug(text: str, maxlen: int = 50) -> str:
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    text = re.sub(r"[^a-zA-Z0-9]+", "-", text).strip("-").lower()
    return text[:maxlen] or "video"


def _used_topics(cfg: dict) -> set[str]:
    p = Path(cfg["_salida"]) / _USED_FILE
    if p.exists():
        return {l.strip() for l in p.read_text(encoding="utf-8").splitlines() if l.strip()}
    return set()


def _mark_used(cfg: dict, topic: str) -> None:
    p = Path(cfg["_salida"]) / _USED_FILE
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "a", encoding="utf-8") as fh:
        fh.write(topic + "\n")


def make_video(cfg: dict, topic: str | None = None, do_upload: bool = False) -> dict:
    """Genera un video completo. Devuelve un dict con las rutas y la metadata."""
    salida = Path(cfg["_salida"])
    salida.mkdir(parents=True, exist_ok=True)

    if not topic:
        topic = ideas.generate_topics(cfg, n=1, used=_used_topics(cfg))[0]
    print(f"→ Tema: {topic}")

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_dir = salida / f"{stamp}_{_slug(topic)}"
    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as work:
        # 1) Guion
        print("→ Escribiendo guion...")
        sc = script.write_script(cfg, topic)

        # 2) Voz + tiempos por palabra
        print("→ Generando voz en off...")
        audio_path = str(out_dir / "voz.mp3")
        audio_path, words = voice.synthesize(cfg, sc["narracion"], audio_path)
        dur = ffutil.media_duration(audio_path) or (words[-1][2] if words else 30.0)

        # 3) Subtítulos animados
        print("→ Creando subtítulos animados...")
        ass_path = captions.build_ass(cfg, words, str(Path(work) / "subs.ass"))

        # 4) Fondo (un poco más largo que el audio; -shortest recorta al final)
        print("→ Generando fondo...")
        bg = visuals.make_background(cfg, topic, dur + 2.0, work, str(Path(work) / "bg.mp4"))

        # 5) Ensamblado final
        print("→ Ensamblando el Short...")
        video_path = str(out_dir / "video.mp4")
        assemble.assemble(cfg, bg, audio_path, ass_path, video_path)

        # 6) Metadata + miniatura
        print("→ Generando metadata SEO y miniatura...")
        meta = metadata.build_metadata(cfg, topic, sc["narracion"])
        thumb_path = str(out_dir / "miniatura.png")
        thumbnail.make_thumbnail(cfg, topic, thumb_path)

    # Guardamos todo el paquete
    (out_dir / "guion.txt").write_text(sc["narracion"], encoding="utf-8")
    (out_dir / "metadata.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    _mark_used(cfg, topic)

    result = {
        "topic": topic,
        "carpeta": str(out_dir),
        "video": video_path,
        "miniatura": thumb_path,
        "metadata": meta,
        "duracion": round(dur, 1),
    }

    if do_upload:
        print("→ Subiendo a YouTube...")
        from . import upload

        result["url"] = upload.upload_video(cfg, video_path, meta)

    print(f"✓ Listo: {out_dir}")
    return result


def make_batch(cfg: dict, count: int, do_upload: bool = False) -> list[dict]:
    used = _used_topics(cfg)
    topics = ideas.generate_topics(cfg, n=count, used=used)
    results = []
    for t in topics:
        try:
            results.append(make_video(cfg, topic=t, do_upload=do_upload))
        except Exception as exc:
            print(f"✗ Error con '{t}': {exc}")
    return results
