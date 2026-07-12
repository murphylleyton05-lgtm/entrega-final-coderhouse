#!/usr/bin/env python3
"""CLI del canal de YouTube con IA.

Ejemplos:
  python main.py generar                      # un Short con tema automático
  python main.py generar --tema "Por qué el cielo es azul"
  python main.py lote --cantidad 5            # 5 Shorts de una
  python main.py generar --subir              # generar y publicar en YouTube
  python main.py ideas --cantidad 10          # solo listar ideas de tema
  python main.py voces                        # listar voces disponibles (español)
"""
from __future__ import annotations

import argparse
import asyncio
import sys

from src import backgrounds as bg_mod
from src import ideas as ideas_mod
from src import voice
from src.config import load_config
from src.pipeline import make_batch, make_video


def main(argv=None):
    parser = argparse.ArgumentParser(description="Canal de YouTube con IA (Shorts).")
    parser.add_argument("--config", default=None, help="Ruta a config.yaml alternativo")
    sub = parser.add_subparsers(dest="cmd", required=True)

    g = sub.add_parser("generar", help="Genera un Short")
    g.add_argument("--tema", default=None, help="Tema específico (si no, se genera solo)")
    g.add_argument("--subir", action="store_true", help="Subir a YouTube al terminar")

    b = sub.add_parser("lote", help="Genera varios Shorts")
    b.add_argument("--cantidad", type=int, default=3)
    b.add_argument("--subir", action="store_true")

    i = sub.add_parser("ideas", help="Lista ideas de tema")
    i.add_argument("--cantidad", type=int, default=10)

    sub.add_parser("voces", help="Lista voces de edge-tts en español")

    f = sub.add_parser("fondos", help="Gestión de clips de fondo (gameplay)")
    f.add_argument("--dividir", metavar="ARCHIVO", help="Corta un gameplay en segmentos")
    f.add_argument("--segundos", type=float, default=15.0, help="Duración de cada segmento")
    f.add_argument("--saltar-inicio", type=float, default=0.0, help="Segundos a saltar del principio (intro)")
    f.add_argument("--saltar-fin", type=float, default=0.0, help="Segundos a saltar del final (outro)")
    f.add_argument("--listar", action="store_true", help="Lista los clips de fondo")

    args = parser.parse_args(argv)
    cfg = load_config(args.config)

    if args.cmd == "generar":
        make_video(cfg, topic=args.tema, do_upload=args.subir)
    elif args.cmd == "lote":
        res = make_batch(cfg, count=args.cantidad, do_upload=args.subir)
        print(f"\n✓ {len(res)} videos generados en {cfg['_salida']}")
    elif args.cmd == "ideas":
        for t in ideas_mod.generate_topics(cfg, n=args.cantidad):
            print(f"• {t}")
    elif args.cmd == "voces":
        voces = asyncio.run(voice.list_voices("es"))
        if not voces:
            print("edge-tts no disponible o sin red. Instalá: pip install edge-tts")
        for v in voces:
            print(v)
    elif args.cmd == "fondos":
        if args.dividir:
            bg_mod.split_source(
                cfg, args.dividir, seg_len=args.segundos,
                skip_start=args.saltar_inicio, skip_end=args.saltar_fin,
            )
        elif args.listar or True:
            clips = bg_mod.list_clips(cfg)
            if not clips:
                print("No hay clips en assets/backgrounds/. Cortá un gameplay con:")
                print("  python main.py fondos --dividir tu_gameplay.mp4")
            for c in clips:
                print(f"• {c.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
