# 🤖 Canal de YouTube con IA — Pipeline de Shorts (faceless)

Sistema en **Python** que produce Shorts verticales de curiosidades de forma
100% automática: elige el tema → escribe el guion con **Claude** → genera la
**voz en off** → arma los **subtítulos animados palabra por palabra** → compone
el video vertical 1080×1920 → genera **título/descripción/tags con SEO** → y
opcionalmente lo **sube a YouTube**.

Está diseñado para **funcionar sin claves de pago**: la voz es gratis
(edge-tts) y, si no hay clave de IA ni conexión, cae en plantillas offline.
Con tus APIs conectadas escala a producción en serie.

> Frame real generado por el pipeline (subtítulos karaoke sincronizados):
> el fondo con gradiente + el resaltado amarillo que sigue la narración.

---

## Arquitectura

```
                 config.yaml (nicho, voz, estilo, YouTube)
                              │
   ┌──────────┐   ┌───────────┴───────────┐   ┌────────────┐
   │ ideas.py │──▶│      pipeline.py      │──▶│  output/   │
   │  (temas) │   │     (orquestador)     │   │ video.mp4  │
   └──────────┘   └───────────┬───────────┘   │ metadata   │
        ▲                     │               │ miniatura  │
   ┌────┴─────┐   ┌───────────┼───────────┐   └────────────┘
   │  llm.py  │   │           │           │
   │ (Claude) │   ▼           ▼           ▼
   └──────────┘  script.py  voice.py   captions.py
                    │          │           │
                    │      (edge-tts)  (ASS karaoke)
                    ▼          ▼           ▼
                 metadata.py  visuals.py  assemble.py ──▶ ffmpeg (libx264+libass)
                            (gradiente/Pexels)  │
                                                ▼
                                        upload.py (YouTube Data API v3)
```

| Módulo | Qué hace |
|--------|----------|
| `src/ideas.py` | Genera temas de video (IA + banco offline) |
| `src/script.py` | Escribe la narración de 30-45s (IA + plantilla) |
| `src/voice.py` | Voz en off con edge-tts + tiempos por palabra |
| `src/captions.py` | Subtítulos `.ass` con resaltado karaoke |
| `src/visuals.py` | Fondo: gradiente animado (Ken Burns) o b-roll de Pexels |
| `src/assemble.py` | Une fondo + voz + subtítulos con ffmpeg |
| `src/metadata.py` | Título, descripción y tags con SEO |
| `src/thumbnail.py` | Miniatura del video |
| `src/upload.py` | Subida automática a YouTube (OAuth) |
| `src/pipeline.py` | Orquesta todo y guarda el paquete final |
| `main.py` | CLI (`generar`, `lote`, `ideas`, `voces`) |

---

## Instalación

```bash
cd youtube_ai_channel
python -m venv .venv && source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env                                   # completá tus claves (opcional)
```

`ffmpeg` viene incluido vía `imageio-ffmpeg` — **no** hace falta instalarlo aparte.

## Uso

```bash
python main.py generar                                 # un Short con tema automático
python main.py generar --tema "Por qué el cielo es azul"
python main.py lote --cantidad 5                       # 5 Shorts de una
python main.py ideas --cantidad 10                     # solo ver ideas de tema
python main.py voces                                   # voces disponibles en español
python main.py generar --subir                         # generar y publicar en YouTube
```

Cada video se guarda en `output/<fecha>_<tema>/` con: `video.mp4`,
`metadata.json`, `guion.txt`, `miniatura.png` y `voz.mp3`.

---

## Claves (todas opcionales)

| Variable | Para qué | Sin ella… |
|----------|----------|-----------|
| `ANTHROPIC_API_KEY` | Guiones y metadata con Claude | Usa plantillas offline |
| `PEXELS_API_KEY` | B-roll de video real de fondo | Usa gradiente animado |
| `credentials/client_secret.json` | Subir a YouTube (OAuth) | Solo genera archivos locales |

Ver **[GUIA_RAPIDA.md](./GUIA_RAPIDA.md)** para el paso a paso de cada una y
**[ESTRATEGIA.md](./ESTRATEGIA.md)** para cómo llevar el canal a monetización.

---

## Notas honestas

- YouTube monetiza recién con **1000 subs + 4000 horas** (o 10M vistas de Shorts
  en 90 días). Este sistema produce el contenido; el crecimiento requiere
  constancia y buenos temas.
- Revisá siempre los datos antes de publicar: la IA puede equivocarse en un dato.
- Empezá con privacidad `private` en `config.yaml` para revisar antes de publicar.
