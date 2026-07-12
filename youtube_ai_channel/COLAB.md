# ▶️ Escuchar/generar el video CON VOZ (sin instalar nada)

Este entorno de desarrollo bloquea los servicios de voz, por eso los videos de
demo salieron mudos. Para generar el Short **con la voz en off**, usá Google
Colab (corre en el navegador, gratis, con internet libre):

## Opción A — Google Colab (recomendado, sin instalar nada)

1. Abrí este link (necesitás una cuenta de Google):

   **https://colab.research.google.com/github/murphylleyton05-lgtm/entrega-final-coderhouse/blob/claude/ai-youtube-channel-ndwiq9/youtube_ai_channel/notebooks/canal_ia_colab.ipynb**

2. Corré las celdas de arriba hacia abajo (botón ▶️).
3. En el paso 3 subí tu gameplay (opcional). En el paso 5 se reproduce el video
   **con audio** y lo podés descargar.

## Opción B — En tu computadora

```bash
cd youtube_ai_channel
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python main.py fondos --dividir mi_gameplay.mp4 --segundos 15 --saltar-fin 6.5
python main.py generar --tema "Por qué la miel nunca se echa a perder"
```

El video queda en `output/…/video.mp4` con la voz incluida.

> La voz usa **edge-tts** (gratis, sin clave). Solo necesita conexión a internet,
> que en tu PC o en Colab está disponible sin restricciones.
