# 🚀 Guía rápida — de cero al primer Short

Esta guía asume que **no tenés nada configurado**. Al final vas a tener videos
generándose solos y, si querés, subiéndose a tu canal.

## 1. Preparar el entorno (5 min)

```bash
cd youtube_ai_channel
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Probá que funcione **sin ninguna clave** (usa voz gratis + plantillas):

```bash
python main.py generar --tema "Por qué la miel nunca se echa a perder"
```

Vas a encontrar el video en `output/…/video.mp4`. ¡Ya tenés tu primer Short!

> Nota: la voz `edge-tts` necesita conexión a internet. En algunos entornos
> corporativos o sandboxes puede estar bloqueada; ahí el sistema genera una
> pista de respaldo sin voz para que puedas ver el armado. En tu PC normal
> funciona directo.

## 2. Guiones y metadata con IA (opcional pero recomendado)

1. Creá una cuenta en https://console.anthropic.com y generá una API key.
2. Copiá `.env.example` a `.env` y pegá tu clave en `ANTHROPIC_API_KEY=`.
3. Listo: ahora los guiones y el SEO los escribe Claude, con temas ilimitados.

Para bajar costos podés cambiar el modelo en `config.yaml` a
`claude-haiku-4-5` (más barato, ideal para volumen).

## 3. Fondo de video (estilo gameplay/satisfying) — recomendado

El canal ya viene configurado en modo `gameplay`. Para que se vea pro:

1. Conseguí 1-3 clips de fondo (gameplay propio, o "satisfying"/loops de licencia
   libre en https://www.pexels.com/videos/ o https://pixabay.com/videos/).
2. Guardalos en `assets/backgrounds/` (formato `.mp4`).
3. Listo: el sistema los usa, los recorta a 9:16 y los pone en loop solo.

> ⚠️ No uses gameplay con copyright ajeno (Minecraft/GTA de otros, etc.):
> YouTube puede reclamar el video. Ver `assets/backgrounds/LEEME.md`.

**Alternativa (b-roll del tema):** sacá una API key gratis en
https://www.pexels.com/api/, ponela en `.env` como `PEXELS_API_KEY=`, y cambiá
`fondo.tipo` a `pexels` en `config.yaml` para que baje video relacionado a cada tema.

## 4. Subir a YouTube automáticamente (opcional)

1. Entrá a https://console.cloud.google.com y creá un proyecto.
2. Habilitá **"YouTube Data API v3"**.
3. Creá credenciales **OAuth 2.0** tipo **"App de escritorio"** y descargá el JSON.
4. Guardalo como `credentials/client_secret.json` dentro de este proyecto.
5. Corré:

   ```bash
   python main.py generar --subir
   ```

   La primera vez se abre el navegador para autorizar tu cuenta. Se guarda un
   `credentials/token.json` y las próximas subidas ya no piden nada.

> Empezá con `privacidad: private` en `config.yaml`. Revisá el video en tu canal
> y recién ahí cambialo a `public`. La API tiene una cuota diaria (~6 subidas/día
> por defecto); pedí ampliación en Google Cloud si vas a publicar más.

## 5. Producción en serie

```bash
python main.py lote --cantidad 7        # una semana de contenido
python main.py lote --cantidad 7 --subir
```

El sistema recuerda los temas ya usados (`output/temas_usados.txt`) para no
repetir. Combinalo con el cron/Task Scheduler de tu sistema para publicar solo.

## Personalizar el canal

Todo se cambia en **`config.yaml`** sin tocar código:
- `canal.nombre` y `canal.descripcion_nicho` → tema del canal
- `voz.voz` → probá otras con `python main.py voces`
- `video.fondo.color_a / color_b` → paleta del gradiente
- `video.subtitulos` → fuente, tamaño y colores de los subtítulos
- `youtube.privacidad`, `youtube.categoria_id`, `youtube.hashtags_fijos`
