# 💰 Estrategia de crecimiento y monetización

Sé honesto con vos mismo: **no hay botón mágico de dinero**. Este pipeline
resuelve la parte más pesada (producir contenido consistente y de buena calidad
técnica). El resultado depende de constancia, nicho y calidad de temas. Acá va
un plan realista.

## Las reglas de monetización de YouTube (2025)

Para entrar al Programa de Socios (YPP) y empezar a cobrar necesitás **una** de:

- **1000 suscriptores + 4000 horas** de visualización pública (últimos 12 meses), **o**
- **1000 suscriptores + 10 millones de vistas de Shorts** válidas en 90 días.

Con Shorts, la vía rápida suele ser la segunda. Publicando 1-2 Shorts por día y
pegando algunos virales, es alcanzable en 3-6 meses. No antes.

## De dónde sale el dinero

1. **Ad revenue de Shorts** (el pozo de anuncios se reparte por vistas). El RPM
   de Shorts es bajo (~$0.05-0.15 por 1000 vistas), así que el volumen manda.
2. **Videos largos**: RPM mucho más alto ($2-8 / 1000 vistas). Estrategia común:
   crecer con Shorts, derivar a largos. El sistema ya soporta cambiar el formato.
3. **Afiliados / productos propios / patrocinios**: lo más rentable una vez que
   hay audiencia. Un link de afiliado en la descripción escala mejor que el RPM.

## Plan de 90 días

| Semana | Objetivo |
|--------|----------|
| 1-2 | Definir nicho y voz. Generar 14 Shorts con `lote`. Publicar 1/día. Revisar retención en YouTube Studio. |
| 3-4 | Duplicar los formatos que retienen. Ajustar ganchos (primeros 3 segundos). |
| 5-8 | 2 Shorts/día. Empezar a probar miniaturas y títulos. Analizar los 3 mejores y clonar su estructura. |
| 9-12 | Introducir 1 video largo/semana con los temas que funcionaron. Sumar link de afiliado. |

## Palancas de calidad (donde este sistema te ayuda)

- **Gancho fuerte**: el guion arranca con una pregunta/afirmación de curiosidad.
  Editá el prompt en `src/script.py` si querés otro estilo.
- **Subtítulos animados**: suben la retención en móvil (ya vienen sincronizados).
- **Consistencia visual**: paleta y voz fijas → identidad de canal reconocible.
- **SEO**: título con gancho + tags relevantes (los genera `metadata.py`).

## Errores que matan canales de IA

1. **Datos falsos**: revisá cada dato antes de publicar. Un error viral te quema.
2. **Contenido re-subido/robado**: YouTube penaliza el contenido no original y
   sin valor agregado. La voz + guion + edición propios evitan esto.
3. **Publicar y desaparecer**: el algoritmo premia la constancia. Usá `lote` +
   un programador de tareas para no depender de tu ánimo.
4. **Sonar 100% robot**: probá varias voces (`python main.py voces`) y elegí la
   más natural para tu idioma/país.

## Costo operativo aproximado

- Voz (edge-tts): **gratis**.
- Fondos (gradiente): **gratis**. Pexels: **gratis** con API key.
- IA (Claude) para guion + metadata: unos centavos por video. Con
  `claude-haiku-4-5` baja aún más. 30 videos/mes ≈ pocos dólares.
- YouTube: gratis.

En resumen: podés operar el canal por **casi $0** hasta que empiece a generar.
