# 🎙️ EchoNotes — Notas por voz

App web que transcribe lo que hablás **en tiempo real** y lo guarda como notas, directamente en el navegador. Sin backend, sin cuentas, sin claves de API. Construida desde cero con **React + Vite**.

## ✨ Características

- **Dictado por voz en vivo** usando la Web Speech API (`SpeechRecognition`) — ves el texto mientras hablás.
- **Multi-idioma**: Español (AR/ES/MX), Inglés y Portugués.
- **Edición manual**: si tu navegador no soporta voz, podés escribir las notas igual.
- **Persistencia local**: tus notas se guardan en `localStorage` (privadas, nunca salen de tu equipo).
- **Buscador** instantáneo entre tus notas.
- **Estadísticas**: total de notas y palabras.
- **Exportar** todas las notas a un archivo `.txt`.
- **Tema claro / oscuro** con un toque.
- **Diseño responsive** y animaciones suaves.

## 🚀 Cómo ejecutarla

```bash
cd echonotes
npm install
npm run dev      # servidor de desarrollo (http://localhost:5173)
```

Para generar la versión de producción:

```bash
npm run build    # genera la carpeta dist/
npm run preview  # sirve el build localmente
```

## 🧠 ¿Cómo funciona el dictado?

El hook `useSpeechRecognition` (en `src/hooks/`) envuelve la API nativa del navegador:

- `interimResults` muestra el texto parcial mientras hablás.
- `continuous` mantiene la escucha activa; si el navegador corta la sesión, se reinicia sola.
- Maneja permisos de micrófono y errores comunes con avisos claros.

> **Nota de compatibilidad:** el reconocimiento de voz funciona mejor en Chrome y Edge (de escritorio). En navegadores sin soporte, la app sigue siendo 100% usable escribiendo a mano.

## 🗂️ Estructura

```
echonotes/
├── index.html
├── package.json
├── vite.config.js
├── public/
│   └── mic.svg
└── src/
    ├── main.jsx
    ├── App.jsx                  # estado global, persistencia, tema, export
    ├── App.css
    ├── index.css                # variables de tema y estilos base
    ├── hooks/
    │   └── useSpeechRecognition.js
    └── components/
        ├── Recorder.jsx         # micrófono + transcripción en vivo
        ├── NoteCard.jsx         # tarjeta de nota (copiar/editar/borrar)
        └── Stats.jsx            # contadores
```

## 🛠️ Stack

React 18 · Vite 5 · Web Speech API · CSS puro (sin librerías de UI)

---

Hecho con 💜 por Lleyton Murphy
