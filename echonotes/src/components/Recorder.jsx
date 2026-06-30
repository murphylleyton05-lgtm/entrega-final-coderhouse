import { useState, useEffect } from 'react'
import { useSpeechRecognition } from '../hooks/useSpeechRecognition.js'

const LANGS = [
  { code: 'es-AR', label: 'Espanol (AR)' },
  { code: 'es-ES', label: 'Espanol (ES)' },
  { code: 'es-MX', label: 'Espanol (MX)' },
  { code: 'en-US', label: 'English (US)' },
  { code: 'pt-BR', label: 'Portugues (BR)' },
]

export default function Recorder({ onSave }) {
  const [lang, setLang] = useState('es-AR')
  const {
    supported,
    listening,
    finalTranscript,
    interimTranscript,
    error,
    start,
    stop,
    reset,
    setFinalTranscript,
  } = useSpeechRecognition({ lang })

  const [draft, setDraft] = useState('')

  // Mantenemos el textarea sincronizado con lo que se va transcribiendo
  useEffect(() => {
    setDraft(finalTranscript)
  }, [finalTranscript])

  const liveText = (draft + (interimTranscript ? ' ' + interimTranscript : '')).trim()
  const wordCount = liveText ? liveText.split(/\s+/).filter(Boolean).length : 0

  const handleSave = () => {
    const text = draft.trim()
    if (!text) return
    onSave(text)
    setDraft('')
    reset()
    if (listening) stop()
  }

  const handleClear = () => {
    setDraft('')
    reset()
  }

  return (
    <section className="recorder">
      <div className="recorder__toolbar">
        <label className="lang-select">
          <span>Idioma</span>
          <select value={lang} onChange={(e) => setLang(e.target.value)} disabled={listening}>
            {LANGS.map((l) => (
              <option key={l.code} value={l.code}>{l.label}</option>
            ))}
          </select>
        </label>
        <span className="word-count">{wordCount} palabras</span>
      </div>

      <div className="mic-wrap">
        <button
          className={`mic-btn ${listening ? 'mic-btn--active' : ''}`}
          onClick={listening ? stop : start}
          disabled={!supported}
          aria-label={listening ? 'Detener grabacion' : 'Iniciar grabacion'}
          title={supported ? '' : 'Tu navegador no soporta reconocimiento de voz'}
        >
          <MicIcon />
          {listening && <span className="pulse-ring" />}
        </button>
        <p className="mic-status">
          {!supported
            ? 'Tu navegador no soporta voz — escribi la nota a mano'
            : listening
              ? 'Escuchando... habla con naturalidad'
              : 'Toca el microfono y empeza a hablar'}
        </p>
      </div>

      {error && <p className="error-msg">{error}</p>}

      <div className="transcript">
        <textarea
          value={draft}
          onChange={(e) => {
            setDraft(e.target.value)
            setFinalTranscript(e.target.value)
          }}
          placeholder="Aca aparece lo que dictes... tambien podes editarlo a mano."
          rows={5}
        />
        {interimTranscript && (
          <p className="interim">
            <span className="interim__dot" /> {interimTranscript}
          </p>
        )}
      </div>

      <div className="recorder__actions">
        <button className="btn btn--ghost" onClick={handleClear} disabled={!draft && !interimTranscript}>
          Limpiar
        </button>
        <button className="btn btn--primary" onClick={handleSave} disabled={!draft.trim()}>
          Guardar nota
        </button>
      </div>
    </section>
  )
}

function MicIcon() {
  return (
    <svg viewBox="0 0 24 24" width="34" height="34" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M12 2a3 3 0 0 0-3 3v7a3 3 0 0 0 6 0V5a3 3 0 0 0-3-3z" />
      <path d="M19 10v2a7 7 0 0 1-14 0v-2" />
      <line x1="12" y1="19" x2="12" y2="23" />
      <line x1="8" y1="23" x2="16" y2="23" />
    </svg>
  )
}
