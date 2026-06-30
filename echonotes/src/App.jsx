import { useState, useEffect, useMemo } from 'react'
import Recorder from './components/Recorder.jsx'
import NoteCard from './components/NoteCard.jsx'
import Stats from './components/Stats.jsx'
import './App.css'

const STORAGE_KEY = 'echonotes.notes.v1'
const THEME_KEY = 'echonotes.theme'

function loadNotes() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    return raw ? JSON.parse(raw) : []
  } catch {
    return []
  }
}

export default function App() {
  const [notes, setNotes] = useState(loadNotes)
  const [query, setQuery] = useState('')
  const [theme, setTheme] = useState(() => localStorage.getItem(THEME_KEY) || 'dark')

  // Persistencia de notas
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(notes))
  }, [notes])

  // Tema
  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem(THEME_KEY, theme)
  }, [theme])

  const addNote = (text) => {
    const note = {
      id: crypto.randomUUID(),
      text,
      createdAt: Date.now(),
    }
    setNotes((prev) => [note, ...prev])
  }

  const deleteNote = (id) => {
    setNotes((prev) => prev.filter((n) => n.id !== id))
  }

  const editNote = (id, text) => {
    setNotes((prev) => prev.map((n) => (n.id === id ? { ...n, text } : n)))
  }

  const exportNotes = () => {
    const content = notes
      .map((n) => `[${new Date(n.createdAt).toLocaleString('es-AR')}]\n${n.text}\n`)
      .join('\n')
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `echonotes-${new Date().toISOString().slice(0, 10)}.txt`
    a.click()
    URL.revokeObjectURL(url)
  }

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase()
    if (!q) return notes
    return notes.filter((n) => n.text.toLowerCase().includes(q))
  }, [notes, query])

  return (
    <div className="app">
      <header className="app__header">
        <div className="brand">
          <span className="brand__logo">🎙️</span>
          <div>
            <h1 className="brand__name">EchoNotes</h1>
            <p className="brand__tag">Toma notas hablando</p>
          </div>
        </div>
        <button
          className="theme-toggle"
          onClick={() => setTheme((t) => (t === 'dark' ? 'light' : 'dark'))}
          aria-label="Cambiar tema"
        >
          {theme === 'dark' ? '☀️' : '🌙'}
        </button>
      </header>

      <main className="app__main">
        <Recorder onSave={addNote} />

        <section className="notes-section">
          <div className="notes-section__top">
            <Stats notes={notes} />
            <div className="notes-section__tools">
              <input
                className="search"
                type="search"
                placeholder="Buscar en mis notas..."
                value={query}
                onChange={(e) => setQuery(e.target.value)}
              />
              <button className="btn btn--ghost" onClick={exportNotes} disabled={!notes.length}>
                Exportar
              </button>
            </div>
          </div>

          {filtered.length === 0 ? (
            <div className="empty">
              {notes.length === 0
                ? 'Todavia no tenes notas. ¡Toca el microfono y dictá la primera!'
                : 'Ninguna nota coincide con tu busqueda.'}
            </div>
          ) : (
            <div className="notes-grid">
              {filtered.map((note) => (
                <NoteCard key={note.id} note={note} onDelete={deleteNote} onEdit={editNote} />
              ))}
            </div>
          )}
        </section>
      </main>

      <footer className="app__footer">
        <p>
          Hecho con React + Vite · Voz vía Web Speech API · Tus notas se guardan solo en tu navegador
        </p>
      </footer>
    </div>
  )
}
