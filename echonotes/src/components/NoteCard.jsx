import { useState } from 'react'

function formatDate(ts) {
  const d = new Date(ts)
  return d.toLocaleString('es-AR', {
    day: '2-digit',
    month: 'short',
    hour: '2-digit',
    minute: '2-digit',
  })
}

export default function NoteCard({ note, onDelete, onEdit }) {
  const [editing, setEditing] = useState(false)
  const [text, setText] = useState(note.text)
  const [copied, setCopied] = useState(false)

  const words = note.text.split(/\s+/).filter(Boolean).length

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(note.text)
      setCopied(true)
      setTimeout(() => setCopied(false), 1500)
    } catch {
      // clipboard no disponible
    }
  }

  const handleSaveEdit = () => {
    const trimmed = text.trim()
    if (trimmed) onEdit(note.id, trimmed)
    setEditing(false)
  }

  return (
    <article className="note-card">
      <header className="note-card__head">
        <time className="note-card__date">{formatDate(note.createdAt)}</time>
        <span className="note-card__words">{words} pal.</span>
      </header>

      {editing ? (
        <textarea
          className="note-card__edit"
          value={text}
          onChange={(e) => setText(e.target.value)}
          rows={4}
          autoFocus
        />
      ) : (
        <p className="note-card__text">{note.text}</p>
      )}

      <footer className="note-card__actions">
        {editing ? (
          <>
            <button className="icon-btn" onClick={handleSaveEdit} title="Guardar">Guardar</button>
            <button className="icon-btn" onClick={() => { setEditing(false); setText(note.text) }} title="Cancelar">Cancelar</button>
          </>
        ) : (
          <>
            <button className="icon-btn" onClick={handleCopy} title="Copiar">
              {copied ? '✓ Copiado' : 'Copiar'}
            </button>
            <button className="icon-btn" onClick={() => setEditing(true)} title="Editar">Editar</button>
            <button className="icon-btn icon-btn--danger" onClick={() => onDelete(note.id)} title="Borrar">Borrar</button>
          </>
        )}
      </footer>
    </article>
  )
}
