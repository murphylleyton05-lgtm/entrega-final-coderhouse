export default function Stats({ notes }) {
  const totalNotes = notes.length
  const totalWords = notes.reduce(
    (sum, n) => sum + n.text.split(/\s+/).filter(Boolean).length,
    0,
  )

  return (
    <div className="stats">
      <div className="stat">
        <span className="stat__value">{totalNotes}</span>
        <span className="stat__label">Notas</span>
      </div>
      <div className="stat">
        <span className="stat__value">{totalWords}</span>
        <span className="stat__label">Palabras</span>
      </div>
    </div>
  )
}
