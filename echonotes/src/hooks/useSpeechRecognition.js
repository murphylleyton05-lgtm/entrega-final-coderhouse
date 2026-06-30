import { useState, useRef, useEffect, useCallback } from 'react'

// Hook que envuelve la Web Speech API (SpeechRecognition).
// Devuelve el transcript final + el parcial (interim) en tiempo real,
// el estado de escucha y controles para iniciar/detener.
export function useSpeechRecognition({ lang = 'es-AR' } = {}) {
  const SpeechRecognition =
    typeof window !== 'undefined' &&
    (window.SpeechRecognition || window.webkitSpeechRecognition)

  const supported = Boolean(SpeechRecognition)

  const [listening, setListening] = useState(false)
  const [finalTranscript, setFinalTranscript] = useState('')
  const [interimTranscript, setInterimTranscript] = useState('')
  const [error, setError] = useState(null)

  const recognitionRef = useRef(null)
  // Bandera para reiniciar automaticamente mientras el usuario quiera seguir grabando
  const wantListeningRef = useRef(false)

  useEffect(() => {
    if (!supported) return

    const recognition = new SpeechRecognition()
    recognition.continuous = true
    recognition.interimResults = true
    recognition.lang = lang

    recognition.onresult = (event) => {
      let interim = ''
      let finalChunk = ''
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const transcript = event.results[i][0].transcript
        if (event.results[i].isFinal) {
          finalChunk += transcript
        } else {
          interim += transcript
        }
      }
      if (finalChunk) {
        setFinalTranscript((prev) => (prev + ' ' + finalChunk).trim())
      }
      setInterimTranscript(interim)
    }

    recognition.onerror = (event) => {
      // "no-speech" y "aborted" son comunes y no son fatales
      if (event.error === 'not-allowed' || event.error === 'service-not-allowed') {
        setError('Permiso de microfono denegado. Habilitalo en el navegador.')
        wantListeningRef.current = false
        setListening(false)
      } else if (event.error !== 'no-speech' && event.error !== 'aborted') {
        setError('Error de reconocimiento: ' + event.error)
      }
    }

    recognition.onend = () => {
      // El navegador corta la sesion sola cada cierto tiempo: reiniciamos si el usuario sigue grabando
      if (wantListeningRef.current) {
        try {
          recognition.start()
        } catch {
          // ya esta corriendo, ignorar
        }
      } else {
        setListening(false)
        setInterimTranscript('')
      }
    }

    recognitionRef.current = recognition

    return () => {
      wantListeningRef.current = false
      try {
        recognition.stop()
      } catch {
        // ignorar
      }
    }
  }, [supported, lang, SpeechRecognition])

  const start = useCallback(() => {
    if (!supported || !recognitionRef.current) return
    setError(null)
    wantListeningRef.current = true
    try {
      recognitionRef.current.start()
      setListening(true)
    } catch {
      // si ya estaba corriendo no hacemos nada
    }
  }, [supported])

  const stop = useCallback(() => {
    wantListeningRef.current = false
    if (recognitionRef.current) {
      try {
        recognitionRef.current.stop()
      } catch {
        // ignorar
      }
    }
    setListening(false)
  }, [])

  const reset = useCallback(() => {
    setFinalTranscript('')
    setInterimTranscript('')
  }, [])

  return {
    supported,
    listening,
    finalTranscript,
    interimTranscript,
    error,
    start,
    stop,
    reset,
    setFinalTranscript,
  }
}
