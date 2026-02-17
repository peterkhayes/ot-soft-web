import { useRef, useCallback } from 'react'
import { parse_tableau } from '../../pkg/ot_soft.js'
import type { Tableau } from '../../pkg/ot_soft.js'
import { TINY_EXAMPLE } from '../constants.ts'

interface InputPanelProps {
  onTableauLoaded: (tableau: Tableau, text: string, filename: string) => void
  onParseError: (error: string) => void
}

function InputPanel({ onTableauLoaded, onParseError }: InputPanelProps) {
  const fileInputRef = useRef<HTMLInputElement>(null)

  const parseAndLoad = useCallback((text: string, filename: string) => {
    try {
      const tableau = parse_tableau(text)
      onTableauLoaded(tableau, text, filename)
    } catch (err) {
      console.error('Parse error:', err)
      onParseError(String(err))
    }
  }, [onTableauLoaded, onParseError])

  function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (file) {
      file.text().then(text => {
        parseAndLoad(text, file.name)
      })
    }
  }

  function handleLoadExample() {
    parseAndLoad(TINY_EXAMPLE, 'TinyIllustrativeFile.txt')
  }

  return (
    <section className="input-panel">
      <div className="panel-header">
        <h2>Tableau Input</h2>
        <span className="panel-number">01</span>
      </div>
      <p className="input-instruction">Load a tab-delimited OT tableau file or begin with an example</p>

      <div className="file-upload-area">
        <input
          type="file"
          ref={fileInputRef}
          accept=".txt"
          className="file-input-hidden"
          id="fileInput"
          onChange={handleFileChange}
        />
        <label htmlFor="fileInput" className="file-upload-label">
          <svg className="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="17 8 12 3 7 8"></polyline>
            <line x1="12" y1="3" x2="12" y2="15"></line>
          </svg>
          <span className="upload-text">Choose file</span>
          <span className="upload-hint">or drag and drop</span>
        </label>
      </div>

      <div className="divider-with-text">
        <span>or</span>
      </div>

      <button className="secondary-button" onClick={handleLoadExample}>
        Load Example Tableau
      </button>
    </section>
  )
}

export default InputPanel
