import { useState } from 'react'
import { format_factorial_typology_output } from '../../pkg/ot_soft.js'
import type { Tableau } from '../../pkg/ot_soft.js'
import { downloadTextFile, makeOutputFilename } from '../utils.ts'

interface FactorialTypologyPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface FtResultState {
  output: string
  error?: undefined
}

interface FtErrorState {
  error: string
  output?: undefined
}

type FtState = FtResultState | FtErrorState

function FactorialTypologyPanel({ tableauText, inputFilename }: FactorialTypologyPanelProps) {
  const [result, setResult] = useState<FtState | null>(null)
  const [isLoading, setIsLoading] = useState(false)
  const [aprioriText, setAprioriText] = useState('')
  const [aprioriFilename, setAprioriFilename] = useState<string | null>(null)

  function handleAprioriFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    file.text().then(text => {
      setAprioriText(text)
      setAprioriFilename(file.name)
    }).catch(err => {
      console.error('Error reading a priori file:', err)
    })
  }

  function handleRun() {
    setIsLoading(true)
    setTimeout(() => {
      try {
        const filename = inputFilename || 'tableau.txt'
        const output = format_factorial_typology_output(tableauText, filename, aprioriText)
        setResult({ output })
      } catch (err) {
        console.error('Factorial Typology error:', err)
        setResult({ error: String(err) })
      } finally {
        setIsLoading(false)
      }
    }, 0)
  }

  function handleDownload() {
    if (!result?.output) return
    downloadTextFile(result.output, makeOutputFilename(inputFilename, 'FactorialTypology'))
  }

  const successResult = result && !result.error ? result : null

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Factorial Typology</h2>
        <span className="panel-number">05</span>
      </div>

      <div className="action-bar">
        <label className="apriori-upload" title="Optional: load an a priori rankings file">
          <span>{aprioriFilename ?? 'A priori rankings (optional)'}</span>
          <input type="file" accept=".txt" onChange={handleAprioriFile} style={{ display: 'none' }} />
        </label>

        <button
          className={`primary-button${isLoading ? ' primary-button--loading' : ''}`}
          onClick={handleRun}
          disabled={isLoading}
        >
          {isLoading ? (
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M5 22h14"/><path d="M5 2h14"/>
              <path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22"/>
              <path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2"/>
            </svg>
          ) : (
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polygon points="5 3 19 12 5 21 5 3"/>
            </svg>
          )}
          Run Factorial Typology
        </button>

        {successResult && (
          <button className="download-button" onClick={handleDownload}>
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
              <polyline points="7 10 12 15 17 10"/>
              <line x1="12" y1="15" x2="12" y2="3"/>
            </svg>
            Download Results
          </button>
        )}
      </div>

      {result?.error && (
        <div className="rcd-status failure">
          Error running Factorial Typology: {result.error}
        </div>
      )}

      {successResult && (
        <pre className="ft-output">{successResult.output}</pre>
      )}
    </section>
  )
}

export default FactorialTypologyPanel
