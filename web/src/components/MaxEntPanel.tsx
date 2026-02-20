import { useState } from 'react'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { run_maxent, format_maxent_output, MaxEntOptions } from '../../pkg/ot_soft.js'
import type { Tableau } from '../../pkg/ot_soft.js'
import { downloadTextFile, makeOutputFilename } from '../utils.ts'

interface MaxEntPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface MaxEntResultState {
  weights: { abbrev: string; fullName: string; weight: number }[]
  forms: {
    input: string
    candidates: {
      form: string
      obsPct: number
      predPct: number
      violations: number[]
    }[]
  }[]
  logProb: number
  error?: undefined
}

interface MaxEntErrorState {
  error: string
  weights?: undefined
}

type MaxEntState = MaxEntResultState | MaxEntErrorState

interface MaxEntParams { iterations: number; weightMin: number; weightMax: number }
const MAXENT_DEFAULTS: MaxEntParams = { iterations: 100, weightMin: 0, weightMax: 50 }

function MaxEntPanel({ tableau, tableauText, inputFilename }: MaxEntPanelProps) {
  const [params, setParams] = useLocalStorage<MaxEntParams>('otsoft:params:maxent', MAXENT_DEFAULTS)
  const { iterations, weightMin, weightMax } = params
  const [result, setResult] = useState<MaxEntState | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  function handleRun() {
    setIsLoading(true)
    setTimeout(() => {
    try {
      const opts = new MaxEntOptions()
      opts.iterations = iterations
      opts.weight_min = weightMin
      opts.weight_max = weightMax
      const r = run_maxent(tableauText, opts)
      const constraintCount = tableau.constraint_count()
      const formCount = tableau.form_count()

      // Build weights array sorted by weight descending
      const weights = Array.from({ length: constraintCount }, (_, i) => {
        const c = tableau.get_constraint(i)!
        return { abbrev: c.abbrev, fullName: c.full_name, weight: r.get_weight(i), index: i }
      }).sort((a, b) => b.weight - a.weight)

      // Build forms with candidates
      const forms = Array.from({ length: formCount }, (_, formIdx) => {
        const form = tableau.get_form(formIdx)!
        const totalFreq = Array.from({ length: form.candidate_count() }, (_, ci) =>
          form.get_candidate(ci)!.frequency
        ).reduce((a, b) => a + b, 0)

        const candidates = Array.from({ length: form.candidate_count() }, (_, candIdx) => {
          const cand = form.get_candidate(candIdx)!
          return {
            form: cand.form,
            obsPct: totalFreq > 0 ? (cand.frequency / totalFreq) * 100 : 0,
            predPct: r.get_predicted_prob(formIdx, candIdx) * 100,
            violations: Array.from({ length: constraintCount }, (_, ci) =>
              cand.get_violation(ci) ?? 0
            ),
          }
        })

        return { input: form.input, candidates }
      })

      setResult({ weights, forms, logProb: r.log_prob() })
    } catch (err) {
      console.error('MaxEnt error:', err)
      setResult({ error: String(err) })
    } finally {
      setIsLoading(false)
    }
    }, 0)
  }

  function handleDownload() {
    try {
      const opts = new MaxEntOptions()
      opts.iterations = iterations
      opts.weight_min = weightMin
      opts.weight_max = weightMax
      const formattedOutput = format_maxent_output(tableauText, inputFilename || 'tableau.txt', opts)
      downloadTextFile(formattedOutput, makeOutputFilename(inputFilename, 'MaxEntOutput'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  const constraintCount = tableau.constraint_count()
  const constraintAbbrevs = Array.from({ length: constraintCount }, (_, i) =>
    tableau.get_constraint(i)!.abbrev
  )
  const successResult: MaxEntResultState | null = result && !result.error ? result as MaxEntResultState : null

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Maximum Entropy</h2>
        <span className="panel-number">04</span>
      </div>

      <div className="nhg-options-header" style={{ marginBottom: 'var(--space-xs)' }}>
        <button className="reset-button" style={{ marginLeft: 'auto' }} onClick={() => setParams(MAXENT_DEFAULTS)}>Reset to defaults</button>
      </div>
      <div className="maxent-params">
        <label className="param-label">
          Iterations
          <input
            type="number"
            className="param-input"
            value={iterations}
            min={1}
            max={100000}
            onChange={e => setParams({ iterations: Math.max(1, parseInt(e.target.value) || 1) })}
          />
        </label>
        <label className="param-label">
          Weight min
          <input
            type="number"
            className="param-input"
            value={weightMin}
            step={0.1}
            onChange={e => setParams({ weightMin: parseFloat(e.target.value) || 0 })}
          />
        </label>
        <label className="param-label">
          Weight max
          <input
            type="number"
            className="param-input"
            value={weightMax}
            min={0.1}
            step={1}
            onChange={e => setParams({ weightMax: Math.max(0.1, parseFloat(e.target.value) || 50) })}
          />
        </label>
      </div>

      <div className="action-bar">
        <button className={`primary-button${isLoading ? ' primary-button--loading' : ''}`} onClick={handleRun} disabled={isLoading}>
          {isLoading ? (
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M5 22h14"/><path d="M5 2h14"/>
              <path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22"/>
              <path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2"/>
            </svg>
          ) : (
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polygon points="5 3 19 12 5 21 5 3"></polygon>
            </svg>
          )}
          Run MaxEnt
        </button>
        {result && !result.error && (
          <button className="download-button" onClick={handleDownload}>
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Download Results
          </button>
        )}
      </div>

      {result?.error && (
        <div className="rcd-status failure">
          Error running MaxEnt: {result.error}
        </div>
      )}
      {successResult && (
        <div className="maxent-results">
          <div className="maxent-weights">
            <h3 className="results-subheader">Constraint Weights</h3>
            <table className="weights-table">
              <thead>
                <tr>
                  <th>Constraint</th>
                  <th className="weight-col">Weight</th>
                </tr>
              </thead>
              <tbody>
                {successResult.weights.map((w, i) => (
                  <tr key={i}>
                    <td>
                      <span className="abbrev">{w.abbrev}</span>
                      <span className="full-name"> ({w.fullName})</span>
                    </td>
                    <td className="weight-col weight-value">{w.weight.toFixed(3)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="log-prob">
              Log probability of data: {successResult.logProb.toFixed(4)}
            </div>
          </div>

          <div className="maxent-tableaux">
            <h3 className="results-subheader">Predicted Probabilities</h3>
            {successResult.forms.map((form, fi) => (
              <div className="maxent-form" key={fi}>
                <div className="form-label">/{form.input}/</div>
                <table className="predictions-table">
                  <thead>
                    <tr>
                      <th></th>
                      <th className="pct-col">Obs%</th>
                      <th className="pct-col">Pred%</th>
                      {constraintAbbrevs.map((abbrev, ci) => (
                        <th key={ci} className="viol-col">{abbrev}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {form.candidates.map((cand, ci) => (
                      <tr key={ci} className={cand.obsPct > 0 ? 'winner-row' : ''}>
                        <td className="cand-form">
                          {cand.obsPct > 0 && <span className="winner-marker">▶</span>}
                          {cand.form}
                        </td>
                        <td className="pct-col">{cand.obsPct.toFixed(1)}%</td>
                        <td className="pct-col">{cand.predPct.toFixed(1)}%</td>
                        {cand.violations.map((v, vi) => (
                          <td key={vi} className="viol-col">{v > 0 ? v : ''}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ))}
          </div>
        </div>
      )}
    </section>
  )
}

export default MaxEntPanel
