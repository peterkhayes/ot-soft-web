import { useState } from 'react'
import { run_maxent, format_maxent_output } from '../../pkg/ot_soft.js'
import type { Tableau } from '../../pkg/ot_soft.js'

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

function MaxEntPanel({ tableau, tableauText, inputFilename }: MaxEntPanelProps) {
  const [iterations, setIterations] = useState(100)
  const [weightMin, setWeightMin] = useState(0)
  const [weightMax, setWeightMax] = useState(50)
  const [result, setResult] = useState<MaxEntState | null>(null)

  function handleRun() {
    try {
      const r = run_maxent(tableauText, iterations, weightMin, weightMax)
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
    }
  }

  function handleDownload() {
    try {
      let outputFilename: string
      if (inputFilename) {
        const lastDot = inputFilename.lastIndexOf('.')
        outputFilename = lastDot > 0
          ? inputFilename.substring(0, lastDot) + 'MaxEntOutput' + inputFilename.substring(lastDot)
          : inputFilename + 'MaxEntOutput.txt'
      } else {
        outputFilename = 'MaxEntOutput.txt'
      }

      const formattedOutput = format_maxent_output(
        tableauText,
        inputFilename || 'tableau.txt',
        iterations,
        weightMin,
        weightMax
      )

      const blob = new Blob([formattedOutput], { type: 'text/plain;charset=utf-8' })
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = outputFilename
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      URL.revokeObjectURL(url)
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  const constraintCount = tableau.constraint_count()
  const constraintAbbrevs = Array.from({ length: constraintCount }, (_, i) =>
    tableau.get_constraint(i)!.abbrev
  )

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Maximum Entropy</h2>
        <span className="panel-number">04</span>
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
            onChange={e => setIterations(Math.max(1, parseInt(e.target.value) || 1))}
          />
        </label>
        <label className="param-label">
          Weight min
          <input
            type="number"
            className="param-input"
            value={weightMin}
            step={0.1}
            onChange={e => setWeightMin(parseFloat(e.target.value) || 0)}
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
            onChange={e => setWeightMax(Math.max(0.1, parseFloat(e.target.value) || 50))}
          />
        </label>
      </div>

      <div className="action-bar">
        <button className="primary-button" onClick={handleRun}>
          <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <polygon points="5 3 19 12 5 21 5 3"></polygon>
          </svg>
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

      {result && (
        result.error ? (
          <div className="rcd-status failure">
            Error running MaxEnt: {result.error}
          </div>
        ) : (() => {
          const r = result as MaxEntResultState
          return (
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
                    {r.weights.map((w, i) => (
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
                  Log probability of data: {r.logProb.toFixed(4)}
                </div>
              </div>

              <div className="maxent-tableaux">
                <h3 className="results-subheader">Predicted Probabilities</h3>
                {r.forms.map((form, fi) => (
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
          )
        })()
      )}
    </section>
  )
}

export default MaxEntPanel
