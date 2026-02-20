import { useState } from 'react'
import { run_nhg, format_nhg_output, NhgOptions } from '../../pkg/ot_soft.js'
import type { Tableau } from '../../pkg/ot_soft.js'
import { downloadTextFile, makeOutputFilename } from '../utils.ts'

interface NhgPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface NhgResultState {
  weights: { fullName: string; abbrev: string; weight: number }[]
  forms: {
    input: string
    candidates: { form: string; obsPct: number; genPct: number }[]
  }[]
  logLikelihood: number
  error?: undefined
}

interface NhgErrorState {
  error: string
  weights?: undefined
}

type NhgState = NhgResultState | NhgErrorState

function NhgPanel({ tableau, tableauText, inputFilename }: NhgPanelProps) {
  const [cycles, setCycles] = useState(5000)
  const [initialPlasticity, setInitialPlasticity] = useState(2.0)
  const [finalPlasticity, setFinalPlasticity] = useState(0.002)
  const [testTrials, setTestTrials] = useState(2000)

  // Noise variant checkboxes
  const [noiseByCell, setNoiseByCell] = useState(false)
  const [postMultNoise, setPostMultNoise] = useState(false)
  const [noiseForZeroCells, setNoiseForZeroCells] = useState(false)
  const [lateNoise, setLateNoise] = useState(false)
  const [exponentialNhg, setExponentialNhg] = useState(false)
  const [demiGaussians, setDemiGaussians] = useState(false)
  const [negativeWeightsOk, setNegativeWeightsOk] = useState(false)
  const [resolveTiesBySkipping, setResolveTiesBySkipping] = useState(false)

  const [result, setResult] = useState<NhgState | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  function handleRun() {
    setIsLoading(true)
    setTimeout(() => {
    try {
      const opts = new NhgOptions()
      opts.cycles = cycles
      opts.initial_plasticity = initialPlasticity
      opts.final_plasticity = finalPlasticity
      opts.test_trials = testTrials
      opts.noise_by_cell = noiseByCell
      opts.post_mult_noise = postMultNoise
      opts.noise_for_zero_cells = noiseForZeroCells
      opts.late_noise = lateNoise
      opts.exponential_nhg = exponentialNhg
      opts.demi_gaussians = demiGaussians
      opts.negative_weights_ok = negativeWeightsOk
      opts.resolve_ties_by_skipping = resolveTiesBySkipping
      const r = run_nhg(tableauText, opts)

      const constraintCount = tableau.constraint_count()
      const formCount = tableau.form_count()

      const weights = Array.from({ length: constraintCount }, (_, i) => {
        const c = tableau.get_constraint(i)!
        return { fullName: c.full_name, abbrev: c.abbrev, weight: r.get_weight(i) }
      })

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
            genPct: r.get_test_prob(formIdx, candIdx) * 100,
          }
        })

        return { input: form.input, candidates }
      })

      setResult({ weights, forms, logLikelihood: r.log_likelihood() })
    } catch (err) {
      console.error('NHG error:', err)
      setResult({ error: String(err) })
    } finally {
      setIsLoading(false)
    }
    }, 0)
  }

  function handleDownload() {
    try {
      const opts = new NhgOptions()
      opts.cycles = cycles
      opts.initial_plasticity = initialPlasticity
      opts.final_plasticity = finalPlasticity
      opts.test_trials = testTrials
      opts.noise_by_cell = noiseByCell
      opts.post_mult_noise = postMultNoise
      opts.noise_for_zero_cells = noiseForZeroCells
      opts.late_noise = lateNoise
      opts.exponential_nhg = exponentialNhg
      opts.demi_gaussians = demiGaussians
      opts.negative_weights_ok = negativeWeightsOk
      opts.resolve_ties_by_skipping = resolveTiesBySkipping
      const formattedOutput = format_nhg_output(tableauText, inputFilename || 'tableau.txt', opts)
      downloadTextFile(formattedOutput, makeOutputFilename(inputFilename, 'NHGOutput'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  const successResult: NhgResultState | null = result && !result.error ? result as NhgResultState : null

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Noisy Harmonic Grammar</h2>
        <span className="panel-number">04</span>
      </div>

      <div className="maxent-params">
        <label className="param-label">
          Cycles
          <input
            type="number"
            className="param-input"
            value={cycles}
            min={1}
            max={10000000}
            onChange={e => setCycles(Math.max(1, parseInt(e.target.value) || 1))}
          />
        </label>
        <label className="param-label">
          Initial plasticity
          <input
            type="number"
            className="param-input"
            value={initialPlasticity}
            min={0.0001}
            step={0.1}
            onChange={e => setInitialPlasticity(Math.max(0.0001, parseFloat(e.target.value) || 2))}
          />
        </label>
        <label className="param-label">
          Final plasticity
          <input
            type="number"
            className="param-input"
            value={finalPlasticity}
            min={0.000001}
            step={0.001}
            onChange={e => setFinalPlasticity(Math.max(0.000001, parseFloat(e.target.value) || 0.002))}
          />
        </label>
        <label className="param-label">
          Test trials
          <input
            type="number"
            className="param-input"
            value={testTrials}
            min={1}
            max={100000}
            onChange={e => setTestTrials(Math.max(1, parseInt(e.target.value) || 2000))}
          />
        </label>
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Noise variant options:</div>
        <label className="nhg-checkbox">
          <input type="checkbox" checked={noiseByCell} onChange={e => setNoiseByCell(e.target.checked)} />
          Apply noise by tableau cell, not by constraint
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={postMultNoise}
            onChange={e => {
              setPostMultNoise(e.target.checked)
              if (!e.target.checked) setNoiseForZeroCells(false)
            }}
          />
          Apply noise after multiplication of weights by violations
        </label>
        {postMultNoise && (
          <label className="nhg-checkbox nhg-checkbox-indent">
            <input type="checkbox" checked={noiseForZeroCells} onChange={e => setNoiseForZeroCells(e.target.checked)} />
            Include noise even in cells with no violation
          </label>
        )}
        <label className="nhg-checkbox">
          <input type="checkbox" checked={lateNoise} onChange={e => setLateNoise(e.target.checked)} />
          Add noise to candidates, after harmony calculation
        </label>
        <label className="nhg-checkbox">
          <input type="checkbox" checked={exponentialNhg} onChange={e => setExponentialNhg(e.target.checked)} />
          Employ Exponential NHG
        </label>
        <label className="nhg-checkbox">
          <input type="checkbox" checked={demiGaussians} onChange={e => setDemiGaussians(e.target.checked)} />
          Use positive demi-Gaussians
        </label>
        <label className="nhg-checkbox">
          <input type="checkbox" checked={negativeWeightsOk} onChange={e => setNegativeWeightsOk(e.target.checked)} />
          Allow constraint weights to go negative
        </label>
        <label className="nhg-checkbox">
          <input type="checkbox" checked={resolveTiesBySkipping} onChange={e => setResolveTiesBySkipping(e.target.checked)} />
          Resolve ties by skipping trial
        </label>
      </div>

      <div className="action-bar">
        <button className="primary-button" onClick={handleRun} disabled={isLoading}>
          <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <polygon points="5 3 19 12 5 21 5 3"></polygon>
          </svg>
          Run Noisy HG
        </button>
        {isLoading && (
          <div className="loading-indicator" title="Running NHG..."></div>
        )}
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
          Error running NHG: {result.error}
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
              Log likelihood of data: {successResult.logLikelihood.toFixed(4)}
            </div>
          </div>

          <div className="maxent-tableaux">
            <h3 className="results-subheader">Matchup to Input Frequencies</h3>
            {successResult.forms.map((form, fi) => (
              <div className="maxent-form" key={fi}>
                <div className="form-label">/{form.input}/</div>
                <table className="predictions-table">
                  <thead>
                    <tr>
                      <th></th>
                      <th className="pct-col">Obs%</th>
                      <th className="pct-col">Gen%</th>
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
                        <td className="pct-col">{cand.genPct.toFixed(1)}%</td>
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

export default NhgPanel
