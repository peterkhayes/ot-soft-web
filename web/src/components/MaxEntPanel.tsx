import { useCallback } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import { format_maxent_output, MaxEntOptions, MaxEntRunner } from '../../pkg/ot_soft.js'
import { useDownload } from '../contexts/downloadContext.ts'
import { useChunkedRunner } from '../hooks/useChunkedRunner.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import { type MaxEntDefaults, maxentDefaults } from '../wasmDefaults.ts'
import DownloadButton from './DownloadButton.tsx'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'

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
  history?: string
  outputProbHistory?: string
  error?: undefined
}

interface MaxEntErrorState {
  error: string
  weights?: undefined
}

type MaxEntState = MaxEntResultState | MaxEntErrorState

type MaxEntParams = MaxEntDefaults

function MaxEntPanel({ tableau, tableauText, inputFilename }: MaxEntPanelProps) {
  const [params, setParams] = useLocalStorage<MaxEntParams>(
    'otsoft:params:maxent',
    maxentDefaults(),
  )
  const {
    iterations,
    weightMin,
    weightMax,
    usePrior,
    sigmaSquared,
    generateHistory,
    generateOutputProbHistory,
    sortByWeight,
  } = params
  const download = useDownload()

  const createRunner = useCallback(() => {
    const opts = new MaxEntOptions()
    opts.iterations = iterations
    opts.weight_min = weightMin
    opts.weight_max = weightMax
    opts.use_prior = usePrior
    opts.sigma_squared = sigmaSquared
    opts.generate_history = generateHistory
    opts.generate_output_prob_history = generateOutputProbHistory
    return new MaxEntRunner(tableauText, opts)
  }, [
    tableauText,
    iterations,
    weightMin,
    weightMax,
    usePrior,
    sigmaSquared,
    generateHistory,
    generateOutputProbHistory,
  ])

  const extractResult = useCallback(
    (runner: MaxEntRunner): MaxEntState => {
      const r = runner.take_result()
      const constraintCount = tableau.constraint_count()
      const formCount = tableau.form_count()

      const weights = Array.from({ length: constraintCount }, (_, i) => {
        const c = tableau.get_constraint(i)!
        return { abbrev: c.abbrev, fullName: c.full_name, weight: r.get_weight(i), index: i }
      })
      if (sortByWeight) {
        weights.sort((a, b) => b.weight - a.weight)
      }

      const forms = Array.from({ length: formCount }, (_, formIdx) => {
        const form = tableau.get_form(formIdx)!
        const totalFreq = Array.from(
          { length: form.candidate_count() },
          (_, ci) => form.get_candidate(ci)!.frequency,
        ).reduce((a, b) => a + b, 0)

        const candidates = Array.from({ length: form.candidate_count() }, (_, candIdx) => {
          const cand = form.get_candidate(candIdx)!
          return {
            form: cand.form,
            obsPct: totalFreq > 0 ? (cand.frequency / totalFreq) * 100 : 0,
            predPct: r.get_predicted_prob(formIdx, candIdx) * 100,
            violations: Array.from(
              { length: constraintCount },
              (_, ci) => cand.get_violation(ci) ?? 0,
            ),
          }
        })

        return { input: form.input, candidates }
      })

      const state: MaxEntResultState = {
        weights,
        forms,
        logProb: r.log_prob(),
        history: r.history() ?? undefined,
        outputProbHistory: r.output_prob_history() ?? undefined,
      }
      r.free()
      return state
    },
    [tableau, sortByWeight],
  )

  const { state: runnerState, run: handleRun } = useChunkedRunner(createRunner, extractResult)

  const result: MaxEntState | null =
    runnerState.status === 'done'
      ? runnerState.result
      : runnerState.status === 'error'
        ? { error: runnerState.error }
        : null
  const isLoading = runnerState.status === 'running'

  function handleDownload() {
    try {
      const opts = new MaxEntOptions()
      opts.iterations = iterations
      opts.weight_min = weightMin
      opts.weight_max = weightMax
      opts.use_prior = usePrior
      opts.sigma_squared = sigmaSquared
      opts.sort_by_weight = sortByWeight
      const formattedOutput = format_maxent_output(
        tableauText,
        inputFilename || 'tableau.txt',
        opts,
      )
      download(formattedOutput, makeOutputFilename(inputFilename, 'MaxEntOutput'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  const atDefaults = isAtDefaults(params, maxentDefaults())
  const constraintCount = tableau.constraint_count()
  const constraintAbbrevs = Array.from(
    { length: constraintCount },
    (_, i) => tableau.get_constraint(i)!.abbrev,
  )
  const successResult: MaxEntResultState | null =
    result && !result.error ? (result as MaxEntResultState) : null

  function handleDownloadHistory() {
    if (successResult?.history) {
      download(successResult.history, makeOutputFilename(inputFilename, 'HistoryOfWeights'))
    }
  }

  function handleDownloadOutputProbHistory() {
    if (successResult?.outputProbHistory) {
      download(
        successResult.outputProbHistory,
        makeOutputFilename(inputFilename, 'HistoryOfOutputProbabilities'),
      )
    }
  }

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
            onChange={(e) => setParams({ iterations: Math.max(1, parseInt(e.target.value) || 1) })}
          />
        </label>
        <label className="param-label">
          Weight min
          <input
            type="number"
            className="param-input"
            value={weightMin}
            step={0.1}
            onChange={(e) => setParams({ weightMin: parseFloat(e.target.value) || 0 })}
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
            onChange={(e) =>
              setParams({ weightMax: Math.max(0.1, parseFloat(e.target.value) || 50) })
            }
          />
        </label>
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Prior</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={usePrior}
            onChange={(e) => setParams({ usePrior: e.target.checked })}
          />
          Gaussian prior (L2 regularization)
        </label>
        {usePrior && (
          <label className="param-label" style={{ marginLeft: '1.5rem' }}>
            σ²
            <input
              type="number"
              className="param-input"
              value={sigmaSquared}
              min={0.0001}
              step={0.1}
              onChange={(e) =>
                setParams({ sigmaSquared: Math.max(0.0001, parseFloat(e.target.value) || 1) })
              }
            />
          </label>
        )}
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Output options</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={sortByWeight}
            onChange={(e) => setParams({ sortByWeight: e.target.checked })}
          />
          Sort constraints by weight
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={generateHistory}
            onChange={(e) => setParams({ generateHistory: e.target.checked })}
          />
          Generate history of weights
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={generateOutputProbHistory}
            onChange={(e) => setParams({ generateOutputProbHistory: e.target.checked })}
          />
          Generate history of output probabilities
        </label>
      </div>

      <div className="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run MaxEnt" />
        {result && !result.error && (
          <DownloadButton onClick={handleDownload}>Download Results</DownloadButton>
        )}
        {successResult?.history && (
          <DownloadButton onClick={handleDownloadHistory}>Download History</DownloadButton>
        )}
        {successResult?.outputProbHistory && (
          <DownloadButton onClick={handleDownloadOutputProbHistory}>
            Download Output Probability History
          </DownloadButton>
        )}
        <button
          className="reset-button"
          onClick={() => setParams(maxentDefaults())}
          disabled={atDefaults}
        >
          <svg
            className="button-icon"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <polyline points="1 4 1 10 7 10"></polyline>
            <path d="M3.51 15a9 9 0 1 0 .49-4.99"></path>
          </svg>
          Reset to Defaults
        </button>
      </div>

      <RunnerProgressBar state={runnerState} />
      {result?.error && (
        <div className="rcd-status failure">Error running MaxEnt: {result.error}</div>
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
                        <th key={ci} className="viol-col">
                          {abbrev}
                        </th>
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
                          <td key={vi} className="viol-col">
                            {v > 0 ? v : ''}
                          </td>
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
