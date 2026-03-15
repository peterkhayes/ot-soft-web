import { useCallback, useEffect, useState } from 'react'

import type { GlaResult, Tableau } from '../../pkg/ot_soft.js'
import {
  format_gla_output,
  gla_hasse_dot,
  gla_pairwise_probabilities,
  GlaMultipleRunsRunner,
  GlaOptions,
  GlaRunner,
} from '../../pkg/ot_soft.js'
import { DEFAULT_SCHEDULE_TEMPLATE } from '../constants.ts'
import { useDownload } from '../contexts/downloadContext.ts'
import { useChunkedRunner } from '../hooks/useChunkedRunner.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import { type GlaDefaults, glaDefaults } from '../wasmDefaults.ts'
import HasseDiagram from './HasseDiagram.tsx'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'
import TextFileEditor from './TextFileEditor.tsx'

interface GlaPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface GlaResultState {
  values: { fullName: string; abbrev: string; value: number; active: boolean }[]
  forms: {
    input: string
    candidates: { form: string; obsPct: number; genPct: number }[]
  }[]
  logLikelihood: number
  maxentMode: boolean
  hasseDot?: string
  pairwiseTable?: string
  history?: string
  fullHistory?: string
  candidateProbHistory?: string
  error?: undefined
}

interface GlaErrorState {
  error: string
  values?: undefined
}

type GlaState = GlaResultState | GlaErrorState

type GlaParams = GlaDefaults

function GlaPanel({ tableau, tableauText, inputFilename }: GlaPanelProps) {
  const [params, setParams] = useLocalStorage<GlaParams>('otsoft:params:gla', glaDefaults())
  const {
    maxentMode,
    cycles,
    initialPlasticity,
    finalPlasticity,
    testTrials,
    negativeWeightsOk,
    gaussianPrior,
    sigma,
    magriUpdateRule,
    exactProportions,
    showApriori,
    aprioriText,
    aprioriGap,
    generateHistory,
    generateFullHistory,
    generateCandidateProbHistory,
    useCustomSchedule,
    customSchedule,
    multipleRunsCount,
  } = params

  const [scheduleError, setScheduleError] = useState<string | null>(null)
  const download = useDownload()

  // useCallback rather than useMemo: GlaOptions is a WASM heap object with no cleanup
  // phase available in useMemo, so we build it lazily on demand to avoid leaking abandoned instances.
  const buildOpts = useCallback((): GlaOptions => {
    const opts = new GlaOptions()
    opts.maxent_mode = maxentMode
    opts.cycles = cycles
    opts.initial_plasticity = initialPlasticity
    opts.final_plasticity = finalPlasticity
    opts.test_trials = maxentMode ? 0 : testTrials
    opts.negative_weights_ok = negativeWeightsOk
    opts.gaussian_prior = maxentMode && gaussianPrior
    opts.sigma = sigma
    opts.magri_update_rule = !maxentMode && magriUpdateRule
    opts.exact_proportions = exactProportions
    if (!maxentMode && aprioriText.trim()) {
      opts.apriori_text = aprioriText
      opts.apriori_gap = aprioriGap
    }
    opts.generate_history = generateHistory
    opts.generate_full_history = generateFullHistory
    opts.generate_candidate_prob_history = maxentMode && generateCandidateProbHistory
    if (useCustomSchedule) {
      opts.learning_schedule = customSchedule
    }
    return opts
  }, [
    maxentMode,
    cycles,
    initialPlasticity,
    finalPlasticity,
    testTrials,
    negativeWeightsOk,
    gaussianPrior,
    sigma,
    magriUpdateRule,
    exactProportions,
    aprioriText,
    aprioriGap,
    generateHistory,
    generateFullHistory,
    generateCandidateProbHistory,
    useCustomSchedule,
    customSchedule,
  ])

  const createRunner = useCallback(() => {
    setScheduleError(null)
    return new GlaRunner(tableauText, buildOpts())
  }, [tableauText, buildOpts])

  const extractResult = useCallback(
    (runner: GlaRunner): GlaState => {
      const r: GlaResult = runner.take_result()

      const constraintCount = tableau.constraint_count()
      const formCount = tableau.form_count()

      const values = Array.from({ length: constraintCount }, (_, i) => {
        const c = tableau.get_constraint(i)!
        return { fullName: c.full_name, abbrev: c.abbrev, value: r.get_ranking_value(i), active: r.get_active_constraint(i) }
      })

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
            genPct: r.get_test_prob(formIdx, candIdx) * 100,
          }
        })

        return { input: form.input, candidates }
      })

      let hasseDot: string | undefined
      let pairwiseTable: string | undefined
      if (!maxentMode) {
        const rankingValues = new Float64Array(values.map((v) => v.value))
        try {
          hasseDot = gla_hasse_dot(tableauText, rankingValues)
        } catch (e) {
          console.warn('GLA Hasse diagram generation failed:', e)
        }
        try {
          pairwiseTable = gla_pairwise_probabilities(tableauText, rankingValues)
        } catch (e) {
          console.warn('GLA pairwise probabilities generation failed:', e)
        }
      }

      const extracted: GlaResultState = {
        values,
        forms,
        logLikelihood: r.log_likelihood(),
        maxentMode: r.is_maxent_mode(),
        hasseDot,
        pairwiseTable,
        history: r.history() ?? undefined,
        fullHistory: r.full_history() ?? undefined,
        candidateProbHistory: r.candidate_prob_history() ?? undefined,
      }
      r.free()
      return extracted
    },
    [tableau, tableauText, maxentMode],
  )

  const { state: runnerState, run: handleRun } = useChunkedRunner(createRunner, extractResult)

  const createMultipleRunsRunner = useCallback(
    () => new GlaMultipleRunsRunner(tableauText, multipleRunsCount, buildOpts()),
    [tableauText, multipleRunsCount, buildOpts],
  )
  const extractMultipleRunsResult = useCallback((runner: GlaMultipleRunsRunner) => runner.take_result(), [])
  const {
    state: multipleRunsState,
    run: handleMultipleRuns,
    reset: resetMultipleRuns,
  } = useChunkedRunner(createMultipleRunsRunner, extractMultipleRunsResult, 1)

  // Download and reset when multiple runs completes
  useEffect(() => {
    if (multipleRunsState.status === 'done') {
      download(multipleRunsState.result, makeOutputFilename(inputFilename, 'CollateRuns'))
      resetMultipleRuns()
    }
  }, [multipleRunsState, download, inputFilename, resetMultipleRuns])

  const result: GlaState | null =
    runnerState.status === 'done'
      ? runnerState.result
      : runnerState.status === 'error'
        ? { error: runnerState.error }
        : null
  const isLoading = runnerState.status === 'running'
  const isLoadingMultiple = multipleRunsState.status === 'running'

  function handleDownload() {
    try {
      const opts = buildOpts()
      const output = format_gla_output(tableauText, inputFilename || 'tableau.txt', opts)
      const suffix = maxentMode ? 'GLA-MaxEntOutput' : 'GLA-StochasticOTOutput'
      download(output, makeOutputFilename(inputFilename, suffix))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  const atDefaults = isAtDefaults(params, glaDefaults())
  const successResult: GlaResultState | null =
    result && !result.error ? (result as GlaResultState) : null
  const valueLabel = successResult?.maxentMode ? 'Weight' : 'Ranking Value'

  function handleDownloadHistory() {
    if (successResult?.history) {
      download(successResult.history, makeOutputFilename(inputFilename, 'History'))
    }
  }

  function handleDownloadFullHistory() {
    if (successResult?.fullHistory) {
      download(successResult.fullHistory, makeOutputFilename(inputFilename, 'FullHistory'))
    }
  }

  function handleDownloadCandidateProbHistory() {
    if (successResult?.candidateProbHistory) {
      download(
        successResult.candidateProbHistory,
        makeOutputFilename(inputFilename, 'HistoryOfCandidateProbabilities'),
      )
    }
  }

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Gradual Learning Algorithm</h2>
        <span className="panel-number">04</span>
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Framework</div>
        <label className="nhg-checkbox">
          <input
            type="radio"
            name="gla-mode"
            checked={!maxentMode}
            onChange={() => setParams({ maxentMode: false })}
          />
          Stochastic OT (ranking values)
        </label>
        {!maxentMode && (
          <label className="nhg-checkbox nhg-checkbox-indent">
            <input
              type="checkbox"
              checked={magriUpdateRule}
              onChange={(e) => setParams({ magriUpdateRule: e.target.checked })}
            />
            Use the Magri update rule
          </label>
        )}
        <label className="nhg-checkbox">
          <input
            type="radio"
            name="gla-mode"
            checked={maxentMode}
            onChange={() => setParams({ maxentMode: true })}
          />
          Online MaxEnt (weights)
        </label>
        {maxentMode && (
          <>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={negativeWeightsOk}
                onChange={(e) => setParams({ negativeWeightsOk: e.target.checked })}
              />
              Allow constraint weights to go negative
            </label>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={gaussianPrior}
                onChange={(e) => setParams({ gaussianPrior: e.target.checked })}
              />
              Gaussian prior (L2 regularization)
            </label>
            {gaussianPrior && (
              <label className="param-label" style={{ marginLeft: '3rem' }}>
                σ
                <input
                  type="number"
                  className="param-input"
                  value={sigma}
                  min={0.0001}
                  step={0.1}
                  onChange={(e) =>
                    setParams({ sigma: Math.max(0.0001, parseFloat(e.target.value) || 1) })
                  }
                />
              </label>
            )}
          </>
        )}
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
            onChange={(e) => setParams({ cycles: Math.max(1, parseInt(e.target.value) || 1) })}
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
            onChange={(e) =>
              setParams({ initialPlasticity: Math.max(0.0001, parseFloat(e.target.value) || 2) })
            }
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
            onChange={(e) =>
              setParams({
                finalPlasticity: Math.max(0.000001, parseFloat(e.target.value) || 0.001),
              })
            }
          />
        </label>
        {!maxentMode && (
          <label className="param-label">
            Test trials
            <input
              type="number"
              className="param-input"
              value={testTrials}
              min={1}
              max={100000}
              onChange={(e) =>
                setParams({ testTrials: Math.max(1, parseInt(e.target.value) || 10000) })
              }
            />
          </label>
        )}
      </div>

      <div className="options-three-col">
        <div className="nhg-options">
          <div className="nhg-options-label">Learning schedule</div>
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={exactProportions}
              onChange={(e) =>
                setParams(
                  e.target.checked
                    ? { exactProportions: true, useCustomSchedule: false }
                    : { exactProportions: false },
                )
              }
            />
            Present data in exact proportions
          </label>
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={useCustomSchedule}
              onChange={(e) =>
                setParams(
                  e.target.checked
                    ? { useCustomSchedule: true, exactProportions: false }
                    : { useCustomSchedule: false },
                )
              }
            />
            Use custom learning schedule
          </label>
          {useCustomSchedule && (
            <TextFileEditor
              value={customSchedule}
              onChange={(text) => {
                setParams({ customSchedule: text })
                setScheduleError(null)
              }}
              defaultValue={DEFAULT_SCHEDULE_TEMPLATE}
              hint="Columns: Trials, PlastMark, PlastFaith, NoiseMark, NoiseFaith (tab or space separated)"
              error={scheduleError}
              testId="gla-schedule-file-input"
            />
          )}
        </div>

        <div className="nhg-options">
          <div className="nhg-options-label">Multiple runs</div>
          <input
            type="number"
            className="param-input"
            aria-label="Number of runs"
            value={multipleRunsCount}
            min={1}
            onChange={(e) =>
              setParams({ multipleRunsCount: Math.max(1, parseInt(e.target.value) || 1) })
            }
          />
        </div>

        <div className="nhg-options">
          <div className="nhg-options-label">Output options</div>
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={generateHistory}
              onChange={(e) => setParams({ generateHistory: e.target.checked })}
            />
            Generate history of {maxentMode ? 'weights' : 'ranking values'}
          </label>
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={generateFullHistory}
              onChange={(e) => setParams({ generateFullHistory: e.target.checked })}
            />
            Generate full history (with input/output annotations)
          </label>
          {maxentMode && (
            <label className="nhg-checkbox">
              <input
                type="checkbox"
                checked={generateCandidateProbHistory}
                onChange={(e) => setParams({ generateCandidateProbHistory: e.target.checked })}
              />
              Generate history of candidate probabilities
            </label>
          )}
        </div>
      </div>

      {!maxentMode && (
        <div className="nhg-options">
          <div className="nhg-options-label">A priori rankings</div>
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={showApriori}
              onChange={(e) => {
                setParams({ showApriori: e.target.checked })
                if (!e.target.checked) setParams({ aprioriText: '' })
              }}
            />
            Use a priori rankings
          </label>
          {showApriori && (
            <>
              <TextFileEditor
                value={aprioriText}
                onChange={(text) => setParams({ aprioriText: text })}
                hint="Tab-delimited constraint × constraint matrix (abbreviations must match current tableau)."
                placeholder="Load from file or paste content here…"
                testId="gla-apriori-file-input"
              />
              {aprioriText.trim() && (
                <label className="param-label" style={{ marginTop: '0.5rem' }}>
                  Constraints ranked a priori must differ by
                  <input
                    type="number"
                    className="param-input"
                    value={aprioriGap}
                    min={0.001}
                    step={1}
                    onChange={(e) =>
                      setParams({ aprioriGap: Math.max(0.001, parseFloat(e.target.value) || 20) })
                    }
                  />
                </label>
              )}
            </>
          )}
        </div>
      )}

      <div className="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run GLA" />
        {result && !result.error && (
          <button className="download-button" onClick={handleDownload}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Download Results
          </button>
        )}
        {successResult?.history && (
          <button className="download-button" onClick={handleDownloadHistory}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Download History
          </button>
        )}
        {successResult?.fullHistory && (
          <button className="download-button" onClick={handleDownloadFullHistory}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Download Full History
          </button>
        )}
        {successResult?.candidateProbHistory && (
          <button className="download-button" onClick={handleDownloadCandidateProbHistory}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Download Candidate Probability History
          </button>
        )}
        <button
          className={`download-button${isLoadingMultiple ? ' primary-button--loading' : ''}`}
          onClick={handleMultipleRuns}
          disabled={isLoadingMultiple}
        >
          <svg
            className="button-icon"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="7 10 12 15 17 10"></polyline>
            <line x1="12" y1="15" x2="12" y2="3"></line>
          </svg>
          {multipleRunsState.status === 'running'
            ? `Run ${multipleRunsState.completed}/${multipleRunsState.total}…`
            : `Run ${multipleRunsCount} times & Download`}
        </button>
        <button
          className="reset-button"
          onClick={() => setParams(glaDefaults())}
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
      <RunnerProgressBar state={multipleRunsState} />
      {result?.error && <div className="rcd-status failure">Error running GLA: {result.error}</div>}
      {successResult && (
        <div className="maxent-results">
          <div className="maxent-weights">
            <h3 className="results-subheader">Constraint {valueLabel}s</h3>
            <table className="weights-table">
              <thead>
                <tr>
                  <th>Constraint</th>
                  <th className="weight-col">{valueLabel}</th>
                </tr>
              </thead>
              <tbody>
                {successResult.values.map((v, i) => (
                  <tr key={i}>
                    <td>
                      <span className="abbrev">{v.abbrev}</span>
                      <span className="full-name"> ({v.fullName})</span>
                    </td>
                    <td className="weight-col weight-value">{v.value.toFixed(3)}</td>
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
                      <th className="pct-col">{successResult.maxentMode ? 'Prob%' : 'Gen%'}</th>
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
          {successResult.hasseDot && (
            <HasseDiagram
              dotString={successResult.hasseDot}
              downloadName={`${inputFilename ? inputFilename.replace(/\.[^.]+$/, '') : 'tableau'}Hasse`}
            />
          )}
          {successResult.pairwiseTable && (
            <div className="pairwise-probabilities">
              <h3 className="results-subheader">Pairwise Ranking Probabilities</h3>
              <pre className="pairwise-table">{successResult.pairwiseTable}</pre>
            </div>
          )}
          <div className="maxent-weights">
            <h3 className="results-subheader">Active Constraints</h3>
            <p className="active-constraints-note">
              A constraint is active if it causes the winning candidate to defeat a rival in at least
              one competition.
            </p>
            <table className="weights-table">
              <thead>
                <tr>
                  <th>Constraint</th>
                  <th className="weight-col">Status</th>
                </tr>
              </thead>
              <tbody>
                {[...successResult.values]
                  .sort((a, b) => b.value - a.value)
                  .map((v, i) => (
                    <tr key={i}>
                      <td>
                        <span className="abbrev">{v.abbrev}</span>
                        <span className="full-name"> ({v.fullName})</span>
                      </td>
                      <td className="weight-col">{v.active ? 'Active' : 'Inactive'}</td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </section>
  )
}

export default GlaPanel
