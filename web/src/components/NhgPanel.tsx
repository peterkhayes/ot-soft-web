import { useCallback, useState } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import { format_nhg_output, NhgOptions, NhgRunner } from '../../pkg/ot_soft.js'
import { DEFAULT_SCHEDULE_TEMPLATE } from '../constants.ts'
import { useDownload } from '../contexts/downloadContext.ts'
import { useChunkedRunner } from '../hooks/useChunkedRunner.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import { type NhgDefaults, nhgDefaults } from '../wasmDefaults.ts'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'
import TextFileEditor from './TextFileEditor.tsx'

interface NhgPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface NhgResultState {
  weights: { fullName: string; abbrev: string; weight: number }[]
  forms: {
    input: string
    totalFreq: number
    candidates: {
      form: string
      frequency: number
      obsPct: number
      genCount: number
      genPct: number
    }[]
  }[]
  logLikelihood: number
  zeroPredictionWarning: boolean
  history?: string
  fullHistory?: string
  error?: undefined
}

interface NhgErrorState {
  error: string
  weights?: undefined
}

type NhgState = NhgResultState | NhgErrorState

type NhgParams = NhgDefaults

function NhgPanel({ tableau, tableauText, inputFilename }: NhgPanelProps) {
  const [params, setParams] = useLocalStorage<NhgParams>('otsoft:params:nhg', nhgDefaults())
  const {
    cycles,
    initialPlasticity,
    finalPlasticity,
    testTrials,
    noiseByCell,
    postMultNoise,
    noiseForZeroCells,
    lateNoise,
    exponentialNhg,
    demiGaussians,
    negativeWeightsOk,
    resolveTiesBySkipping,
    exactProportions,
    generateHistory,
    generateFullHistory,
    useCustomSchedule,
    customSchedule,
  } = params

  const [scheduleError, setScheduleError] = useState<string | null>(null)
  const download = useDownload()

  // useCallback rather than useMemo: NhgOptions is a WASM heap object with no cleanup
  // phase available in useMemo, so we build it lazily on demand to avoid leaking abandoned instances.
  const buildOpts = useCallback((): NhgOptions => {
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
    opts.exact_proportions = exactProportions
    opts.generate_history = generateHistory
    opts.generate_full_history = generateFullHistory
    if (useCustomSchedule) {
      opts.learning_schedule = customSchedule
    }
    return opts
  }, [
    cycles,
    initialPlasticity,
    finalPlasticity,
    testTrials,
    noiseByCell,
    postMultNoise,
    noiseForZeroCells,
    lateNoise,
    exponentialNhg,
    demiGaussians,
    negativeWeightsOk,
    resolveTiesBySkipping,
    exactProportions,
    generateHistory,
    generateFullHistory,
    useCustomSchedule,
    customSchedule,
  ])

  const createRunner = useCallback(() => {
    setScheduleError(null)
    return new NhgRunner(tableauText, buildOpts())
  }, [tableauText, buildOpts])

  const extractResult = useCallback(
    (runner: NhgRunner): NhgState => {
      const r = runner.take_result()
      const constraintCount = tableau.constraint_count()
      const formCount = tableau.form_count()

      const weights = Array.from({ length: constraintCount }, (_, i) => {
        const c = tableau.get_constraint(i)!
        return { fullName: c.full_name, abbrev: c.abbrev, weight: r.get_weight(i) }
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
            frequency: cand.frequency,
            obsPct: totalFreq > 0 ? (cand.frequency / totalFreq) * 100 : 0,
            genCount: r.get_test_count(formIdx, candIdx),
            genPct: r.get_test_prob(formIdx, candIdx) * 100,
          }
        })

        return { input: form.input, totalFreq, candidates }
      })

      const state: NhgResultState = {
        weights,
        forms,
        logLikelihood: r.log_likelihood(),
        zeroPredictionWarning: r.zero_prediction_warning(),
        history: r.history() ?? undefined,
        fullHistory: r.full_history() ?? undefined,
      }
      r.free()
      return state
    },
    [tableau],
  )

  const { state: runnerState, run: handleRun } = useChunkedRunner(createRunner, extractResult)

  const result: NhgState | null =
    runnerState.status === 'done'
      ? runnerState.result
      : runnerState.status === 'error'
        ? { error: runnerState.error }
        : null
  const isLoading = runnerState.status === 'running'

  // Show schedule errors inline when the runner fails with a schedule-related message
  if (runnerState.status === 'error' && runnerState.error.toLowerCase().includes('schedule')) {
    if (!scheduleError) setScheduleError(runnerState.error)
  }

  function handleDownload() {
    try {
      const opts = buildOpts()
      const formattedOutput = format_nhg_output(tableauText, inputFilename || 'tableau.txt', opts)
      download(formattedOutput, makeOutputFilename(inputFilename, 'NHGOutput'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  const atDefaults = isAtDefaults(params, nhgDefaults())
  const successResult: NhgResultState | null =
    result && !result.error ? (result as NhgResultState) : null

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
                finalPlasticity: Math.max(0.000001, parseFloat(e.target.value) || 0.002),
              })
            }
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
            onChange={(e) =>
              setParams({ testTrials: Math.max(1, parseInt(e.target.value) || 2000) })
            }
          />
        </label>
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Noise variant options</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={noiseByCell}
            onChange={(e) => setParams({ noiseByCell: e.target.checked })}
          />
          Apply noise by tableau cell, not by constraint
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={postMultNoise}
            onChange={(e) => {
              setParams(
                e.target.checked
                  ? { postMultNoise: true }
                  : { postMultNoise: false, noiseForZeroCells: false },
              )
            }}
          />
          Apply noise after multiplication of weights by violations
        </label>
        {postMultNoise && (
          <label className="nhg-checkbox nhg-checkbox-indent">
            <input
              type="checkbox"
              checked={noiseForZeroCells}
              onChange={(e) => setParams({ noiseForZeroCells: e.target.checked })}
            />
            Include noise even in cells with no violation
          </label>
        )}
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={lateNoise}
            onChange={(e) => setParams({ lateNoise: e.target.checked })}
          />
          Add noise to candidates, after harmony calculation
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={exponentialNhg}
            onChange={(e) => setParams({ exponentialNhg: e.target.checked })}
          />
          Employ Exponential NHG
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={demiGaussians}
            onChange={(e) => setParams({ demiGaussians: e.target.checked })}
          />
          Use positive demi-Gaussians
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={negativeWeightsOk}
            onChange={(e) => setParams({ negativeWeightsOk: e.target.checked })}
          />
          Allow constraint weights to go negative
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={resolveTiesBySkipping}
            onChange={(e) => setParams({ resolveTiesBySkipping: e.target.checked })}
          />
          Resolve ties by skipping trial
        </label>
      </div>

      <div className="options-two-col">
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
              testId="nhg-schedule-file-input"
            />
          )}
        </div>

        <div className="nhg-options">
          <div className="nhg-options-label">Output options</div>
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
              checked={generateFullHistory}
              onChange={(e) => setParams({ generateFullHistory: e.target.checked })}
            />
            Generate full history (with input/output annotations)
          </label>
        </div>
      </div>

      <div className="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run Noisy HG" />
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
        <button
          className="reset-button"
          onClick={() => setParams(nhgDefaults())}
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
      {result?.error && <div className="rcd-status failure">Error running NHG: {result.error}</div>}
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
            {successResult.zeroPredictionWarning && (
              <div className="warning">
                Caution: at least one candidate with positive frequency was assigned zero
                probability; since zero has no log this was approximated as .001.
              </div>
            )}
          </div>

          <div className="maxent-tableaux">
            <h3 className="results-subheader">Matchup to Input Frequencies</h3>
            {successResult.forms.map((form, fi) => (
              <div className="maxent-form" key={fi}>
                <div className="form-label">
                  /{form.input}/ <span className="form-freq">({form.totalFreq} cases)</span>
                </div>
                <table className="predictions-table">
                  <thead>
                    <tr>
                      <th></th>
                      <th className="pct-col">Count</th>
                      <th className="pct-col">Obs%</th>
                      <th className="pct-col">Gen</th>
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
                        <td className="pct-col">{cand.frequency}</td>
                        <td className="pct-col">{cand.obsPct.toFixed(1)}%</td>
                        <td className="pct-col">{cand.genCount}</td>
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
