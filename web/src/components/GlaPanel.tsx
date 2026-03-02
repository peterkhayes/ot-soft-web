import { useState } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import {
  format_gla_multiple_runs_output,
  format_gla_output,
  gla_hasse_dot,
  GlaOptions,
  run_gla,
} from '../../pkg/ot_soft.js'
import { DEFAULT_SCHEDULE_TEMPLATE } from '../constants.ts'
import { useDownload } from '../contexts/downloadContext.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import HasseDiagram from './HasseDiagram.tsx'
import TextFileEditor from './TextFileEditor.tsx'

interface GlaPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface GlaResultState {
  values: { fullName: string; abbrev: string; value: number }[]
  forms: {
    input: string
    candidates: { form: string; obsPct: number; genPct: number }[]
  }[]
  logLikelihood: number
  maxentMode: boolean
  hasseDot?: string
  error?: undefined
}

interface GlaErrorState {
  error: string
  values?: undefined
}

type GlaState = GlaResultState | GlaErrorState

interface GlaParams {
  maxentMode: boolean
  cycles: number
  initialPlasticity: number
  finalPlasticity: number
  testTrials: number
  negativeWeightsOk: boolean
  gaussianPrior: boolean
  sigma: number
  magriUpdateRule: boolean
  useCustomSchedule: boolean
  customSchedule: string
  multipleRunsCount: 10 | 100 | 1000
}
const GLA_DEFAULTS: GlaParams = {
  maxentMode: false,
  cycles: 1000000,
  initialPlasticity: 2.0,
  finalPlasticity: 0.001,
  testTrials: 10000,
  negativeWeightsOk: false,
  gaussianPrior: false,
  sigma: 1.0,
  magriUpdateRule: false,
  useCustomSchedule: false,
  customSchedule: DEFAULT_SCHEDULE_TEMPLATE,
  multipleRunsCount: 10,
}

function GlaPanel({ tableau, tableauText, inputFilename }: GlaPanelProps) {
  const [params, setParams] = useLocalStorage<GlaParams>('otsoft:params:gla', GLA_DEFAULTS)
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
    useCustomSchedule,
    customSchedule,
    multipleRunsCount,
  } = params

  const [result, setResult] = useState<GlaState | null>(null)
  const [scheduleError, setScheduleError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)
  const [isLoadingMultiple, setIsLoadingMultiple] = useState(false)
  const download = useDownload()

  function buildOpts(): GlaOptions {
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
    if (useCustomSchedule) {
      opts.learning_schedule = customSchedule
    }
    return opts
  }

  function handleRun() {
    setIsLoading(true)
    setScheduleError(null)
    setTimeout(() => {
      try {
        const opts = buildOpts()
        const r = run_gla(tableauText, opts)

        const constraintCount = tableau.constraint_count()
        const formCount = tableau.form_count()

        const values = Array.from({ length: constraintCount }, (_, i) => {
          const c = tableau.get_constraint(i)!
          return { fullName: c.full_name, abbrev: c.abbrev, value: r.get_ranking_value(i) }
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
        if (!maxentMode) {
          try {
            const rankingValues = new Float64Array(values.map((v) => v.value))
            hasseDot = gla_hasse_dot(tableauText, rankingValues)
          } catch (e) {
            console.warn('GLA Hasse diagram generation failed:', e)
          }
        }

        setResult({
          values,
          forms,
          logLikelihood: r.log_likelihood(),
          maxentMode: r.is_maxent_mode(),
          hasseDot,
        })
      } catch (err) {
        console.error('GLA error:', err)
        const msg = String(err)
        if (msg.toLowerCase().includes('learning schedule')) {
          setScheduleError(msg)
        }
        setResult({ error: msg })
      } finally {
        setIsLoading(false)
      }
    }, 0)
  }

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

  function handleMultipleRuns() {
    setIsLoadingMultiple(true)
    setScheduleError(null)
    setTimeout(() => {
      try {
        const opts = buildOpts()
        const output = format_gla_multiple_runs_output(tableauText, multipleRunsCount, opts)
        download(output, makeOutputFilename(inputFilename, 'CollateRuns'))
      } catch (err) {
        console.error('Multiple runs error:', err)
        alert('Error running multiple GLA runs: ' + err)
      } finally {
        setIsLoadingMultiple(false)
      }
    }, 0)
  }

  const atDefaults = isAtDefaults(params, GLA_DEFAULTS)
  const successResult: GlaResultState | null =
    result && !result.error ? (result as GlaResultState) : null
  const valueLabel = successResult?.maxentMode ? 'Weight' : 'Ranking Value'

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Gradual Learning Algorithm</h2>
        <span className="panel-number">04</span>
      </div>

      <div className="nhg-options" style={{ marginBottom: '1rem' }}>
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
        <label className="nhg-checkbox">
          <input
            type="radio"
            name="gla-mode"
            checked={maxentMode}
            onChange={() => setParams({ maxentMode: true })}
          />
          Online MaxEnt (weights)
        </label>
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

      {!maxentMode && (
        <div className="nhg-options">
          <div className="nhg-options-label">Stochastic OT options</div>
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={magriUpdateRule}
              onChange={(e) => setParams({ magriUpdateRule: e.target.checked })}
            />
            Use the Magri update rule
          </label>
        </div>
      )}

      {maxentMode && (
        <div className="nhg-options">
          <div className="nhg-options-label">MaxEnt options</div>
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
              checked={gaussianPrior}
              onChange={(e) => setParams({ gaussianPrior: e.target.checked })}
            />
            Gaussian prior (L2 regularization)
          </label>
          {gaussianPrior && (
            <label className="param-label" style={{ marginLeft: '1.5rem' }}>
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
        </div>
      )}

      <div className="nhg-options">
        <div className="nhg-options-label">Learning schedule</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={useCustomSchedule}
            onChange={(e) => setParams({ useCustomSchedule: e.target.checked })}
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
        <label className="param-label">
          Number of runs
          <select
            value={multipleRunsCount}
            onChange={(e) =>
              setParams({ multipleRunsCount: parseInt(e.target.value) as 10 | 100 | 1000 })
            }
          >
            <option value={10}>10</option>
            <option value={100}>100</option>
            <option value={1000}>1000</option>
          </select>
        </label>
      </div>

      <div className="action-bar">
        <button
          className={`primary-button${isLoading ? ' primary-button--loading' : ''}`}
          onClick={handleRun}
          disabled={isLoading}
        >
          {isLoading ? (
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M5 22h14" />
              <path d="M5 2h14" />
              <path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22" />
              <path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2" />
            </svg>
          ) : (
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <polygon points="5 3 19 12 5 21 5 3"></polygon>
            </svg>
          )}
          Run GLA
        </button>
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
          {isLoadingMultiple ? 'Running…' : `Run ${multipleRunsCount} times & Download`}
        </button>
        <button
          className="reset-button"
          onClick={() => setParams(GLA_DEFAULTS)}
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
        </div>
      )}
    </section>
  )
}

export default GlaPanel
