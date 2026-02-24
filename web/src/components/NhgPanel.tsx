import { useState } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import { format_nhg_output, NhgOptions, run_nhg } from '../../pkg/ot_soft.js'
import { useDownload } from '../contexts/downloadContext.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import TextFileEditor from './TextFileEditor.tsx'

const DEFAULT_SCHEDULE_TEMPLATE =
  'Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n' +
  '15000\t2\t2\t2\t2\n' +
  '15000\t0.2\t0.2\t2\t2\n' +
  '15000\t0.02\t0.02\t2\t2\n' +
  '15000\t0.002\t0.002\t2\t2'

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

interface NhgParams {
  cycles: number
  initialPlasticity: number
  finalPlasticity: number
  testTrials: number
  noiseByCell: boolean
  postMultNoise: boolean
  noiseForZeroCells: boolean
  lateNoise: boolean
  exponentialNhg: boolean
  demiGaussians: boolean
  negativeWeightsOk: boolean
  resolveTiesBySkipping: boolean
  useCustomSchedule: boolean
  customSchedule: string
}
const NHG_DEFAULTS: NhgParams = {
  cycles: 5000,
  initialPlasticity: 2.0,
  finalPlasticity: 0.002,
  testTrials: 2000,
  noiseByCell: false,
  postMultNoise: false,
  noiseForZeroCells: false,
  lateNoise: false,
  exponentialNhg: false,
  demiGaussians: false,
  negativeWeightsOk: false,
  resolveTiesBySkipping: false,
  useCustomSchedule: false,
  customSchedule: DEFAULT_SCHEDULE_TEMPLATE,
}

function NhgPanel({ tableau, tableauText, inputFilename }: NhgPanelProps) {
  const [params, setParams] = useLocalStorage<NhgParams>('otsoft:params:nhg', NHG_DEFAULTS)
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
    useCustomSchedule,
    customSchedule,
  } = params

  const [result, setResult] = useState<NhgState | null>(null)
  const [scheduleError, setScheduleError] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)
  const download = useDownload()

  function buildOpts(): NhgOptions {
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
        const r = run_nhg(tableauText, opts)

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
              obsPct: totalFreq > 0 ? (cand.frequency / totalFreq) * 100 : 0,
              genPct: r.get_test_prob(formIdx, candIdx) * 100,
            }
          })

          return { input: form.input, candidates }
        })

        setResult({ weights, forms, logLikelihood: r.log_likelihood() })
      } catch (err) {
        console.error('NHG error:', err)
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
      const formattedOutput = format_nhg_output(tableauText, inputFilename || 'tableau.txt', opts)
      download(formattedOutput, makeOutputFilename(inputFilename, 'NHGOutput'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  const atDefaults = isAtDefaults(params, NHG_DEFAULTS)
  const successResult: NhgResultState | null =
    result && !result.error ? (result as NhgResultState) : null

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
            testId="nhg-schedule-file-input"
          />
        )}
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
          Run Noisy HG
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
          className="reset-button"
          onClick={() => setParams(NHG_DEFAULTS)}
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
