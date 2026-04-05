import { useCallback, useState } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import { format_nhg_output, NhgOptions, NhgRunner } from '../../pkg/ot_soft.js'
import { useDownload } from '../contexts/downloadContext.ts'
import { useChunkedRunner } from '../hooks/useChunkedRunner.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import { nhgDefaults } from '../wasmDefaults.ts'
import DownloadMenu from './DownloadMenu.tsx'
import NhgNoiseOptions from './nhg/NhgNoiseOptions.tsx'
import NhgParameterInputs from './nhg/NhgParameterInputs.tsx'
import NhgResults from './nhg/NhgResults.tsx'
import NhgScheduleOptions from './nhg/NhgScheduleOptions.tsx'
import type { NhgParams, NhgResultState, NhgState } from './nhg/types.ts'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'

interface NhgPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

function NhgPanel({ tableau, tableauText, inputFilename }: NhgPanelProps) {
  const [params, setParams] = useLocalStorage<NhgParams>('otsoft:params:nhg', nhgDefaults())

  const [scheduleError, setScheduleError] = useState<string | null>(null)
  const download = useDownload()

  // useCallback rather than useMemo: NhgOptions is a WASM heap object with no cleanup
  // phase available in useMemo, so we build it lazily on demand to avoid leaking abandoned instances.
  const buildOpts = useCallback((): NhgOptions => {
    const opts = new NhgOptions()
    opts.cycles = params.cycles
    opts.initial_plasticity = params.initialPlasticity
    opts.final_plasticity = params.finalPlasticity
    opts.test_trials = params.testTrials
    opts.noise_by_cell = params.noiseByCell
    opts.post_mult_noise = params.postMultNoise
    opts.noise_for_zero_cells = params.noiseForZeroCells
    opts.late_noise = params.lateNoise
    opts.exponential_nhg = params.exponentialNhg
    opts.demi_gaussians = params.demiGaussians
    opts.negative_weights_ok = params.negativeWeightsOk
    opts.resolve_ties_by_skipping = params.resolveTiesBySkipping
    opts.exact_proportions = params.exactProportions
    opts.generate_history = params.generateHistory
    opts.generate_full_history = params.generateFullHistory
    if (params.useCustomSchedule) {
      opts.learning_schedule = params.customSchedule
    }
    return opts
  }, [params])

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

      <NhgParameterInputs params={params} setParams={setParams} />
      <NhgNoiseOptions params={params} setParams={setParams} />
      <NhgScheduleOptions
        params={params}
        setParams={setParams}
        scheduleError={scheduleError}
        setScheduleError={setScheduleError}
      />

      <div className="action-bar" data-testid="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run Noisy HG" />
        {result && !result.error && (
          <DownloadMenu
            items={[
              { label: 'Download Results', onClick: handleDownload },
              ...(successResult?.history
                ? [{ label: 'Download History', onClick: handleDownloadHistory }]
                : []),
              ...(successResult?.fullHistory
                ? [{ label: 'Download Full History', onClick: handleDownloadFullHistory }]
                : []),
            ]}
          />
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
            aria-hidden="true"
          >
            <polyline points="1 4 1 10 7 10"></polyline>
            <path d="M3.51 15a9 9 0 1 0 .49-4.99"></path>
          </svg>
          Reset to Defaults
        </button>
      </div>

      <RunnerProgressBar state={runnerState} />
      {result?.error && <div className="rcd-status failure">Error running NHG: {result.error}</div>}
      {successResult && <NhgResults result={successResult} />}
    </section>
  )
}

export default NhgPanel
