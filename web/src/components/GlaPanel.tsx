import { useCallback, useEffect, useState } from 'react'

import type { GlaResult, Tableau } from '../../pkg/ot_soft.js'
import {
  format_gla_output,
  gla_hasse_dot,
  gla_pairwise_probabilities_json,
  GlaMultipleRunsRunner,
  GlaOptions,
  GlaRunner,
} from '../../pkg/ot_soft.js'
import { useDownload } from '../contexts/downloadContext.ts'
import { useChunkedRunner } from '../hooks/useChunkedRunner.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import { glaDefaults } from '../wasmDefaults.ts'
import DownloadButton from './DownloadButton.tsx'
import GlaAprioriOptions from './gla/GlaAprioriOptions.tsx'
import GlaFrameworkOptions from './gla/GlaFrameworkOptions.tsx'
import GlaOptionsGrid from './gla/GlaOptionsGrid.tsx'
import GlaParameterInputs from './gla/GlaParameterInputs.tsx'
import GlaResults from './gla/GlaResults.tsx'
import type { GlaParams, GlaResultState, GlaState } from './gla/types.ts'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'

interface GlaPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

function GlaPanel({ tableau, tableauText, inputFilename }: GlaPanelProps) {
  const [params, setParams] = useLocalStorage<GlaParams>('otsoft:params:gla', glaDefaults())
  const { maxentMode, multipleRunsCount } = params

  const [scheduleError, setScheduleError] = useState<string | null>(null)
  const download = useDownload()

  // useCallback rather than useMemo: GlaOptions is a WASM heap object with no cleanup
  // phase available in useMemo, so we build it lazily on demand to avoid leaking abandoned instances.
  const buildOpts = useCallback((): GlaOptions => {
    const opts = new GlaOptions()
    opts.maxent_mode = params.maxentMode
    opts.cycles = params.cycles
    opts.initial_plasticity = params.initialPlasticity
    opts.final_plasticity = params.finalPlasticity
    opts.test_trials = params.maxentMode ? 0 : params.testTrials
    opts.negative_weights_ok = params.negativeWeightsOk
    opts.gaussian_prior = params.maxentMode && params.gaussianPrior
    opts.sigma = params.sigma
    opts.magri_update_rule = !params.maxentMode && params.magriUpdateRule
    opts.exact_proportions = params.exactProportions
    if (!params.maxentMode && params.aprioriText.trim()) {
      opts.apriori_text = params.aprioriText
      opts.apriori_gap = params.aprioriGap
    }
    opts.generate_history = params.generateHistory
    opts.generate_full_history = params.generateFullHistory
    opts.generate_candidate_prob_history = params.maxentMode && params.generateCandidateProbHistory
    if (params.useCustomSchedule) {
      opts.learning_schedule = params.customSchedule
    }
    return opts
  }, [params])

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
        return {
          fullName: c.full_name,
          abbrev: c.abbrev,
          value: r.get_ranking_value(i),
          active: r.get_active_constraint(i),
        }
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
      let pairwiseData: { headers: string[]; matrix: string[][] } | undefined
      if (!maxentMode) {
        const rankingValues = new Float64Array(values.map((v) => v.value))
        try {
          hasseDot = gla_hasse_dot(tableauText, rankingValues)
        } catch (e) {
          console.warn('GLA Hasse diagram generation failed:', e)
        }
        try {
          pairwiseData = JSON.parse(
            gla_pairwise_probabilities_json(tableauText, rankingValues),
          ) as { headers: string[]; matrix: string[][] }
        } catch (e) {
          console.warn('GLA pairwise probabilities data generation failed:', e)
        }
      }

      const extracted: GlaResultState = {
        values,
        forms,
        logLikelihood: r.log_likelihood(),
        maxentMode: r.is_maxent_mode(),
        hasseDot,
        pairwiseData,
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
  const extractMultipleRunsResult = useCallback(
    (runner: GlaMultipleRunsRunner) => runner.take_result(),
    [],
  )
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

      <GlaFrameworkOptions params={params} setParams={setParams} />
      <GlaParameterInputs params={params} setParams={setParams} />
      <GlaOptionsGrid
        params={params}
        setParams={setParams}
        scheduleError={scheduleError}
        setScheduleError={setScheduleError}
      />
      {!maxentMode && <GlaAprioriOptions params={params} setParams={setParams} />}

      <div className="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run GLA" />
        {result && !result.error && (
          <DownloadButton onClick={handleDownload}>Download Results</DownloadButton>
        )}
        {successResult?.history && (
          <DownloadButton onClick={handleDownloadHistory}>Download History</DownloadButton>
        )}
        {successResult?.fullHistory && (
          <DownloadButton onClick={handleDownloadFullHistory}>Download Full History</DownloadButton>
        )}
        {successResult?.candidateProbHistory && (
          <DownloadButton onClick={handleDownloadCandidateProbHistory}>
            Download Candidate Probability History
          </DownloadButton>
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
            aria-hidden="true"
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
            aria-hidden="true"
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
      {successResult && <GlaResults result={successResult} inputFilename={inputFilename} />}
    </section>
  )
}

export default GlaPanel
