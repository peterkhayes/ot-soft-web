import { useCallback } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import { format_maxent_output, MaxEntOptions, MaxEntRunner } from '../../pkg/ot_soft.js'
import { useDownload } from '../contexts/downloadContext.ts'
import { useChunkedRunner } from '../hooks/useChunkedRunner.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import { maxentDefaults } from '../wasmDefaults.ts'
import DownloadMenu, { type DownloadMenuItem } from './DownloadMenu.tsx'
import MaxEntOptionsComponent from './maxent/MaxEntOptions.tsx'
import MaxEntParameterInputs from './maxent/MaxEntParameterInputs.tsx'
import MaxEntResults from './maxent/MaxEntResults.tsx'
import type { MaxEntParams, MaxEntResultState, MaxEntState } from './maxent/types.ts'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'

interface MaxEntPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

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

      <MaxEntParameterInputs params={params} setParams={setParams} />
      <MaxEntOptionsComponent params={params} setParams={setParams} />

      <div className="action-bar" data-testid="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run MaxEnt" />
        {result && !result.error && (
          <DownloadMenu
            items={[
              { label: 'Download Results', onClick: handleDownload },
              ...(successResult?.history
                ? [{ label: 'Download History', onClick: handleDownloadHistory }]
                : []),
              ...(successResult?.outputProbHistory
                ? [
                    {
                      label: 'Download Output Probability History',
                      onClick: handleDownloadOutputProbHistory,
                    } satisfies DownloadMenuItem,
                  ]
                : []),
            ]}
          />
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
            aria-hidden="true"
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
        <MaxEntResults result={successResult} constraintAbbrevs={constraintAbbrevs} />
      )}
    </section>
  )
}

export default MaxEntPanel
