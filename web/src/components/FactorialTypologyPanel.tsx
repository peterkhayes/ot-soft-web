import { useCallback, useState } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import {
  format_compact_sum_output,
  format_factorial_typology_output,
  format_ft_sum,
  FtOptions,
  FtRunner,
} from '../../pkg/ot_soft.js'
import { useDownload } from '../contexts/downloadContext.ts'
import { useChunkedRunner } from '../hooks/useChunkedRunner.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { makeOutputFilename } from '../utils.ts'
import { ftDefaults } from '../wasmDefaults.ts'
import DownloadButton from './DownloadButton.tsx'
import FtOptionsComponent from './ft/FtOptions.tsx'
import FtResults from './ft/FtResults.tsx'
import type { FtParams, FtResultData, FtState, FtTOrderEntry, FtWinnersRow } from './ft/types.ts'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'

interface FactorialTypologyPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

function FactorialTypologyPanel({
  tableau,
  tableauText,
  inputFilename,
}: FactorialTypologyPanelProps) {
  const [params, setParams] = useLocalStorage<FtParams>('otsoft:params:ft', ftDefaults())
  const download = useDownload()
  const { includeFullListing, includeFtsum, includeCompactSum } = params
  const [aprioriText, setAprioriText] = useState('')
  const [showApriori, setShowApriori] = useState(false)

  const createRunner = useCallback(() => {
    return new FtRunner(tableauText, aprioriText)
  }, [tableauText, aprioriText])

  const extractResult = useCallback(
    (runner: FtRunner): FtState => {
      const ftResult = runner.take_result()
      const patternCount = ftResult.pattern_count()
      const formCount = tableau.form_count()
      const constraintCount = tableau.constraint_count()

      const formInputs: string[] = []
      for (let fi = 0; fi < formCount; fi++) {
        formInputs.push(tableau.get_form(fi)!.input)
      }

      const patterns: { candidates: string[]; isWinner: boolean[] }[] = []
      for (let pi = 0; pi < patternCount; pi++) {
        const candidates: string[] = []
        const isWinner: boolean[] = []
        for (let fi = 0; fi < formCount; fi++) {
          const ci = ftResult.get_pattern_candidate(pi, fi) ?? 0
          const form = tableau.get_form(fi)!
          const cand = form.get_candidate(ci)!
          candidates.push(cand.form)
          isWinner.push(cand.frequency > 0)
        }
        patterns.push({ candidates, isWinner })
      }

      const winners: FtWinnersRow[] = []
      for (let fi = 0; fi < formCount; fi++) {
        const form = tableau.get_form(fi)!
        const row: FtWinnersRow = { formInput: form.input, candidates: [] }
        for (let ci = 0; ci < form.candidate_count(); ci++) {
          const cand = form.get_candidate(ci)!
          row.candidates.push({
            form: cand.form,
            isWinner: cand.frequency > 0,
            derivable: ftResult.is_candidate_derivable(fi, ci),
          })
        }
        winners.push(row)
      }

      const torder: FtTOrderEntry[] = []
      for (let i = 0; i < ftResult.torder_count(); i++) {
        const impFi = ftResult.get_torder_implicator_form(i)
        const impCi = ftResult.get_torder_implicator_candidate(i)
        const tedFi = ftResult.get_torder_implicated_form(i)
        const tedCi = ftResult.get_torder_implicated_candidate(i)
        torder.push({
          implicatorInput: tableau.get_form(impFi)!.input,
          implicatorCandidate: tableau.get_form(impFi)!.get_candidate(impCi)!.form,
          implicatedInput: tableau.get_form(tedFi)!.input,
          implicatedCandidate: tableau.get_form(tedFi)!.get_candidate(tedCi)!.form,
        })
      }

      const alwaysWinners: { input: string; candidate: string }[] = []
      for (const row of winners) {
        const derivable = row.candidates.filter((c) => c.derivable)
        if (derivable.length === 1) {
          alwaysWinners.push({ input: row.formInput, candidate: derivable[0].form })
        }
      }

      const implicatorKeys = new Set<string>()
      for (let i = 0; i < ftResult.torder_count(); i++) {
        implicatorKeys.add(
          `${ftResult.get_torder_implicator_form(i)}:${ftResult.get_torder_implicator_candidate(i)}`,
        )
      }
      const nonImplicators: { input: string; candidate: string }[] = []
      for (let fi = 0; fi < formCount; fi++) {
        const form = tableau.get_form(fi)!
        for (let ci = 0; ci < form.candidate_count(); ci++) {
          if (!implicatorKeys.has(`${fi}:${ci}`)) {
            nonImplicators.push({ input: form.input, candidate: form.get_candidate(ci)!.form })
          }
        }
      }

      ftResult.free()
      return {
        data: {
          patternCount,
          constraintCount,
          formInputs,
          patterns,
          winners,
          torder,
          alwaysWinners,
          nonImplicators,
        },
      }
    },
    [tableau],
  )

  const { state: runnerState, run: handleRun } = useChunkedRunner(createRunner, extractResult)

  const result: FtState | null =
    runnerState.status === 'done'
      ? runnerState.result
      : runnerState.status === 'error'
        ? { error: runnerState.error }
        : null
  const isLoading = runnerState.status === 'running'

  const successResult: FtResultData | null = result && !result.error ? (result.data ?? null) : null

  function handleDownload() {
    try {
      const filename = inputFilename || 'tableau.txt'
      const opts = new FtOptions()
      opts.include_full_listing = includeFullListing
      const output = format_factorial_typology_output(tableauText, filename, aprioriText, opts)
      download(output, makeOutputFilename(inputFilename, 'FactorialTypology'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  function handleDownloadFtsum() {
    try {
      const output = format_ft_sum(tableauText, aprioriText)
      download(output, makeOutputFilename(inputFilename, 'FTSum'))
    } catch (err) {
      console.error('FTSum download error:', err)
      alert('Error generating FTSum download: ' + err)
    }
  }

  function handleDownloadCompactSum() {
    try {
      const output = format_compact_sum_output(tableauText, aprioriText)
      download(output, makeOutputFilename(inputFilename, 'CompactSum'))
    } catch (err) {
      console.error('CompactSum download error:', err)
      alert('Error generating CompactSum download: ' + err)
    }
  }

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Factorial Typology</h2>
        <span className="panel-number">05</span>
      </div>

      <FtOptionsComponent
        params={params}
        setParams={setParams}
        aprioriText={aprioriText}
        setAprioriText={setAprioriText}
        showApriori={showApriori}
        setShowApriori={setShowApriori}
      />

      <div className="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run Factorial Typology" />

        {successResult && (
          <DownloadButton onClick={handleDownload}>Download Results</DownloadButton>
        )}

        {successResult && includeFtsum && (
          <DownloadButton onClick={handleDownloadFtsum}>Download FTSum</DownloadButton>
        )}

        {successResult && includeCompactSum && (
          <DownloadButton onClick={handleDownloadCompactSum}>Download CompactSum</DownloadButton>
        )}
      </div>

      <RunnerProgressBar state={runnerState} />
      {result?.error && (
        <div className="rcd-status failure">Error running Factorial Typology: {result.error}</div>
      )}

      {successResult && <FtResults result={successResult} />}
    </section>
  )
}

export default FactorialTypologyPanel
