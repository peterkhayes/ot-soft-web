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
import { type FtDefaults, ftDefaults } from '../wasmDefaults.ts'
import RunButton from './RunButton.tsx'
import RunnerProgressBar from './RunnerProgressBar.tsx'
import TextFileEditor from './TextFileEditor.tsx'

interface FactorialTypologyPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface FtCandidate {
  form: string
  isWinner: boolean
  derivable: boolean
}

interface FtWinnersRow {
  formInput: string
  candidates: FtCandidate[]
}

interface FtTOrderEntry {
  implicatorInput: string
  implicatorCandidate: string
  implicatedInput: string
  implicatedCandidate: string
}

interface FtResultData {
  patternCount: number
  constraintCount: number
  formInputs: string[]
  patterns: { candidates: string[]; isWinner: boolean[] }[]
  winners: FtWinnersRow[]
  torder: FtTOrderEntry[]
  alwaysWinners: { input: string; candidate: string }[]
  nonImplicators: { input: string; candidate: string }[]
}

interface FtResultState {
  data: FtResultData
  error?: undefined
}

interface FtErrorState {
  error: string
  data?: undefined
}

type FtState = FtResultState | FtErrorState

type FtParams = FtDefaults

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

  const successResult = result && !result.error ? result.data : null

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

      <div className="nhg-options">
        <div className="nhg-options-label">Output options</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeFullListing}
            onChange={(e) => setParams({ includeFullListing: e.target.checked })}
          />
          Include rankings in results
        </label>

        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeFtsum}
            onChange={(e) => setParams({ includeFtsum: e.target.checked })}
          />
          Generate FTSum file
        </label>

        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeCompactSum}
            onChange={(e) => setParams({ includeCompactSum: e.target.checked })}
          />
          Generate CompactSum file
        </label>
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">A priori rankings</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={showApriori}
            onChange={(e) => {
              setShowApriori(e.target.checked)
              if (!e.target.checked) setAprioriText('')
            }}
          />
          Use a priori rankings
        </label>
        {showApriori && (
          <TextFileEditor
            value={aprioriText}
            onChange={setAprioriText}
            hint="Tab-delimited constraint × constraint matrix (abbreviations must match current tableau)."
            placeholder="Load from file or paste content here…"
            testId="ft-apriori-file-input"
          />
        )}
      </div>

      <div className="action-bar">
        <RunButton isLoading={isLoading} onClick={handleRun} label="Run Factorial Typology" />

        {successResult && (
          <button className="download-button" onClick={handleDownload}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="7 10 12 15 17 10" />
              <line x1="12" y1="15" x2="12" y2="3" />
            </svg>
            Download Results
          </button>
        )}

        {successResult && includeFtsum && (
          <button className="download-button" onClick={handleDownloadFtsum}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="7 10 12 15 17 10" />
              <line x1="12" y1="15" x2="12" y2="3" />
            </svg>
            Download FTSum
          </button>
        )}

        {successResult && includeCompactSum && (
          <button className="download-button" onClick={handleDownloadCompactSum}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="7 10 12 15 17 10" />
              <line x1="12" y1="15" x2="12" y2="3" />
            </svg>
            Download CompactSum
          </button>
        )}
      </div>

      <RunnerProgressBar state={runnerState} />
      {result?.error && (
        <div className="rcd-status failure">Error running Factorial Typology: {result.error}</div>
      )}

      {successResult && (
        <div className="ft-results">
          {/* Summary */}
          <div className="ft-summary">
            <span className="ft-summary-count">{successResult.patternCount}</span> output{' '}
            {successResult.patternCount === 1 ? 'pattern' : 'patterns'} found
          </div>

          {successResult.patternCount > 0 && (
            <>
              {/* Patterns table */}
              <div className="ft-section">
                <div className="ft-section-header">Output Patterns</div>
                <p className="ft-section-note">
                  Forms marked as winners in the input are marked with ›.
                </p>
                <div className="ft-patterns-scroll">
                  <table className="ft-patterns-table">
                    <thead>
                      <tr>
                        <th className="ft-input-col">Input</th>
                        {successResult.patterns.map((_, pi) => (
                          <th key={pi} className="ft-pattern-col">
                            Output #{pi + 1}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {successResult.formInputs.map((input, fi) => (
                        <tr key={fi}>
                          <td className="ft-input-cell">/{input}/</td>
                          {successResult.patterns.map((pattern, pi) => (
                            <td
                              key={pi}
                              className={`ft-pattern-cell${pattern.isWinner[fi] ? ' ft-pattern-cell--winner' : ''}`}
                            >
                              {pattern.isWinner[fi] && <span className="ft-winner-marker">›</span>}
                              {pattern.candidates[fi]}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* List of winners */}
              <div className="ft-section">
                <div className="ft-section-header">List of Winners</div>
                <p className="ft-section-note">
                  For each candidate, whether there is at least one ranking that derives it.
                </p>
                <div className="ft-winners">
                  {successResult.winners.map((row, fi) => (
                    <div className="ft-winners-row" key={fi}>
                      <div className="ft-winners-input">/{row.formInput}/</div>
                      <div className="ft-winners-candidates">
                        {row.candidates.map((cand, ci) => (
                          <div className="ft-winners-candidate" key={ci}>
                            {cand.isWinner && <span className="ft-winner-marker">›</span>}
                            <span className="ft-cand-form">{cand.form}</span>
                            <span
                              className={`ft-derivable-badge${cand.derivable ? ' ft-derivable-badge--yes' : ' ft-derivable-badge--no'}`}
                            >
                              {cand.derivable ? 'derivable' : 'not derivable'}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* T-Order */}
              <div className="ft-section">
                <div className="ft-section-header">T-Order</div>
                <p className="ft-section-note">
                  The set of implications in the factorial typology.
                </p>

                {successResult.alwaysWinners.length > 0 && (
                  <div className="ft-torder-note">
                    <strong>No competition:</strong> the following forms have only one derivable
                    output and are not listed in the t-order.
                    <div className="ft-always-winners">
                      {successResult.alwaysWinners.map((aw, i) => (
                        <span className="ft-always-winner-item" key={i}>
                          /{aw.input}/ → [{aw.candidate}]
                        </span>
                      ))}
                    </div>
                  </div>
                )}

                {successResult.torder.length === 0 ? (
                  <div className="ft-torder-empty">No t-order implications were found.</div>
                ) : (
                  <table className="ft-torder-table">
                    <thead>
                      <tr>
                        <th>If this input</th>
                        <th>has this output</th>
                        <th>then this input</th>
                        <th>has this output</th>
                      </tr>
                    </thead>
                    <tbody>
                      {successResult.torder.map((entry, i) => (
                        <tr key={i}>
                          <td className="ft-torder-input">/{entry.implicatorInput}/</td>
                          <td className="ft-torder-cand">[{entry.implicatorCandidate}]</td>
                          <td className="ft-torder-input">/{entry.implicatedInput}/</td>
                          <td className="ft-torder-cand">[{entry.implicatedCandidate}]</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}

                {successResult.nonImplicators.length > 0 && (
                  <div className="ft-torder-note ft-torder-note--below">
                    <strong>Nothing implied by:</strong>
                    <div className="ft-always-winners">
                      {successResult.nonImplicators.map((ni, i) => (
                        <span className="ft-always-winner-item" key={i}>
                          /{ni.input}/ → [{ni.candidate}]
                        </span>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      )}
    </section>
  )
}

export default FactorialTypologyPanel
