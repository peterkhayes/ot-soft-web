import { useState } from 'react'

import type { Tableau } from '../../pkg/ot_soft.js'
import {
  AxisMode,
  clear_log,
  format_bcd_html_output,
  format_bcd_output,
  format_lfcd_html_output,
  format_lfcd_output,
  format_rcd_html_output,
  format_rcd_output,
  format_sorted_input_file,
  fred_hasse_dot,
  FredOptions,
  get_log,
  run_bcd,
  run_lfcd,
  run_rcd,
} from '../../pkg/ot_soft.js'
import { useDownload } from '../contexts/downloadContext.ts'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { isAtDefaults, makeOutputFilename } from '../utils.ts'
import { rcdDefaults } from '../wasmDefaults.ts'
import DownloadButton from './DownloadButton.tsx'
import RcdOptions from './rcd/RcdOptions.tsx'
import RcdResults from './rcd/RcdResults.tsx'
import { ALGORITHM_LABELS, type RcdParams, type RcdState } from './rcd/types.ts'

interface RcdPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
  axisMode?: AxisMode
}

function RcdPanel({
  tableau,
  tableauText,
  inputFilename,
  axisMode = AxisMode.SwitchAll,
}: RcdPanelProps) {
  const [rcdResult, setRcdResult] = useState<RcdState | null>(null)
  const [params, setParams] = useLocalStorage<RcdParams>('otsoft:params:rcd', rcdDefaults())
  const { algorithm, includeFred, useMib, showDetails, includeMiniTableaux, diagnostics } = params
  const [aprioriText, setAprioriText] = useState<string>('')
  const [showApriori, setShowApriori] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const download = useDownload()

  const supportsApriori = algorithm === 'rcd' || algorithm === 'lfcd'
  const atDefaults = isAtDefaults(params, rcdDefaults())

  function handleRun() {
    setIsLoading(true)
    setTimeout(() => {
      try {
        const apriori = supportsApriori ? aprioriText : ''
        clear_log()
        const result =
          algorithm === 'rcd'
            ? run_rcd(tableauText, apriori)
            : algorithm === 'lfcd'
              ? run_lfcd(tableauText, apriori)
              : run_bcd(tableauText, algorithm === 'bcd-specific')
        const log = get_log()

        const numStrata = result.num_strata()
        const constraintCount = tableau.constraint_count()

        const strata: { constraints: { abbrev: string; fullName: string }[] }[] = []
        for (let s = 0; s < numStrata; s++) {
          strata.push({ constraints: [] })
        }

        for (let i = 0; i < constraintCount; i++) {
          const stratum = result.get_stratum(i)
          if (stratum !== undefined && stratum >= 1 && stratum <= numStrata) {
            const constraint = tableau.get_constraint(i)!
            strata[stratum - 1].constraints.push({
              abbrev: constraint.abbrev,
              fullName: constraint.full_name,
            })
          }
        }

        let hasseDot: string | undefined
        if (includeFred) {
          try {
            hasseDot = fred_hasse_dot(tableauText, apriori, useMib)
          } catch (e) {
            console.warn('Hasse diagram generation failed:', e)
          }
        }

        setRcdResult({
          success: result.success(),
          strata,
          tieWarning: result.tie_warning(),
          hasseDot,
          log,
        })
      } catch (err) {
        console.error('Algorithm error:', err)
        setRcdResult({ error: String(err) })
      } finally {
        setIsLoading(false)
      }
    }, 0)
  }

  function buildFredOpts(): FredOptions {
    const fredOpts = new FredOptions()
    fredOpts.include_fred = includeFred
    fredOpts.use_mib = useMib
    fredOpts.show_details = showDetails
    fredOpts.include_mini_tableaux = includeMiniTableaux
    fredOpts.diagnostics = diagnostics
    return fredOpts
  }

  function handleDownload() {
    try {
      const apriori = supportsApriori ? aprioriText : ''
      const fredOpts = buildFredOpts()
      const formattedOutput =
        algorithm === 'rcd'
          ? format_rcd_output(tableauText, inputFilename || 'tableau.txt', apriori, fredOpts)
          : algorithm === 'lfcd'
            ? format_lfcd_output(tableauText, inputFilename || 'tableau.txt', apriori, fredOpts)
            : format_bcd_output(
                tableauText,
                inputFilename || 'tableau.txt',
                algorithm === 'bcd-specific',
                fredOpts,
              )

      download(formattedOutput, makeOutputFilename(inputFilename, 'Output'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  function handleDownloadHtml() {
    try {
      const apriori = supportsApriori ? aprioriText : ''
      const fredOpts = buildFredOpts()
      const htmlContent =
        algorithm === 'rcd'
          ? format_rcd_html_output(
              tableauText,
              inputFilename || 'tableau.txt',
              apriori,
              fredOpts,
              axisMode,
            )
          : algorithm === 'lfcd'
            ? format_lfcd_html_output(
                tableauText,
                inputFilename || 'tableau.txt',
                apriori,
                fredOpts,
                axisMode,
              )
            : format_bcd_html_output(
                tableauText,
                inputFilename || 'tableau.txt',
                algorithm === 'bcd-specific',
                fredOpts,
                axisMode,
              )

      download(htmlContent, makeOutputFilename(inputFilename, 'Output', '.html'))
    } catch (err) {
      console.error('HTML download error:', err)
      alert('Error generating HTML download: ' + err)
    }
  }

  function handleDownloadLog() {
    if (rcdResult && !('error' in rcdResult) && rcdResult.log) {
      download(rcdResult.log, makeOutputFilename(inputFilename, 'HowIRanked'))
    }
  }

  function handleDownloadSortedInput() {
    try {
      const apriori = supportsApriori ? aprioriText : ''
      const alg = algorithm === 'bcd-specific' ? 'bcd-specific' : algorithm
      const content = format_sorted_input_file(tableauText, apriori, alg)
      download(content, makeOutputFilename(inputFilename, 'Sorted'))
    } catch (err) {
      console.error('Sorted input download error:', err)
      alert('Error generating sorted input: ' + err)
    }
  }

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Constraint Ranking</h2>
        <span className="panel-number">04</span>
      </div>

      <RcdOptions
        params={params}
        setParams={setParams}
        aprioriText={aprioriText}
        setAprioriText={setAprioriText}
        showApriori={showApriori}
        setShowApriori={setShowApriori}
      />

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
              aria-hidden="true"
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
              aria-hidden="true"
            >
              <polygon points="5 3 19 12 5 21 5 3"></polygon>
            </svg>
          )}
          Run {ALGORITHM_LABELS[algorithm]} Algorithm
        </button>
        {rcdResult && !rcdResult.error && (
          <>
            <DownloadButton onClick={handleDownload}>Download Results</DownloadButton>
            <DownloadButton onClick={handleDownloadHtml}>Download HTML</DownloadButton>
            <DownloadButton onClick={handleDownloadSortedInput}>
              Download Sorted Input
            </DownloadButton>
            <DownloadButton onClick={handleDownloadLog}>Download Log</DownloadButton>
          </>
        )}
        <button
          className="reset-button"
          onClick={() => setParams(rcdDefaults())}
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

      {rcdResult &&
        (rcdResult.error ? (
          <div className="rcd-status failure">Error running algorithm: {rcdResult.error}</div>
        ) : (
          <RcdResults
            result={rcdResult as import('./rcd/types.ts').RcdResultState}
            inputFilename={inputFilename}
          />
        ))}
    </section>
  )
}

export default RcdPanel
