import { useState, useRef } from 'react'
import { useLocalStorage } from '../hooks/useLocalStorage.ts'
import { run_rcd, format_rcd_output, run_bcd, format_bcd_output, run_lfcd, format_lfcd_output, FredOptions } from '../../pkg/ot_soft.js'
import type { Tableau } from '../../pkg/ot_soft.js'
import { downloadTextFile, makeOutputFilename, isAtDefaults } from '../utils.ts'

interface RcdPanelProps {
  tableau: Tableau
  tableauText: string
  inputFilename: string | null
}

interface StratumData {
  constraints: { abbrev: string; fullName: string }[]
}

interface RcdResultState {
  success: boolean
  strata: StratumData[]
  tieWarning: boolean
  error?: undefined
}

interface RcdErrorState {
  error: string
  success?: undefined
  strata?: undefined
  tieWarning?: undefined
}

type RcdState = RcdResultState | RcdErrorState

type Algorithm = 'rcd' | 'bcd' | 'bcd-specific' | 'lfcd'

interface RcdParams { algorithm: Algorithm; includeFred: boolean; useMib: boolean; showDetails: boolean; includeMiniTableaux: boolean }
const RCD_DEFAULTS: RcdParams = { algorithm: 'rcd', includeFred: true, useMib: false, showDetails: true, includeMiniTableaux: true }

const ALGORITHM_LABELS: Record<Algorithm, string> = {
  'rcd': 'RCD',
  'bcd': 'BCD',
  'bcd-specific': 'BCD (Specific)',
  'lfcd': 'LFCD',
}

const ALGORITHM_DESCRIPTIONS: Record<Algorithm, string> = {
  'rcd': 'Recursive Constraint Demotion',
  'bcd': 'Biased Constraint Demotion',
  'bcd-specific': 'Biased Constraint Demotion (favors specific faithfulness)',
  'lfcd': 'Low Faithfulness Constraint Demotion',
}

function RcdPanel({ tableau, tableauText, inputFilename }: RcdPanelProps) {
  const [rcdResult, setRcdResult] = useState<RcdState | null>(null)
  const [params, setParams] = useLocalStorage<RcdParams>('otsoft:params:rcd', RCD_DEFAULTS)
  const { algorithm, includeFred, useMib, showDetails, includeMiniTableaux } = params
  const [aprioriText, setAprioriText] = useState<string>('')
  const [aprioriFilename, setAprioriFilename] = useState<string | null>(null)
  const [isAprioriDragging, setIsAprioriDragging] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const aprioriInputRef = useRef<HTMLInputElement>(null)

  const supportsApriori = algorithm === 'rcd' || algorithm === 'lfcd'
  const atDefaults = isAtDefaults(params, RCD_DEFAULTS)

  function loadAprioriFile(file: File) {
    file.text().then(text => {
      setAprioriText(text)
      setAprioriFilename(file.name)
    }).catch(err => {
      console.error('Error reading a priori file:', err)
    })
  }

  function handleAprioriFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (file) loadAprioriFile(file)
  }

  function handleAprioriDrop(e: React.DragEvent) {
    e.preventDefault()
    setIsAprioriDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file) loadAprioriFile(file)
  }

  function handleAprioriClear() {
    setAprioriText('')
    setAprioriFilename(null)
    if (aprioriInputRef.current) aprioriInputRef.current.value = ''
  }

  function handleRun() {
    setIsLoading(true)
    setTimeout(() => {
    try {
      const apriori = supportsApriori ? aprioriText : ''
      const result = algorithm === 'rcd'
        ? run_rcd(tableauText, apriori)
        : algorithm === 'lfcd'
          ? run_lfcd(tableauText, apriori)
          : run_bcd(tableauText, algorithm === 'bcd-specific')

      const numStrata = result.num_strata()
      const constraintCount = tableau.constraint_count()

      const strata: StratumData[] = []
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

      setRcdResult({
        success: result.success(),
        strata,
        tieWarning: result.tie_warning(),
      })
    } catch (err) {
      console.error('Algorithm error:', err)
      setRcdResult({ error: String(err) })
    } finally {
      setIsLoading(false)
    }
    }, 0)
  }

  function handleDownload() {
    try {
      const apriori = supportsApriori ? aprioriText : ''
      const fredOpts = new FredOptions()
      fredOpts.include_fred = includeFred
      fredOpts.use_mib = useMib
      fredOpts.show_details = showDetails
      fredOpts.include_mini_tableaux = includeMiniTableaux
      const formattedOutput = algorithm === 'rcd'
        ? format_rcd_output(tableauText, inputFilename || 'tableau.txt', apriori, fredOpts)
        : algorithm === 'lfcd'
          ? format_lfcd_output(tableauText, inputFilename || 'tableau.txt', apriori, fredOpts)
          : format_bcd_output(tableauText, inputFilename || 'tableau.txt', algorithm === 'bcd-specific', fredOpts)

      downloadTextFile(formattedOutput, makeOutputFilename(inputFilename, 'Output'))
    } catch (err) {
      console.error('Download error:', err)
      alert('Error generating download: ' + err)
    }
  }

  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Constraint Ranking</h2>
        <span className="panel-number">04</span>
      </div>

      <div className="action-bar">
        <select
          className="algorithm-select"
          value={algorithm}
          onChange={(e) => setParams({ algorithm: e.target.value as Algorithm })}
          title={ALGORITHM_DESCRIPTIONS[algorithm]}
        >
          <option value="rcd">RCD</option>
          <option value="bcd">BCD</option>
          <option value="bcd-specific">BCD (Specific)</option>
          <option value="lfcd">LFCD</option>
        </select>
        {supportsApriori && (
          <div
            className={`apriori-upload${isAprioriDragging ? ' apriori-upload--dragging' : ''}${aprioriFilename ? ' apriori-upload--loaded' : ''}`}
            title="Optional: load an a priori rankings file"
            onClick={() => aprioriInputRef.current?.click()}
            onDragOver={(e) => { e.preventDefault(); setIsAprioriDragging(true) }}
            onDragLeave={() => setIsAprioriDragging(false)}
            onDrop={handleAprioriDrop}
          >
            <input ref={aprioriInputRef} type="file" accept=".txt" onChange={handleAprioriFile} style={{ display: 'none' }} />
            <span className="apriori-upload-label">{aprioriFilename ?? 'A priori rankings (optional)'}</span>
            {aprioriFilename && (
              <button
                className="apriori-clear"
                onClick={(e) => { e.stopPropagation(); handleAprioriClear() }}
                title="Clear a priori rankings"
              >×</button>
            )}
          </div>
        )}
        <button className={`primary-button${isLoading ? ' primary-button--loading' : ''}`} onClick={handleRun} disabled={isLoading}>
          {isLoading ? (
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M5 22h14"/><path d="M5 2h14"/>
              <path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22"/>
              <path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2"/>
            </svg>
          ) : (
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <polygon points="5 3 19 12 5 21 5 3"></polygon>
            </svg>
          )}
          Run {ALGORITHM_LABELS[algorithm]} Algorithm
        </button>
        {rcdResult && !rcdResult.error && (
          <button className="download-button" onClick={handleDownload}>
            <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Download Results
          </button>
        )}
        <button className="reset-button" onClick={() => setParams(RCD_DEFAULTS)} disabled={atDefaults}>
          <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <polyline points="1 4 1 10 7 10"></polyline>
            <path d="M3.51 15a9 9 0 1 0 .49-4.99"></path>
          </svg>
          Reset to Defaults
        </button>
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Ranking Argumentation</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeFred}
            onChange={(e) => setParams({ includeFred: e.target.checked })}
          />
          Include ranking arguments
        </label>
        {includeFred && (
          <>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={useMib}
                onChange={(e) => setParams({ useMib: e.target.checked })}
              />
              Use Most Informative Basis
            </label>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={showDetails}
                onChange={(e) => setParams({ showDetails: e.target.checked })}
              />
              Show details of argumentation
            </label>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={includeMiniTableaux}
                onChange={(e) => setParams({ includeMiniTableaux: e.target.checked })}
              />
              Include illustrative mini-tableaux
            </label>
          </>
        )}
      </div>

      {rcdResult && (
        rcdResult.error ? (
          <div className="rcd-status failure">
            Error running algorithm: {rcdResult.error}
          </div>
        ) : (
          <div className="rcd-results">
            <div className={`rcd-status ${rcdResult.success ? 'success' : 'failure'}`}>
              {rcdResult.success
                ? 'A ranking was found that generates the correct outputs'
                : 'Failed to find a valid ranking'}
            </div>

            {rcdResult.tieWarning && (
              <div className="rcd-status warning">
                Caution: BCD selected arbitrarily among tied faithfulness constraint subsets.
                Try changing the order of faithfulness constraints in the input file to see
                whether this results in a different ranking.
              </div>
            )}

            {rcdResult.strata!.map((stratum, s) => (
              <div className="stratum" key={s}>
                <div className="stratum-header">Stratum {s + 1}</div>
                <div className="constraint-list">
                  {stratum.constraints.map((constraint, i) => (
                    <div className="constraint-item" key={i}>
                      <span className="abbrev">{constraint.abbrev}</span>
                      <span className="full-name">({constraint.fullName})</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )
      )}
    </section>
  )
}

export default RcdPanel
