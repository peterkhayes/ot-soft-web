import { useState } from 'react'
import { run_rcd, format_rcd_output, run_bcd, format_bcd_output, run_lfcd, format_lfcd_output, FredOptions } from '../../pkg/ot_soft.js'
import type { Tableau } from '../../pkg/ot_soft.js'
import { downloadTextFile, makeOutputFilename } from '../utils.ts'

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
  const [algorithm, setAlgorithm] = useState<Algorithm>('rcd')
  const [aprioriText, setAprioriText] = useState<string>('')
  const [aprioriFilename, setAprioriFilename] = useState<string | null>(null)
  const [isLoading, setIsLoading] = useState(false)

  // Ranking argumentation options (match VB6 defaults)
  const [includeFred, setIncludeFred] = useState<boolean>(true)
  const [useMib, setUseMib] = useState<boolean>(false)
  const [showDetails, setShowDetails] = useState<boolean>(true)
  const [includeMiniTableaux, setIncludeMiniTableaux] = useState<boolean>(true)

  const supportsApriori = algorithm === 'rcd' || algorithm === 'lfcd'

  function handleAprioriFile(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    file.text().then(text => {
      setAprioriText(text)
      setAprioriFilename(file.name)
    }).catch(err => {
      console.error('Error reading a priori file:', err)
    })
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
          onChange={(e) => setAlgorithm(e.target.value as Algorithm)}
          title={ALGORITHM_DESCRIPTIONS[algorithm]}
        >
          <option value="rcd">RCD</option>
          <option value="bcd">BCD</option>
          <option value="bcd-specific">BCD (Specific)</option>
          <option value="lfcd">LFCD</option>
        </select>
        {supportsApriori && (
          <label className="apriori-upload" title="Optional: load an a priori rankings file">
            <span>{aprioriFilename ?? 'A priori rankings (optional)'}</span>
            <input type="file" accept=".txt" onChange={handleAprioriFile} style={{ display: 'none' }} />
          </label>
        )}
        <button className="primary-button" onClick={handleRun} disabled={isLoading}>
          <svg className="button-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <polygon points="5 3 19 12 5 21 5 3"></polygon>
          </svg>
          Run {ALGORITHM_LABELS[algorithm]} Algorithm
        </button>
        {isLoading && (
          <div className="loading-indicator" title={`Running ${ALGORITHM_LABELS[algorithm]}...`}></div>
        )}
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
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Ranking Argumentation</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeFred}
            onChange={(e) => setIncludeFred(e.target.checked)}
          />
          Include ranking arguments
        </label>
        {includeFred && (
          <>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={useMib}
                onChange={(e) => setUseMib(e.target.checked)}
              />
              Use Most Informative Basis
            </label>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={showDetails}
                onChange={(e) => setShowDetails(e.target.checked)}
              />
              Show details of argumentation
            </label>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={includeMiniTableaux}
                onChange={(e) => setIncludeMiniTableaux(e.target.checked)}
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
