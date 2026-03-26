import HasseDiagram from '../HasseDiagram.tsx'
import type { RcdResultState } from './types.ts'

interface RcdResultsProps {
  result: RcdResultState
  inputFilename: string | null
}

function RcdResults({ result, inputFilename }: RcdResultsProps) {
  return (
    <div className="rcd-results">
      <div className={`rcd-status ${result.success ? 'success' : 'failure'}`}>
        {result.success
          ? 'A ranking was found that generates the correct outputs'
          : 'Failed to find a valid ranking'}
      </div>

      {result.tieWarning && (
        <div className="rcd-status warning">
          Caution: BCD selected arbitrarily among tied faithfulness constraint subsets. Try changing
          the order of faithfulness constraints in the input file to see whether this results in a
          different ranking.
        </div>
      )}

      {result.strata.map((stratum, s) => (
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
      {result.hasseDot && (
        <HasseDiagram
          dotString={result.hasseDot}
          downloadName={`${inputFilename ? inputFilename.replace(/\.[^.]+$/, '') : 'tableau'}Hasse`}
        />
      )}
    </div>
  )
}

export default RcdResults
