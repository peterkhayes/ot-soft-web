import HasseDiagram from '../HasseDiagram.tsx'
import type { GlaResultState } from './types.ts'

interface GlaResultsProps {
  result: GlaResultState
  inputFilename: string | null
}

function GlaResults({ result, inputFilename }: GlaResultsProps) {
  const valueLabel = result.maxentMode ? 'Weight' : 'Ranking Value'

  return (
    <div className="maxent-results">
      <div className="maxent-weights">
        <h3 className="results-subheader">Constraint {valueLabel}s</h3>
        <table className="weights-table">
          <thead>
            <tr>
              <th>Constraint</th>
              <th className="weight-col">{valueLabel}</th>
            </tr>
          </thead>
          <tbody>
            {result.values.map((v, i) => (
              <tr key={i}>
                <td>
                  <span className="abbrev">{v.abbrev}</span>
                  <span className="full-name"> ({v.fullName})</span>
                </td>
                <td className="weight-col weight-value">{v.value.toFixed(3)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="log-prob">Log likelihood of data: {result.logLikelihood.toFixed(4)}</div>
      </div>

      <div className="maxent-tableaux">
        <h3 className="results-subheader">Matchup to Input Frequencies</h3>
        {result.forms.map((form, fi) => (
          <div className="maxent-form" key={fi}>
            <div className="form-label">/{form.input}/</div>
            <table className="predictions-table">
              <thead>
                <tr>
                  <th></th>
                  <th className="pct-col">Obs%</th>
                  <th className="pct-col">{result.maxentMode ? 'Prob%' : 'Gen%'}</th>
                </tr>
              </thead>
              <tbody>
                {form.candidates.map((cand, ci) => (
                  <tr key={ci} className={cand.obsPct > 0 ? 'winner-row' : ''}>
                    <td className="cand-form">
                      {cand.obsPct > 0 && <span className="winner-marker">&#x25B6;</span>}
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

      {result.hasseDot && (
        <HasseDiagram
          dotString={result.hasseDot}
          downloadName={`${inputFilename ? inputFilename.replace(/\.[^.]+$/, '') : 'tableau'}Hasse`}
        />
      )}

      {result.pairwiseData && (
        <div className="pairwise-probabilities">
          <h3 className="results-subheader">Pairwise Ranking Probabilities</h3>
          <p className="pairwise-note">
            The computed ranking values imply the pairwise ranking probabilities given below. In the
            table, the probability given is that of the row constraint outranking the column
            constraint.
          </p>
          <div className="table-scroll-wrapper">
            <table className="pairwise-table">
              <thead>
                <tr>
                  <th></th>
                  {result.pairwiseData.headers.slice(1).map((h) => (
                    <th key={h}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {result.pairwiseData.matrix.map((row, i) => (
                  <tr key={i}>
                    <th>{result.pairwiseData!.headers[i]}</th>
                    {row.map((cell, j) => (
                      <td key={j} className={cell === '' ? 'pairwise-empty' : ''}>
                        {cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <div className="maxent-weights">
        <h3 className="results-subheader">Active Constraints</h3>
        <p className="active-constraints-note">
          A constraint is active if it causes the winning candidate to defeat a rival in at least
          one competition.
        </p>
        <table className="weights-table">
          <thead>
            <tr>
              <th>Constraint</th>
              <th className="weight-col">Status</th>
            </tr>
          </thead>
          <tbody>
            {[...result.values]
              .sort((a, b) => b.value - a.value)
              .map((v, i) => (
                <tr key={i}>
                  <td>
                    <span className="abbrev">{v.abbrev}</span>
                    <span className="full-name"> ({v.fullName})</span>
                  </td>
                  <td className="weight-col">{v.active ? 'Active' : 'Inactive'}</td>
                </tr>
              ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

export default GlaResults
