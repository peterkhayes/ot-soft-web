import type { MaxEntResultState } from './types.ts'

interface MaxEntResultsProps {
  result: MaxEntResultState
  constraintAbbrevs: string[]
}

function MaxEntResults({ result, constraintAbbrevs }: MaxEntResultsProps) {
  return (
    <div className="maxent-results">
      <div className="maxent-weights">
        <h3 className="results-subheader">Constraint Weights</h3>
        <table className="weights-table">
          <thead>
            <tr>
              <th>Constraint</th>
              <th className="weight-col">Weight</th>
            </tr>
          </thead>
          <tbody>
            {result.weights.map((w, i) => (
              <tr key={i}>
                <td>
                  <span className="abbrev">{w.abbrev}</span>
                  <span className="full-name"> ({w.fullName})</span>
                </td>
                <td className="weight-col weight-value">{w.weight.toFixed(3)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="log-prob">Log probability of data: {result.logProb.toFixed(4)}</div>
      </div>

      <div className="maxent-tableaux">
        <h3 className="results-subheader">Predicted Probabilities</h3>
        {result.forms.map((form, fi) => (
          <div className="maxent-form" key={fi}>
            <div className="form-label">/{form.input}/</div>
            <table className="predictions-table">
              <thead>
                <tr>
                  <th></th>
                  <th className="pct-col">Obs%</th>
                  <th className="pct-col">Pred%</th>
                  {constraintAbbrevs.map((abbrev, ci) => (
                    <th key={ci} className="viol-col">
                      {abbrev}
                    </th>
                  ))}
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
                    <td className="pct-col">{cand.predPct.toFixed(1)}%</td>
                    {cand.violations.map((v, vi) => (
                      <td key={vi} className="viol-col">
                        {v > 0 ? v : ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ))}
      </div>
    </div>
  )
}

export default MaxEntResults
