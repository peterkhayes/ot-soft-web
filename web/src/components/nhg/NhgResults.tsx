import type { NhgResultState } from './types.ts'

interface NhgResultsProps {
  result: NhgResultState
}

function NhgResults({ result }: NhgResultsProps) {
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
        <div className="log-prob">Log likelihood of data: {result.logLikelihood.toFixed(4)}</div>
        {result.zeroPredictionWarning && (
          <div className="warning">
            Caution: at least one candidate with positive frequency was assigned zero probability;
            since zero has no log this was approximated as .001.
          </div>
        )}
      </div>

      <div className="maxent-tableaux">
        <h3 className="results-subheader">Matchup to Input Frequencies</h3>
        {result.forms.map((form, fi) => (
          <div className="maxent-form" key={fi}>
            <div className="form-label">
              /{form.input}/ <span className="form-freq">({form.totalFreq} cases)</span>
            </div>
            <table className="predictions-table">
              <thead>
                <tr>
                  <th></th>
                  <th className="pct-col">Count</th>
                  <th className="pct-col">Obs%</th>
                  <th className="pct-col">Gen</th>
                  <th className="pct-col">Gen%</th>
                </tr>
              </thead>
              <tbody>
                {form.candidates.map((cand, ci) => (
                  <tr key={ci} className={cand.obsPct > 0 ? 'winner-row' : ''}>
                    <td className="cand-form">
                      {cand.obsPct > 0 && <span className="winner-marker">&#x25B6;</span>}
                      {cand.form}
                    </td>
                    <td className="pct-col">{cand.frequency}</td>
                    <td className="pct-col">{cand.obsPct.toFixed(1)}%</td>
                    <td className="pct-col">{cand.genCount}</td>
                    <td className="pct-col">{cand.genPct.toFixed(1)}%</td>
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

export default NhgResults
