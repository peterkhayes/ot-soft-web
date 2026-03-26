import type { FtResultData } from './types.ts'

interface FtResultsProps {
  result: FtResultData
}

function FtResults({ result }: FtResultsProps) {
  return (
    <div className="ft-results">
      {/* Summary */}
      <div className="ft-summary">
        <span className="ft-summary-count">{result.patternCount}</span> output{' '}
        {result.patternCount === 1 ? 'pattern' : 'patterns'} found
      </div>

      {result.patternCount > 0 && (
        <>
          {/* Patterns table */}
          <div className="ft-section">
            <div className="ft-section-header">Output Patterns</div>
            <p className="ft-section-note">
              Forms marked as winners in the input are marked with &#x203A;.
            </p>
            <div className="ft-patterns-scroll">
              <table className="ft-patterns-table">
                <thead>
                  <tr>
                    <th className="ft-input-col">Input</th>
                    {result.patterns.map((_, pi) => (
                      <th key={pi} className="ft-pattern-col">
                        Output #{pi + 1}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {result.formInputs.map((input, fi) => (
                    <tr key={fi}>
                      <td className="ft-input-cell">/{input}/</td>
                      {result.patterns.map((pattern, pi) => (
                        <td
                          key={pi}
                          className={`ft-pattern-cell${pattern.isWinner[fi] ? ' ft-pattern-cell--winner' : ''}`}
                        >
                          {pattern.isWinner[fi] && (
                            <span className="ft-winner-marker">&#x203A;</span>
                          )}
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
              {result.winners.map((row, fi) => (
                <div className="ft-winners-row" key={fi}>
                  <div className="ft-winners-input">/{row.formInput}/</div>
                  <div className="ft-winners-candidates">
                    {row.candidates.map((cand, ci) => (
                      <div className="ft-winners-candidate" key={ci}>
                        {cand.isWinner && <span className="ft-winner-marker">&#x203A;</span>}
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
            <p className="ft-section-note">The set of implications in the factorial typology.</p>

            {result.alwaysWinners.length > 0 && (
              <div className="ft-torder-note">
                <strong>No competition:</strong> the following forms have only one derivable output
                and are not listed in the t-order.
                <div className="ft-always-winners">
                  {result.alwaysWinners.map((aw, i) => (
                    <span className="ft-always-winner-item" key={i}>
                      /{aw.input}/ &rarr; [{aw.candidate}]
                    </span>
                  ))}
                </div>
              </div>
            )}

            {result.torder.length === 0 ? (
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
                  {result.torder.map((entry, i) => (
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

            {result.nonImplicators.length > 0 && (
              <div className="ft-torder-note ft-torder-note--below">
                <strong>Nothing implied by:</strong>
                <div className="ft-always-winners">
                  {result.nonImplicators.map((ni, i) => (
                    <span className="ft-always-winner-item" key={i}>
                      /{ni.input}/ &rarr; [{ni.candidate}]
                    </span>
                  ))}
                </div>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  )
}

export default FtResults
