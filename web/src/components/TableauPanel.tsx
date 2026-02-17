import type { Tableau, Constraint, InputForm } from '../../pkg/ot_soft.js'

interface TableauPanelProps {
  tableau: Tableau
}

function TableauPanel({ tableau }: TableauPanelProps) {
  const constraintCount = tableau.constraint_count()
  const formCount = tableau.form_count()

  const constraints: Constraint[] = []
  for (let i = 0; i < constraintCount; i++) {
    constraints.push(tableau.get_constraint(i)!)
  }

  const forms: InputForm[] = []
  for (let i = 0; i < formCount; i++) {
    forms.push(tableau.get_form(i)!)
  }

  return (
    <div className="tableau-container">
      <table className="tableau-table">
        <thead>
          <tr className="header-row">
            <th></th>
            <th></th>
            <th></th>
            {constraints.map((c, i) => (
              <th key={i}>{c.full_name}</th>
            ))}
          </tr>
          <tr className="subheader-row">
            <th>Input</th>
            <th>Candidate</th>
            <th>Freq</th>
            {constraints.map((c, i) => (
              <th key={i}>{c.abbrev}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {forms.flatMap((form, formIdx) => {
            const candidateCount = form.candidate_count()
            const rows = []
            for (let j = 0; j < candidateCount; j++) {
              const candidate = form.get_candidate(j)!
              rows.push(
                <tr className="data-row" key={`${formIdx}-${j}`}>
                  <td className="input-cell">
                    {j === 0 ? form.input : ''}
                  </td>
                  <td className="candidate-cell">{candidate.form}</td>
                  <td className="frequency-cell">{candidate.frequency}</td>
                  {constraints.map((_, k) => {
                    const violation = candidate.get_violation(k)
                    const violationStr = (violation === 0 || violation === undefined) ? '' : String(violation)
                    return (
                      <td className="violation-cell" key={k}>{violationStr}</td>
                    )
                  })}
                </tr>
              )
            }
            return rows
          })}
        </tbody>
      </table>
    </div>
  )
}

export default TableauPanel
