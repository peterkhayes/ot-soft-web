import type { Candidate, Constraint, InputForm, Tableau } from '../../pkg/ot_soft.js'
import { AxisMode } from '../../pkg/ot_soft.js'

interface TableauPanelProps {
  tableau: Tableau
  axisMode?: AxisMode
}

/** VB6 heuristic: should this form use transposed layout in "some" mode? */
function shouldTranspose(form: InputForm, totalConstraintLength: number): boolean {
  if (totalConstraintLength <= 75) return false
  const candidateCount = form.candidate_count()
  let totalCandidateLength = 0
  for (let j = 0; j < candidateCount; j++) {
    const candidate = form.get_candidate(j)!
    totalCandidateLength += candidate.form.length + 2
  }
  return totalCandidateLength < totalConstraintLength + 5
}

function TableauPanel({ tableau, axisMode = AxisMode.SwitchAll }: TableauPanelProps) {
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

  // Precompute total constraint abbreviation length for "some" mode
  const totalConstraintLength = constraints.reduce((sum, c) => sum + c.abbrev.length + 1, 0)

  if (axisMode === AxisMode.NeverSwitch) {
    return <NormalTableau constraints={constraints} forms={forms} />
  }

  if (axisMode === AxisMode.SwitchAll) {
    return <TransposedTableau constraints={constraints} forms={forms} />
  }

  // SwitchWhereNeeded: render each form individually, choosing layout per the VB6 heuristic
  return (
    <div className="tableau-container">
      {forms.map((form, formIdx) => {
        const transposed = shouldTranspose(form, totalConstraintLength)
        if (transposed) {
          return <TransposedFormTable key={formIdx} form={form} constraints={constraints} />
        }
        return (
          <NormalFormTable key={formIdx} form={form} formIdx={formIdx} constraints={constraints} />
        )
      })}
    </div>
  )
}

/** Collect candidates from a form into an array. */
function getCandidates(form: InputForm): Candidate[] {
  const candidates: Candidate[] = []
  const count = form.candidate_count()
  for (let j = 0; j < count; j++) {
    candidates.push(form.get_candidate(j)!)
  }
  return candidates
}

/** Format a violation value for display. */
function formatViolation(violation: number | undefined): string {
  return violation === 0 || violation === undefined ? '' : String(violation)
}

// ── Normal (traditional) layout ──────────────────────────────────────────

function NormalTableau({ constraints, forms }: { constraints: Constraint[]; forms: InputForm[] }) {
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
            const candidates = getCandidates(form)
            return candidates.map((candidate, j) => (
              <tr className="data-row" key={`${formIdx}-${j}`}>
                <td className="input-cell">{j === 0 ? form.input : ''}</td>
                <td className="candidate-cell">{candidate.form}</td>
                <td className="frequency-cell">{candidate.frequency}</td>
                {constraints.map((_, k) => (
                  <td className="violation-cell" key={k}>
                    {formatViolation(candidate.get_violation(k))}
                  </td>
                ))}
              </tr>
            ))
          })}
        </tbody>
      </table>
    </div>
  )
}

/** A single form rendered in normal layout (for "some" mode). */
function NormalFormTable({
  form,
  formIdx,
  constraints,
}: {
  form: InputForm
  formIdx: number
  constraints: Constraint[]
}) {
  const candidates = getCandidates(form)
  return (
    <table className="tableau-table tableau-table--per-form">
      <thead>
        <tr className="subheader-row">
          <th>/{form.input}/</th>
          {constraints.map((c, i) => (
            <th key={i}>{c.abbrev}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {candidates.map((candidate, j) => (
          <tr className="data-row" key={`${formIdx}-${j}`}>
            <td className="candidate-cell">
              {candidate.frequency > 0 ? '\u261E ' : '\u00A0\u00A0\u00A0'}
              {candidate.form}
            </td>
            {constraints.map((_, k) => (
              <td className="violation-cell" key={k}>
                {formatViolation(candidate.get_violation(k))}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

// ── Transposed layout ────────────────────────────────────────────────────

function TransposedTableau({
  constraints,
  forms,
}: {
  constraints: Constraint[]
  forms: InputForm[]
}) {
  return (
    <div className="tableau-container">
      {forms.map((form, formIdx) => (
        <TransposedFormTable key={formIdx} form={form} constraints={constraints} />
      ))}
    </div>
  )
}

/** A single form rendered in transposed layout: candidates as columns, constraints as rows. */
function TransposedFormTable({
  form,
  constraints,
}: {
  form: InputForm
  constraints: Constraint[]
}) {
  const candidates = getCandidates(form)

  return (
    <table className="tableau-table tableau-table--transposed">
      <thead>
        <tr className="subheader-row">
          <th>/{form.input}/</th>
          {candidates.map((candidate, j) => (
            <th key={j} className="candidate-cell">
              {candidate.frequency > 0 ? '\u261E ' : ''}
              {candidate.form}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {constraints.map((constraint, k) => (
          <tr className="data-row" key={k}>
            <td className="constraint-label-cell">{constraint.abbrev}</td>
            {candidates.map((candidate, j) => (
              <td className="violation-cell" key={j}>
                {formatViolation(candidate.get_violation(k))}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

export default TableauPanel
