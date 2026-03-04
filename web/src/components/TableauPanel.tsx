import type { Candidate, Constraint, InputForm, Tableau } from '../../pkg/ot_soft.js'
import { AxisMode } from '../../pkg/ot_soft.js'

interface TableauPanelProps {
  tableau: Tableau
  axisMode?: AxisMode
}

/** VB6 heuristic: should this form use transposed layout in "switch where needed" mode? */
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

/**
 * Renders a tableau as per-form tables, mirroring VB6 behavior.
 *
 * VB6 always renders one table per form. The axis mode controls
 * whether each form's table is transposed (constraints as rows,
 * candidates as columns) or normal (constraints as columns,
 * candidates as rows).
 */
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

  // Precompute total constraint abbreviation length for "switch where needed" heuristic
  const totalConstraintLength = constraints.reduce((sum, c) => sum + c.abbrev.length + 1, 0)

  return (
    <div className="tableau-container">
      {forms.map((form, formIdx) => {
        const transposed =
          axisMode === AxisMode.SwitchAll ||
          (axisMode === AxisMode.SwitchWhereNeeded && shouldTranspose(form, totalConstraintLength))

        if (transposed) {
          return <TransposedFormTable key={formIdx} form={form} constraints={constraints} />
        }
        return <NormalFormTable key={formIdx} form={form} constraints={constraints} />
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

// ── Normal (non-transposed) per-form table ───────────────────────────────

/** A single form: constraints as columns, candidates as rows. */
function NormalFormTable({ form, constraints }: { form: InputForm; constraints: Constraint[] }) {
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
          <tr className="data-row" key={j}>
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

// ── Transposed per-form table ────────────────────────────────────────────

/** A single form: candidates as columns, constraints as rows. */
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
