---
status: blocked
type: bug
priority: high
depends_on: []
---

# `rcd_apriori` — Missing "A Priori Rankings" section + wrong section numbers + incomplete mass deletion message + wrong constraint necessity ordering

**Symptom:** VB6 output has "2. A Priori Rankings" between Result and Tableaux, shifting all subsequent section numbers up by 1. Our output skips this section entirely.

This is a multi-part fix:

## A. Missing "A Priori Rankings" section — DONE

VB6 `Main.frm:7004-7040` (`PrintOutTheAprioriRankings`) prints a table showing which constraints dominate others. The table uses `s.PrintTable` (`s.bas:244`) with `CenterCells=True`, which calls `PrintTextTable` (`s.bas:373`) using per-column widths and `CenteredFillout` (`s.bas:430`) for alignment (centered with leftward error for odd spacing). The VB6 array stores data as `Table(col, row)` (reversed indices, per `s.bas:264`), with axis-switching: `gAPrioriRankingsTable(i,j)=True` (i dominates j) places "yes" at `Table(j+1, i+1)` = row i, col j in the output. Our `apriori.rs` stores `table[i][j]=true` meaning i dominates j, so display should use `apriori[row][col]`. Note: `gla.rs:458` already has an a priori table formatter but uses `apriori[col][row]` (inverted) — that may be a bug in GLA (no golden file to test against yet).

**Implementation:** Added the section to both `format_output_with_options` and `format_html_output_full` in `rcd.rs`. Exact column alignment doesn't matter because the conformance test normalizer collapses multiple spaces to one.

## B. Dynamic section numbering — DONE

VB6 uses a global auto-incrementing counter `gLevel1HeadingNumber` (`Module1.bas:175`). Our code hardcodes section numbers (1, 2, 3, 4, 5) in `rcd.rs` and `fred.rs:463`. When a priori is present, sections shift: 1=Result, 2=A Priori Rankings, 3=Tableaux, 4=Status, 5=FRed, 6=Mini-Tableaux. Need a mutable counter, and `fred.rs:format_section4` must accept the section number as a parameter.

**Implementation:** Replaced all hardcoded section numbers with a `let mut section = 0usize;` counter that increments before each heading. `fred.rs:format_section4` renamed to `format_section_fred(section_num)`. The `apriori` table is threaded through the format functions as a parameter:
- `format_output` → `format_output_with_options(..., apriori)`
- `format_html_output_with_options` → `format_html_output_full(..., apriori)`
- Callers in `lib.rs` updated for RCD, BCD (passes `&[]`), and LFCD.

## C. Mass deletion message — missing "individual but not mass" case — DONE

VB6 `Main.frm:6147-6163` has two messages gated on `NumberOfDeletableConstraints >= 2`: if mass deletion succeeds, print success message; otherwise, print "although the grammar will still work with the removal of ANY ONE... will NOT work if they are removed en masse." `NumberOfDeletableConstraints` counts constraints where `TrulyNeeded=False` (`Main.frm:6183-6188`). Our code (`rcd.rs:463-470`) only handles the success case and emits `\n\n` for the failure case — it doesn't print the "individual but not mass" message.

**Implementation:** Added the failure message for both text and HTML output. Also gated both messages on `num_deletable >= 2` to match VB6.

## D. Constraint necessity ordering — DONE

VB6 `Main.frm:6087-6141` outputs constraints grouped by category in three separate loops: first all Necessary constraints (in original order), then all "Not necessary (but included to show Faithfulness violations)" (in order), then all "Not necessary" (in order). Our code (`rcd.rs:447`) iterates by constraint index, which produces a different order when categories are interleaved.

**Implementation:** Changed both text and HTML output to iterate three times (one per `ConstraintNecessity` variant) instead of once.

## E. Apriori table must use abbreviations, not full names — DONE (discovered during implementation)

The example file `examples/TinyIllustrativeFile_apriori.txt` used full constraint names (`*No Onset`, `Max(t)`) but VB6's `ReadAPrioriRankingsAsTable` validates column/row labels against `Abbrev()` — the abbreviation array. Fixed the example file to use abbreviations (`*NoOns`, `Max`).

## F. Necessity/mass-deletion checks must pass a priori rankings through — DONE

`compute_constraint_necessity`, `is_constraint_necessary`, and `check_mass_deletion` all ran internal RCD with `&[]` (no a priori). VB6 uses `FastRCDWithAPrioriRankings` when the a priori menu option is checked. Fixed by threading `apriori: &[Vec<bool>]` through `compute_extra_analyses` → `compute_constraint_necessity` → `is_constraint_necessary`, and through `check_mass_deletion`. All callers updated (RCD, BCD passes `&[]`, LFCD passes its apriori).

## G. Necessity results still differ from VB6 — OPEN BUG

Even after threading a priori through, our code computes `*NoOns` and `*Coda` as Necessary, while VB6 marks ALL four constraints as "Not necessary". The root cause:

When `*NoOns` violations are zeroed and RCD is run with a priori `*NoOns >> Max`:
- Pair `/a/` (winner `?a` vs loser `a`) becomes `(0,0,0,1)` vs `(0,0,0,0)` — only `Dep` distinguishes them, and `Dep` prefers the **loser** (loser has fewer violations). No constraint prefers the winner. RCD fails because this pair can never be eliminated.

VB6 uses `FastRCDWithAPrioriRankings` (Main.frm:6318-6477) for this check instead of full RCD. There's a **suspected VB6 bug at Main.frm:5974**: when a priori rankings are enabled, the result is written to `TrulyNeeded(OuterConstraintIndex)` instead of `TrulyNeeded(ConstraintUnderAssessment)`. Since `OuterConstraintIndex` is a different loop variable (used in mass-deletion checking, not in the per-constraint loop), this likely writes to the **wrong index**, causing all constraints to appear unnecessary when a priori rankings are active.

**Next step:** Verify this VB6 bug theory. If confirmed, we need to reproduce the buggy behavior to match the golden file. Note: the `FastRCDWithAPrioriRankings` logic itself appears identical to our `run_rcd_internal` with a priori — the only relevant difference would be this variable name bug.
