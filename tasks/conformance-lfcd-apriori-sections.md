---
status: open
type: bug
priority: medium
depends_on: []
---

# Conformance: LFCD with a priori rankings has wrong section ordering

## Affected cases
- `TinyIllustrativeFile_lfcd_apriori`

## Description

Rust and VB6 output different section orderings for LFCD with a priori rankings:

- **Expected (VB6)**: Section 2 is `"Tableaux"`
- **Actual (Rust)**: Section 2 is `"A Priori Rankings"`

Line count also differs: expected 95, actual 107.

This differs from the RCD apriori case (which already has an `ignore_sections` workaround) — for LFCD the section ordering itself is wrong, not just the content of one section.

## Fix

Inspect `format_lfcd_output` when `apriori_file` is non-empty. Compare section order against `conformance/golden/TinyIllustrativeFile/lfcd_apriori.txt` and reorder sections to match VB6. Also investigate the line count discrepancy.

## Acceptance Criteria
- [ ] `TinyIllustrativeFile_lfcd_apriori` passes without a skip
- [ ] `make conformance-test` reports 0 failures for this case
