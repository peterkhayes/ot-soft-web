---
status: open
type: bug
priority: medium
depends_on: []
---

# Conformance: MaxEnt output differs from VB6 (title typo + structural differences)

## Affected cases
- `TinyIllustrativeFile_maxent_defaults`

## Description

Two known differences:

1. **Title typo**: VB6 outputs `"Result of Applying Maximum Entropy..."` (singular), Rust outputs `"Results of Applying Maximum Entropy..."` (plural).

2. **Line count mismatch**: Expected 72 lines, actual 44 lines. This suggests significant structural differences beyond just the title — sections may be missing or formatted differently.

## Fix

Compare `conformance/golden/TinyIllustrativeFile/maxent_defaults.txt` line-by-line against `format_maxent_output` output to identify all differences:
- Fix the title to match VB6 (`Result of` not `Results of`)
- Identify and fix any missing sections or structural mismatches

## Acceptance Criteria
- [ ] `TinyIllustrativeFile_maxent_defaults` passes without a skip
- [ ] `make conformance-test` reports 0 failures for this case
