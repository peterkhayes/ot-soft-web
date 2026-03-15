---
status: open
type: testing
priority: high
depends_on: []
---

# Collect remaining VB6 golden files

## Description

5 of 31 conformance test cases skip at runtime because golden files are missing.
All 5 are MaxEnt cases that need to be run in VB6 OTSoft to capture output.

## Missing golden files

| Case ID | Golden file |
|---------|-------------|
| `TinyIllustrativeFile_maxent_prior` | `conformance/golden/TinyIllustrativeFile/maxent_prior.txt` |
| `TinyIllustrativeFile_maxent_sigma10` | `conformance/golden/TinyIllustrativeFile/maxent_sigma10.txt` |
| `ilokano_maxent_defaults` | `conformance/golden/IlokanoHiatusResolution/maxent_defaults.txt` |
| `ilokano_maxent_prior` | `conformance/golden/IlokanoHiatusResolution/maxent_prior.txt` |
| `ilokano_maxent_sigma10` | `conformance/golden/IlokanoHiatusResolution/maxent_sigma10.txt` |

Note: An additional 12 cases have explicit `skip` fields for known divergences (not missing files):
- `TinyIllustrativeFile_maxent_defaults` — output structure mismatch (see `tasks/conformance-maxent-output.md`)
- 7× Ilokano RCD/BCD/LFCD variants — constraint ordering within strata (see `tasks/conformance-ilokano-constraint-order.md`)
- 4× HTML cases — VB6 HTML lacks `<table>` wrappers (see `tasks/conformance-html-table-extraction.md`)

## Acceptance Criteria
- [ ] All 5 missing MaxEnt golden files collected from VB6
- [ ] No cases skip due to missing golden files
