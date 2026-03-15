---
status: open
type: testing
priority: medium
depends_on: []
---

# Collect golden files for AnttilaFinnishGenitivePlurals

## Description

A new input file has been added to `examples/`:
- `AnttilaFinnishGenitivePlurals.txt` — Finnish genitive plurals (frequency data)

Golden files need to be collected from VB6 OTSoft on Windows for deterministic conformance tests. Structural conformance tests for stochastic algorithms (GLA-SOT, GLA-MaxEnt, NHG) have already been added to `conformance.rs`.

## Golden files needed

### AnttilaFinnishGenitivePlurals.txt (deterministic only)
- `maxent_defaults.txt` — batch MaxEnt with default params (5 iterations, no prior)

## Steps

1. Use the conformance automation scripts (`conformance/automation/`) on the Windows laptop
2. Run VB6 OTSoft with the input file and parameter combination
3. Save output to `conformance/golden/AnttilaFinnishGenitivePlurals/`
4. Add manifest entry and verify with `make conformance-test`
