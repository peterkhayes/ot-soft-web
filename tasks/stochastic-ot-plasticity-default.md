---
status: done
type: bug
priority: medium
depends_on: []
---

# Stochastic OT: initial plasticity default mismatch

## Resolution

**No code change needed.** Our defaults already match the VB6 form defaults:

| Parameter | VB6 Form Default | Our Default |
|-----------|-----------------|-------------|
| GLA initial plasticity | 2 (`boersma.frm` line 87) | 2.0 |
| GLA final plasticity | .001 (`boersma.frm` line 78) | 0.001 |

The original bug report compared against values from `OTSoftRememberUserChoices.txt`
(0.01 / 0.0001), which is a saved user state file — not the application's built-in
defaults. Per `source/CLAUDE.md`, the `.frm` source files are the authoritative
source for VB6 defaults.

The VB6 also defines separate Mark/Faith defaults in `Module1.bas`:
- `DefaultUpperFaith = 2`, `DefaultLowerFaith = 0.002`
- `DefaultUpperMark = 0.2`, `DefaultLowerMark = 0.002`

These are used when custom M/F plasticity is enabled, not for the standard schedule.

The accuracy discrepancy between our output and BPH's output is likely because BPH
ran VB6 with custom plasticity settings (0.01/0.0001), not the VB6 defaults.

## Original Description

Our Stochastic OT uses an initial plasticity of 0.100, while the VB6 version uses 0.0100 — a 10x difference. This likely explains the significant accuracy discrepancy noted by BPH (our log likelihood: -845.87 vs VB6: -1007.64).

The higher plasticity may produce better results, but it doesn't match VB6 behavior. Need to determine the correct default and whether the VB6 value is intentional.

## Reference

- BPH comparison file: `ComparisonsForPeter/5. StochasticOT/BPH_AnttilaFinnishGenitivePluralsDraftOutput.txt`
- PKH comparison file: `ComparisonsForPeter/5. StochasticOT/PKH_AnttilaFinnishGenitivePluralsGLA-StochasticOTOutput.txt`
- VB6 GLA parameter schedule: Stage 1 PlastMark=0.0100, PlastFaith=0.0100

## Related

- See also: [Stochastic OT accuracy discrepancy](stochastic-ot-accuracy.md)

## Acceptance Criteria

- [x] Verify current initial plasticity default — matches VB6 form defaults (2.0 / 0.001)
- [x] Match VB6 default (0.01) or document why a different value is better — 0.01 was not the VB6 default; our 2.0 matches
