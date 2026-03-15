---
status: open
type: bug
priority: medium
depends_on: []
---

# Stochastic OT: initial plasticity default mismatch

## Description

Our Stochastic OT uses an initial plasticity of 0.100, while the VB6 version uses 0.0100 — a 10x difference. This likely explains the significant accuracy discrepancy noted by BPH (our log likelihood: -845.87 vs VB6: -1007.64).

The higher plasticity may produce better results, but it doesn't match VB6 behavior. Need to determine the correct default and whether the VB6 value is intentional.

**Note:** This may have already been fixed since the comparison was made (2026-03-07). Verify current default before changing.

## Reference

- BPH comparison file: `ComparisonsForPeter/5. StochasticOT/BPH_AnttilaFinnishGenitivePluralsDraftOutput.txt`
- PKH comparison file: `ComparisonsForPeter/5. StochasticOT/PKH_AnttilaFinnishGenitivePluralsGLA-StochasticOTOutput.txt`
- VB6 GLA parameter schedule: Stage 1 PlastMark=0.0100, PlastFaith=0.0100

## Related

- See also: [Stochastic OT accuracy discrepancy](stochastic-ot-accuracy.md)

## Acceptance Criteria

- [ ] Verify current initial plasticity default
- [ ] Match VB6 default (0.01) or document why a different value is better
