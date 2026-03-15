---
status: open
type: bug
priority: medium
depends_on: []
---

# Stochastic OT: investigate accuracy discrepancy with VB6

## Description

The original author reports that our Stochastic OT produces different (potentially better) results than the VB6 version. This could indicate:

1. A bug in our implementation that happens to produce plausible-looking but incorrect output
2. A bug in the VB6 implementation that we've inadvertently fixed
3. A difference in algorithm parameters, random seed handling, or convergence criteria

The author also notes differences in output display format ("distinct, subset output display").

## Investigation Plan

- Run both versions on the same input and compare outputs side-by-side
- Check GLA learning rate schedule, noise distribution, and plasticity parameters against VB6
- Compare random number generation approach (VB6 uses `Rnd`, we use `rand` crate)
- Verify the evaluation/sampling loop matches VB6's `GenerateSOTOutput` logic
- Check if the number of test trials or averaging method differs
- Review output formatting differences noted by the author

## Acceptance Criteria

- [ ] Root cause of the discrepancy identified
- [ ] Either fix our implementation or document the VB6 bug
- [ ] Output display format reviewed and aligned where appropriate
