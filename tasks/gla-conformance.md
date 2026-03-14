---
status: done
type: testing
priority: medium
depends_on: []
---

# GLA/Stochastic OT conformance cases

## Description
GLA and Stochastic OT are non-deterministic, so exact output matching isn't possible. Create conformance tests using structural assertions (section headings present, value ranges, format correctness).

## Acceptance Criteria
- [x] Structural conformance tests for GLA output
- [x] Structural conformance tests for Stochastic OT output

## Implementation

Two new test functions added to `rust/tests/conformance.rs`:

**`gla_structural_conformance`** — tests both example files in both modes:
- Asserts expected section headings are present in formatted output
- Verifies SOT-only sections (`Testing the Grammar: Details`, `Ranking Value to Ranking Probability Conversion`) are absent from MaxEnt output and vice versa
- Runs a short 500-cycle training run for speed

**`gla_convergence_tiny_illustrative`** — convergence correctness test for TinyIllustrativeFile:
- SOT: asserts ranking values satisfy the known RCD stratum ordering (*NoOns and *Coda strictly above Max and Dep) and that predicted winner probabilities exceed 0.5 for each form
- MaxEnt: asserts weight ordering satisfies the grammar's correctness conditions (`w_Dep < w_*NoOns`, `w_Max < w_*Coda`) derived from the training data

**Bonus: MaxEnt update sign bug fixed** — conformance testing revealed that the MaxEnt GLA update direction was inverted (using `gen_v - obs_v` instead of `obs_v - gen_v`, causing faithfulness constraints to be promoted instead of demoted). Fixed to match VB6 `RankingValueAdjustment`.
