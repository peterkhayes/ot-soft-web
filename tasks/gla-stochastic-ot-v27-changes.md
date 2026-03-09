# GLA Stochastic OT — VB6 v2.7 behavior changes

These reflect recent VB6 refactors that may have changed behavior. Verify Rust matches.

## Full history logging

Full history now logs per-constraint adjustment amounts and resulting values (not just final values).

## A priori ranking enforcement refactored

Ranking value adjustment refactored — a priori ranking enforcement (`AdjustAPrioriRankings_Up/Down`) removed from within the per-constraint loop; verify Rust implementation matches new behavior.

## RankingValueChange accumulator removed

`RankingValueChange` accumulator removed — adjustments now applied inline to `mRankingValue` directly.
