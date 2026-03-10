# GLA Stochastic OT — VB6 v2.7 behavior changes

These reflect recent VB6 refactors that may have changed behavior. Verify Rust matches.

## Full history logging

Full history now logs per-constraint adjustment amounts and resulting values (not just final values).

**VB6 details**: In `RankingValueAdjustment` (StochasticOT branch, `boersma.frm:2794-2838`), the full history writes two columns per constraint per mismatch trial:
- Promoted: `\t{plasticity}\t{new_value}` (positive delta)
- Demoted: `\t-{plasticity}\t{new_value}` (negative delta)
- No change: `\t\t` (two empty tab-separated columns)

These are written from inside the per-constraint loop (with VB6 `;` to stay on the same line), then `GLACore` (`boersma.frm:2416-2428`) appends trial#, input, generated, heard, and final values.

**Rust status**: `gla.rs:873-884` only writes one column per constraint (the final value). Needs to add per-constraint delta+value for StochasticOT.

## A priori ranking enforcement refactored

Ranking value adjustment refactored — a priori ranking enforcement (`AdjustAPrioriRankings_Up/Down`) removed from within the per-constraint loop; verify Rust implementation matches new behavior.

**VB6 details**: Old code (commented out, `boersma.frm:2867-2889`) called `AdjustAPrioriRankings_Up` after each promoted constraint and `_Down` after each demoted constraint inside the per-constraint loop. New code only calls `AdjustAPrioriRankings_Up` once during initialization (`boersma.frm:1436`), with no enforcement during learning.

**Rust status**: `gla.rs:856-858` calls `adjust_apriori_up/down` after every trial's updates (outside constraint loop but inside trial loop). Needs removal — only keep the initialization call at `gla.rs:685-688`.

## RankingValueChange accumulator removed

`RankingValueChange` accumulator removed — adjustments now applied inline to `mRankingValue` directly.

**VB6 details**: Old code (commented out, `boersma.frm:2855-2882`) accumulated into `RankingValueChange` then applied. New code applies directly: `mRankingValue(i) = mRankingValue(i) + MyPlastFaith` etc. (`boersma.frm:2786-2831`). Note: `RankingValueChange` is still used in the MaxEnt branch (`boersma.frm:2758`).

**Rust status**: Already matches v2.7 — applies inline with `*rv += plast * promo_scale` etc. **No change needed.**
