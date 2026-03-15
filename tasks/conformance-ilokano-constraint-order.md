---
status: in_progress
type: bug
priority: high
depends_on: []
---

# Conformance: Ilokano constraint ordering within strata differs from VB6

## Affected cases
- `ilokano_rcd_defaults` ✓ FIXED
- `ilokano_rcd_no_fred` ✓ FIXED
- `ilokano_rcd_mib` ✓ FIXED
- `ilokano_bcd_defaults` ✓ FIXED
- `ilokano_bcd_specific` ← still failing
- `ilokano_lfcd_defaults` ✓ FIXED
- `ilokano_lfcd_mib` ✓ FIXED

## Description

All RCD/BCD/LFCD conformance cases for the Ilokano Hiatus Resolution dataset fail because the ordering of constraints within strata differs between VB6 and Rust.

## Fixes applied

### 1. VB6 unstable selection sort for constraint ordering (`sorted_constraint_indices`)
VB6's `SortTheConstraints` is an O(n²) unstable selection sort that orders constraints by stratum. Implemented `sorted_constraint_indices` (and helper `vb6_sort_constraint_slice`) to replicate this exactly in `rcd.rs`. Used in all text and HTML tableau output functions.

### 2. Candidate sorting by harmony (`sorted_candidate_indices`)
VB6's `SortTheCandidates` sorts rivals by harmony (lexicographic violation comparison through sorted constraints, winner always first). Implemented `sorted_candidate_indices` in `rcd.rs`.

### 3. Constraint necessity check — identical-pair step
VB6's `FindUnnecessaryConstraints` first checks if removing a constraint makes any winner-rival pair identical (same violations on all remaining constraints) before running FastRCD. Added this preliminary check to `is_constraint_necessary`.

### 4. Mini-tableau constraint sorting
VB6's `PrepareMiniTableaux` collects constraints in original input order then calls `SortTheConstraints`. Applied `vb6_sort_constraint_slice` to `mini.included_constraints` in `format_mini_tableau`.

## Remaining issue: `ilokano_bcd_specific`

**Symptom:** Expected `Max` at line 18 (stratum 2), actual has `Max-stem` at line 18, and actual has 1 extra line (158 vs 157).

**Root cause:** Rust's specific BCD correctly computes that Max-stem's violations are a subset of Max's violations (`Subset(Max-stem, Max) = True`), so it marks Max as "subsetted" (blocked by more-specific Max-stem). This excludes Max from stratum 2's rankable faithfulness set, pushing Max to stratum 3.

But VB6 specific BCD produces *identical output* to non-specific BCD for Ilokano — both put Max and Max-stem in stratum 2 together. This suggests VB6's subsetted logic doesn't exclude Max here, possibly because in the "no markedness released" case VB6 handles stratum membership differently, or the subset computation differs subtly.

**Investigation needed:**
- Trace VB6's `FindMinFaithSet` when `BestMarkCount = 0` — does it rank ALL faith (including subsetted) or only rankable?
- Check if VB6's `Subset(Max-stem, Max)` actually evaluates to `True` for this dataset.
- The VB6 code at `FindMinFaithSet` line ~520: when BestMarkCount=0, it sets BestSubset to all of RankableFaith (excludes subsetted) and ranks `For ConstraintIndex = 1 To SubsetSize`. Does `SubsetSize` equal `RFCount` here, or something else?

## Acceptance Criteria
- [ ] All 7 Ilokano conformance cases pass without a skip
- [ ] `make conformance-test` reports 0 failures for Ilokano cases
