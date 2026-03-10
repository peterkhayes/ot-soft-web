# OTSoft: Task Tracker

## Bugs

### Conformance test failures

- [x] [`bcd_specific` — wrong title](tasks/conformance-bcd-specific.md)
- [x] [`rcd_no_fred` — extra trailing blank line](tasks/conformance-rcd-no-fred.md)
- [ ] [`rcd_apriori` — missing section, numbering, mass deletion, ordering](tasks/conformance-rcd-apriori.md) *(partially done: A–F complete, G open — suspected VB6 bug in necessity with a priori)*

### Other bugs

- [x] NHG: fix zero-prediction warning text (missing the word "frequency")
- [x] NHG: suppress infinite-loop MsgBox for 100 consecutive ties (VB6 now silently handles this)

### VB6 v2.7 behavior changes

- [ ] [GLA Stochastic OT — verify Rust matches v2.7 refactors](tasks/gla-stochastic-ot-v27-changes.md)

---

## UX Polish

- [ ] Add "Done" indicator when an algorithm finishes
- [ ] Progress reporting for long-running jobs (requires WASM progress callbacks)
- [x] Add sorted/unsorted tableaux toggle (VB6 has this; web always shows unsorted)
- [ ] Make MaxEnt constraint sorting by weight optional (add toggle, default on)
- [x] NHG: show actual counts alongside percentages
- [ ] Add "Diagnostics if ranking fails" option
- [ ] Add HTML output options for shading customization

### Investigation needed

- [ ] Investigate Stochastic OT output display differences vs VB6
- [ ] Investigate MaxEnt control parameter differences vs VB6

---

## New Features

- [ ] Excel file parsing (`.xlsx`)
- [ ] Praat export (`.OTGrammar`, `.PairDistribution`)
- [ ] R export (logistic regression format)
- [ ] HowIRanked log file

---

## Testing

- [ ] Collect remaining VB6 golden files (24 of 31 cases currently skip)
- [ ] Edge case examples (ranking failures, ties, large tableaux)
- [ ] GLA/Stochastic OT conformance cases (structural assertions only)

---

## Completed

### Input Parsing
- [x] Tab-delimited text file parsing
- [x] Negative violation counts
- [x] Abbreviation row auto-detection
- [x] A priori rankings file parsing

### Categorical Ranking Algorithms
- [x] RCD, BCD, LFCD with all options
- [x] A priori ranking enforcement
- [x] FRed (Skeletal Basis + Most Informative Basis)
- [x] FRed integrated into RCD/BCD/LFCD output
- [x] Hasse diagram DOT generation (FRed + GLA)

### Probabilistic Learning Algorithms
- [x] GLA (Stochastic OT + MaxEnt modes)
- [x] Batch Maximum Entropy (GIS)
- [x] Noisy Harmonic Grammar (8 noise variants)
- [x] Learning schedule, Magri update, Gaussian prior
- [x] Multiple runs, all history file formats

### Factorial Typology
- [x] Core computation, FastRCD, T-order
- [x] All output formats (FTSum, CompactSum, full listing)

### Output & Web
- [x] Text + HTML output formatting
- [x] Full web interface with all controls
- [x] Conformance test infrastructure
- [x] Web flow tests
