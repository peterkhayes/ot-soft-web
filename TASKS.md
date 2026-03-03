# OTSoft Port: Task Tracker

This file tracks the status of porting features from the VB6 source to Rust/Wasm + Web.

## Legend

- [x] Complete
- [~] Partially complete
- [ ] Not started

---

## Input Parsing

- [x] Tab-delimited text file parsing (`.txt`)
- [ ] Excel file parsing (`.xlsx`)
- [ ] Legacy Ranker format (`.in`)
- [ ] Structural descriptions (automatic violation computation)
- [x] A priori rankings file parsing
- [ ] Natural class file parsing

## Categorical Ranking Algorithms

- [x] Recursive Constraint Demotion (RCD)
- [x] RCD ranking arguments (FRed Skeletal Basis)
- [x] RCD mini-tableaux generation
- [x] RCD constraint necessity analysis
- [x] Biased Constraint Demotion (BCD)
- [x] BCD specificity option (`mnuSpecificBCD`)
- [x] Low Faithfulness Constraint Demotion (LFCD)
- [x] A priori ranking enforcement in RCD/BCD/LFCD

## Ranking Argumentation (FRed)

- [x] FRed — full ERC fusion algebra and recursive basis computation
- [x] Skeletal Basis mode (default)
- [x] Most Informative Basis mode
- [x] FRed integrated into RCD/BCD/LFCD Section 4 output
- [x] Standalone `run_fred` and `format_fred_output` WASM exports
- [x] Detailed argumentation output (verbose recursion tree)
- [x] Hasse diagram DOT generation in Rust (`rust/src/hasse.rs`, WASM export `fred_hasse_dot`)
- [x] Hasse diagram DOT generation for GLA/Stochastic OT (`gla_hasse_dot`, probability-labeled edges)

## Probabilistic Learning Algorithms

- [x] Gradual Learning Algorithm (GLA) — Stochastic OT mode
- [x] Gradual Learning Algorithm (GLA) — online MaxEnt mode
- [x] Batch Maximum Entropy (GIS optimizer)
- [x] Noisy Harmonic Grammar (NHG) — 8 noise variants
- [x] Learning schedule (multi-stage plasticity interpolation)
- [x] Custom learning schedule from file
- [x] Magri update rule (Stochastic OT)
- [x] Gaussian prior (MaxEnt)
- [x] Exact proportions data presentation
- [x] Multiple runs with collated results
- [x] Pairwise ranking probabilities (used internally by GLA Hasse diagram)
- [x] History file output (simple weight/ranking value history per iteration)
- [ ] Full annotated history (FullHistory.txt — trial + input/generated/heard + values)
- [ ] Candidate probability history (HistoryOfCandidateProbabilities.txt — GLA MaxEnt mode only)
- [ ] Output probability history (HistoryOfOutputProbabilities.txt — Batch MaxEnt)

## Factorial Typology

- [x] Core factorial typology computation
- [x] FastRCD (streamlined RCD for derivability testing)
- [x] T-order computation (typological implications)
- [x] FTSum output file
- [x] CompactSum output file
- [x] Full listing with grammars per output pattern

## Output Formatting

- [x] Formatted text output (matching VB6 draft output style)
- [x] Tableaux with fatal violation markers
- [x] Downloadable results file
- [x] HTML tableaux with configurable shading
- [x] Sorted input file (constraints/candidates reordered by rank)
- [ ] Praat export (`.OTGrammar`, `.PairDistribution`)
- [ ] R export (logistic regression format)
- [ ] HowIRanked log file

## Web Interface

- [x] File upload (drag-and-drop and file picker)
- [x] Load example tableau
- [x] Tableau display (interactive table)
- [x] RCD results display (stratified constraints)
- [x] Download results button
- [x] Framework selection (Classical OT / MaxEnt / NHG / Stochastic OT)
- [x] Algorithm variant selection (RCD / BCD / LFCD)
- [x] Ranking argumentation options (MIB, details, minitableaux)
- [x] Factorial typology button and options
- [x] A priori rankings controls
- [x] Parameter inputs for probabilistic algorithms
- [x] Progress indicator for long computations
- [ ] Tableau axis switching (for crowded tableaux)
- [x] Hasse diagram viewer (`HasseDiagram.tsx` component using `@viz-js/viz`, SVG + PNG export)
- [x] Hasse diagram integrated into RcdPanel (FRed Hasse, shown when FRed is enabled)
- [x] Hasse diagram integrated into GlaPanel (Stochastic OT mode only)
- [x] Settings persistence

## Testing & Examples

- [x] Conformance test infrastructure (`conformance/manifest.json`, `rust/tests/conformance.rs`) — compares Rust output against VB6 golden files with skip-on-missing
- [x] Tiny example conformance cases (RCD, BCD, LFCD, MaxEnt, FT — text + HTML)
- [x] Ilokano Hiatus Resolution conformance cases (RCD, BCD, LFCD, MaxEnt, FT — text + HTML)
- [x] HTML conformance: semantic cell-grid comparison (tolerates structural HTML differences between VB6 and Rust)
- [x] Web flow tests (`web/tests/flows/`) for RCD, MaxEnt, GLA, NHG, Factorial Typology, Hasse diagrams
- [ ] Collect remaining VB6 golden files (most conformance cases skip due to missing golden files)
- [ ] Edge case examples (ranking failures, ties, large tableaux)
- [ ] GLA/Stochastic OT conformance cases (stochastic — structural assertions only)

---

## Suggested Next Tasks

Roughly ordered by value and dependency:

1. **Collect VB6 golden files** — Run VB6 OTSoft on Windows following `conformance/CHECKLIST.md` to populate the missing golden files. Most conformance test cases currently skip.

2. **Extended history outputs** — Full annotated history (FullHistory.txt), candidate probability history, and output probability history files.

3. **Exact proportions data presentation** — Display data as exact proportions rather than frequency counts.

6. **Praat export** — Generate `.OTGrammar` and `.PairDistribution` files for use in Praat.

7. **Excel file parsing** — Support `.xlsx` input files in addition to tab-delimited `.txt`.
