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
- [ ] Hasse diagram DOT generation for GLA/Stochastic OT (`gla_hasse_dot`, probability-labeled edges)

## Probabilistic Learning Algorithms

- [x] Gradual Learning Algorithm (GLA) — Stochastic OT mode
- [x] Gradual Learning Algorithm (GLA) — online MaxEnt mode
- [x] Batch Maximum Entropy (GIS optimizer)
- [x] Noisy Harmonic Grammar (NHG) — 8 noise variants
- [ ] Learning schedule (multi-stage plasticity interpolation)
- [ ] Custom learning schedule from file
- [ ] Magri update rule (Stochastic OT)
- [x] Gaussian prior (MaxEnt)
- [ ] Exact proportions data presentation
- [ ] Multiple runs with collated results
- [ ] Pairwise ranking probabilities (used internally by GLA Hasse diagram)
- [ ] History file output (weights/ranking values over time)

## Factorial Typology

- [x] Core factorial typology computation
- [x] FastRCD (streamlined RCD for derivability testing)
- [x] T-order computation (typological implications)
- [ ] FTSum output file
- [ ] CompactSum output file
- [x] Full listing with grammars per output pattern

## Output Formatting

- [x] Formatted text output (matching VB6 draft output style)
- [x] Tableaux with fatal violation markers
- [x] Downloadable results file
- [ ] HTML tableaux with configurable shading
- [ ] Sorted input file (constraints/candidates reordered by rank)
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
- [ ] Hasse diagram integrated into GlaPanel (Stochastic OT mode only)
- [x] Settings persistence

## Testing & Examples

- [x] Tiny example (`examples/tiny/`) with expected RCD output
- [ ] BCD example with expected output
- [ ] LFCD example with expected output
- [x] FRed example with expected output (tiny example verified)
- [ ] GLA/MaxEnt example with expected output
- [~] NHG example with expected output
- [~] Factorial typology example with expected output
- [ ] Edge case examples (ranking failures, ties, large tableaux)

---

## Suggested Next Tasks

Roughly ordered by value and dependency:

1. **FT full listing** — For each output pattern in the factorial typology, run RCD and show the ranking that produces it (plus optional tableaux and FRed arguments). Depends on factorial typology core (done).

2. **FTSum / CompactSum output** — Generate the tab-delimited `FTSum.txt` (one row per pattern) and `CompactSum.txt` (collapsed by surface outputs, deduplicated) files.

3. **Hasse diagrams** — Generate and display visual ranking hierarchy from FRed and GLA output.
   See `source/CHARTS.md` for full requirements and technology choice.
   Technology: **`@viz-js/viz`** (GraphViz compiled to WASM; renders DOT → SVG in the browser).
   Implementation steps:
   - **Rust**: Add `rust/src/hasse.rs` with two WASM exports:
     - `fred_hasse_dot(tableau_text, apriori_text, use_mib)` — DOT string from FRed Valhalla ERCs
       (ports `Fred.bas:PrepareHasseDiagram`; certain edges solid, disjunctive edges dotted with "or" label)
     - `gla_hasse_dot(tableau_text, ranking_values)` — DOT string from GLA ranking values
       (ports `boersma.frm:PrintPairwiseRankingProbabilities`; edges labeled with probability, dotted if P < 0.95)
   - **Web**: Add `@viz-js/viz` npm dependency; create `web/src/components/HasseDiagram.tsx`
     (lazy-loads viz.js WASM, renders DOT → inline SVG, provides SVG and PNG download buttons)
   - **Web**: Wire `HasseDiagram` into `RcdPanel.tsx` and `GlaPanel.tsx`
   - **Tests**: `web/tests/flows/hasse.test.tsx` — verify SVG renders and export buttons work

4. **Learning schedule** — Multi-stage plasticity interpolation for GLA/NHG. Allows the plasticity to follow a custom schedule rather than a simple linear interpolation.

5. **Multiple runs with collated results** — Run probabilistic algorithms N times and aggregate weights/ranking values. Useful for assessing stability.

6. **HTML tableaux** — Render tableaux as HTML with configurable shading, rather than plain text in the download.

7. **Praat export** — Generate `.OTGrammar` and `.PairDistribution` files for use in Praat.

8. **Excel file parsing** — Support `.xlsx` input files in addition to tab-delimited `.txt`.
