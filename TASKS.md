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
- [ ] A priori rankings file parsing
- [ ] Natural class file parsing

## Categorical Ranking Algorithms

- [x] Recursive Constraint Demotion (RCD)
- [x] RCD ranking arguments (transitive reduction basis)
- [x] RCD mini-tableaux generation
- [x] RCD constraint necessity analysis
- [x] Biased Constraint Demotion (BCD)
- [x] BCD specificity option (`mnuSpecificBCD`)
- [ ] Low Faithfulness Constraint Demotion (LFCD)
- [ ] A priori ranking enforcement in RCD/BCD/LFCD

## Ranking Argumentation (FRed)

- [~] FRed — partial implementation exists within RCD output, but standalone FRed with full ERC fusion/recursion is not implemented
- [ ] Skeletal Basis mode
- [ ] Most Informative Basis mode
- [ ] Detailed argumentation output
- [ ] Hasse diagram generation (GraphViz DOT output)

## Probabilistic Learning Algorithms

- [ ] Gradual Learning Algorithm (GLA) — Stochastic OT mode
- [ ] Gradual Learning Algorithm (GLA) — online MaxEnt mode
- [x] Batch Maximum Entropy (GIS optimizer)
- [ ] Noisy Harmonic Grammar (NHG) — 8 noise variants
- [ ] Learning schedule (multi-stage plasticity interpolation)
- [ ] Custom learning schedule from file
- [ ] Magri update rule (Stochastic OT)
- [ ] Gaussian prior (MaxEnt)
- [ ] Exact proportions data presentation
- [ ] Multiple runs with collated results
- [ ] Pairwise ranking probabilities
- [ ] History file output (weights/ranking values over time)

## Factorial Typology

- [ ] Core factorial typology computation
- [ ] FastRCD (streamlined RCD for derivability testing)
- [ ] T-order computation (typological implications)
- [ ] FTSum output file
- [ ] CompactSum output file
- [ ] Full listing with grammars per output pattern

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
- [~] Algorithm variant selection (RCD / BCD / LFCD)
- [ ] Ranking argumentation options (MIB, details, minitableaux)
- [ ] Factorial typology button and options
- [ ] A priori rankings controls
- [ ] Parameter inputs for probabilistic algorithms
- [ ] Progress indicator for long computations
- [ ] Tableau axis switching (for crowded tableaux)
- [ ] Hasse diagram viewer
- [ ] Settings persistence

## Testing & Examples

- [x] Tiny example (`examples/tiny/`) with expected RCD output
- [ ] BCD example with expected output
- [ ] LFCD example with expected output
- [ ] FRed example with expected output
- [ ] GLA/MaxEnt example with expected output
- [ ] NHG example with expected output
- [ ] Factorial typology example with expected output
- [ ] Edge case examples (ranking failures, ties, large tableaux)

---

## Suggested Next Tasks

Roughly ordered by value and dependency:

1. **BCD algorithm** — Second categorical ranking algorithm. Shares core demotion logic with RCD but adds Faithfulness delay, activeness, and subset selection heuristics.

2. **LFCD algorithm** — Third categorical ranking algorithm. Adds superset filtering, specificity, and autonomy (helper counting).

3. **Framework selection UI** — Add radio buttons to choose between Classical OT / MaxEnt / NHG / Stochastic OT. Wire up algorithm variant selection (RCD/BCD/LFCD) for Classical OT.

4. **FRed (full standalone)** — Complete ERC fusion algebra, recursive basis computation, entailment checks. Both Skeletal Basis and MIB modes.

5. **A priori rankings** — Parse ranking files, enforce during RCD/BCD/LFCD, convert to ERCs for FRed.

6. **Batch MaxEnt** — GIS optimizer. Simpler than GLA since it's a batch algorithm with no noise or sampling.

7. **GLA (Stochastic OT + online MaxEnt)** — Online error-driven learner with plasticity schedule and noise.

8. **NHG** — Online learner with 8 noise variant configurations.

9. **Factorial Typology** — Requires FastRCD. Incremental cross-classification of output patterns.

10. **Hasse diagrams** — Generate DOT format from ranking arguments. Could use a JS graph library instead of GraphViz.
