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
- [x] Full annotated history (FullHistory.txt — trial + input/generated/heard + values)
- [x] Candidate probability history (HistoryOfCandidateProbabilities.txt — GLA MaxEnt mode only)
- [x] Output probability history (HistoryOfOutputProbabilities.txt — Batch MaxEnt)
- [x] History of weights (HistoryOfWeights.txt — Batch MaxEnt, per-iteration weight log)

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
- [x] Tableau axis switching (for crowded tableaux)
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

## VB6 UI Parity

Discrepancies between the VB6 UI defaults/options and the web UI:

- [x] Default framework should be Maximum Entropy (VB6: `optMaximumEntropy` checked; web: Classical OT)
- [x] "Show details of argumentation" should default to checked (VB6: checked; web: unchecked)
- [x] MaxEnt iterations default should be 5 (VB6: 5; web: 100)
- [x] Add a priori rankings support to GLA panel (VB6 `boersma.frm` supports it; web doesn't expose it)
- [ ] Add "Diagnostics if ranking fails" option (VB6: `chkDiagnosticTableaux`, checked by default; web: not exposed)
- [x] Rename Factorial Typology "Include grammar listing" to "Include rankings in results" (VB6 terminology)
- [ ] Add HTML output options for shading customization (VB6: dialog for shading darkness and custom color; web: hardcoded)

---

## Conformance Test Failures

Of 31 conformance cases in the manifest, 24 skip (golden file not yet collected), 4 pass, and 3 fail:

### 1. `bcd_specific` — Wrong title

**Symptom:** Our title says "Biased Constraint Demotion (Specific)", VB6 says "Biased Constraint Demotion".

**VB6 source:** `Main.frm:5655-5656` always sets `gAlgorithmName = "Biased Constraint Demotion"` regardless of specificity mode. The specificity note is a separate paragraph after the header (`Main.frm:5678-5681`): "Version of BCD: specific Faithfulness constraints get priority."

**Fix:** In `lib.rs`, `format_bcd_output` and `format_bcd_html_output`: always use `"Biased Constraint Demotion"` as the algorithm name. Optionally add the separate specificity paragraph to match VB6.

### 2. `rcd_no_fred` — Extra trailing blank line

**Symptom:** 53 actual lines vs 52 expected — one extra trailing blank line.

**VB6 source:** `Main.frm:6349-6351` uses `PrintPara` for the mass deletion message, which ends with two `Print` statements (`Module1.bas:243-244`) producing exactly `\n\n` after the text. Our code (`rcd.rs:467`) ends with `\n\n\n` — one extra newline that's invisible when more sections follow, but creates an extra blank line when mass deletion is the last section.

**Fix:** In `rcd.rs`, change mass deletion success trailing from `\n\n\n` to `\n\n`, and add a leading `\n` to section 4 (FRed) and section 5 (mini-tableaux) headers to preserve inter-section spacing. Alternatively, normalize trailing newlines at the end of `format_output_with_algorithm`.

### 3. `rcd_apriori` — Missing "A Priori Rankings" section + wrong section numbers + incomplete mass deletion message + wrong constraint necessity ordering

**Symptom:** VB6 output has "2. A Priori Rankings" between Result and Tableaux, shifting all subsequent section numbers up by 1. Our output skips this section entirely.

This is a multi-part fix:

**3a. Missing "A Priori Rankings" section.** VB6 `Main.frm:7004-7040` (`PrintOutTheAprioriRankings`) prints a table showing which constraints dominate others. The table uses `s.PrintTable` (`s.bas:244`) with `CenterCells=True`, which calls `PrintTextTable` (`s.bas:373`) using per-column widths and `CenteredFillout` (`s.bas:430`) for alignment (centered with leftward error for odd spacing). The VB6 array stores data as `Table(col, row)` (reversed indices, per `s.bas:264`), with axis-switching: `gAPrioriRankingsTable(i,j)=True` (i dominates j) places "yes" at `Table(j+1, i+1)` = row i, col j in the output. Our `apriori.rs` stores `table[i][j]=true` meaning i dominates j, so display should use `apriori[row][col]`. Note: `gla.rs:458` already has an a priori table formatter but uses `apriori[col][row]` (inverted) — that may be a bug in GLA (no golden file to test against yet).

**3b. Dynamic section numbering.** VB6 uses a global auto-incrementing counter `gLevel1HeadingNumber` (`Module1.bas:175`). Our code hardcodes section numbers (1, 2, 3, 4, 5) in `rcd.rs` and `fred.rs:463`. When a priori is present, sections shift: 1=Result, 2=A Priori Rankings, 3=Tableaux, 4=Status, 5=FRed, 6=Mini-Tableaux. Need a mutable counter, and `fred.rs:format_section4` must accept the section number as a parameter.

**3c. Mass deletion message — missing "individual but not mass" case.** VB6 `Main.frm:6347-6358` has two messages gated on `NumberOfDeletableConstraints >= 2`: if mass deletion succeeds, print success message; otherwise, print "although the grammar will still work with the removal of ANY ONE... will NOT work if they are removed en masse." `NumberOfDeletableConstraints` counts constraints where `TrulyNeeded=False` (`Main.frm:6183-6188`). Our code (`rcd.rs:463-470`) only handles the success case and emits `\n\n` for the failure case — it doesn't print the "individual but not mass" message.

**3d. Constraint necessity ordering.** VB6 `Main.frm:6287-6341` outputs constraints grouped by category in three separate loops: first all Necessary constraints (in original order), then all "Not necessary (but included to show Faithfulness violations)" (in order), then all "Not necessary" (in order). Our code (`rcd.rs:447`) iterates by constraint index, which produces a different order when categories are interleaved.

---

## Suggested Next Tasks

Roughly ordered by value and dependency:

1. **Fix conformance failures** — Address the three failures documented above.

2. **Collect VB6 golden files** — Run VB6 OTSoft on Windows following `conformance/CHECKLIST.md` to populate the missing golden files. Most conformance test cases currently skip.

3. **Praat export** — Generate `.OTGrammar` and `.PairDistribution` files for use in Praat.

4. **Excel file parsing** — Support `.xlsx` input files in addition to tab-delimited `.txt`.
