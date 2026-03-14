//! Factorial Typology
//!
//! Computes the full factorial typology: the set of all possible output
//! patterns derivable by some total ranking of the constraints.
//!
//! Algorithm (matches VB6 FactorialTypology.bas):
//!
//! 1. Pre-filter: for each form, test each candidate independently to see
//!    if it can be derived at all. Candidates that can never win become
//!    "permanent losers" and are excluded from the search (but shown in output).
//!
//! 2. Initialize Valhalla with all possible candidates for the first form.
//!
//! 3. Incremental construction: for each subsequent form, cross-classify
//!    existing patterns with the new form's possible candidates. Keep only
//!    combinations that are jointly derivable by some ranking.
//!
//! 4. Derive candidate_derivable and T-order from the final pattern set.

use wasm_bindgen::prelude::*;
use crate::tableau::Tableau;

/// A single typological implication in the t-order.
/// "If input[implicator_form] → candidate[implicator_candidate],
///  then input[implicated_form] → candidate[implicated_candidate]."
#[derive(Debug, Clone)]
pub struct TOrderEntry {
    pub implicator_form: usize,
    pub implicator_candidate: usize,
    pub implicated_form: usize,
    pub implicated_candidate: usize,
}

/// Result of a factorial typology run.
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct FactorialTypologyResult {
    /// Each pattern is a vec of candidate indices (0-indexed), one per form.
    /// patterns[pattern_idx][form_idx] = candidate_idx into form.candidates
    #[wasm_bindgen(skip)]
    pub patterns: Vec<Vec<usize>>,

    /// For each form, for each candidate: whether it appears in any pattern.
    /// candidate_derivable[form_idx][cand_idx]
    #[wasm_bindgen(skip)]
    pub candidate_derivable: Vec<Vec<bool>>,

    /// T-order implications.
    #[wasm_bindgen(skip)]
    pub torder: Vec<TOrderEntry>,
}

#[wasm_bindgen]
impl FactorialTypologyResult {
    /// Number of derivable output patterns.
    pub fn pattern_count(&self) -> usize {
        self.patterns.len()
    }

    /// Candidate index (into the form's candidates) for a given pattern and form.
    pub fn get_pattern_candidate(&self, pattern_idx: usize, form_idx: usize) -> Option<usize> {
        self.patterns.get(pattern_idx)?.get(form_idx).copied()
    }

    /// Whether a given candidate is derivable by any ranking.
    pub fn is_candidate_derivable(&self, form_idx: usize, cand_idx: usize) -> bool {
        self.candidate_derivable
            .get(form_idx)
            .and_then(|row| row.get(cand_idx))
            .copied()
            .unwrap_or(false)
    }

    /// Number of t-order implications.
    pub fn torder_count(&self) -> usize {
        self.torder.len()
    }

    pub fn get_torder_implicator_form(&self, i: usize) -> usize {
        self.torder[i].implicator_form
    }

    pub fn get_torder_implicator_candidate(&self, i: usize) -> usize {
        self.torder[i].implicator_candidate
    }

    pub fn get_torder_implicated_form(&self, i: usize) -> usize {
        self.torder[i].implicated_form
    }

    pub fn get_torder_implicated_candidate(&self, i: usize) -> usize {
        self.torder[i].implicated_candidate
    }
}

// ─── FastRCD ─────────────────────────────────────────────────────────────────

/// Returns true if there exists some constraint ranking that derives all
/// winner→rival preferences simultaneously across all forms.
///
/// winner_viols[form_idx][constraint_idx]
/// rival_viols[form_idx][rival_idx][constraint_idx]
/// apriori[i][j] = true means constraint i must outrank constraint j
fn fast_rcd(
    winner_viols: &[&[i32]],
    rival_viols: &[Vec<Vec<i32>>],
    nc: usize,
    apriori: &[Vec<bool>],
) -> bool {
    let nf = winner_viols.len();
    debug_assert_eq!(rival_viols.len(), nf);

    let mut stratum = vec![0usize; nc];

    // informative[form_idx][rival_idx]
    let mut informative: Vec<Vec<bool>> = rival_viols
        .iter()
        .map(|rivals| vec![true; rivals.len()])
        .collect();

    let mut current_stratum = 0usize;

    loop {
        current_stratum += 1;

        let mut demotable = vec![false; nc];

        // AVOID PREFERENCE FOR LOSERS
        for form_idx in 0..nf {
            let wv = winner_viols[form_idx];
            for (rival_idx, rv) in rival_viols[form_idx].iter().enumerate() {
                if !informative[form_idx][rival_idx] {
                    continue;
                }
                for c_idx in 0..nc {
                    if stratum[c_idx] != 0 {
                        continue;
                    }
                    if wv[c_idx] > rv[c_idx] {
                        demotable[c_idx] = true;
                    }
                }
            }
        }

        // A PRIORI RANKINGS
        if !apriori.is_empty() {
            for outer in 0..nc {
                if stratum[outer] == 0 {
                    for inner in 0..nc {
                        if apriori[outer][inner] {
                            demotable[inner] = true;
                        }
                    }
                }
            }
        }

        // Assign non-demotable unranked constraints to this stratum
        let mut added_any = false;
        for c_idx in 0..nc {
            if stratum[c_idx] == 0 && !demotable[c_idx] {
                stratum[c_idx] = current_stratum;
                added_any = true;
            }
        }

        // Check termination
        let all_ranked = stratum.iter().all(|&s| s != 0);
        if all_ranked {
            return true; // Case III: success
        }
        if !added_any {
            return false; // Case II: all remaining unranked are demotable → failure
        }
        // Case I: some added, some remain demotable → continue

        // UPDATE INFORMATIVENESS
        for c_idx in 0..nc {
            if stratum[c_idx] == current_stratum {
                for form_idx in 0..nf {
                    let wv = winner_viols[form_idx];
                    for (rival_idx, rv) in rival_viols[form_idx].iter().enumerate() {
                        if rv[c_idx] > wv[c_idx] {
                            informative[form_idx][rival_idx] = false;
                        }
                    }
                }
            }
        }
    }
}

// ─── T-Order ─────────────────────────────────────────────────────────────────

fn compute_torder(patterns: &[Vec<usize>], nf: usize) -> Vec<TOrderEntry> {
    if patterns.is_empty() {
        return Vec::new();
    }

    // For each form, how many distinct candidates appear in FT patterns?
    let mut winners_count: Vec<usize> = vec![0; nf];
    for form_idx in 0..nf {
        let mut seen: Vec<usize> = Vec::new();
        for pattern in patterns {
            let c = pattern[form_idx];
            if !seen.contains(&c) {
                seen.push(c);
            }
        }
        winners_count[form_idx] = seen.len();
    }

    let mut torder = Vec::new();

    for implicator_form in 0..nf {
        // Collect all candidates that appear in the FT for this form
        let mut ft_candidates: Vec<usize> = Vec::new();
        for pattern in patterns {
            let c = pattern[implicator_form];
            if !ft_candidates.contains(&c) {
                ft_candidates.push(c);
            }
        }

        for implicator_cand in ft_candidates {
            for implicated_form in 0..nf {
                if implicated_form == implicator_form {
                    continue;
                }
                // Skip forms where only one candidate ever wins (trivially implied)
                if winners_count[implicated_form] <= 1 {
                    continue;
                }

                // Find if implicator_form → implicator_cand uniquely determines
                // implicated_form's output across all patterns.
                let mut implicated_cand: Option<usize> = None;
                let mut consistent = true;

                for pattern in patterns {
                    if pattern[implicator_form] == implicator_cand {
                        match implicated_cand {
                            None => implicated_cand = Some(pattern[implicated_form]),
                            Some(prev) => {
                                if pattern[implicated_form] != prev {
                                    consistent = false;
                                    break;
                                }
                            }
                        }
                    }
                }

                if consistent {
                    if let Some(imp_cand) = implicated_cand {
                        torder.push(TOrderEntry {
                            implicator_form,
                            implicator_candidate: implicator_cand,
                            implicated_form,
                            implicated_candidate: imp_cand,
                        });
                    }
                }
            }
        }
    }

    torder
}

// ─── Main Algorithm ───────────────────────────────────────────────────────────

impl Tableau {
    /// Run the factorial typology algorithm.
    ///
    /// Returns the complete set of derivable output patterns and their T-order.
    pub fn run_factorial_typology(&self, apriori: &[Vec<bool>]) -> FactorialTypologyResult {
        let nc = self.constraints.len();
        let nf = self.forms.len();

        if nf == 0 {
            return FactorialTypologyResult {
                patterns: Vec::new(),
                candidate_derivable: Vec::new(),
                torder: Vec::new(),
            };
        }

        // ── Step 1: Pre-filter ────────────────────────────────────────────────
        // For each form, find which candidates can be derived in isolation.

        let mut possible: Vec<Vec<usize>> = Vec::with_capacity(nf);

        for form in &self.forms {
            let ncands = form.candidates.len();
            let mut form_possible = Vec::new();

            for cand_idx in 0..ncands {
                // Build winner/rival views for this single-form test
                let winner_v: &[i32] = &form.candidates[cand_idx].violations;
                let rivals: Vec<Vec<i32>> = form
                    .candidates
                    .iter()
                    .enumerate()
                    .filter(|(j, _)| *j != cand_idx)
                    .map(|(_, c)| c.violations.clone())
                    .collect();

                let winner_slice: &[&[i32]] = &[winner_v];
                let rival_slice: &[Vec<Vec<i32>>] = &[rivals];

                if fast_rcd(winner_slice, rival_slice, nc, apriori) {
                    form_possible.push(cand_idx);
                }
            }

            possible.push(form_possible);
        }

        // ── Step 2: Initialize Valhalla for first form ────────────────────────
        let mut patterns: Vec<Vec<usize>> = possible[0]
            .iter()
            .map(|&cand_idx| vec![cand_idx])
            .collect();

        // ── Step 3: Incremental Construction ─────────────────────────────────
        for (form_idx, possible_cands) in possible.iter().enumerate().skip(1) {
            let mut new_patterns: Vec<Vec<usize>> = Vec::new();

            for old_pattern in &patterns {
                for &new_cand in possible_cands {
                    // Build the test data: for each form in 0..=form_idx,
                    // the selected candidate is the winner and all others are rivals.

                    let test_n = form_idx + 1;
                    let mut winner_vecs: Vec<Vec<i32>> = Vec::with_capacity(test_n);
                    let mut rival_vecs: Vec<Vec<Vec<i32>>> = Vec::with_capacity(test_n);

                    // Forms already in pattern
                    for (fi, &selected) in old_pattern.iter().enumerate() {
                        let form = &self.forms[fi];
                        winner_vecs.push(form.candidates[selected].violations.clone());
                        let rivals: Vec<Vec<i32>> = form
                            .candidates
                            .iter()
                            .enumerate()
                            .filter(|(j, _)| *j != selected)
                            .map(|(_, c)| c.violations.clone())
                            .collect();
                        rival_vecs.push(rivals);
                    }

                    // New form
                    let new_form = &self.forms[form_idx];
                    winner_vecs.push(new_form.candidates[new_cand].violations.clone());
                    let rivals: Vec<Vec<i32>> = new_form
                        .candidates
                        .iter()
                        .enumerate()
                        .filter(|(j, _)| *j != new_cand)
                        .map(|(_, c)| c.violations.clone())
                        .collect();
                    rival_vecs.push(rivals);

                    // Build slices for fast_rcd
                    let winner_slices: Vec<&[i32]> =
                        winner_vecs.iter().map(|v| v.as_slice()).collect();

                    if fast_rcd(&winner_slices, &rival_vecs, nc, apriori) {
                        let mut new_p = old_pattern.clone();
                        new_p.push(new_cand);
                        new_patterns.push(new_p);
                    }
                }
            }

            patterns = new_patterns;
        }

        // ── Step 4: Derive candidate_derivable ───────────────────────────────
        let mut candidate_derivable: Vec<Vec<bool>> = self
            .forms
            .iter()
            .map(|f| vec![false; f.candidates.len()])
            .collect();

        for pattern in &patterns {
            for (form_idx, &cand_idx) in pattern.iter().enumerate() {
                candidate_derivable[form_idx][cand_idx] = true;
            }
        }

        // ── Step 5: T-order ───────────────────────────────────────────────────
        let torder = compute_torder(&patterns, nf);

        FactorialTypologyResult {
            patterns,
            candidate_derivable,
            torder,
        }
    }
}

// ─── Output Formatting ───────────────────────────────────────────────────────

impl FactorialTypologyResult {
    /// Generate formatted text output for factorial typology results.
    ///
    /// Matches VB6 FactorialTypology.bas output format exactly.
    ///
    /// `apriori` is the parsed a priori rankings table (empty slice = none).
    /// `include_full_listing` controls the "Summary results appear at end of file" note
    /// and should match whether `format_full_listing` will be appended.
    pub fn format_output(
        &self,
        tableau: &Tableau,
        filename: &str,
        apriori: &[Vec<bool>],
        include_full_listing: bool,
    ) -> String {
        let mut out = String::new();
        let nf = tableau.forms.len();
        let nc = tableau.constraints.len();
        let np = self.patterns.len();
        let has_apriori = !apriori.is_empty();

        // Section counter (incremented before each PrintLevel1Header call)
        let mut section = 0usize;

        // ── Header ────────────────────────────────────────────────────────────
        // VB6: PrintTopLevelHeader (no filename), then PrintPara for date/version/source
        out.push_str("Results of Factorial Typology Search\n\n");

        let now = chrono::Local::now();
        out.push_str(&format!(
            "{}\n\n",
            now.format("%-m-%-d-%Y, %-I:%M %p")
                .to_string()
                .to_lowercase()
        ));
        out.push_str("OTSoft 2.7, release date 2/1/2026\n\n");
        out.push_str(&format!("Source file:  {}\n\n\n", filename));

        // ── 1. Constraints ────────────────────────────────────────────────────
        // VB6: PrintLevel1Header("Constraints") then s.PrintTable with 3-col table
        section += 1;
        out.push_str(&format!("\n{}. Constraints\n\n", section));

        // VB6 table col widths: max(header, data) per column.
        // Col 1: number labels ("N."), col 2: full names, col 3: abbreviations.
        let num_digits = nc.to_string().len();
        let col1_w = num_digits + 1; // e.g. "1." = 2, "10." = 3
        let max_full_name = tableau
            .constraints
            .iter()
            .map(|c| c.full_name().len())
            .max()
            .unwrap_or(0)
            .max("Full Name".len());
        let max_abbrev = tableau
            .constraints
            .iter()
            .map(|c| c.abbrev().len())
            .max()
            .unwrap_or(0)
            .max("Abbr.".len());

        // Header row: empty col1, then "Full Name" and "Abbr." headers
        out.push_str(&format!(
            "{:<cw$}  {:<fw$}  {:<aw$}\n",
            "",
            "Full Name",
            "Abbr.",
            cw = col1_w,
            fw = max_full_name,
            aw = max_abbrev,
        ));
        // Data rows
        for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
            let label = format!("{}.", c_idx + 1);
            out.push_str(&format!(
                "{:<cw$}  {:<fw$}  {:<aw$}\n",
                label,
                constraint.full_name(),
                constraint.abbrev(),
                cw = col1_w,
                fw = max_full_name,
                aw = max_abbrev,
            ));
        }
        // s.PrintTable adds 1 trailing blank line
        out.push('\n');

        // ── Post-constraints: a priori section or "All rankings were considered." ──
        if has_apriori {
            // VB6: Form1.PrintOutTheAprioriRankings — uses PrintLevel1Header then table
            section += 1;
            out.push_str(&format!("\n{}. A Priori Rankings\n\n", section));
            out.push_str(
                "In the following table, \"yes\" means that the constraint of the indicated row \n\
                 was marked a priori to dominate the constraint in the given column.\n\n",
            );

            // Build the apriori table. Due to a VB6 bug, all data cells are always blank.
            // Col 1: row labels (abbrevs), col 2..N: column headers (abbrevs), cells = "".
            let abbrevs: Vec<String> =
                tableau.constraints.iter().map(|c| c.abbrev()).collect();
            // Col widths: col1 = max abbrev length; each data col = that col's abbrev length.
            let col1_abbrev_w = abbrevs.iter().map(|a| a.len()).max().unwrap_or(0);
            // Header row
            out.push_str(&format!("{:<w$}", "", w = col1_abbrev_w));
            for abbrev in &abbrevs {
                out.push_str(&format!("  {:<w$}", abbrev, w = abbrev.len()));
            }
            out.push('\n');
            // Data rows (all blank — VB6 bug)
            for row_abbrev in &abbrevs {
                out.push_str(&format!("{:<w$}", row_abbrev, w = col1_abbrev_w));
                for abbrev in &abbrevs {
                    out.push_str(&format!("  {:<w$}", "", w = abbrev.len()));
                }
                out.push('\n');
            }
            out.push_str("\n\n\n");
        } else {
            // VB6: PrintPara("All rankings were considered.") then Print #mTmpFile,
            out.push_str("All rankings were considered.\n");
        }

        // VB6: if full listing, print "Summary results appear at end of file." note
        if include_full_listing {
            out.push_str(
                "\n\nSummary results appear at end of file.  \n\
                 Immediately below are reports on individual patterns generated.\n",
            );
        }

        // ── N. Summary Information ────────────────────────────────────────────
        section += 1;
        out.push_str(&format!("\n\n\n{}. Summary Information\n\n", section));

        let max_grammars = factorial(nc);
        match max_grammars {
            Some(n) => out.push_str(&format!(
                "With {} constraints, the number of logically possible grammars is {}.\n",
                nc, n
            )),
            None => out.push_str(&format!(
                "With {} constraints, the number of logically possible grammars is too large to compute.\n",
                nc
            )),
        }

        out.push_str(&format!("\nThere were {} different output patterns.\n", np));

        if np == 0 {
            out.push_str("No derivable output patterns were found.\n");
            return out;
        }

        // Pattern table — 4 patterns per block
        out.push_str("\nForms marked as winners in the input file are marked with >.\n\n");

        // First col width: "/" + max_input + "/" padded to max_input + 3
        let max_input_width = tableau
            .forms
            .iter()
            .map(|f| f.input.len())
            .max()
            .unwrap_or(0);
        let first_col_width = max_input_width + 4;

        let block_size = 4;
        let mut block_start = 0;
        while block_start < np {
            let block_end = (block_start + block_size).min(np);
            let block = &self.patterns[block_start..block_end];

            // Per-column widths: max(header len, max candidate name) + 2
            let col_widths: Vec<usize> = block
                .iter()
                .enumerate()
                .map(|(bi, _)| {
                    let header_len = format!("Output #{}", block_start + bi + 1).len();
                    let max_cand = block[bi]
                        .iter()
                        .enumerate()
                        .map(|(fi, &ci)| tableau.forms[fi].candidates[ci].form.len())
                        .max()
                        .unwrap_or(0);
                    header_len.max(max_cand) + 2
                })
                .collect();

            // Column headers
            out.push_str(&" ".repeat(first_col_width));
            for (bi, _) in block.iter().enumerate() {
                let header = format!("Output #{}", block_start + bi + 1);
                if bi + 1 < block.len() {
                    out.push_str(&format!("{:<width$}", header, width = col_widths[bi]));
                } else {
                    out.push_str(&header);
                }
            }
            out.push('\n');

            // One row per input form
            for form_idx in 0..nf {
                let form = &tableau.forms[form_idx];
                let input_display = format!("/{}/", form.input);
                out.push_str(&format!(
                    "{:<width$}",
                    input_display,
                    width = first_col_width
                ));

                for (bi, pattern) in block.iter().enumerate() {
                    let cand_idx = pattern[form_idx];
                    let cand_name = &form.candidates[cand_idx].form;
                    // VB6: marks ">" when Valhalla index == 1 (the first/winner candidate).
                    // After InstallTheWinnersAsMereCandidates, position 1 = the first-listed
                    // candidate (index 0 in Rust). This is a known VB6 bug (comment in source:
                    // "This needs fixed: show winners if frequency > 0") — we reproduce it.
                    let marker = if cand_idx == 0 { ">" } else { " " };
                    let cell = format!("{}{}", marker, cand_name);
                    if bi + 1 < block.len() {
                        out.push_str(&format!("{:<width$}", cell, width = col_widths[bi]));
                    } else {
                        out.push_str(&cell);
                    }
                }
                out.push('\n');
            }

            out.push('\n');
            block_start += block_size;
        }

        // ── N. List of Winners ─────────────────────────────────────────────────
        // VB6: PrintLevel1Header("List of Winners"), then per-form / per-candidate output
        section += 1;
        out.push_str(&format!("\n{}. List of Winners\n\n", section));
        out.push_str(
            "The following specifies for each candidate whether there is at least one ranking that derives it:\n\n",
        );

        for (form_idx, form) in tableau.forms.iter().enumerate() {
            // VB6: Print "/input/:  " (with 2 trailing spaces)
            out.push_str(&format!("/{}/:\n", form.input));

            // VB6 iterates candidates in two passes: derivable (yes) first, then
            // non-derivable (no), each group in original file order.
            let print_candidate = |out: &mut String, cand: &crate::tableau::Candidate, derivable: bool| {
                let status = if derivable { "yes" } else { "no" };
                // VB6: "   [rival]:  " + spaces from Len(rival) to 10 + yes/no
                let spaces = 11usize.saturating_sub(cand.form.len());
                out.push_str(&format!(
                    "   [{}]:  {}{}\n",
                    cand.form,
                    " ".repeat(spaces),
                    status,
                ));
            };
            // Derivable candidates first
            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                if self.candidate_derivable[form_idx][cand_idx] {
                    print_candidate(&mut out, cand, true);
                }
            }
            // Non-derivable candidates second
            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                if !self.candidate_derivable[form_idx][cand_idx] {
                    print_candidate(&mut out, cand, false);
                }
            }
            // VB6: no blank line between forms in this section
        }

        // ── N. T-orders ────────────────────────────────────────────────────────
        // VB6 section title is "T-orders" (lowercase s) and uses "factorical" (typo)
        section += 1;
        out.push_str(&format!("\n{}. T-orders\n\n", section));
        out.push_str("The t-order is the set of implications in a factorical typology.\n\n");

        // Always-winning candidates (only one derivable candidate for that form):
        // VB6 uses col headers "Input" and "Output" with raw names (no slashes/brackets)
        let mut always_winners: Vec<(usize, usize)> = Vec::new();
        for form_idx in 0..nf {
            let winning_cands: Vec<usize> = tableau.forms[form_idx]
                .candidates
                .iter()
                .enumerate()
                .filter(|(ci, _)| self.candidate_derivable[form_idx][*ci])
                .map(|(ci, _)| ci)
                .collect();
            if winning_cands.len() == 1 {
                always_winners.push((form_idx, winning_cands[0]));
            }
        }

        if !always_winners.is_empty() {
            out.push_str("For the following input-output pairs, no other candidate ever wins, so they are not reported separately in the t-order:\n\n");
            // VB6 table: col "Input" and "Output", raw names without slashes/brackets
            let col1_w = always_winners
                .iter()
                .map(|(fi, _)| tableau.forms[*fi].input.len())
                .max()
                .unwrap_or(0)
                .max("Input".len());
            let col2_w = always_winners
                .iter()
                .map(|(fi, ci)| tableau.forms[*fi].candidates[*ci].form.len())
                .max()
                .unwrap_or(0)
                .max("Output".len());
            out.push_str(&format!(
                "{:<cw$}  {:<ow$}\n",
                "Input",
                "Output",
                cw = col1_w,
                ow = col2_w,
            ));
            for (form_idx, cand_idx) in &always_winners {
                let form = &tableau.forms[*form_idx];
                out.push_str(&format!(
                    "{:<cw$}  {:<ow$}\n",
                    form.input,
                    form.candidates[*cand_idx].form,
                    cw = col1_w,
                    ow = col2_w,
                ));
            }
            out.push('\n');
        }

        // T-order implication table (VB6: s.PrintTable with 4 cols, dynamic widths)
        if self.torder.is_empty() {
            out.push_str("No t-order implications were found.\n");
        } else {
            let col1_w = self
                .torder
                .iter()
                .map(|e| format!("/{}/", tableau.forms[e.implicator_form].input).len())
                .max()
                .unwrap_or(0)
                .max("If this input".len());
            let col2_w = self
                .torder
                .iter()
                .map(|e| {
                    let f = &tableau.forms[e.implicator_form];
                    format!("[{}]", f.candidates[e.implicator_candidate].form).len()
                })
                .max()
                .unwrap_or(0)
                .max("has this output".len());
            let col3_w = self
                .torder
                .iter()
                .map(|e| format!("/{}/", tableau.forms[e.implicated_form].input).len())
                .max()
                .unwrap_or(0)
                .max("then this input".len());
            let col4_w = self
                .torder
                .iter()
                .map(|e| {
                    let f = &tableau.forms[e.implicated_form];
                    format!("[{}]", f.candidates[e.implicated_candidate].form).len()
                })
                .max()
                .unwrap_or(0)
                .max("has this output".len());

            out.push_str(&format!(
                "{:<c1$}  {:<c2$}  {:<c3$}  {:<c4$}\n",
                "If this input",
                "has this output",
                "then this input",
                "has this output",
                c1 = col1_w,
                c2 = col2_w,
                c3 = col3_w,
                c4 = col4_w,
            ));
            for entry in &self.torder {
                let imp_form = &tableau.forms[entry.implicator_form];
                let imp_cand = &imp_form.candidates[entry.implicator_candidate];
                let ted_form = &tableau.forms[entry.implicated_form];
                let ted_cand = &ted_form.candidates[entry.implicated_candidate];
                out.push_str(&format!(
                    "{:<c1$}  {:<c2$}  {:<c3$}  {:<c4$}\n",
                    format!("/{}/", imp_form.input),
                    format!("[{}]", imp_cand.form),
                    format!("/{}/", ted_form.input),
                    format!("[{}]", ted_cand.form),
                    c1 = col1_w,
                    c2 = col2_w,
                    c3 = col3_w,
                    c4 = col4_w,
                ));
            }
        }

        // Non-implicators: VB6 includes ALL candidates (not just derivable ones)
        // that are not implicators, using raw names without slashes/brackets.
        let implicator_set: Vec<(usize, usize)> = self
            .torder
            .iter()
            .map(|e| (e.implicator_form, e.implicator_candidate))
            .collect();

        let non_implicators: Vec<(usize, usize)> = (0..nf)
            .flat_map(|fi| {
                let imp_set_ref = &implicator_set;
                tableau.forms[fi]
                    .candidates
                    .iter()
                    .enumerate()
                    .filter(move |(ci, _)| !imp_set_ref.contains(&(fi, *ci)))
                    .map(move |(ci, _)| (fi, ci))
                    .collect::<Vec<_>>()
            })
            .collect();

        if !non_implicators.is_empty() {
            out.push_str("\nNothing is implicated by these input-output pairs:\n\n");
            // VB6 table: col "Input" and "Candidate", raw names without slashes/brackets
            let col1_w = non_implicators
                .iter()
                .map(|(fi, _)| tableau.forms[*fi].input.len())
                .max()
                .unwrap_or(0)
                .max("Input".len());
            let col2_w = non_implicators
                .iter()
                .map(|(fi, ci)| tableau.forms[*fi].candidates[*ci].form.len())
                .max()
                .unwrap_or(0)
                .max("Candidate".len());
            out.push_str(&format!(
                "{:<cw$}  {:<ow$}\n",
                "Input",
                "Candidate",
                cw = col1_w,
                ow = col2_w,
            ));
            // VB6 prints non-implicators sorted per form: derivable (yes) first,
            // then non-derivable (no), each group in original file order.
            let print_ni_row = |out: &mut String, form_idx: usize, cand_idx: usize| {
                let form = &tableau.forms[form_idx];
                out.push_str(&format!(
                    "{:<cw$}  {:<ow$}\n",
                    form.input,
                    form.candidates[cand_idx].form,
                    cw = col1_w,
                    ow = col2_w,
                ));
            };
            // Two passes per form (derivable first, then non-derivable)
            for fi in 0..nf {
                for &(f, ci) in &non_implicators {
                    if f == fi && self.candidate_derivable[fi][ci] {
                        print_ni_row(&mut out, fi, ci);
                    }
                }
                for &(f, ci) in &non_implicators {
                    if f == fi && !self.candidate_derivable[fi][ci] {
                        print_ni_row(&mut out, fi, ci);
                    }
                }
            }
            // VB6 adds 2 blank lines after the non-implicators table
            out.push_str("\n\n");
        }

        out
    }

    /// Append the complete listing section: for each pattern, run RCD and show the grammar.
    ///
    /// `section_num` is the 1-based section number for the heading (e.g. 5 or 6 depending
    /// on whether an a priori section was included).
    pub fn format_full_listing(
        &self,
        tableau: &Tableau,
        apriori: &[Vec<bool>],
        section_num: usize,
    ) -> String {
        let nc = tableau.constraints.len();
        let np = self.patterns.len();

        if np == 0 {
            return String::new();
        }

        let mut out = String::new();

        // VB6: PrintLevel1Header adds "\n" before + "N. text\n" + "\n" after
        out.push_str(&format!(
            "\n{}. Complete Listing of Output Patterns\n\n",
            section_num
        ));

        // Width for aligning input column (VB6 uses raw input length, not "/input/" length)
        let max_input_width = tableau.forms.iter().map(|f| f.input.len()).max().unwrap_or(0);

        for (pat_idx, pattern) in self.patterns.iter().enumerate() {
            if pat_idx > 0 {
                out.push_str("\n\n------------------------------------------------------------------------------\n");
            }

            // VB6: PrintPara("OUTPUT SET #N:") adds blank line after
            out.push_str(&format!("OUTPUT SET #{}:\n\n", pat_idx + 1));

            // VB6: "These are the winning outputs.  PARA> specifies..." splits at PARA
            // into two PrintPara calls: first prints the part before PARA + blank line,
            // second prints "> specifies..." + blank line.
            out.push_str(
                "These are the winning outputs.  \n\
                 > specifies outputs marked as winning candidates in the input file.\n\n",
            );

            for (fi, &ci) in pattern.iter().enumerate() {
                let form = &tableau.forms[fi];
                let cand = &form.candidates[ci];
                let input_padded = format!("/{}/", form.input);
                // VB6: ">" is always printed unconditionally (line 1198 of FactorialTypology.bas)
                // VB6: (actual) when candidate index == last candidate (Valhalla == mNumberOfRivals).
                // This only triggers when VB6's InstallTheWinnersAsMereCandidates ran properly,
                // which requires the first-listed candidate (mWinner) to have frequency > 0.
                // When the first-listed candidate has freq=0, the install doesn't mark a winner
                // and the condition is never met.
                let is_actual = ci == form.candidates.len() - 1
                    && form.candidates[0].frequency > 0;
                let actual_label = if is_actual { " (actual)" } else { "" };
                // VB6: candidate + " " + optional(" (actual)") — note trailing space always present
                out.push_str(&format!(
                    "   {:<width$} -->  >{} {}\n",
                    input_padded,
                    cand.form,
                    actual_label,
                    width = max_input_width + 4,
                ));
            }

            out.push('\n');

            // Run RCD with this pattern's candidates as winners
            let rcd = tableau.run_rcd_with_winner_indices(pattern, apriori);

            out.push_str("Grammar:\n\n");
            for stratum in 1..=rcd.num_strata() {
                out.push_str(&format!("   Stratum #{}\n", stratum));
                for c_idx in 0..nc {
                    if rcd.get_stratum(c_idx) == Some(stratum) {
                        let c = &tableau.constraints[c_idx];
                        out.push_str(&format!(
                            "      {:<42}[= {}]\n",
                            c.full_name(),
                            c.abbrev()
                        ));
                    }
                }
            }
        }

        out
    }

    /// Format the FTSum tab-delimited output.
    ///
    /// Matches VB6 `FSSummary`:
    /// - Header row: `/input1/\t/input2/\t...` (no trailing tab)
    /// - One row per pattern with candidate names (no trailing tab)
    pub fn format_ftsum(&self, tableau: &Tableau) -> String {
        let nf = tableau.forms.len();
        let mut out = String::new();

        // Header row
        for (fi, form) in tableau.forms.iter().enumerate() {
            out.push('/');
            out.push_str(&form.input);
            out.push('/');
            if fi + 1 < nf {
                out.push('\t');
            }
        }
        out.push('\n');

        // One row per pattern
        for pattern in &self.patterns {
            for (fi, &ci) in pattern.iter().enumerate() {
                out.push_str(&tableau.forms[fi].candidates[ci].form);
                if fi + 1 < nf {
                    out.push('\t');
                }
            }
            out.push('\n');
        }

        out
    }

    /// Format the CompactSum tab-delimited output.
    ///
    /// Matches VB6 `CompactFTFile`:
    /// - Collates patterns by distinct surface output sets (ignoring which input each came from)
    /// - Deduplicates rows with identical compact representations
    /// - Each row: `count\tout1\tout2\t...` (trailing tab after each output, matching VB6)
    pub fn format_compact_sum(&self, tableau: &Tableau) -> String {
        let mut compact_valhalla: Vec<String> = Vec::new();

        for pattern in &self.patterns {
            let mut buffer = String::new();
            let mut count = 0usize;

            for (fi, &ci) in pattern.iter().enumerate() {
                let cand_name = &tableau.forms[fi].candidates[ci].form;
                // Check if this output was already recorded from an earlier form in this pattern
                let already_recorded = (0..fi).any(|inner_fi| {
                    let inner_ci = pattern[inner_fi];
                    &tableau.forms[inner_fi].candidates[inner_ci].form == cand_name
                });
                if !already_recorded {
                    buffer.push_str(cand_name);
                    buffer.push('\t');
                    count += 1;
                }
            }

            let full_buffer = format!("{}\t{}", count, buffer);
            if !compact_valhalla.contains(&full_buffer) {
                compact_valhalla.push(full_buffer);
            }
        }

        let mut out = String::new();
        for entry in compact_valhalla {
            out.push_str(&entry);
            out.push('\n');
        }
        out
    }
}

/// Compute n! returning None if overflow would occur (n > 20)
fn factorial(n: usize) -> Option<u64> {
    if n > 20 {
        return None;
    }
    let mut result = 1u64;
    for i in 2..=(n as u64) {
        result = result.checked_mul(i)?;
    }
    Some(result)
}

// ─── Chunked FT Runner ────────────────────────────────────────────────────────

/// Execution phase of the FT runner.
enum FtPhase {
    /// Running the pre-filter: testing each candidate in isolation.
    PreFilter,
    /// Incremental construction: cross-classifying patterns form by form.
    Construction,
    /// All computation complete.
    Done,
}

/// Chunked Factorial Typology runner for interactive progress reporting.
///
/// Processes one form at a time during the incremental construction phase,
/// yielding control to the browser between forms.
#[wasm_bindgen]
pub struct FtRunner {
    // ── Immutable config ──────────────────────────────────────────────────────
    tableau: Tableau,
    apriori: Vec<Vec<bool>>,
    total_forms: usize,

    // ── Pre-filter state ──────────────────────────────────────────────────────
    possible: Vec<Vec<usize>>,
    pre_filter_form_idx: usize,

    // ── Construction state ────────────────────────────────────────────────────
    patterns: Vec<Vec<usize>>,
    construction_form_idx: usize,

    // ── Phase ─────────────────────────────────────────────────────────────────
    phase: FtPhase,

    // ── Final result ───────────────────────────────────────────────────────────
    result: Option<FactorialTypologyResult>,
}

#[wasm_bindgen]
impl FtRunner {
    #[wasm_bindgen(constructor)]
    pub fn new(text: &str, apriori_text: &str) -> Result<FtRunner, String> {
        let tableau = Tableau::parse(text)?;
        let apriori = if apriori_text.trim().is_empty() {
            Vec::new()
        } else {
            let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
            crate::apriori::parse_apriori(apriori_text, &abbrevs)?
        };
        let total_forms = tableau.forms.len();
        Ok(FtRunner {
            apriori,
            total_forms,
            possible: Vec::new(),
            pre_filter_form_idx: 0,
            patterns: Vec::new(),
            construction_form_idx: 1,
            phase: FtPhase::PreFilter,
            result: None,
            tableau,
        })
    }

    /// Advance computation. Returns true when complete.
    /// Processes one form per call during the pre-filter and construction phases.
    pub fn run_chunk(&mut self, _max_work: usize) -> bool {
        match self.phase {
            FtPhase::PreFilter => self.run_prefilter_step(),
            FtPhase::Construction => self.run_construction_step(),
            FtPhase::Done => true,
        }
    }

    /// Progress as [completed, total].
    pub fn progress(&self) -> Vec<f64> {
        let nf = self.total_forms;
        if nf == 0 {
            return vec![1.0, 1.0];
        }
        match self.phase {
            FtPhase::PreFilter => vec![self.pre_filter_form_idx as f64, nf as f64],
            FtPhase::Construction => vec![
                (nf + self.construction_form_idx - 1) as f64,
                (2 * nf - 1) as f64,
            ],
            FtPhase::Done => vec![1.0, 1.0],
        }
    }

    /// Extract the final result. Only valid after `run_chunk` returns true.
    pub fn take_result(&mut self) -> FactorialTypologyResult {
        self.result.take().expect("FtRunner: take_result called before completion")
    }
}

impl FtRunner {
    /// Process one form in the pre-filter phase.
    fn run_prefilter_step(&mut self) -> bool {
        let nf = self.total_forms;
        if nf == 0 {
            self.finalize();
            return true;
        }

        if self.pre_filter_form_idx >= nf {
            return self.start_construction();
        }

        let form_idx = self.pre_filter_form_idx;
        let nc = self.tableau.constraints.len();
        let form = &self.tableau.forms[form_idx];
        let mut form_possible = Vec::new();

        for cand_idx in 0..form.candidates.len() {
            let winner_v: &[i32] = &form.candidates[cand_idx].violations;
            let rivals: Vec<Vec<i32>> = form.candidates.iter().enumerate()
                .filter(|(j, _)| *j != cand_idx)
                .map(|(_, c)| c.violations.clone())
                .collect();
            if fast_rcd(&[winner_v], &[rivals], nc, &self.apriori) {
                form_possible.push(cand_idx);
            }
        }

        self.possible.push(form_possible);
        self.pre_filter_form_idx += 1;

        if self.pre_filter_form_idx >= nf {
            self.start_construction()
        } else {
            false
        }
    }

    /// Transition from pre-filter to construction (or directly to done if 0/1 forms).
    fn start_construction(&mut self) -> bool {
        let nf = self.total_forms;

        if nf == 0 {
            self.finalize();
            return true;
        }

        // Initialize patterns from form 0's possible candidates
        self.patterns = self.possible[0].iter().map(|&c| vec![c]).collect();

        if nf == 1 {
            self.finalize();
            return true;
        }

        self.construction_form_idx = 1;
        self.phase = FtPhase::Construction;
        false
    }

    /// Process one form in the construction phase.
    fn run_construction_step(&mut self) -> bool {
        let nf = self.total_forms;
        let nc = self.tableau.constraints.len();

        if self.construction_form_idx >= nf {
            self.finalize();
            return true;
        }

        let form_idx = self.construction_form_idx;
        let possible_cands = &self.possible[form_idx];
        let mut new_patterns: Vec<Vec<usize>> = Vec::new();

        for old_pattern in &self.patterns {
            for &new_cand in possible_cands {
                let test_n = form_idx + 1;
                let mut winner_vecs: Vec<Vec<i32>> = Vec::with_capacity(test_n);
                let mut rival_vecs: Vec<Vec<Vec<i32>>> = Vec::with_capacity(test_n);

                for (fi, &selected) in old_pattern.iter().enumerate() {
                    let form = &self.tableau.forms[fi];
                    winner_vecs.push(form.candidates[selected].violations.clone());
                    let rivals: Vec<Vec<i32>> = form.candidates.iter().enumerate()
                        .filter(|(j, _)| *j != selected)
                        .map(|(_, c)| c.violations.clone())
                        .collect();
                    rival_vecs.push(rivals);
                }

                let new_form = &self.tableau.forms[form_idx];
                winner_vecs.push(new_form.candidates[new_cand].violations.clone());
                let rivals: Vec<Vec<i32>> = new_form.candidates.iter().enumerate()
                    .filter(|(j, _)| *j != new_cand)
                    .map(|(_, c)| c.violations.clone())
                    .collect();
                rival_vecs.push(rivals);

                let winner_slices: Vec<&[i32]> = winner_vecs.iter().map(|v| v.as_slice()).collect();
                if fast_rcd(&winner_slices, &rival_vecs, nc, &self.apriori) {
                    let mut new_p = old_pattern.clone();
                    new_p.push(new_cand);
                    new_patterns.push(new_p);
                }
            }
        }

        self.patterns = new_patterns;
        self.construction_form_idx += 1;

        if self.construction_form_idx >= nf {
            self.finalize();
            true
        } else {
            false
        }
    }

    fn finalize(&mut self) {
        let nf = self.total_forms;

        let mut candidate_derivable: Vec<Vec<bool>> = self.tableau.forms.iter()
            .map(|f| vec![false; f.candidates.len()])
            .collect();
        for pattern in &self.patterns {
            for (form_idx, &cand_idx) in pattern.iter().enumerate() {
                candidate_derivable[form_idx][cand_idx] = true;
            }
        }

        let torder = compute_torder(&self.patterns, nf);

        self.phase = FtPhase::Done;
        self.result = Some(FactorialTypologyResult {
            patterns: std::mem::take(&mut self.patterns),
            candidate_derivable,
            torder,
        });
    }
}

// ─── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tableau::Tableau;

    fn parse(text: &str) -> Tableau {
        Tableau::parse(text).expect("parse failed")
    }

    // Minimal 2-form, 2-constraint, 2-candidate-each tableau.
    // Input A: winner=a1, rival=a2
    //   a1: C1=0, C2=1
    //   a2: C1=1, C2=0
    // Input B: winner=b1, rival=b2
    //   b1: C1=0, C2=1
    //   b2: C1=1, C2=0
    //
    // Possible rankings:
    //   C1>>C2: A→a1, B→b1
    //   C2>>C1: A→a2, B→b2
    //
    // Expected patterns: [{a1,b1}, {a2,b2}]
    //
    // Note: rivals must have empty first column (blank input field) to be
    // parsed as continuation of the same input form.
    const TINY_FT: &str = "\t\t\tC1\tC2\n\t\t\tC1\tC2\nA\ta1\t1\t0\t1\n\ta2\t0\t1\t0\nB\tb1\t1\t0\t1\n\tb2\t0\t1\t0\n";

    #[test]
    fn test_fast_rcd_basic() {
        // Single form: winner=[0,1], rivals=[[1,0]] → C2>>C1 → success
        let winner = vec![0i32, 1];
        let rivals = vec![vec![1i32, 0]];
        assert!(fast_rcd(&[&winner], &[rivals], 2, &[]));
    }

    #[test]
    fn test_fast_rcd_failure() {
        // Single form: winner=[1,0,0], rivals=[[0,1,0],[0,0,1]]
        // winner violates C1 more than rival1 AND C2 more than rival2
        // This creates a cycle → failure
        let _winner = vec![1i32, 1, 0];
        let _rival1 = vec![0i32, 0, 1];
        let _rival2 = vec![0i32, 2, 0];
        // C1 demotes because winner violates C1 more than rival1
        // C2 demotes because winner violates C2 more than rival2 (1 > 0... wait, winner C2=1, rival2 C2=2)
        // Let me think: winner=[1,1,0], rival1=[0,0,1], rival2=[0,2,0]
        // Pair winner vs rival1: winner violates C1 more (1>0) → C1 demotable
        // Pair winner vs rival2: winner violates C2 more? (1 < 2) → no. C3: (0 < 0) no.
        // Hmm, this is derivable. Let me make a true failure case.
        // winner=[1,0], rival1=[0,1], rival2=[1,0] where rival2 is same as winner
        // winner=[1,1], rival1=[0,2], rival2=[2,0]
        // vs rival1: C1 winner(1) > rival1(0) → C1 demotable
        // vs rival2: C2 winner(1) > rival2(0) → C2 demotable
        // All demotable → failure
        let winner2 = vec![1i32, 1];
        let rival1_2 = vec![0i32, 2];
        let rival2_2 = vec![2i32, 0];
        assert!(!fast_rcd(
            &[&winner2],
            &[vec![rival1_2, rival2_2]],
            2,
            &[]
        ));
    }

    #[test]
    fn test_factorial_typology_two_form() {
        let tableau = parse(TINY_FT);
        let result = tableau.run_factorial_typology(&[]);
        // Should find exactly 2 patterns
        assert_eq!(result.patterns.len(), 2);
        // All candidates should be derivable
        for form_derivable in &result.candidate_derivable {
            for &d in form_derivable {
                assert!(d, "all candidates should be derivable in this symmetric case");
            }
        }
    }

    #[test]
    fn test_factorial_typology_tiny_example() {
        // Use the actual tiny example from examples/
        let text = include_str!("../../examples/TinyIllustrativeFile.txt");
        let tableau = Tableau::parse(text).expect("parse failed");
        let result = tableau.run_factorial_typology(&[]);
        // Tiny example has 4 constraints and several candidates per input.
        // We just check it runs without panicking and produces some patterns.
        assert!(!result.patterns.is_empty());
    }

    #[test]
    fn test_format_ftsum_symmetric() {
        // TINY_FT has 2 forms (A, B) and 2 patterns: {a1,b1} and {a2,b2}
        let tableau = parse(TINY_FT);
        let result = tableau.run_factorial_typology(&[]);
        let ftsum = result.format_ftsum(&tableau);

        let lines: Vec<&str> = ftsum.lines().collect();
        assert_eq!(lines[0], "/A/\t/B/");
        // Two patterns — order may vary but each line has tab-separated values
        assert_eq!(lines.len(), 3); // header + 2 patterns
        for line in &lines[1..] {
            let parts: Vec<&str> = line.split('\t').collect();
            assert_eq!(parts.len(), 2);
        }
    }

    #[test]
    fn test_format_compact_sum_deduplication() {
        // Tableau where two different inputs both map to "x" in all patterns
        // should collapse to 1 distinct output.
        // Input A: winner=x, rival=y;  Input B: winner=x, rival=y
        // Only pattern where both produce x is allowed (C1>>C2 or C2>>C1 same result).
        // Use TINY_FT and check: patterns {a1,b1} → both distinct; {a2,b2} → both distinct.
        let tableau = parse(TINY_FT);
        let result = tableau.run_factorial_typology(&[]);
        let compact = result.format_compact_sum(&tableau);

        let lines: Vec<&str> = compact.lines().collect();
        // Two patterns, both have 2 distinct outputs (a1 != b1, a2 != b2)
        assert_eq!(lines.len(), 2);
        for line in &lines {
            assert!(line.starts_with("2\t"));
        }
    }

    #[test]
    fn test_format_ftsum_tiny_example() {
        let text = include_str!("../../examples/TinyIllustrativeFile.txt");
        let tableau = Tableau::parse(text).expect("parse failed");
        let result = tableau.run_factorial_typology(&[]);
        let ftsum = result.format_ftsum(&tableau);

        let lines: Vec<&str> = ftsum.lines().collect();
        // Header: /a/\t/tat/\t/at/
        assert_eq!(lines[0], "/a/\t/tat/\t/at/");
        // 4 patterns
        assert_eq!(lines.len(), 5); // header + 4 patterns
        for line in &lines[1..] {
            let parts: Vec<&str> = line.split('\t').collect();
            assert_eq!(parts.len(), 3);
        }
    }

    #[test]
    fn test_format_compact_sum_tiny_example() {
        let text = include_str!("../../examples/TinyIllustrativeFile.txt");
        let tableau = Tableau::parse(text).expect("parse failed");
        let result = tableau.run_factorial_typology(&[]);
        let compact = result.format_compact_sum(&tableau);

        // Each line should start with a count and a tab
        let lines: Vec<&str> = compact.lines().collect();
        assert_eq!(lines.len(), 4); // 4 patterns, all distinct compact forms
        for line in lines {
            let mut parts = line.splitn(2, '\t');
            let count_str = parts.next().unwrap();
            let _rest = parts.next().unwrap();
            let count: usize = count_str.parse().expect("count should be integer");
            assert!(count >= 1);
        }
    }
}
