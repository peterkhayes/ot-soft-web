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
    winner_viols: &[&[usize]],
    rival_viols: &[Vec<Vec<usize>>],
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
                let winner_v: &[usize] = &form.candidates[cand_idx].violations;
                let rivals: Vec<Vec<usize>> = form
                    .candidates
                    .iter()
                    .enumerate()
                    .filter(|(j, _)| *j != cand_idx)
                    .map(|(_, c)| c.violations.clone())
                    .collect();

                let winner_slice: &[&[usize]] = &[winner_v];
                let rival_slice: &[Vec<Vec<usize>>] = &[rivals];

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
                    let mut winner_vecs: Vec<Vec<usize>> = Vec::with_capacity(test_n);
                    let mut rival_vecs: Vec<Vec<Vec<usize>>> = Vec::with_capacity(test_n);

                    // Forms already in pattern
                    for (fi, &selected) in old_pattern.iter().enumerate() {
                        let form = &self.forms[fi];
                        winner_vecs.push(form.candidates[selected].violations.clone());
                        let rivals: Vec<Vec<usize>> = form
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
                    let rivals: Vec<Vec<usize>> = new_form
                        .candidates
                        .iter()
                        .enumerate()
                        .filter(|(j, _)| *j != new_cand)
                        .map(|(_, c)| c.violations.clone())
                        .collect();
                    rival_vecs.push(rivals);

                    // Build slices for fast_rcd
                    let winner_slices: Vec<&[usize]> =
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
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        let mut out = String::new();
        let nf = tableau.forms.len();
        let nc = tableau.constraints.len();
        let np = self.patterns.len();

        // ── Header ────────────────────────────────────────────────────────────
        out.push_str(&format!(
            "Results of Factorial Typology Search for {}\n\n\n",
            filename
        ));

        let now = chrono::Local::now();
        out.push_str(&format!(
            "{}\n\n",
            now.format("%-m-%-d-%Y, %-I:%M %p")
                .to_string()
                .to_lowercase()
        ));
        out.push_str("OTSoft 2.7, release date 2/1/2026\n");
        out.push_str(&format!("Source file:  {}\n\n\n", filename));

        // Constraint list
        out.push_str("Constraints\n\n");
        for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
            out.push_str(&format!(
                "   {}. {:<42}{}\n",
                c_idx + 1,
                constraint.full_name(),
                constraint.abbrev()
            ));
        }
        out.push_str("\n\n");

        // ── Summary Information ───────────────────────────────────────────────
        out.push_str("Summary Information\n\n");

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

        out.push_str(&format!("There were {} different output patterns.\n\n", np));

        if np == 0 {
            out.push_str("No derivable output patterns were found.\n");
            return out;
        }

        // Pattern table — 4 patterns per block
        out.push_str("Forms marked as winners in the input file are marked with >.\n\n");

        // Find max input width for column alignment
        let max_input_width = tableau
            .forms
            .iter()
            .map(|f| f.input.len())
            .max()
            .unwrap_or(0);
        let first_col_width = max_input_width + 4; // "/ " + input + " /" + padding

        let block_size = 4;
        let mut block_start = 0;
        while block_start < np {
            let block_end = (block_start + block_size).min(np);
            let block = &self.patterns[block_start..block_end];

            // Calculate per-column widths
            let col_widths: Vec<usize> = block
                .iter()
                .enumerate()
                .map(|(bi, _)| {
                    // "Output #N" header
                    let header_len = format!("Output #{}", block_start + bi + 1).len();
                    // Max candidate name in this column
                    let max_cand: usize = block[bi]
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
                    // Mark with ">" if this candidate is the original winner (frequency > 0)
                    let marker = if form.candidates[cand_idx].frequency > 0 {
                        ">"
                    } else {
                        " "
                    };
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

        // ── List of Winners ────────────────────────────────────────────────────
        out.push_str("\n\nList of Winners\n\n");
        out.push_str(
            "The following specifies for each candidate whether there is at least one ranking that derives it:\n\n",
        );

        for (form_idx, form) in tableau.forms.iter().enumerate() {
            out.push_str(&format!("/{}/:\n", form.input));
            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                let derivable = self.candidate_derivable[form_idx][cand_idx];
                let marker = if cand.frequency > 0 { ">" } else { " " };
                let status = if derivable { "yes" } else { "no" };
                out.push_str(&format!(
                    "   {}[{:<12}]  {}\n",
                    marker, cand.form, status
                ));
            }
            out.push('\n');
        }

        // ── T-Order ────────────────────────────────────────────────────────────
        out.push_str("\n\nT-Orders\n\n");
        out.push_str("The t-order is the set of implications in a factorial typology.\n\n");

        // Find always-winning candidates (only one candidate wins for that form)
        let mut always_winners: Vec<(usize, usize)> = Vec::new(); // (form_idx, cand_idx)
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
            out.push_str(&format!("{:<20}  {}\n", "Input", "Candidate"));
            for (form_idx, cand_idx) in &always_winners {
                let form = &tableau.forms[*form_idx];
                out.push_str(&format!(
                    "/{:<18}/  [{}]\n",
                    form.input, form.candidates[*cand_idx].form
                ));
            }
            out.push('\n');
        }

        if self.torder.is_empty() {
            out.push_str("No t-order implications were found.\n");
        } else {
            out.push_str(&format!(
                "{:<22}  {:<22}  {:<22}  {}\n",
                "If this input", "has this output", "then this input", "has this output"
            ));
            for entry in &self.torder {
                let imp_form = &tableau.forms[entry.implicator_form];
                let imp_cand = &imp_form.candidates[entry.implicator_candidate];
                let ted_form = &tableau.forms[entry.implicated_form];
                let ted_cand = &ted_form.candidates[entry.implicated_candidate];
                out.push_str(&format!(
                    "/{:<20}/  [{:<20}]  /{:<20}/  [{}]\n",
                    imp_form.input, imp_cand.form, ted_form.input, ted_cand.form
                ));
            }
        }

        // Non-implicators
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
                    .filter(move |(ci, _)| {
                        self.candidate_derivable[fi][*ci]
                            && !imp_set_ref.contains(&(fi, *ci))
                    })
                    .map(move |(ci, _)| (fi, ci))
                    .collect::<Vec<_>>()
            })
            .collect();

        if !non_implicators.is_empty() {
            out.push_str("\nNothing is implicated by these input-output pairs:\n\n");
            out.push_str(&format!("{:<20}  {}\n", "Input", "Candidate"));
            for (form_idx, cand_idx) in &non_implicators {
                let form = &tableau.forms[*form_idx];
                out.push_str(&format!(
                    "/{:<18}/  [{}]\n",
                    form.input, form.candidates[*cand_idx].form
                ));
            }
        }

        out
    }

    /// Append the complete listing section: for each pattern, run RCD and show the grammar.
    pub fn format_full_listing(&self, tableau: &Tableau, apriori: &[Vec<bool>]) -> String {
        let nc = tableau.constraints.len();
        let np = self.patterns.len();

        if np == 0 {
            return String::new();
        }

        let mut out = String::new();

        out.push_str("\n\nComplete Listing of Output Patterns\n\n");

        // Width for aligning input column
        let max_input_width = tableau.forms.iter().map(|f| f.input.len()).max().unwrap_or(0);

        for (pat_idx, pattern) in self.patterns.iter().enumerate() {
            if pat_idx > 0 {
                out.push_str("\n\n------------------------------------------------------------------------------\n");
            }

            out.push_str(&format!("OUTPUT SET #{}:\n", pat_idx + 1));
            out.push_str(
                "These are the winning outputs.  > specifies outputs marked as winning candidates in the input file.\n\n",
            );

            for (fi, &ci) in pattern.iter().enumerate() {
                let form = &tableau.forms[fi];
                let cand = &form.candidates[ci];
                let input_padded = format!("/{}/", form.input);
                let is_actual = cand.frequency > 0;
                let marker = if is_actual { ">" } else { " " };
                let actual_label = if is_actual { "  (actual)" } else { "" };
                out.push_str(&format!(
                    "   {:<width$}  -->  {}{}{}\n",
                    input_padded,
                    marker,
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

        out.push('\n');
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
        let winner = vec![0usize, 1usize];
        let rivals = vec![vec![1usize, 0usize]];
        assert!(fast_rcd(&[&winner], &[rivals], 2, &[]));
    }

    #[test]
    fn test_fast_rcd_failure() {
        // Single form: winner=[1,0,0], rivals=[[0,1,0],[0,0,1]]
        // winner violates C1 more than rival1 AND C2 more than rival2
        // This creates a cycle → failure
        let _winner = vec![1usize, 1usize, 0usize];
        let _rival1 = vec![0usize, 0usize, 1usize];
        let _rival2 = vec![0usize, 2usize, 0usize];
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
        let winner2 = vec![1usize, 1usize];
        let rival1_2 = vec![0usize, 2usize];
        let rival2_2 = vec![2usize, 0usize];
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
        let text = include_str!("../../examples/tiny/input.txt");
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
        let text = include_str!("../../examples/tiny/input.txt");
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
        let text = include_str!("../../examples/tiny/input.txt");
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
