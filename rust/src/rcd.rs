//! Recursive Constraint Demotion (RCD) algorithm
//!
//! This module implements the RCD algorithm for finding stratified constraint
//! rankings in Optimality Theory. The algorithm iteratively ranks constraints
//! into strata by identifying which constraints never prefer losers.

use wasm_bindgen::prelude::*;
use crate::tableau::Tableau;
use crate::fred::FRedResult;
use crate::AxisMode;

/// Broken bar (¦, U+00A6) — used by VB6 as the separator in ranked constraint lists.
const BROKEN_BAR: char = '\u{00A6}';

// ── HTML output constants ────────────────────────────────────────────────────

/// Embedded CSS for the HTML tableau document.
const HTML_STYLE: &str = r#"<style>
body { font-family: sans-serif; padding: 1.5em; max-width: 1200px; margin: 0 auto; }
h1 { font-size: 1.1em; font-weight: bold; }
h2 { font-size: 1em; font-weight: bold; margin-top: 1.5em; }
pre { font-family: monospace; white-space: pre-wrap; font-size: 0.9em; }
table { border-collapse: collapse; margin: 1em 0; }
td, th { padding: 3px 10px; border: 1px solid #ddd; text-align: center;
         font-family: monospace; white-space: nowrap; }
th { font-weight: bold; background-color: #f5f5f5; }
td:first-child, th:first-child { text-align: left; border: 1px solid #ddd; }
.necessity-table td { border: none; padding: 1px 12px 1px 0; }
.cl4 { border-right: 2px solid #888; background-color: #CCCCCC; }
.cl8 { border-right: 2px solid #888; }
.cl9 { background-color: #CCCCCC; }
.cl10 { }
.stratum-break { border-top: 2px solid #888; }
.success { color: #2d7a2d; }
.failure { color: #cc0000; }
.warning { background: #fff8dc; border-left: 4px solid #f0c040; padding: 0.5em 1em; }
</style>"#;

/// Escape a string for safe inclusion in HTML content.
fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
     .replace('"', "&quot;")
}

/// Choose the CSS class for a violation cell based on position and shading state.
///
/// Mirrors VB6 classes cl4/cl8/cl9/cl10:
/// - cl4: shaded + right border (stratum separator)
/// - cl8: unshaded + right border (stratum separator)
/// - cl9: shaded, no right border
/// - cl10: unshaded, no right border
fn html_cell_class(is_shaded: bool, is_last_col: bool, has_right_border: bool) -> &'static str {
    match (is_last_col, is_shaded, has_right_border) {
        (true, true, _)   => "cl9",
        (true, false, _)  => "cl10",
        (false, true, true)  => "cl4",
        (false, false, true) => "cl8",
        (false, true, false)  => "cl9",
        (false, false, false) => "cl10",
    }
}

/// Format a violation count as HTML content using VB6's asterisk notation.
///
/// - 0 violations → `&nbsp;`
/// - 1–9 violations → repeated asterisks (e.g., `**` for 2)
/// - ≥10 violations → the number
/// - Fatal violations → asterisks up to `winner_viols + 1`, then `!`, then remaining
fn format_html_viol(viols: i32, winner_viols: i32, is_fatal: bool) -> String {
    if viols == 0 {
        return "&nbsp;".to_string();
    }
    if !(0..10).contains(&viols) {
        return if is_fatal { format!("{}!", viols) } else { viols.to_string() };
    }
    // Safe: we already returned for viols < 0 above
    let v = viols as usize;
    if is_fatal {
        let before = (winner_viols + 1).max(0) as usize;
        let after = v.saturating_sub(before);
        format!("{}!{}", "*".repeat(before), "*".repeat(after))
    } else {
        "*".repeat(v)
    }
}

// ────────────────────────────────────────────────────────────────────────────

/// Format a violation value centered in a column, matching VB6's centering formula.
///
/// VB6 centering: `leading = floor(col_width/2) - digit_count`, value, then
/// `trailing = col_width - floor(col_width/2)` for non-fatal, or `-1` for fatal (to
/// accommodate the `!`).
fn format_violation(col_width: usize, viols: i32, is_fatal: bool) -> String {
    if viols == 0 {
        return " ".repeat(col_width);
    }
    let effective_width = col_width.max(2);
    let half = effective_width / 2;
    let viol_str = viols.to_string();
    let digit_count = viol_str.len();
    let leading = half.saturating_sub(digit_count);
    let content = if is_fatal {
        let trailing = effective_width.saturating_sub(half + 1);
        format!("{}{viol_str}!{}", " ".repeat(leading), " ".repeat(trailing))
    } else {
        let trailing = effective_width - half;
        format!("{}{viol_str}{}", " ".repeat(leading), " ".repeat(trailing))
    };
    // Ensure we hit exactly col_width (pad or truncate trailing if rounding differs)
    if content.len() < col_width {
        format!("{}{}", content, " ".repeat(col_width - content.len()))
    } else {
        content[..col_width].to_string()
    }
}

/// Classification of constraint necessity
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ConstraintNecessity {
    Necessary,
    UnnecessaryButShownForFaithfulness,
    CompletelyUnnecessary,
}

/// A minimal-pair ranking argument: constraint `higher` must dominate `lower`,
/// evidenced by (form_index, rival_index).
struct MinimalPairEvidence {
    form_idx: usize,
    rival_idx: usize,
}

/// One form's filtered data for diagnostic tableaux.
struct DiagnosticEntry {
    form_idx: usize,
    winner_idx: usize,
    rival_indices: Vec<usize>,
    constraint_indices: Vec<usize>,
}

/// A mini-tableau showing a simplified winner-loser comparison
#[derive(Debug, Clone)]
pub struct MiniTableau {
    pub form_index: usize,
    pub winner_index: usize,
    pub loser_index: usize,
    pub included_constraints: Vec<usize>,
}

/// Result of running RCD algorithm
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct RCDResult {
    /// Stratum number for each constraint (1-indexed)
    constraint_strata: Vec<usize>,
    /// Total number of strata
    num_strata: usize,
    /// Whether a valid ranking was found
    success: bool,
    /// Necessity classification for each constraint (not exposed to WASM)
    #[wasm_bindgen(skip)]
    constraint_necessity: Vec<ConstraintNecessity>,
    /// FRed ranking-argumentation result (not exposed to WASM)
    #[wasm_bindgen(skip)]
    fred_result: Option<FRedResult>,
    /// Mini-tableaux for ranking arguments (not exposed to WASM)
    #[wasm_bindgen(skip)]
    mini_tableaux: Vec<MiniTableau>,
    /// Whether BCD encountered tied faithfulness subsets (arbitrary choice made)
    tie_warning: bool,
}

#[wasm_bindgen]
impl RCDResult {
    pub fn num_strata(&self) -> usize {
        self.num_strata
    }

    pub fn success(&self) -> bool {
        self.success
    }

    pub fn get_stratum(&self, constraint_index: usize) -> Option<usize> {
        self.constraint_strata.get(constraint_index).copied()
    }

    pub fn tie_warning(&self) -> bool {
        self.tie_warning
    }
}

impl RCDResult {
    /// Create a new RCDResult (used by BCD to construct results)
    pub(crate) fn new(constraint_strata: Vec<usize>, num_strata: usize, success: bool) -> Self {
        RCDResult {
            constraint_strata,
            num_strata,
            success,
            constraint_necessity: Vec::new(),
            fred_result: None,
            mini_tableaux: Vec::new(),
            tie_warning: false,
        }
    }

    /// Compute additional analyses (necessity, FRed ranking arguments, mini-tableaux)
    /// Compute constraint necessity, FRed, and mini-tableaux.
    ///
    /// `include_apriori_in_fred`: whether to feed apriori ERCs into FRed.
    /// RCD passes `true` (VB6 includes them); LFCD passes `false` (VB6 does not).
    pub(crate) fn compute_extra_analyses(
        &mut self,
        tableau: &Tableau,
        apriori: &[Vec<bool>],
        include_apriori_in_fred: bool,
    ) {
        self.constraint_necessity = tableau.compute_constraint_necessity(self, apriori);
        self.fred_result = Some(if apriori.is_empty() || !include_apriori_in_fred {
            tableau.run_fred(false)
        } else {
            tableau.run_fred_with_apriori(false, apriori)
        });
        self.mini_tableaux = self.generate_mini_tableaux(tableau);
    }

    pub(crate) fn set_tie_warning(&mut self, value: bool) {
        self.tie_warning = value;
    }

    /// Override FRed options after the initial computation.
    ///
    /// Called by format functions to apply user-specified argumentation options.
    pub(crate) fn apply_fred_options(
        &mut self,
        tableau: &Tableau,
        opts: &crate::FredOptions,
    ) {
        if !opts.include_fred {
            self.fred_result = None;
            self.mini_tableaux = Vec::new();
            return;
        }

        // Re-run FRed if options differ from the default (SB, no verbose).
        // Default was computed in compute_extra_analyses as run_fred(false).
        if opts.use_mib || opts.show_details {
            self.fred_result = Some(tableau.run_fred_verbose(opts.use_mib, opts.show_details));
        }

        if !opts.include_mini_tableaux {
            self.mini_tableaux = Vec::new();
        }
    }

    /// Generate mini-tableaux showing simplified ranking arguments
    fn generate_mini_tableaux(&self, tableau: &Tableau) -> Vec<MiniTableau> {
        let mut mini_tableaux = Vec::new();

        for (form_idx, form) in tableau.forms.iter().enumerate() {
            // Find the winner
            let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
                Some(idx) => idx,
                None => continue,
            };

            let winner = &form.candidates[winner_idx];

            // Check each loser
            for (loser_idx, loser) in form.candidates.iter().enumerate() {
                if loser.frequency > 0 {
                    continue; // Skip winners
                }

                // Count constraints preferring each candidate
                let mut winner_preferring = Vec::new();
                let mut loser_preferring = Vec::new();
                let mut included_constraints = Vec::new();

                for c_idx in 0..tableau.constraints.len() {
                    let w_viol = winner.violations[c_idx];
                    let l_viol = loser.violations[c_idx];

                    // Include column if either candidate violates
                    if w_viol == 0 && l_viol == 0 {
                        continue;
                    }

                    included_constraints.push(c_idx);

                    if w_viol < l_viol {
                        winner_preferring.push(c_idx);
                    } else if l_viol < w_viol {
                        loser_preferring.push(c_idx);
                    }
                }

                // Include if exactly one winner-preferring and at least one loser-preferring
                if winner_preferring.len() == 1 && !loser_preferring.is_empty() {
                    mini_tableaux.push(MiniTableau {
                        form_index: form_idx,
                        winner_index: winner_idx,
                        loser_index: loser_idx,
                        included_constraints,
                    });
                }
            }
        }

        mini_tableaux
    }

    /// Returns constraint indices reordered by stratum, matching VB6's exact
    /// `PrintTableaux.bas:SortTheConstraints` algorithm.
    ///
    /// VB6 uses an O(n²) selection-sort that swaps whenever an inner element has
    /// a strictly smaller stratum than the current outer element. This is **not**
    /// a stable sort — equal-stratum constraints can end up in a different relative
    /// order than they appear in the input file (see task
    /// `conformance-ilokano-constraint-order.md`).
    pub(crate) fn sorted_constraint_indices(&self, num_constraints: usize) -> Vec<usize> {
        let mut indices: Vec<usize> = (0..num_constraints).collect();
        self.vb6_sort_constraint_slice(&mut indices);
        indices
    }

    /// Sort a slice of global constraint indices in-place using VB6's unstable selection sort
    /// (stratum ascending). Equivalent to VB6's `SortTheConstraints`.
    fn vb6_sort_constraint_slice(&self, indices: &mut [usize]) {
        let strata = &self.constraint_strata;
        let n = indices.len();
        for outer in 0..n {
            for inner in (outer + 1)..n {
                let s_outer = strata.get(indices[outer]).copied().unwrap_or(usize::MAX);
                let s_inner = strata.get(indices[inner]).copied().unwrap_or(usize::MAX);
                if s_inner < s_outer {
                    indices.swap(outer, inner);
                }
            }
        }
    }

    /// Returns candidate indices sorted for tableau display: winner first (index 0),
    /// followed by rivals in harmony order (fewer violations through sorted constraints).
    ///
    /// Matches VB6's `PrintTableaux.bas:SortTheCandidates` behaviour.
    pub(crate) fn sorted_candidate_indices(
        &self,
        form: &crate::tableau::InputForm,
        sorted_constraints: &[usize],
    ) -> Vec<usize> {
        let winner_idx = form.candidates.iter().position(|c| c.frequency > 0);

        let mut rival_indices: Vec<usize> = (0..form.candidates.len())
            .filter(|&i| Some(i) != winner_idx)
            .collect();

        rival_indices.sort_by(|&a, &b| {
            for &c_idx in sorted_constraints {
                let va = form.candidates[a].violations[c_idx];
                let vb = form.candidates[b].violations[c_idx];
                match va.cmp(&vb) {
                    std::cmp::Ordering::Less => return std::cmp::Ordering::Less,
                    std::cmp::Ordering::Greater => return std::cmp::Ordering::Greater,
                    std::cmp::Ordering::Equal => continue,
                }
            }
            std::cmp::Ordering::Equal
        });

        let mut result = Vec::with_capacity(form.candidates.len());
        if let Some(wi) = winner_idx {
            result.push(wi);
        }
        result.extend(rival_indices);
        result
    }

    /// Generate formatted text output for the RCD analysis
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        self.format_output_with_options(tableau, filename, "Recursive Constraint Demotion", &[], true, false)
    }

    /// Generate formatted text output with a configurable algorithm name and a priori table.
    /// Does not include the A Priori Rankings section (used by BCD, which takes no a priori input).
    pub(crate) fn format_output_with_algorithm(&self, tableau: &Tableau, filename: &str, algorithm_name: &str, diagnostics: bool) -> String {
        self.format_output_with_options(tableau, filename, algorithm_name, &[], false, diagnostics)
    }

    /// Generate formatted text output with full options.
    ///
    /// `show_apriori_section`: whether to include the "A Priori Rankings" section.
    /// RCD includes it; LFCD does not (VB6 behaviour).
    /// `diagnostics`: whether to include diagnostic output when ranking fails.
    pub(crate) fn format_output_with_options(
        &self,
        tableau: &Tableau,
        filename: &str,
        algorithm_name: &str,
        apriori: &[Vec<bool>],
        show_apriori_section: bool,
        diagnostics: bool,
    ) -> String {
        let mut output = String::new();
        let mut section = 0usize;

        self.fmt_text_header(&mut output, algorithm_name, filename);
        self.fmt_text_result(&mut output, &mut section, tableau);

        if show_apriori_section && !apriori.is_empty() {
            Self::fmt_text_apriori(&mut output, &mut section, tableau, apriori);
        }

        self.fmt_text_tableaux(&mut output, &mut section, tableau);

        if !self.constraint_necessity.is_empty() {
            self.fmt_text_necessity(&mut output, &mut section, tableau, apriori);
        }

        if let Some(ref fred) = self.fred_result {
            section += 1;
            output.push_str(&fred.format_section_fred(section));
        }

        if !self.mini_tableaux.is_empty() {
            self.fmt_text_mini_tableaux(&mut output, &mut section, tableau);
        }

        if !self.success && diagnostics {
            self.fmt_text_diagnostics(&mut output, &mut section, tableau);
        }

        // Normalize trailing newlines to match VB6 output (exactly one trailing blank line)
        let trimmed = output.trim_end_matches('\n');
        format!("{}\n\n", trimmed)
    }

    fn fmt_text_header(&self, out: &mut String, algorithm_name: &str, filename: &str) {
        out.push_str(&format!("Results of Applying {} to {}\n", algorithm_name, filename));
        out.push_str("\n\n");

        let now = chrono::Local::now();
        out.push_str(&format!("{}\n\n", now.format("%-m-%-d-%Y, %-I:%M %p").to_string().to_lowercase()));
        out.push_str(crate::VERSION_STRING);
        out.push('\n');
        out.push_str("\n\n");

        if self.tie_warning {
            out.push_str("Caution: The BCD algorithm has selected arbitrarily among tied Faithfulness constraint subsets.\n");
            out.push_str("You may wish to try changing the order of the Faithfulness constraints in the input file,\n");
            out.push_str("to see whether this results in a different ranking.\n\n\n");
        }
    }

    fn fmt_text_result(&self, out: &mut String, section: &mut usize, tableau: &Tableau) {
        *section += 1;
        out.push_str(&format!("{}. Result\n\n", section));

        if self.success {
            out.push_str("A ranking was found that generates the correct outputs.\n\n");
        } else {
            out.push_str("No ranking was found.\n\n");
        }

        for stratum in 1..=self.num_strata {
            out.push_str(&format!("   Stratum #{}\n", stratum));
            for (c_idx, &c_stratum) in self.constraint_strata.iter().enumerate() {
                if c_stratum == stratum {
                    if let Some(constraint) = tableau.get_constraint(c_idx) {
                        out.push_str(&format!("      {:<42}{}\n", constraint.full_name(), constraint.abbrev()));
                    }
                }
            }
        }
        out.push('\n');
    }

    fn fmt_text_apriori(out: &mut String, section: &mut usize, tableau: &Tableau, apriori: &[Vec<bool>]) {
        *section += 1;
        out.push_str(&format!("{}. A Priori Rankings\n\n", section));
        out.push_str("In the following table, \"yes\" means that the constraint of the indicated row \n");
        out.push_str("was marked a priori to dominate the constraint in the given column.\n\n");

        let nc = tableau.constraints.len();
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        let col_widths: Vec<usize> = abbrevs.iter().map(|a| a.len().max(3)).collect();
        let row_label_width = abbrevs.iter().map(|a| a.len()).max().unwrap_or(0);

        // Header row
        out.push_str(&format!("{:width$}", "", width = row_label_width));
        for (j, abbrev) in abbrevs.iter().enumerate() {
            out.push_str(&format!("  {:^width$}", abbrev, width = col_widths[j]));
        }
        out.push('\n');

        // Data rows
        for i in 0..nc {
            out.push_str(&format!("{:<width$}", abbrevs[i], width = row_label_width));
            for j in 0..nc {
                let cell = if apriori[i][j] { "yes" } else { "" };
                out.push_str(&format!("  {:^width$}", cell, width = col_widths[j]));
            }
            out.push('\n');
        }
        out.push_str("\n\n");
    }

    fn fmt_text_tableaux(&self, out: &mut String, section: &mut usize, tableau: &Tableau) {
        *section += 1;
        out.push_str(&format!("{}. Tableaux\n\n", section));

        let sorted = self.sorted_constraint_indices(tableau.constraints.len());

        for form in &tableau.forms {
            out.push('\n');
            out.push_str(&format!("/{}/: \n", form.input));

            let max_cand_width = form.candidates.iter()
                .map(|c| c.form.len())
                .max()
                .unwrap_or(0)
                .max(2);

            // Build constraint header with stratum separators
            let mut header = String::new();
            for _ in 0..max_cand_width + 2 {
                header.push(' ');
            }
            for (pos, &c_idx) in sorted.iter().enumerate() {
                if pos > 0 {
                    let prev_stratum = self.constraint_strata[sorted[pos - 1]];
                    if self.constraint_strata[c_idx] != prev_stratum {
                        header.push('|');
                    } else {
                        header.push(BROKEN_BAR);
                    }
                }
                header.push_str(&tableau.constraints[c_idx].abbrev());
            }
            out.push_str(&header);
            out.push('\n');

            let winner_idx = form.candidates.iter().position(|c| c.frequency > 0);
            let sorted_cands = self.sorted_candidate_indices(form, &sorted);

            for &orig_cand_idx in &sorted_cands {
                let candidate = &form.candidates[orig_cand_idx];
                let is_winner = Some(orig_cand_idx) == winner_idx;
                let marker = if is_winner { ">" } else { " " };

                let first_fatal_idx = if !is_winner {
                    winner_idx.and_then(|wi| {
                        let winner = &form.candidates[wi];
                        sorted.iter().position(|&c_idx| candidate.violations[c_idx] > winner.violations[c_idx])
                    })
                } else {
                    None
                };

                out.push_str(&format!("{}{:<width$} ", marker, candidate.form, width = max_cand_width));

                for (pos, &c_idx) in sorted.iter().enumerate() {
                    let viols = candidate.violations[c_idx];
                    let col_width = tableau.constraints[c_idx].abbrev().len();

                    if pos > 0 {
                        let prev_stratum = self.constraint_strata[sorted[pos - 1]];
                        if self.constraint_strata[c_idx] != prev_stratum {
                            out.push('|');
                        } else {
                            out.push(BROKEN_BAR);
                        }
                    }

                    let is_fatal = first_fatal_idx == Some(pos);
                    out.push_str(&format_violation(col_width, viols, is_fatal));
                }
                out.push('\n');
            }
            out.push('\n');
        }
        out.push('\n');
    }

    fn fmt_text_necessity(&self, out: &mut String, section: &mut usize, tableau: &Tableau, apriori: &[Vec<bool>]) {
        *section += 1;
        out.push_str(&format!("{}. Status of Proposed Constraints:  Necessary or Unnecessary\n\n", section));

        let max_abbrev_width = tableau.constraints.iter()
            .map(|c| c.abbrev().len())
            .max()
            .unwrap_or(0);

        let category_order = [
            ConstraintNecessity::Necessary,
            ConstraintNecessity::UnnecessaryButShownForFaithfulness,
            ConstraintNecessity::CompletelyUnnecessary,
        ];
        for category in &category_order {
            for (c_idx, necessity) in self.constraint_necessity.iter().enumerate() {
                if necessity != category {
                    continue;
                }
                let constraint = &tableau.constraints[c_idx];
                let status = match necessity {
                    ConstraintNecessity::Necessary => "Necessary",
                    ConstraintNecessity::UnnecessaryButShownForFaithfulness =>
                        "Not necessary (but included to show Faithfulness violations\n              of a winning candidate)",
                    ConstraintNecessity::CompletelyUnnecessary => "Not necessary",
                };
                out.push_str(&format!(
                    "   {:<width$}  {}\n",
                    constraint.abbrev(),
                    status,
                    width = max_abbrev_width,
                ));
            }
        }

        let num_deletable = self.constraint_necessity.iter()
            .filter(|n| **n != ConstraintNecessity::Necessary)
            .count();
        if num_deletable >= 2 {
            let mass_deletion_possible = self.check_mass_deletion(tableau, apriori);
            if mass_deletion_possible {
                out.push_str("\nA check has determined that the grammar will still work even if the \n");
                out.push_str("constraints marked above as unnecessary are removed en masse.\n\n\n");
            } else {
                out.push_str("\n\nA check has determined that, although the grammar will still work with the\n");
                out.push_str("removal of ANY ONE of the constraints marked above as unnecessary, the\n");
                out.push_str("grammar will NOT work if they are removed en masse.\n\n\n");
            }
        } else {
            out.push_str("\n\n");
        }
    }

    fn fmt_text_mini_tableaux(&self, out: &mut String, section: &mut usize, tableau: &Tableau) {
        *section += 1;
        out.push_str(&format!("{}. Mini-Tableaux\n\n", section));
        out.push_str("The following small tableaux may be useful in presenting ranking arguments. \n");
        out.push_str("They include all winner-rival comparisons in which there is just one \n");
        out.push_str("winner-preferring constraint and at least one loser-preferring constraint.  \n");
        out.push_str("Constraints not violated by either candidate are omitted.\n\n");

        for mini in &self.mini_tableaux {
            self.format_mini_tableau(tableau, mini, out);
        }
    }

    /// Check if mass deletion of unnecessary constraints still allows RCD to succeed
    fn check_mass_deletion(&self, tableau: &Tableau, apriori: &[Vec<bool>]) -> bool {
        use crate::tableau::{InputForm, Candidate};

        // Create a modified tableau with all unnecessary constraints removed
        let forms = tableau.forms.iter().map(|form| {
            let candidates = form.candidates.iter().map(|cand| {
                let mut violations = cand.violations.clone();
                for (c_idx, necessity) in self.constraint_necessity.iter().enumerate() {
                    if *necessity != ConstraintNecessity::Necessary {
                        violations[c_idx] = 0;
                    }
                }
                Candidate {
                    form: cand.form.clone(),
                    frequency: cand.frequency,
                    violations,
                }
            }).collect();

            InputForm {
                input: form.input.clone(),
                candidates,
            }
        }).collect();

        let modified_tableau = Tableau {
            constraints: tableau.constraints.clone(),
            forms,
        };

        // Run RCD on modified tableau (without computing extra analyses to avoid recursion)
        let test_result = modified_tableau.run_rcd_internal(false, apriori);
        test_result.success
    }

    /// Format a mini-tableau
    fn format_mini_tableau(&self, tableau: &Tableau, mini: &MiniTableau, output: &mut String) {
        let form = &tableau.forms[mini.form_index];
        let winner = &form.candidates[mini.winner_index];
        let loser = &form.candidates[mini.loser_index];

        output.push_str(&format!("\n/{}/: \n", form.input));

        // Sort included constraints by stratum using VB6's unstable selection sort,
        // matching VB6's PrepareMiniTableaux which collects in input order then calls SortTheConstraints.
        let mut sorted_constraints = mini.included_constraints.clone();
        self.vb6_sort_constraint_slice(&mut sorted_constraints);

        // Build header with only included constraints

        let max_cand_width = winner.form.len().max(loser.form.len()).max(2);

        let mut header = String::new();
        for _ in 0..max_cand_width + 2 {
            header.push(' ');
        }

        for (i, &c_idx) in sorted_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];

            // Add separator before this constraint
            if i > 0 {
                let prev_c_idx = sorted_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    header.push('|');
                } else {
                    header.push(BROKEN_BAR);
                }
            }

            header.push_str(&constraint.abbrev());
        }
        output.push_str(&header);
        output.push('\n');

        // Output winner
        output.push_str(&format!(">{:<width$} ", winner.form, width = max_cand_width));
        for (i, &c_idx) in sorted_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];
            let col_width = constraint.abbrev().len();

            // Add separator
            if i > 0 {
                let prev_c_idx = sorted_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    output.push('|');
                } else {
                    output.push(BROKEN_BAR);
                }
            }

            output.push_str(&format_violation(col_width, winner.violations[c_idx], false));
        }
        output.push('\n');

        // Output loser (without fatal violation markers in mini-tableaux)
        output.push_str(&format!(" {:<width$} ", loser.form, width = max_cand_width));

        for (i, &c_idx) in sorted_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];
            let col_width = constraint.abbrev().len();

            // Add separator
            if i > 0 {
                let prev_c_idx = sorted_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    output.push('|');
                } else {
                    output.push(BROKEN_BAR);
                }
            }

            output.push_str(&format_violation(col_width, loser.violations[c_idx], false));
        }
        output.push_str("\n\n");
    }

    // ── HTML output ──────────────────────────────────────────────────────────

    /// Generate an HTML document containing styled tableaux for this RCD result.
    pub fn format_html_output(&self, tableau: &Tableau, filename: &str) -> String {
        self.format_html_output_full(tableau, filename, "Recursive Constraint Demotion", AxisMode::default(), &[], false)
    }

    /// Generate an HTML document with configurable algorithm name and axis mode.
    pub(crate) fn format_html_output_with_options(
        &self,
        tableau: &Tableau,
        filename: &str,
        algorithm_name: &str,
        axis_mode: AxisMode,
        diagnostics: bool,
    ) -> String {
        self.format_html_output_full(tableau, filename, algorithm_name, axis_mode, &[], diagnostics)
    }

    /// Generate an HTML document with full options including a priori data.
    pub(crate) fn format_html_output_full(
        &self,
        tableau: &Tableau,
        filename: &str,
        algorithm_name: &str,
        axis_mode: AxisMode,
        apriori: &[Vec<bool>],
        diagnostics: bool,
    ) -> String {
        let mut out = String::new();
        let mut section = 0usize;

        self.fmt_html_header(&mut out, algorithm_name, filename);
        self.fmt_html_result(&mut out, &mut section, tableau);

        if !apriori.is_empty() {
            Self::fmt_html_apriori(&mut out, &mut section, tableau, apriori);
        }

        self.fmt_html_tableaux(&mut out, &mut section, tableau, axis_mode);

        if !self.constraint_necessity.is_empty() {
            self.fmt_html_necessity(&mut out, &mut section, tableau, apriori);
        }

        if let Some(ref fred) = self.fred_result {
            Self::fmt_html_fred(&mut out, &mut section, fred);
        }

        if !self.mini_tableaux.is_empty() {
            self.fmt_html_mini_tableaux(&mut out, &mut section, tableau);
        }

        if !self.success && diagnostics {
            self.fmt_html_diagnostics(&mut out, &mut section, tableau);
        }

        out.push_str("</body>\n</html>\n");
        out
    }

    fn fmt_html_header(&self, out: &mut String, algorithm_name: &str, filename: &str) {
        out.push_str("<!DOCTYPE html>\n<html>\n<head>\n");
        out.push_str("<meta charset=\"UTF-8\">\n");
        out.push_str(&format!(
            "<title>{} {}</title>\n",
            crate::VERSION_STRING,
            html_escape(filename)
        ));
        out.push_str(HTML_STYLE);
        out.push_str("\n</head>\n<body>\n");

        out.push_str(&format!(
            "<h1>Results of Applying {} to {}</h1>\n",
            html_escape(algorithm_name),
            html_escape(filename),
        ));
        let now = chrono::Local::now();
        let timestamp = now.format("%-m-%-d-%Y, %-I:%M %p").to_string().to_lowercase();
        out.push_str(&format!("<p>{}</p>\n", html_escape(&timestamp)));
        out.push_str(&format!("<p>{}</p>\n", crate::VERSION_STRING));

        if self.tie_warning {
            out.push_str(
                "<p class=\"warning\">Caution: The BCD algorithm has selected arbitrarily \
                 among tied Faithfulness constraint subsets. You may wish to try changing \
                 the order of the Faithfulness constraints in the input file, to see whether \
                 this results in a different ranking.</p>\n",
            );
        }
    }

    fn fmt_html_result(&self, out: &mut String, section: &mut usize, tableau: &Tableau) {
        *section += 1;
        out.push_str(&format!("<h2>{}. Result</h2>\n", section));
        if self.success {
            out.push_str("<p class=\"success\">A ranking was found that generates the correct outputs.</p>\n");
        } else {
            out.push_str("<p class=\"failure\">No ranking was found.</p>\n");
        }
        out.push_str("<table>\n");
        out.push_str("  <tr><td><b>Stratum</b></td><td><b>Constraint Name</b></td><td><b>Abbreviation</b></td></tr>\n");
        for stratum in 1..=self.num_strata {
            let mut first_in_stratum = true;
            for (c_idx, &c_stratum) in self.constraint_strata.iter().enumerate() {
                if c_stratum == stratum {
                    if let Some(constraint) = tableau.get_constraint(c_idx) {
                        let stratum_label = if first_in_stratum {
                            format!("Stratum #{stratum}")
                        } else {
                            "&nbsp;".to_string()
                        };
                        out.push_str(&format!(
                            "  <tr><td>{}</td><td>{}</td><td>{}</td></tr>\n",
                            stratum_label,
                            html_escape(&constraint.full_name()),
                            html_escape(&constraint.abbrev()),
                        ));
                        first_in_stratum = false;
                    }
                }
            }
        }
        out.push_str("</table>\n");
    }

    fn fmt_html_apriori(out: &mut String, section: &mut usize, tableau: &Tableau, apriori: &[Vec<bool>]) {
        *section += 1;
        out.push_str(&format!("<h2>{}. A Priori Rankings</h2>\n", section));
        out.push_str(
            "<p>In the following table, &quot;yes&quot; means that the constraint of the indicated row \
             was marked a priori to dominate the constraint in the given column.</p>\n",
        );
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        out.push_str("<table>\n  <tr>\n    <th></th>\n");
        for abbrev in &abbrevs {
            out.push_str(&format!("    <th>{}</th>\n", html_escape(abbrev)));
        }
        out.push_str("  </tr>\n");
        for (i, row) in apriori.iter().enumerate() {
            out.push_str(&format!("  <tr>\n    <th>{}</th>\n", html_escape(&abbrevs[i])));
            for &dominates in row {
                let cell = if dominates { "yes" } else { "" };
                out.push_str(&format!("    <td>{}</td>\n", cell));
            }
            out.push_str("  </tr>\n");
        }
        out.push_str("</table>\n");
    }

    fn fmt_html_tableaux(&self, out: &mut String, section: &mut usize, tableau: &Tableau, axis_mode: AxisMode) {
        *section += 1;
        out.push_str(&format!("<h2>{}. Tableaux</h2>\n", section));

        let total_constraint_length: usize = tableau.constraints.iter()
            .map(|c| c.abbrev().len() + 1)
            .sum();

        let sorted = self.sorted_constraint_indices(tableau.constraints.len());
        for form in &tableau.forms {
            let winner_idx = form.candidates.iter().position(|c| c.frequency > 0);
            let use_reversed = match axis_mode {
                AxisMode::SwitchAll => true,
                AxisMode::NeverSwitch => false,
                AxisMode::SwitchWhereNeeded => {
                    if total_constraint_length <= 75 {
                        false
                    } else {
                        let total_candidate_length: usize = form.candidates.iter()
                            .map(|c| c.form.len() + 2)
                            .sum();
                        total_candidate_length < total_constraint_length + 5
                    }
                }
            };
            if use_reversed {
                out.push_str(&self.format_html_reversed_form_table(tableau, form, winner_idx, &sorted));
            } else {
                out.push_str(&self.format_html_form_table(tableau, form, winner_idx, &sorted));
            }
        }
    }

    fn fmt_html_necessity(&self, out: &mut String, section: &mut usize, tableau: &Tableau, apriori: &[Vec<bool>]) {
        *section += 1;
        out.push_str(&format!(
            "<h2>{}. Status of Proposed Constraints: Necessary or Unnecessary</h2>\n",
            section,
        ));
        out.push_str("<table class=\"necessity-table\">\n<tbody>\n");
        out.push_str("  <tr><td><b>Constraint</b></td><td><b>Status</b></td></tr>\n");
        let category_order = [
            ConstraintNecessity::Necessary,
            ConstraintNecessity::UnnecessaryButShownForFaithfulness,
            ConstraintNecessity::CompletelyUnnecessary,
        ];
        for category in &category_order {
            for (c_idx, necessity) in self.constraint_necessity.iter().enumerate() {
                if necessity != category {
                    continue;
                }
                let constraint = &tableau.constraints[c_idx];
                let status = match necessity {
                    ConstraintNecessity::Necessary => "Necessary",
                    ConstraintNecessity::UnnecessaryButShownForFaithfulness =>
                        "Not necessary (but included to show Faithfulness violations of a winning candidate)",
                    ConstraintNecessity::CompletelyUnnecessary => "Not necessary",
                };
                out.push_str(&format!(
                    "  <tr><td>{}</td><td>{}</td></tr>\n",
                    html_escape(&constraint.full_name()),
                    html_escape(status),
                ));
            }
        }
        out.push_str("</tbody>\n</table>\n");
        let num_deletable = self.constraint_necessity.iter()
            .filter(|n| **n != ConstraintNecessity::Necessary)
            .count();
        if num_deletable >= 2 {
            if self.check_mass_deletion(tableau, apriori) {
                out.push_str(
                    "<p>A check has determined that the grammar will still work even if the \
                     constraints marked above as unnecessary are removed en masse.</p>\n",
                );
            } else {
                out.push_str(
                    "<p>A check has determined that, although the grammar will still work with the \
                     removal of ANY ONE of the constraints marked above as unnecessary, the \
                     grammar will NOT work if they are removed en masse.</p>\n",
                );
            }
        }
    }

    fn fmt_html_fred(out: &mut String, section: &mut usize, fred: &FRedResult) {
        *section += 1;
        out.push_str(&format!(
            "<h2>{}. Ranking Arguments, based on the Fusional Reduction Algorithm</h2>\n",
            section,
        ));

        let basis_name = if fred.use_skeletal_basis() {
            "Skeletal Basis"
        } else {
            "Most Informative Basis"
        };
        let purpose = if fred.use_skeletal_basis() {
            "keep each final ranking argument as pithy as possible"
        } else {
            "minimize the set of final ranking arguments"
        };
        out.push_str(&format!(
            "<p>This run sought to obtain the {basis_name}, intended to {purpose}.</p>\n"
        ));

        if fred.failure() {
            out.push_str(
                "<p>The constraints cannot be ranked to yield the desired outcomes.</p>\n",
            );
        }

        if !fred.detail_text.is_empty() {
            let basis_label = if fred.use_skeletal_basis() {
                "Skeletal Basis"
            } else {
                "Most Informative Basis"
            };
            out.push_str(&format!(
                "<pre>{}</pre>\n<p>Ranking argumentation: Final result</p>\n\
                 <p>The following set of ERCs forms the {} for the ERC set as a whole, \
                 and thus encapsulates the available ranking information.</p>\n",
                html_escape(&fred.detail_text),
                html_escape(basis_label),
            ));
        }

        out.push_str("<p>The final rankings obtained are as follows:</p>\n");
        out.push_str("<table>\n");
        for ranking in fred.ranking_strings() {
            out.push_str(&format!(
                "  <tr><td>{}</td><td>&nbsp;</td></tr>\n",
                html_escape(&ranking)
            ));
        }
        out.push_str("</table>\n");
    }

    fn fmt_html_mini_tableaux(&self, out: &mut String, section: &mut usize, tableau: &Tableau) {
        *section += 1;
        out.push_str(&format!("<h2>{}. Mini-Tableaux</h2>\n", section));
        out.push_str(
            "<p>The following small tableaux may be useful in presenting ranking arguments. \
             They include all winner-rival comparisons in which there is just one \
             winner-preferring constraint and at least one loser-preferring constraint. \
             Constraints not violated by either candidate are omitted.</p>\n",
        );
        for mini in &self.mini_tableaux {
            out.push_str(&self.format_html_mini_tableau(tableau, mini));
        }
    }

    /// Render a single input form as an HTML tableau table.
    ///
    /// `sorted` is the list of constraint indices sorted by stratum (from
    /// `sorted_constraint_indices`), matching VB6's `SortTheConstraints`.
    fn format_html_form_table(
        &self,
        tableau: &Tableau,
        form: &crate::tableau::InputForm,
        winner_idx: Option<usize>,
        sorted: &[usize],
    ) -> String {
        let num_constraints = sorted.len();
        let mut out = String::new();

        out.push_str("<table>\n");

        // Header row: input form + constraint abbreviations (in sorted stratum order)
        out.push_str("  <tr>\n");
        out.push_str(&format!("    <th>/{}/</th>\n", html_escape(&form.input)));
        for (pos, &c_idx) in sorted.iter().enumerate() {
            let is_last = pos + 1 == num_constraints;
            let next_stratum = if is_last { usize::MAX } else { self.constraint_strata[sorted[pos + 1]] };
            let has_border = !is_last && self.constraint_strata[c_idx] != next_stratum;
            let class = html_cell_class(false, is_last, has_border);
            out.push_str(&format!(
                "    <th class=\"{}\">{}</th>\n",
                class,
                html_escape(&tableau.constraints[c_idx].abbrev()),
            ));
        }
        out.push_str("  </tr>\n");

        // Sort candidates: winner first, rivals by harmony (matches VB6 SortTheCandidates)
        let sorted_cands = self.sorted_candidate_indices(form, sorted);

        // Compute winner shading point (sorted position where all losers are dead).
        // Cells strictly after this position are shaded in the winner row.
        let winner_shading_point: usize = if let Some(wi) = winner_idx {
            let winner = &form.candidates[wi];
            let losers: Vec<usize> = sorted_cands.iter().copied()
                .filter(|&i| i != wi)
                .collect();
            if losers.is_empty() {
                usize::MAX
            } else {
                let mut dead = vec![false; losers.len()];
                let mut found = usize::MAX;
                for (pos, &c_idx) in sorted.iter().enumerate() {
                    for (l_pos, &l_idx) in losers.iter().enumerate() {
                        if form.candidates[l_idx].violations[c_idx] > winner.violations[c_idx] {
                            dead[l_pos] = true;
                        }
                    }
                    if dead.iter().all(|&d| d) {
                        found = pos;
                        break;
                    }
                }
                found
            }
        } else {
            usize::MAX
        };

        // Candidate rows (in sorted order: winner first, rivals by harmony)
        for &orig_cand_idx in &sorted_cands {
            let candidate = &form.candidates[orig_cand_idx];
            let is_winner = Some(orig_cand_idx) == winner_idx;
            out.push_str("  <tr>\n");

            // Candidate label cell
            let label = if is_winner {
                format!("&#x261E;&nbsp;{}", html_escape(&candidate.form))
            } else {
                format!("&nbsp;&nbsp;&nbsp;{}", html_escape(&candidate.form))
            };
            out.push_str(&format!("    <td>{label}</td>\n"));

            if is_winner {
                // Winner row: shade cells strictly after the winner shading point
                for (pos, &c_idx) in sorted.iter().enumerate() {
                    let is_shaded = pos > winner_shading_point;
                    let is_last = pos + 1 == num_constraints;
                    let next_stratum = if is_last { usize::MAX } else { self.constraint_strata[sorted[pos + 1]] };
                    let has_border = !is_last && self.constraint_strata[c_idx] != next_stratum;
                    let class = html_cell_class(is_shaded, is_last, has_border);
                    let content = format_html_viol(candidate.violations[c_idx], 0, false);
                    out.push_str(&format!("    <td class=\"{class}\">{content}</td>\n"));
                }
            } else {
                // Loser row: find first fatal violation (in sorted order), shade cells after it.
                // The fatal cell itself is NOT shaded (style chosen before flag is set).
                let first_fatal = if let Some(wi) = winner_idx {
                    let winner = &form.candidates[wi];
                    sorted.iter().position(|&c_idx| candidate.violations[c_idx] > winner.violations[c_idx])
                } else {
                    None
                };

                let mut fatal_seen = false;
                for (pos, &c_idx) in sorted.iter().enumerate() {
                    let is_fatal = first_fatal == Some(pos);
                    // Shade based on fatal_seen BEFORE updating it (matches VB6 behavior)
                    let is_shaded = fatal_seen;
                    let is_last = pos + 1 == num_constraints;
                    let next_stratum = if is_last { usize::MAX } else { self.constraint_strata[sorted[pos + 1]] };
                    let has_border = !is_last && self.constraint_strata[c_idx] != next_stratum;
                    let class = html_cell_class(is_shaded, is_last, has_border);
                    let winner_viols = winner_idx
                        .map(|wi| form.candidates[wi].violations[c_idx])
                        .unwrap_or(0);
                    let content =
                        format_html_viol(candidate.violations[c_idx], winner_viols, is_fatal);
                    out.push_str(&format!("    <td class=\"{class}\">{content}</td>\n"));
                    if is_fatal {
                        fatal_seen = true;
                    }
                }
            }

            out.push_str("  </tr>\n");
        }

        out.push_str("</table>\n\n");
        out
    }

    /// Render a single input form as a transposed HTML tableau table.
    ///
    /// In this layout, the first row shows the input + candidate names as columns,
    /// and each subsequent row shows one constraint with violations per candidate.
    /// `sorted` is the list of constraint indices sorted by stratum (from
    /// `sorted_constraint_indices`), matching VB6's `SortTheConstraints`.
    fn format_html_reversed_form_table(
        &self,
        tableau: &Tableau,
        form: &crate::tableau::InputForm,
        winner_idx: Option<usize>,
        sorted: &[usize],
    ) -> String {
        let mut out = String::new();

        // Sort candidates: winner first, rivals by harmony
        let sorted_cands = self.sorted_candidate_indices(form, sorted);

        out.push_str("<table>\n");

        // Header row: input form + candidate names as columns (sorted)
        out.push_str("  <tr>\n");
        out.push_str(&format!("    <th>/{}/</th>\n", html_escape(&form.input)));
        for &orig_cand_idx in &sorted_cands {
            let candidate = &form.candidates[orig_cand_idx];
            let is_winner = Some(orig_cand_idx) == winner_idx;
            let label = if is_winner {
                format!("&#x261E;&nbsp;{}", html_escape(&candidate.form))
            } else {
                html_escape(&candidate.form)
            };
            out.push_str(&format!("    <th>{label}</th>\n"));
        }
        out.push_str("  </tr>\n");

        // One row per constraint (in sorted stratum order)
        for (pos, &c_idx) in sorted.iter().enumerate() {
            // Stratum separator: add a visual break between strata via a class
            let stratum_break = pos > 0
                && self.constraint_strata[c_idx] != self.constraint_strata[sorted[pos - 1]];

            out.push_str("  <tr>\n");
            let th_class = if stratum_break { " class=\"stratum-break\"" } else { "" };
            out.push_str(&format!(
                "    <th{}>{}</th>\n",
                th_class,
                html_escape(&tableau.constraints[c_idx].abbrev()),
            ));

            for &orig_cand_idx in &sorted_cands {
                let viols = form.candidates[orig_cand_idx].violations[c_idx];
                let winner_viols = winner_idx
                    .map(|wi| form.candidates[wi].violations[c_idx])
                    .unwrap_or(0);
                // In transposed layout we don't mark fatal violations (matches VB6 comment:
                // "this needs work ... HTML output does not have stratal separators")
                let content = format_html_viol(viols, winner_viols, false);
                let td_class = if stratum_break { " class=\"stratum-break\"" } else { "" };
                out.push_str(&format!("    <td{td_class}>{content}</td>\n"));
            }
            out.push_str("  </tr>\n");
        }

        out.push_str("</table>\n\n");
        out
    }

    /// Render a mini-tableau as an HTML table.
    fn format_html_mini_tableau(&self, tableau: &Tableau, mini: &MiniTableau) -> String {
        let form = &tableau.forms[mini.form_index];
        let winner = &form.candidates[mini.winner_index];
        let loser = &form.candidates[mini.loser_index];
        // Sort included constraints by stratum using VB6's unstable selection sort,
        // matching VB6's PrintTableaux.Main which calls SortTheConstraints.
        let mut sorted_constraints = mini.included_constraints.clone();
        self.vb6_sort_constraint_slice(&mut sorted_constraints);
        let included = &sorted_constraints;
        let num_included = included.len();

        let mut out = String::new();
        out.push_str("<table>\n");

        // Header row
        out.push_str("  <tr>\n");
        out.push_str(&format!("    <th>/{}/</th>\n", html_escape(&form.input)));
        for (i, &c_idx) in included.iter().enumerate() {
            let is_last = i + 1 == num_included;
            let has_border = !is_last
                && self.constraint_strata[c_idx] != self.constraint_strata[included[i + 1]];
            let class = html_cell_class(false, is_last, has_border);
            out.push_str(&format!(
                "    <th class=\"{}\">{}</th>\n",
                class,
                html_escape(&tableau.constraints[c_idx].abbrev()),
            ));
        }
        out.push_str("  </tr>\n");

        // Winner row (no fatal markers in mini-tableaux, per text formatter)
        out.push_str("  <tr>\n");
        out.push_str(&format!(
            "    <td>&#x261E;&nbsp;{}</td>\n",
            html_escape(&winner.form)
        ));
        for (i, &c_idx) in included.iter().enumerate() {
            let is_last = i + 1 == num_included;
            let has_border = !is_last
                && self.constraint_strata[c_idx] != self.constraint_strata[included[i + 1]];
            let class = html_cell_class(false, is_last, has_border);
            let content = format_html_viol(winner.violations[c_idx], 0, false);
            out.push_str(&format!("    <td class=\"{class}\">{content}</td>\n"));
        }
        out.push_str("  </tr>\n");

        // Loser row (no fatal markers in mini-tableaux)
        out.push_str("  <tr>\n");
        out.push_str(&format!(
            "    <td>&nbsp;&nbsp;&nbsp;{}</td>\n",
            html_escape(&loser.form)
        ));
        for (i, &c_idx) in included.iter().enumerate() {
            let is_last = i + 1 == num_included;
            let has_border = !is_last
                && self.constraint_strata[c_idx] != self.constraint_strata[included[i + 1]];
            let class = html_cell_class(false, is_last, has_border);
            let content = format_html_viol(loser.violations[c_idx], 0, false);
            out.push_str(&format!("    <td class=\"{class}\">{content}</td>\n"));
        }
        out.push_str("  </tr>\n");

        out.push_str("</table>\n\n");
        out
    }

    // ── Diagnostics when ranking fails ───────────────────────────────────────

    /// Search for contradictory minimal pairs and format as text.
    /// Returns true if contradictions were found (matching VB6 `LookForMinimalPairs`).
    #[allow(clippy::needless_range_loop)] // bidirectional indexing: evidence[outer][inner] AND evidence[inner][outer]
    fn fmt_text_contradictions(
        &self,
        out: &mut String,
        section: &mut usize,
        tableau: &Tableau,
    ) -> bool {
        let evidence = Self::find_minimal_pair_evidence(tableau);
        let nc = tableau.constraints.len();
        let mut found = false;

        for outer in 0..nc.saturating_sub(1) {
            for inner in (outer + 1)..nc {
                if evidence[outer][inner].is_empty() || evidence[inner][outer].is_empty() {
                    continue;
                }
                // First contradiction: print header
                if !found {
                    *section += 1;
                    out.push_str(&format!("\n\n{}. Contradiction Located\n\n", section));
                    out.push_str("The problem can be localized in the form of one or more minimal pairs that cannot be consistently ranked.\n\n");
                }
                found = true;

                out.push_str("The following is a contradiction:\n\n");

                // outer >> inner
                let c_outer = &tableau.constraints[outer];
                let c_inner = &tableau.constraints[inner];
                out.push_str(&format!(
                    "The ranking {} >> {} is needed, because ",
                    c_outer.full_name(), c_inner.full_name()
                ));
                for ev in &evidence[outer][inner] {
                    Self::fmt_text_derivation(out, tableau, ev.form_idx, ev.rival_idx);
                }

                // inner >> outer
                out.push_str(&format!(
                    "\nThe ranking {} >> {} is needed, because ",
                    c_inner.full_name(), c_outer.full_name()
                ));
                for ev in &evidence[inner][outer] {
                    Self::fmt_text_derivation(out, tableau, ev.form_idx, ev.rival_idx);
                }
                out.push('\n');
            }
        }
        found
    }

    /// Search for contradictory minimal pairs and format as HTML.
    /// Returns true if contradictions were found.
    #[allow(clippy::needless_range_loop)] // bidirectional indexing: evidence[outer][inner] AND evidence[inner][outer]
    fn fmt_html_contradictions(
        &self,
        out: &mut String,
        section: &mut usize,
        tableau: &Tableau,
    ) -> bool {
        let evidence = Self::find_minimal_pair_evidence(tableau);
        let nc = tableau.constraints.len();
        let mut found = false;

        for outer in 0..nc.saturating_sub(1) {
            for inner in (outer + 1)..nc {
                if evidence[outer][inner].is_empty() || evidence[inner][outer].is_empty() {
                    continue;
                }
                if !found {
                    *section += 1;
                    out.push_str(&format!("<h2>{}. Contradiction Located</h2>\n", section));
                    out.push_str("<p>The problem can be localized in the form of one or more minimal pairs that cannot be consistently ranked.</p>\n");
                }
                found = true;

                out.push_str("<p>The following is a contradiction:</p>\n");

                let c_outer = &tableau.constraints[outer];
                let c_inner = &tableau.constraints[inner];

                out.push_str(&format!(
                    "<p>The ranking {} &gt;&gt; {} is needed, because ",
                    html_escape(&c_outer.full_name()),
                    html_escape(&c_inner.full_name()),
                ));
                for ev in &evidence[outer][inner] {
                    Self::fmt_html_derivation(out, tableau, ev.form_idx, ev.rival_idx);
                    // Mini-tableau for this evidence
                    self.fmt_html_contradiction_mini_tableau(out, tableau, ev.form_idx, ev.rival_idx, outer, inner);
                }
                out.push_str("</p>\n");

                out.push_str(&format!(
                    "<p>The ranking {} &gt;&gt; {} is needed, because ",
                    html_escape(&c_inner.full_name()),
                    html_escape(&c_outer.full_name()),
                ));
                for ev in &evidence[inner][outer] {
                    Self::fmt_html_derivation(out, tableau, ev.form_idx, ev.rival_idx);
                    self.fmt_html_contradiction_mini_tableau(out, tableau, ev.form_idx, ev.rival_idx, inner, outer);
                }
                out.push_str("</p>\n");
            }
        }
        found
    }

    /// Find all minimal-pair ranking arguments.
    ///
    /// Returns `evidence[higher][lower]` = list of (form, rival) pairs that prove
    /// constraint `higher` must dominate constraint `lower` via a true minimal pair.
    /// Matches VB6 `LookForMinimalPairs` logic.
    #[allow(clippy::needless_range_loop)] // bidirectional indexing into evidence[higher][lower]
    fn find_minimal_pair_evidence(tableau: &Tableau) -> Vec<Vec<Vec<MinimalPairEvidence>>> {
        let nc = tableau.constraints.len();
        let mut evidence: Vec<Vec<Vec<MinimalPairEvidence>>> =
            (0..nc).map(|_| (0..nc).map(|_| Vec::new()).collect()).collect();

        for higher in 0..nc {
            for lower in 0..nc {
                if higher == lower {
                    continue;
                }
                for (form_idx, form) in tableau.forms.iter().enumerate() {
                    let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
                        Some(i) => i,
                        None => continue,
                    };
                    let winner = &form.candidates[winner_idx];

                    for (cand_idx, rival) in form.candidates.iter().enumerate() {
                        if cand_idx == winner_idx {
                            continue;
                        }
                        // Check: rival violates `higher` more than winner,
                        //        winner violates `lower` more than rival
                        if rival.violations[higher] <= winner.violations[higher] {
                            continue;
                        }
                        if rival.violations[lower] >= winner.violations[lower] {
                            continue;
                        }
                        // Check all other constraints are identical
                        let is_minimal = (0..nc).all(|c| {
                            c == higher || c == lower
                                || winner.violations[c] == rival.violations[c]
                        });
                        if is_minimal && evidence[higher][lower].len() < 5 {
                            evidence[higher][lower].push(MinimalPairEvidence {
                                form_idx,
                                rival_idx: cand_idx,
                            });
                        }
                    }
                }
            }
        }
        evidence
    }

    /// Format a derivation line for text output:
    /// `/input/ --> [winner], not *[loser]`
    fn fmt_text_derivation(out: &mut String, tableau: &Tableau, form_idx: usize, rival_idx: usize) {
        let form = &tableau.forms[form_idx];
        let winner = form.candidates.iter().find(|c| c.frequency > 0);
        let rival = &form.candidates[rival_idx];
        if let Some(winner) = winner {
            out.push_str(&format!(
                "/{}/  -->  [{}], not *[{}]\n",
                form.input, winner.form, rival.form
            ));
        }
    }

    /// Format a derivation line for HTML output.
    fn fmt_html_derivation(out: &mut String, tableau: &Tableau, form_idx: usize, rival_idx: usize) {
        let form = &tableau.forms[form_idx];
        let winner = form.candidates.iter().find(|c| c.frequency > 0);
        let rival = &form.candidates[rival_idx];
        if let Some(winner) = winner {
            out.push_str(&format!(
                "/{}/  &rarr;  [{}], not *[{}]",
                html_escape(&form.input),
                html_escape(&winner.form),
                html_escape(&rival.form),
            ));
        }
    }

    /// Format a 2-constraint mini-tableau for a contradiction in HTML.
    ///
    /// Shows only the two constraints involved, with the winner and rival.
    /// `higher` is assigned stratum 1, `lower` stratum 2 (matching VB6 FakeStrata).
    fn fmt_html_contradiction_mini_tableau(
        &self,
        out: &mut String,
        tableau: &Tableau,
        form_idx: usize,
        rival_idx: usize,
        higher: usize,
        lower: usize,
    ) {
        let form = &tableau.forms[form_idx];
        let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
            Some(i) => i,
            None => return,
        };
        let winner = &form.candidates[winner_idx];
        let rival = &form.candidates[rival_idx];
        let constraints = [higher, lower];

        out.push_str("<table>\n  <tr>\n");
        out.push_str(&format!("    <th>{}</th>\n", html_escape(&form.input)));
        for &c_idx in &constraints {
            out.push_str(&format!("    <th>{}</th>\n", html_escape(&tableau.constraints[c_idx].abbrev())));
        }
        out.push_str("  </tr>\n");

        // Winner row
        out.push_str("  <tr>\n");
        out.push_str(&format!("    <td>&#x261E;&nbsp;{}</td>\n", html_escape(&winner.form)));
        for (i, &c_idx) in constraints.iter().enumerate() {
            let is_last = i + 1 == constraints.len();
            let has_border = !is_last; // stratum break between the two
            let class = html_cell_class(false, is_last, has_border);
            let content = format_html_viol(winner.violations[c_idx], 0, false);
            out.push_str(&format!("    <td class=\"{class}\">{content}</td>\n"));
        }
        out.push_str("  </tr>\n");

        // Rival row
        out.push_str("  <tr>\n");
        out.push_str(&format!("    <td>&nbsp;&nbsp;&nbsp;{}</td>\n", html_escape(&rival.form)));
        let winner_higher_viols = winner.violations[higher];
        for (i, &c_idx) in constraints.iter().enumerate() {
            let is_last = i + 1 == constraints.len();
            let has_border = !is_last;
            let is_fatal = rival.violations[c_idx] > winner.violations[c_idx];
            let class = html_cell_class(is_fatal, is_last, has_border);
            let content = format_html_viol(
                rival.violations[c_idx],
                if c_idx == higher { winner_higher_viols } else { 0 },
                is_fatal,
            );
            out.push_str(&format!("    <td class=\"{class}\">{content}</td>\n"));
        }
        out.push_str("  </tr>\n</table>\n");
    }

    /// Format diagnostic tableaux for text output.
    ///
    /// Shows a filtered tableau containing only candidates that survive to the
    /// final (problematic) stratum, with only the constraints from that stratum
    /// that distinguish winners from losers. Matches VB6 `PrepareDiagnosticTableaux`.
    fn fmt_text_diagnostic_tableaux(
        &self,
        out: &mut String,
        section: &mut usize,
        tableau: &Tableau,
    ) {
        let diag = self.compute_diagnostic_filter(tableau);
        if diag.is_empty() {
            return;
        }

        *section += 1;
        out.push_str(&format!("\n\n{}. Diagnostic Tableaux\n\n", section));
        out.push_str("The following tables are provided for diagnosis. They omit all constraints \
            that were ranked before the algorithm crashed, and all data that are explained \
            by rankable constraints. They also exclude all constraints that prefer neither \
            winners nor losers in the remaining data.\n\n");

        for entry in &diag {
            let form = &tableau.forms[entry.form_idx];
            let winner = &form.candidates[entry.winner_idx];

            // Header with constraint abbreviations
            let first_col_width = std::iter::once(winner.form.len())
                .chain(entry.rival_indices.iter().map(|&ri| form.candidates[ri].form.len()))
                .max()
                .unwrap_or(0)
                + 4; // padding for ">" prefix and spacing

            out.push_str(&format!("{:width$}", form.input, width = first_col_width));
            for &c_idx in &entry.constraint_indices {
                out.push_str(&format!("  {}", tableau.constraints[c_idx].abbrev()));
            }
            out.push('\n');

            // Winner row
            out.push_str(&format!(">{:<width$}", winner.form, width = first_col_width - 1));
            for &c_idx in &entry.constraint_indices {
                let v = winner.violations[c_idx];
                let col_w = tableau.constraints[c_idx].abbrev().len().max(2);
                if v == 0 {
                    out.push_str(&format!("  {:>width$}", "", width = col_w));
                } else {
                    out.push_str(&format!("  {:>width$}", v, width = col_w));
                }
            }
            out.push('\n');

            // Rival rows
            for &rival_idx in &entry.rival_indices {
                let rival = &form.candidates[rival_idx];
                out.push_str(&format!(" {:<width$}", rival.form, width = first_col_width - 1));
                for &c_idx in &entry.constraint_indices {
                    let v = rival.violations[c_idx];
                    let col_w = tableau.constraints[c_idx].abbrev().len().max(2);
                    if v == 0 {
                        out.push_str(&format!("  {:>width$}", "", width = col_w));
                    } else {
                        out.push_str(&format!("  {:>width$}", v, width = col_w));
                    }
                }
                out.push('\n');
            }
            out.push('\n');
        }
    }

    /// Format diagnostic tableaux for HTML output.
    fn fmt_html_diagnostic_tableaux(
        &self,
        out: &mut String,
        section: &mut usize,
        tableau: &Tableau,
    ) {
        let diag = self.compute_diagnostic_filter(tableau);
        if diag.is_empty() {
            return;
        }

        *section += 1;
        out.push_str(&format!("<h2>{}. Diagnostic Tableaux</h2>\n", section));
        out.push_str(
            "<p>The following tables are provided for diagnosis. They omit all constraints \
             that were ranked before the algorithm crashed, and all data that are explained \
             by rankable constraints. They also exclude all constraints that prefer neither \
             winners nor losers in the remaining data.</p>\n",
        );

        // Group into a single table
        out.push_str("<table>\n  <tr>\n    <th></th>\n");
        // Use constraints from the first entry (all entries share the same constraint set)
        if let Some(first) = diag.first() {
            for &c_idx in &first.constraint_indices {
                out.push_str(&format!(
                    "    <th>{}</th>\n",
                    html_escape(&tableau.constraints[c_idx].abbrev())
                ));
            }
        }
        out.push_str("  </tr>\n");

        for entry in &diag {
            let form = &tableau.forms[entry.form_idx];
            let winner = &form.candidates[entry.winner_idx];

            // Winner row
            out.push_str("  <tr>\n");
            out.push_str(&format!(
                "    <td>/{}/  &#x261E;&nbsp;{}</td>\n",
                html_escape(&form.input),
                html_escape(&winner.form),
            ));
            for (i, &c_idx) in entry.constraint_indices.iter().enumerate() {
                let is_last = i + 1 == entry.constraint_indices.len();
                let content = format_html_viol(winner.violations[c_idx], 0, false);
                out.push_str(&format!(
                    "    <td class=\"{}\">{}</td>\n",
                    html_cell_class(false, is_last, false),
                    content,
                ));
            }
            out.push_str("  </tr>\n");

            // Rival rows
            for &rival_idx in &entry.rival_indices {
                let rival = &form.candidates[rival_idx];
                out.push_str("  <tr>\n");
                out.push_str(&format!(
                    "    <td>&nbsp;&nbsp;&nbsp;{}</td>\n",
                    html_escape(&rival.form),
                ));
                for (i, &c_idx) in entry.constraint_indices.iter().enumerate() {
                    let is_last = i + 1 == entry.constraint_indices.len();
                    let is_fatal = rival.violations[c_idx] > winner.violations[c_idx];
                    let content = format_html_viol(
                        rival.violations[c_idx],
                        winner.violations[c_idx],
                        is_fatal,
                    );
                    out.push_str(&format!(
                        "    <td class=\"{}\">{}</td>\n",
                        html_cell_class(false, is_last, false),
                        content,
                    ));
                }
                out.push_str("  </tr>\n");
            }
        }
        out.push_str("</table>\n");
    }

    /// Compute the filtered set of forms, rivals, and constraints for diagnostic tableaux.
    ///
    /// Mirrors VB6 `PrepareDiagnosticTableaux` filtering logic:
    /// 1. Exclude rivals eliminated by constraints ranked before the final stratum.
    /// 2. Exclude rivals not preferred by any constraint in the final stratum.
    /// 3. Include only forms that have surviving rivals.
    /// 4. Include only constraints in the final stratum that distinguish winners from losers.
    fn compute_diagnostic_filter(&self, tableau: &Tableau) -> Vec<DiagnosticEntry> {
        let nc = tableau.constraints.len();
        let last_stratum = self.num_strata;

        // Step 1: For each form/rival, check if any ranked constraint (stratum < last)
        // eliminates the rival (rival violates it more than winner).
        let mut rival_ok: Vec<Vec<bool>> = tableau.forms.iter().map(|form| {
            let winner_idx = form.candidates.iter().position(|c| c.frequency > 0);
            form.candidates.iter().enumerate().map(|(cand_idx, rival)| {
                let Some(wi) = winner_idx else { return false };
                if cand_idx == wi { return false; } // Not a rival
                // Check: no ranked constraint eliminates this rival
                !(0..nc).any(|c| {
                    self.constraint_strata[c] < last_stratum
                        && self.constraint_strata[c] > 0
                        && rival.violations[c] > form.candidates[wi].violations[c]
                })
            }).collect()
        }).collect();

        // Step 2: Among surviving rivals, keep only those preferred by at least one
        // constraint in the final stratum.
        for (form_idx, form) in tableau.forms.iter().enumerate() {
            let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
                Some(i) => i,
                None => continue,
            };
            let winner = &form.candidates[winner_idx];
            for (cand_idx, rival) in form.candidates.iter().enumerate() {
                if !rival_ok[form_idx][cand_idx] {
                    continue;
                }
                let preferred_by_any = (0..nc).any(|c| {
                    self.constraint_strata[c] == last_stratum
                        && winner.violations[c] > rival.violations[c]
                });
                if !preferred_by_any {
                    rival_ok[form_idx][cand_idx] = false;
                }
            }
        }

        // Step 3: Determine which forms have surviving rivals.
        let form_ok: Vec<bool> = rival_ok.iter().map(|rivals| rivals.iter().any(|&ok| ok)).collect();

        // Step 4: Determine which constraints in the final stratum distinguish any
        // winner/rival pair among the relevant data.
        let mut constraint_ok = vec![false; nc];
        for (form_idx, form) in tableau.forms.iter().enumerate() {
            if !form_ok[form_idx] {
                continue;
            }
            let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
                Some(i) => i,
                None => continue,
            };
            let winner = &form.candidates[winner_idx];
            for (cand_idx, rival) in form.candidates.iter().enumerate() {
                if !rival_ok[form_idx][cand_idx] {
                    continue;
                }
                for (c, ok) in constraint_ok.iter_mut().enumerate() {
                    if self.constraint_strata[c] == last_stratum
                        && winner.violations[c] != rival.violations[c]
                    {
                        *ok = true;
                    }
                }
            }
        }

        let constraint_indices: Vec<usize> = (0..nc).filter(|&c| constraint_ok[c]).collect();

        // Build entries
        let mut entries = Vec::new();
        for (form_idx, form) in tableau.forms.iter().enumerate() {
            if !form_ok[form_idx] {
                continue;
            }
            let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
                Some(i) => i,
                None => continue,
            };
            let rival_indices: Vec<usize> = (0..form.candidates.len())
                .filter(|&ci| rival_ok[form_idx][ci])
                .collect();
            entries.push(DiagnosticEntry {
                form_idx,
                winner_idx,
                rival_indices,
                constraint_indices: constraint_indices.clone(),
            });
        }
        entries
    }

    /// Top-level text diagnostics: try minimal pairs first, fall back to diagnostic tableaux.
    fn fmt_text_diagnostics(
        &self,
        out: &mut String,
        section: &mut usize,
        tableau: &Tableau,
    ) {
        if !self.fmt_text_contradictions(out, section, tableau) {
            self.fmt_text_diagnostic_tableaux(out, section, tableau);
        }
    }

    /// Top-level HTML diagnostics: try minimal pairs first, fall back to diagnostic tableaux.
    fn fmt_html_diagnostics(
        &self,
        out: &mut String,
        section: &mut usize,
        tableau: &Tableau,
    ) {
        if !self.fmt_html_contradictions(out, section, tableau) {
            self.fmt_html_diagnostic_tableaux(out, section, tableau);
        }
    }
}

impl Tableau {
    /// Run Recursive Constraint Demotion to find a ranking
    pub fn run_rcd(&self) -> RCDResult {
        self.run_rcd_internal(true, &[])
    }

    /// Run RCD enforcing a priori constraint rankings.
    ///
    /// `apriori[i][j] = true` means constraint i must rank above constraint j.
    pub fn run_rcd_with_apriori(&self, apriori: &[Vec<bool>]) -> RCDResult {
        self.run_rcd_internal(true, apriori)
    }

    /// Run RCD with explicitly specified winner indices (one per form).
    ///
    /// Used internally for factorial typology full listing.
    /// Returns a minimal RCDResult without FRed or necessity analysis.
    pub(crate) fn run_rcd_with_winner_indices(
        &self,
        winner_indices: &[usize],
        apriori: &[Vec<bool>],
    ) -> RCDResult {
        debug_assert_eq!(winner_indices.len(), self.forms.len());
        let num_constraints = self.constraints.len();
        let mut constraint_strata = vec![0usize; num_constraints];
        let mut current_stratum = 0;

        let mut informative_pairs: Vec<(usize, usize, usize)> = Vec::new();
        for (form_idx, (&winner_idx, form)) in winner_indices.iter().zip(&self.forms).enumerate() {
            for (loser_idx, _) in form.candidates.iter().enumerate() {
                if loser_idx != winner_idx {
                    informative_pairs.push((form_idx, winner_idx, loser_idx));
                }
            }
        }

        loop {
            current_stratum += 1;

            let mut demotable = vec![false; num_constraints];
            for &(form_idx, winner_idx, loser_idx) in &informative_pairs {
                let winner = &self.forms[form_idx].candidates[winner_idx];
                let loser = &self.forms[form_idx].candidates[loser_idx];
                for c_idx in 0..num_constraints {
                    if constraint_strata[c_idx] != 0 {
                        continue;
                    }
                    if loser.violations[c_idx] < winner.violations[c_idx] {
                        demotable[c_idx] = true;
                    }
                }
            }

            if !apriori.is_empty() {
                for outer in 0..num_constraints {
                    if constraint_strata[outer] == 0 {
                        for inner in 0..num_constraints {
                            if apriori[outer][inner] {
                                demotable[inner] = true;
                            }
                        }
                    }
                }
            }

            let mut added_any = false;
            for c_idx in 0..num_constraints {
                if constraint_strata[c_idx] == 0 && !demotable[c_idx] {
                    constraint_strata[c_idx] = current_stratum;
                    added_any = true;
                }
            }

            let all_ranked = constraint_strata.iter().all(|&s| s != 0);
            if all_ranked {
                return RCDResult::new(constraint_strata, current_stratum, true);
            }

            if !added_any {
                return RCDResult::new(constraint_strata, current_stratum - 1, false);
            }

            informative_pairs.retain(|&(form_idx, winner_idx, loser_idx)| {
                let winner = &self.forms[form_idx].candidates[winner_idx];
                let loser = &self.forms[form_idx].candidates[loser_idx];
                for (c_idx, &c_stratum) in constraint_strata.iter().enumerate() {
                    if c_stratum == current_stratum && winner.violations[c_idx] < loser.violations[c_idx] {
                        return false;
                    }
                }
                true
            });

            if informative_pairs.is_empty() {
                if !constraint_strata.iter().all(|&s| s != 0) {
                    for s in constraint_strata.iter_mut() {
                        if *s == 0 {
                            *s = current_stratum + 1;
                        }
                    }
                    current_stratum += 1;
                }
                return RCDResult::new(constraint_strata, current_stratum, true);
            }

            if current_stratum > num_constraints {
                return RCDResult::new(constraint_strata, current_stratum, false);
            }
        }
    }

    /// Log a "Results so far" summary matching VB6's PrintResultsOfRankingSoFar.
    pub(crate) fn log_results_so_far(&self, strata: &[usize], current_stratum: usize) {
        crate::ot_log!("");
        crate::ot_log!("Results so far:");
        // Already-ranked strata
        for s in 1..current_stratum {
            let mut first = true;
            for (c_idx, &c_stratum) in strata.iter().enumerate() {
                if c_stratum == s {
                    if first {
                        crate::ot_log!("  Stratum {} (already ranked):  {}", s, self.constraints[c_idx].abbrev());
                        first = false;
                    } else {
                        crate::ot_log!("                               {}", self.constraints[c_idx].abbrev());
                    }
                }
            }
        }
        // Newly ranked
        let mut first = true;
        for (c_idx, &c_stratum) in strata.iter().enumerate() {
            if c_stratum == current_stratum {
                if first {
                    crate::ot_log!("  Stratum {} (newly ranked):    {}", current_stratum, self.constraints[c_idx].abbrev());
                    first = false;
                } else {
                    crate::ot_log!("                               {}", self.constraints[c_idx].abbrev());
                }
            }
        }
        // Unranked markedness
        crate::ot_log!("  Markedness constraints still unranked:");
        let unranked_mark: Vec<_> = self.constraints.iter().enumerate()
            .filter(|(i, c)| strata[*i] == 0 && !c.is_faithfulness())
            .collect();
        if unranked_mark.is_empty() {
            crate::ot_log!("    (none)");
        } else {
            for (_, c) in &unranked_mark {
                crate::ot_log!("    {}", c.abbrev());
            }
        }
        // Unranked faithfulness
        crate::ot_log!("  Faithfulness constraints still unranked:");
        let unranked_faith: Vec<_> = self.constraints.iter().enumerate()
            .filter(|(i, c)| strata[*i] == 0 && c.is_faithfulness())
            .collect();
        if unranked_faith.is_empty() {
            crate::ot_log!("    (none)");
        } else {
            for (_, c) in &unranked_faith {
                crate::ot_log!("    {}", c.abbrev());
            }
        }
    }

    /// Internal RCD implementation
    fn run_rcd_internal(&self, compute_extra_analyses: bool, apriori: &[Vec<bool>]) -> RCDResult {
        let num_constraints = self.constraints.len();
        let mut constraint_strata = vec![0; num_constraints];
        let mut current_stratum = 0;

        // Track which winner-loser pairs are still informative
        let mut informative_pairs: Vec<(usize, usize, usize)> = Vec::new();

        // Build list of all winner-loser pairs
        // (form_index, winner_index, loser_index)
        for (form_idx, form) in self.forms.iter().enumerate() {
            // Winner is the candidate with non-zero frequency
            if let Some(winner_idx) = form.candidates.iter().position(|c| c.frequency > 0) {
                for (loser_idx, candidate) in form.candidates.iter().enumerate() {
                    if loser_idx != winner_idx && candidate.frequency == 0 {
                        informative_pairs.push((form_idx, winner_idx, loser_idx));
                    }
                }
            }
        }

        crate::ot_log!("****** Application of Constraint Demotion ******");
        crate::ot_log!("Starting RCD with {} pairs", informative_pairs.len());

        // RCD main loop
        loop {
            current_stratum += 1;
            crate::ot_log!("");
            crate::ot_log!("****** Now doing Stratum #{} ******", current_stratum);

            // Find constraints that are "demotable" (prefer a loser in any informative pair)
            let mut demotable = vec![false; num_constraints];

            for &(form_idx, winner_idx, loser_idx) in &informative_pairs {
                let winner = &self.forms[form_idx].candidates[winner_idx];
                let loser = &self.forms[form_idx].candidates[loser_idx];

                for c_idx in 0..num_constraints {
                    if constraint_strata[c_idx] != 0 || demotable[c_idx] {
                        continue;
                    }

                    if loser.violations[c_idx] < winner.violations[c_idx] {
                        demotable[c_idx] = true;
                        crate::ot_log!("  {} is excluded from stratum; prefers loser *[{}] for /{}/ over winner [{}].",
                            self.constraints[c_idx].abbrev(),
                            loser.form, self.forms[form_idx].input, winner.form);
                    }
                }
            }

            if !demotable.iter().enumerate().any(|(i, &d)| d && constraint_strata[i] == 0) {
                crate::ot_log!("  Search found no unranked constraints that prefer losers.");
            }

            // ENFORCE A PRIORI RANKINGS
            if !apriori.is_empty() {
                let mut found_apriori = false;
                for outer in 0..num_constraints {
                    if constraint_strata[outer] == 0 {
                        for inner in 0..num_constraints {
                            if apriori[outer][inner] && !demotable[inner] {
                                crate::ot_log!("  {} is excluded from stratum; dominated a priori by {}.",
                                    self.constraints[inner].abbrev(), self.constraints[outer].abbrev());
                                found_apriori = true;
                            }
                            if apriori[outer][inner] {
                                demotable[inner] = true;
                            }
                        }
                    }
                }
                if !found_apriori {
                    crate::ot_log!("  Search found no constraints that must be demoted due to an a priori ranking.");
                }
            }

            // All non-demotable, unranked constraints go into current stratum
            let mut added_any = false;
            for c_idx in 0..num_constraints {
                if constraint_strata[c_idx] == 0 && !demotable[c_idx] {
                    constraint_strata[c_idx] = current_stratum;
                    added_any = true;
                    crate::ot_log!("  {} favors no losers, joins stratum #{}.",
                        self.constraints[c_idx].abbrev(), current_stratum);
                }
            }

            self.log_results_so_far(&constraint_strata, current_stratum);
            crate::ot_log!("After stratum {}: {} pairs remaining",
                current_stratum, informative_pairs.len());

            // Check if all constraints are ranked
            let all_ranked = constraint_strata.iter().all(|&s| s != 0);

            // If all constraints ranked, we're done (success even if pairs remain - those are ties)
            if all_ranked {
                crate::ot_log!("");
                crate::ot_log!("Ranking is complete and yields successful grammar.");
                crate::ot_log!("RCD SUCCEEDED: all constraints ranked in {} strata ({} pairs unresolved - ties)",
                    current_stratum, informative_pairs.len());

                // Create initial result
                let mut result = RCDResult {
                    constraint_strata,
                    num_strata: current_stratum,
                    success: true,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };

                // Compute additional analyses only if requested
                if compute_extra_analyses {
                    result.compute_extra_analyses(self, apriori, true);
                }

                return result;
            }

            // If no constraints added but some still unranked, algorithm failed
            if !added_any {
                crate::ot_log!("");
                crate::ot_log!("Ranking has failed. This constraint set cannot derive only winners.");
                crate::ot_log!("RCD FAILED: no constraints added to stratum {}", current_stratum);
                return RCDResult {
                    constraint_strata,
                    num_strata: current_stratum - 1,
                    success: false,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };
            }

            // Remove pairs that are now decided by current stratum
            informative_pairs.retain(|&(form_idx, winner_idx, loser_idx)| {
                let winner = &self.forms[form_idx].candidates[winner_idx];
                let loser = &self.forms[form_idx].candidates[loser_idx];

                // Check if any constraint in current stratum decides this pair
                for (c_idx, &c_stratum) in constraint_strata.iter().enumerate() {
                    if c_stratum == current_stratum {
                        let winner_viols = winner.violations[c_idx];
                        let loser_viols = loser.violations[c_idx];

                        // If this constraint prefers the winner, pair is decided
                        if winner_viols < loser_viols {
                            return false; // Remove this pair
                        }
                    }
                }
                true // Keep this pair
            });

            // If all pairs decided, we're done
            if informative_pairs.is_empty() {
                // Check if all constraints are ranked
                let all_ranked = constraint_strata.iter().all(|&s| s != 0);

                // Unranked constraints go in final stratum
                if !all_ranked {
                    current_stratum += 1;
                    for (c_idx, s) in constraint_strata.iter_mut().enumerate() {
                        if *s == 0 {
                            *s = current_stratum;
                            crate::ot_log!("  {} joins stratum #{} (no remaining pairs).",
                                self.constraints[c_idx].abbrev(), current_stratum);
                        }
                    }
                    self.log_results_so_far(&constraint_strata, current_stratum);
                }

                crate::ot_log!("");
                crate::ot_log!("Ranking is complete and yields successful grammar.");
                crate::ot_log!("RCD SUCCEEDED with {} strata", current_stratum);

                // Create initial result
                let mut result = RCDResult {
                    constraint_strata,
                    num_strata: current_stratum,
                    success: true,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };

                // Compute additional analyses only if requested
                if compute_extra_analyses {
                    result.compute_extra_analyses(self, apriori, true);
                }

                return result;
            }

            // Safety check: avoid infinite loop
            if current_stratum > num_constraints {
                return RCDResult {
                    constraint_strata,
                    num_strata: current_stratum,
                    success: false,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };
            }
        }
    }

    /// Compute constraint necessity for each constraint
    pub(crate) fn compute_constraint_necessity(&self, rcd_result: &RCDResult, apriori: &[Vec<bool>]) -> Vec<ConstraintNecessity> {
        let mut necessity = vec![ConstraintNecessity::Necessary; self.constraints.len()];

        // Only analyze if RCD succeeded
        if !rcd_result.success {
            return necessity;
        }

        for (c_idx, nec) in necessity.iter_mut().enumerate() {
            // Test if constraint is necessary
            if !self.is_constraint_necessary(c_idx, apriori) {
                // Constraint is unnecessary - check if violated by any winner
                if self.is_violated_by_winner(c_idx) {
                    *nec = ConstraintNecessity::UnnecessaryButShownForFaithfulness;
                } else {
                    *nec = ConstraintNecessity::CompletelyUnnecessary;
                }
            }
        }

        necessity
    }

    /// Test if a constraint is necessary by running RCD without it.
    ///
    /// Mirrors VB6's `FindUnnecessaryConstraints`:
    /// 1. First check if removing this constraint makes any winner-rival pair
    ///    identical (same violations on all constraints). If so, the constraint
    ///    is necessary and we skip the RCD check (VB6 lines 5942–5957).
    /// 2. Otherwise, run RCD without the constraint; if RCD fails the constraint
    ///    is necessary.
    fn is_constraint_necessary(&self, constraint_idx: usize, apriori: &[Vec<bool>]) -> bool {
        let nc = self.constraints.len();

        // Step 1: check for any winner-rival pair that becomes identical after zeroing
        for form in &self.forms {
            let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                Some(w) => w,
                None => continue,
            };
            for rival in form.candidates.iter().filter(|c| c.frequency == 0) {
                let identical = (0..nc).all(|c_idx| {
                    let w_v = if c_idx == constraint_idx { 0 } else { winner.violations[c_idx] };
                    let r_v = if c_idx == constraint_idx { 0 } else { rival.violations[c_idx] };
                    w_v == r_v
                });
                if identical {
                    return true;
                }
            }
        }

        // Step 2: run RCD without the constraint
        let modified_tableau = self.clone_with_constraint_removed(constraint_idx);
        let test_result = modified_tableau.run_rcd_internal(false, apriori);
        !test_result.success
    }

    /// Check if any winner violates this constraint
    fn is_violated_by_winner(&self, constraint_idx: usize) -> bool {
        for form in &self.forms {
            // Find the winner (candidate with non-zero frequency)
            if let Some(winner) = form.candidates.iter().find(|c| c.frequency > 0) {
                if winner.violations[constraint_idx] > 0 {
                    return true;
                }
            }
        }
        false
    }

    /// Clone the tableau with a constraint's violations set to zero
    fn clone_with_constraint_removed(&self, constraint_idx: usize) -> Tableau {
        use crate::tableau::{InputForm, Candidate};

        let forms = self.forms.iter().map(|form| {
            let candidates = form.candidates.iter().map(|cand| {
                let mut violations = cand.violations.clone();
                violations[constraint_idx] = 0;
                Candidate {
                    form: cand.form.clone(),
                    frequency: cand.frequency,
                    violations,
                }
            }).collect();

            InputForm {
                input: form.input.clone(),
                candidates,
            }
        }).collect();

        Tableau {
            constraints: self.constraints.clone(),
            forms,
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::tableau::Tableau;

    fn load_tiny_example() -> String {
        std::fs::read_to_string("../examples/TinyIllustrativeFile.txt")
            .expect("Failed to load examples/TinyIllustrativeFile.txt")
    }

    #[test]
    fn test_rcd_tiny_example() {
        let tiny_example = load_tiny_example();
        let tableau = Tableau::parse(&tiny_example).expect("Failed to parse tiny example");

        let result = tableau.run_rcd();

        // Should succeed
        assert!(result.success, "RCD should find a valid ranking");

        // Should have 2 strata according to expected output
        assert_eq!(result.num_strata, 2, "Should have 2 strata");

        // Check constraint rankings
        // Expected from full_output.txt:
        // Stratum #1: *NoOns, *Coda
        // Stratum #2: Max, Dep

        let constraint_0 = tableau.get_constraint(0).unwrap(); // *NoOns
        let constraint_1 = tableau.get_constraint(1).unwrap(); // *Coda
        let constraint_2 = tableau.get_constraint(2).unwrap(); // Max
        let constraint_3 = tableau.get_constraint(3).unwrap(); // Dep

        println!("Constraint 0 ({}): stratum {}", constraint_0.abbrev(), result.get_stratum(0).unwrap());
        println!("Constraint 1 ({}): stratum {}", constraint_1.abbrev(), result.get_stratum(1).unwrap());
        println!("Constraint 2 ({}): stratum {}", constraint_2.abbrev(), result.get_stratum(2).unwrap());
        println!("Constraint 3 ({}): stratum {}", constraint_3.abbrev(), result.get_stratum(3).unwrap());

        // *NoOns and *Coda should be in stratum 1
        assert_eq!(result.get_stratum(0).unwrap(), 1, "*NoOns should be in stratum 1");
        assert_eq!(result.get_stratum(1).unwrap(), 1, "*Coda should be in stratum 1");

        // Max and Dep should be in stratum 2
        assert_eq!(result.get_stratum(2).unwrap(), 2, "Max should be in stratum 2");
        assert_eq!(result.get_stratum(3).unwrap(), 2, "Dep should be in stratum 2");
    }

    #[test]
    fn test_rcd_log_capture() {
        crate::clear_log();
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let _result = tableau.run_rcd();
        let log = crate::get_log();
        // Header
        assert!(log.contains("Application of Constraint Demotion"), "should have algorithm header");
        // Stratum headers
        assert!(log.contains("Stratum #1"), "should log stratum 1");
        assert!(log.contains("Stratum #2"), "should log stratum 2");
        // Demotable constraints with evidence
        assert!(log.contains("Max is excluded from stratum"), "should log Max demotion");
        assert!(log.contains("Dep is excluded from stratum"), "should log Dep demotion");
        // Constraints joining
        assert!(log.contains("*NoOns favors no losers, joins stratum #1"), "should log *NoOns joining");
        assert!(log.contains("*Coda favors no losers, joins stratum #1"), "should log *Coda joining");
        // Results summary
        assert!(log.contains("Results so far:"), "should have results summary");
        assert!(log.contains("(newly ranked)"), "should show newly ranked");
        // Success
        assert!(log.contains("Ranking is complete and yields successful grammar."), "should log success");
    }

    /// Build a contradictory tableau where C1 >> C2 and C2 >> C1 are both required.
    ///
    /// Two forms create a contradiction:
    /// - Form 1: winner has 0 C1-viols, 1 C2-viol; rival has 1 C1-viol, 0 C2-viols → needs C1 >> C2
    /// - Form 2: winner has 1 C1-viol, 0 C2-viols; rival has 0 C1-viols, 1 C2-viol → needs C2 >> C1
    fn contradictory_tableau() -> &'static str {
        // Row 1: constraint names (3 leading tabs for input/candidate/frequency columns)
        // Data rows: input\tcandidate\tfrequency\tviols...
        // Empty input = same form as previous row
        "\t\t\tC1\tC2\n\
         input1\twinnerA\t1\t0\t1\n\
         \trivalA\t0\t1\t0\n\
         input2\twinnerB\t1\t1\t0\n\
         \trivalB\t0\t0\t1"
    }

    #[test]
    fn test_diagnostics_finds_contradictions() {
        let tableau = Tableau::parse(contradictory_tableau()).expect("Failed to parse");
        let result = tableau.run_rcd();
        assert!(!result.success, "RCD should fail on contradictory data");

        // Text diagnostics should contain "Contradiction Located"
        let text = result.format_output_with_options(
            &tableau, "test.txt", "RCD", &[], false, true,
        );
        assert!(text.contains("Contradiction Located"), "Text output should contain contradiction header");
        assert!(text.contains("C1 >> C2"), "Should report C1 >> C2 ranking requirement");
        assert!(text.contains("C2 >> C1"), "Should report C2 >> C1 ranking requirement");

        // HTML diagnostics should also contain contradiction info
        let html = result.format_html_output_full(
            &tableau, "test.txt", "RCD", crate::AxisMode::default(), &[], true,
        );
        assert!(html.contains("Contradiction Located"), "HTML output should contain contradiction header");
    }

    #[test]
    fn test_diagnostics_disabled_no_output() {
        let tableau = Tableau::parse(contradictory_tableau()).expect("Failed to parse");
        let result = tableau.run_rcd();
        assert!(!result.success);

        let text = result.format_output_with_options(
            &tableau, "test.txt", "RCD", &[], false, false,
        );
        assert!(!text.contains("Contradiction Located"), "Diagnostics should not appear when disabled");
    }

}
