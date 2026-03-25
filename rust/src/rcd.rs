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
    /// Called by format functions to apply user-specified argumentation options:
    /// - `include_fred`: if false, suppress FRed output and mini-tableaux
    /// - `use_mib`: use Most Informative Basis instead of Skeletal Basis
    /// - `verbose`: include verbose recursion tree in FRed output
    /// - `include_mini_tableaux`: if false, suppress mini-tableaux
    pub(crate) fn apply_fred_options(
        &mut self,
        tableau: &Tableau,
        include_fred: bool,
        use_mib: bool,
        verbose: bool,
        include_mini_tableaux: bool,
    ) {
        if !include_fred {
            self.fred_result = None;
            self.mini_tableaux = Vec::new();
            return;
        }

        // Re-run FRed if options differ from the default (SB, no verbose).
        // Default was computed in compute_extra_analyses as run_fred(false).
        if use_mib || verbose {
            self.fred_result = Some(tableau.run_fred_verbose(use_mib, verbose));
        }

        if !include_mini_tableaux {
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
        self.format_output_with_options(tableau, filename, "Recursive Constraint Demotion", &[], true)
    }

    /// Generate formatted text output with a configurable algorithm name and a priori table.
    /// Does not include the A Priori Rankings section (used by BCD, which takes no a priori input).
    pub(crate) fn format_output_with_algorithm(&self, tableau: &Tableau, filename: &str, algorithm_name: &str) -> String {
        self.format_output_with_options(tableau, filename, algorithm_name, &[], false)
    }

    /// Generate formatted text output with full options.
    ///
    /// `show_apriori_section`: whether to include the "A Priori Rankings" section.
    /// RCD includes it; LFCD does not (VB6 behaviour).
    pub(crate) fn format_output_with_options(
        &self,
        tableau: &Tableau,
        filename: &str,
        algorithm_name: &str,
        apriori: &[Vec<bool>],
        show_apriori_section: bool,
    ) -> String {
        let mut output = String::new();

        // Header
        output.push_str(&format!("Results of Applying {} to {}\n", algorithm_name, filename));
        output.push_str("\n\n");

        // Date and version (current date/time and version)
        let now = chrono::Local::now();
        output.push_str(&format!("{}\n\n", now.format("%-m-%-d-%Y, %-I:%M %p").to_string().to_lowercase()));
        output.push_str(crate::VERSION_STRING);
        output.push('\n');
        output.push_str("\n\n");

        if self.tie_warning {
            output.push_str("Caution: The BCD algorithm has selected arbitrarily among tied Faithfulness constraint subsets.\n");
            output.push_str("You may wish to try changing the order of the Faithfulness constraints in the input file,\n");
            output.push_str("to see whether this results in a different ranking.\n\n\n");
        }

        // Auto-incrementing section counter (matches VB6 gLevel1HeadingNumber)
        let mut section = 0usize;

        // Section: Result
        section += 1;
        output.push_str(&format!("{}. Result\n\n", section));

        if self.success {
            output.push_str("A ranking was found that generates the correct outputs.\n\n");
        } else {
            output.push_str("No ranking was found.\n\n");
        }

        // List strata
        for stratum in 1..=self.num_strata {
            output.push_str(&format!("   Stratum #{}\n", stratum));

            for (c_idx, &c_stratum) in self.constraint_strata.iter().enumerate() {
                if c_stratum == stratum {
                    if let Some(constraint) = tableau.get_constraint(c_idx) {
                        let full_name = constraint.full_name();
                        let abbrev = constraint.abbrev();
                        // Format: left-align full name in ~42 chars, then abbreviation
                        output.push_str(&format!("      {:<42}{}\n", full_name, abbrev));
                    }
                }
            }
        }
        output.push('\n');

        // Section: A Priori Rankings (only for algorithms that show it, e.g. RCD but not LFCD)
        if show_apriori_section && !apriori.is_empty() {
            section += 1;
            output.push_str(&format!("{}. A Priori Rankings\n\n", section));
            output.push_str("In the following table, \"yes\" means that the constraint of the indicated row \n");
            output.push_str("was marked a priori to dominate the constraint in the given column.\n\n");

            let nc = tableau.constraints.len();
            let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();

            // Compute column widths (max of header abbreviation and "yes")
            let col_widths: Vec<usize> = abbrevs.iter().map(|a| a.len().max(3)).collect();
            let row_label_width = abbrevs.iter().map(|a| a.len()).max().unwrap_or(0);

            // Header row: row label padding, then column abbreviations centered
            output.push_str(&format!("{:width$}", "", width = row_label_width));
            for (j, abbrev) in abbrevs.iter().enumerate() {
                output.push_str(&format!("  {:^width$}", abbrev, width = col_widths[j]));
            }
            output.push('\n');

            // Data rows
            for i in 0..nc {
                output.push_str(&format!("{:<width$}", abbrevs[i], width = row_label_width));
                for j in 0..nc {
                    let cell = if apriori[i][j] { "yes" } else { "" };
                    output.push_str(&format!("  {:^width$}", cell, width = col_widths[j]));
                }
                output.push('\n');
            }
            output.push_str("\n\n");
        }

        // Section: Tableaux
        section += 1;
        output.push_str(&format!("{}. Tableaux\n\n", section));

        for form in &tableau.forms {
            output.push('\n');
            output.push_str(&format!("/{}/: \n", form.input));

            // Build constraint header with stratum separators
            // Use broken bar character (U+00A6) for within-stratum, | for between-strata
            let mut header = String::new();

            // Calculate max candidate width for alignment
            let max_cand_width = form.candidates.iter()
                .map(|c| c.form.len())
                .max()
                .unwrap_or(0)
                .max(2); // At least 2 chars for marker + space

            // Add initial spacing
            for _ in 0..max_cand_width + 2 {
                header.push(' ');
            }

            // Sort constraints by stratum (matches VB6 SortTheConstraints)
            let sorted = self.sorted_constraint_indices(tableau.constraints.len());

            for (pos, &c_idx) in sorted.iter().enumerate() {
                let c_stratum = self.constraint_strata[c_idx];

                // Add separator before this constraint
                if pos > 0 {
                    let prev_stratum = self.constraint_strata[sorted[pos - 1]];
                    if c_stratum != prev_stratum {
                        header.push('|');
                    } else {
                        header.push(BROKEN_BAR);
                    }
                }

                header.push_str(&tableau.constraints[c_idx].abbrev());
            }
            output.push_str(&header);
            output.push('\n');

            // Find winner and sort candidates (winner first, rivals by harmony)
            let winner_idx = form.candidates.iter().position(|c| c.frequency > 0);
            let sorted_cands = self.sorted_candidate_indices(form, &sorted);

            // Output each candidate (in sorted order)
            for &orig_cand_idx in &sorted_cands {
                let candidate = &form.candidates[orig_cand_idx];
                let is_winner = Some(orig_cand_idx) == winner_idx;
                let marker = if is_winner { ">" } else { " " };

                // Find position (in sorted order) of the first fatal violation for this loser
                let first_fatal_idx = if !is_winner {
                    winner_idx.and_then(|wi| {
                        let winner = &form.candidates[wi];
                        sorted.iter().position(|&c_idx| candidate.violations[c_idx] > winner.violations[c_idx])
                    })
                } else {
                    None
                };

                // Candidate surface form (right-aligned to match expected output)
                output.push_str(&format!("{}{:<width$} ", marker, candidate.form, width = max_cand_width));

                // Violations with stratum separators (in sorted constraint order)
                for (pos, &c_idx) in sorted.iter().enumerate() {
                    let viols = candidate.violations[c_idx];
                    let c_stratum = self.constraint_strata[c_idx];
                    let col_width = tableau.constraints[c_idx].abbrev().len();

                    // Add separator
                    if pos > 0 {
                        let prev_stratum = self.constraint_strata[sorted[pos - 1]];
                        if c_stratum != prev_stratum {
                            output.push('|');
                        } else {
                            output.push(BROKEN_BAR);
                        }
                    }

                    // Mark violation as fatal only if it's the FIRST fatal violation
                    let is_fatal = first_fatal_idx == Some(pos);
                    output.push_str(&format_violation(col_width, viols, is_fatal));
                }
                output.push('\n');
            }
            output.push('\n');
        }
        output.push('\n');

        // Section: Status of Proposed Constraints
        if !self.constraint_necessity.is_empty() {
            section += 1;
            output.push_str(&format!("{}. Status of Proposed Constraints:  Necessary or Unnecessary\n\n", section));

            let max_abbrev_width = tableau.constraints.iter()
                .map(|c| c.abbrev().len())
                .max()
                .unwrap_or(0);

            // VB6 outputs constraints grouped by category in three passes:
            // 1. Necessary, 2. UnnecessaryButFaithfulness, 3. CompletelyUnnecessary
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
                    output.push_str(&format!(
                        "   {:<width$}  {}\n",
                        constraint.abbrev(),
                        status,
                        width = max_abbrev_width,
                    ));
                }
            }

            // Check if mass deletion is possible (only relevant when >=2 deletable constraints)
            let num_deletable = self.constraint_necessity.iter()
                .filter(|n| **n != ConstraintNecessity::Necessary)
                .count();
            if num_deletable >= 2 {
                let mass_deletion_possible = self.check_mass_deletion(tableau, apriori);
                if mass_deletion_possible {
                    output.push_str("\nA check has determined that the grammar will still work even if the \n");
                    output.push_str("constraints marked above as unnecessary are removed en masse.\n\n\n");
                } else {
                    output.push_str("\n\nA check has determined that, although the grammar will still work with the\n");
                    output.push_str("removal of ANY ONE of the constraints marked above as unnecessary, the\n");
                    output.push_str("grammar will NOT work if they are removed en masse.\n\n\n");
                }
            } else {
                output.push_str("\n\n");
            }
        }

        // Section: Ranking Arguments (FRed)
        if let Some(ref fred) = self.fred_result {
            section += 1;
            output.push_str(&fred.format_section_fred(section));
        }

        // Section: Mini-Tableaux
        if !self.mini_tableaux.is_empty() {
            section += 1;
            output.push_str(&format!("{}. Mini-Tableaux\n\n", section));
            output.push_str("The following small tableaux may be useful in presenting ranking arguments. \n");
            output.push_str("They include all winner-rival comparisons in which there is just one \n");
            output.push_str("winner-preferring constraint and at least one loser-preferring constraint.  \n");
            output.push_str("Constraints not violated by either candidate are omitted.\n\n");

            for mini in &self.mini_tableaux {
                self.format_mini_tableau(tableau, mini, &mut output);
            }
        }

        // Normalize trailing newlines to match VB6 output (exactly one trailing blank line)
        let trimmed = output.trim_end_matches('\n');
        format!("{}\n\n", trimmed)
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
        self.format_html_output_full(tableau, filename, "Recursive Constraint Demotion", AxisMode::default(), &[])
    }

    /// Generate an HTML document with configurable algorithm name and axis mode.
    pub(crate) fn format_html_output_with_options(
        &self,
        tableau: &Tableau,
        filename: &str,
        algorithm_name: &str,
        axis_mode: AxisMode,
    ) -> String {
        self.format_html_output_full(tableau, filename, algorithm_name, axis_mode, &[])
    }

    /// Generate an HTML document with full options including a priori data.
    pub(crate) fn format_html_output_full(
        &self,
        tableau: &Tableau,
        filename: &str,
        algorithm_name: &str,
        axis_mode: AxisMode,
        apriori: &[Vec<bool>],
    ) -> String {
        let mut out = String::new();

        // DOCTYPE, head, CSS
        out.push_str("<!DOCTYPE html>\n<html>\n<head>\n");
        out.push_str("<meta charset=\"UTF-8\">\n");
        out.push_str(&format!(
            "<title>{} {}</title>\n",
            crate::VERSION_STRING,
            html_escape(filename)
        ));
        out.push_str(HTML_STYLE);
        out.push_str("\n</head>\n<body>\n");

        // Document header
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

        // Auto-incrementing section counter (matches VB6 gLevel1HeadingNumber)
        let mut section = 0usize;

        // Section: Result
        section += 1;
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

        // Section: A Priori Rankings (only when a priori data is present)
        if !apriori.is_empty() {
            section += 1;
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

        // Section: Tableaux as HTML tables
        section += 1;
        out.push_str(&format!("<h2>{}. Tableaux</h2>\n", section));

        // Precompute total constraint abbreviation length for "switch where needed" mode
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

        // Section: Constraint necessity
        if !self.constraint_necessity.is_empty() {
            section += 1;
            out.push_str(&format!(
                "<h2>{}. Status of Proposed Constraints: Necessary or Unnecessary</h2>\n",
                section,
            ));
            out.push_str("<table class=\"necessity-table\">\n<tbody>\n");
            out.push_str("  <tr><td><b>Constraint</b></td><td><b>Status</b></td></tr>\n");
            // VB6 groups by category: Necessary, UnnecessaryButFaithfulness, CompletelyUnnecessary
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

        // Section: Ranking Arguments (FRed)
        if let Some(ref fred) = self.fred_result {
            section += 1;
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

            // VB6 renders rankings as a 2-column table (ranking | &nbsp;)
            out.push_str("<table>\n");
            for ranking in fred.ranking_strings() {
                out.push_str(&format!(
                    "  <tr><td>{}</td><td>&nbsp;</td></tr>\n",
                    html_escape(&ranking)
                ));
            }
            out.push_str("</table>\n");
        }

        // Section: Mini-Tableaux
        if !self.mini_tableaux.is_empty() {
            section += 1;
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

        out.push_str("</body>\n</html>\n");
        out
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

        crate::ot_log!("Starting RCD with {} pairs", informative_pairs.len());

        // RCD main loop
        loop {
            current_stratum += 1;

            // Find constraints that are "demotable" (prefer a loser in any informative pair)
            let mut demotable = vec![false; num_constraints];

            for &(form_idx, winner_idx, loser_idx) in &informative_pairs {
                let winner = &self.forms[form_idx].candidates[winner_idx];
                let loser = &self.forms[form_idx].candidates[loser_idx];

                for c_idx in 0..num_constraints {
                    // Skip constraints already ranked
                    if constraint_strata[c_idx] != 0 {
                        continue;
                    }

                    let winner_viols = winner.violations[c_idx];
                    let loser_viols = loser.violations[c_idx];

                    // Constraint prefers loser if loser has fewer violations
                    if loser_viols < winner_viols {
                        demotable[c_idx] = true;
                    }
                }
            }

            // ENFORCE A PRIORI RANKINGS
            // Any constraint that is a priori dominated by an unranked constraint
            // cannot join the current stratum.
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

            // All non-demotable, unranked constraints go into current stratum
            let mut added_any = false;
            for c_idx in 0..num_constraints {
                if constraint_strata[c_idx] == 0 && !demotable[c_idx] {
                    constraint_strata[c_idx] = current_stratum;
                    added_any = true;
                }
            }

            crate::ot_log!("After stratum {}: {} pairs remaining",
                current_stratum, informative_pairs.len());

            // Check if all constraints are ranked
            let all_ranked = constraint_strata.iter().all(|&s| s != 0);

            // If all constraints ranked, we're done (success even if pairs remain - those are ties)
            if all_ranked {
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
                    for s in constraint_strata.iter_mut() {
                        if *s == 0 {
                            *s = current_stratum + 1;
                        }
                    }
                    current_stratum += 1;
                }

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

}
