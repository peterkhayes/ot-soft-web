//! Recursive Constraint Demotion (RCD) algorithm
//!
//! This module implements the RCD algorithm for finding stratified constraint
//! rankings in Optimality Theory. The algorithm iteratively ranks constraints
//! into strata by identifying which constraints never prefer losers.

use wasm_bindgen::prelude::*;
use crate::tableau::Tableau;
use crate::fred::FRedResult;

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
fn format_html_viol(viols: usize, winner_viols: usize, is_fatal: bool) -> String {
    if viols == 0 {
        return "&nbsp;".to_string();
    }
    if viols >= 10 {
        return if is_fatal { format!("{}!", viols) } else { viols.to_string() };
    }
    if is_fatal {
        let before = winner_viols + 1;
        let after = viols.saturating_sub(before);
        format!("{}!{}", "*".repeat(before), "*".repeat(after))
    } else {
        "*".repeat(viols)
    }
}

// ────────────────────────────────────────────────────────────────────────────

/// Format a violation value centered in a column, matching VB6's centering formula.
///
/// VB6 centering: `leading = floor(col_width/2) - digit_count`, value, then
/// `trailing = col_width - floor(col_width/2)` for non-fatal, or `-1` for fatal (to
/// accommodate the `!`).
fn format_violation(col_width: usize, viols: usize, is_fatal: bool) -> String {
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
    pub(crate) fn compute_extra_analyses(&mut self, tableau: &Tableau) {
        self.constraint_necessity = tableau.compute_constraint_necessity(self);
        self.fred_result = Some(tableau.run_fred(false)); // Skeletal Basis mode
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
        apriori: &[Vec<bool>],
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
            self.fred_result = Some(
                if apriori.is_empty() {
                    tableau.run_fred_verbose(use_mib, verbose)
                } else {
                    tableau.run_fred_with_apriori_verbose(use_mib, apriori, verbose)
                }
            );
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

    /// Generate formatted text output for the RCD analysis
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        self.format_output_with_algorithm(tableau, filename, "Recursive Constraint Demotion")
    }

    /// Generate formatted text output with a configurable algorithm name
    pub(crate) fn format_output_with_algorithm(&self, tableau: &Tableau, filename: &str, algorithm_name: &str) -> String {
        let mut output = String::new();

        // Header
        output.push_str(&format!("Results of Applying {} to {}\n", algorithm_name, filename));
        output.push_str("\n\n");

        // Date and version (current date/time and version)
        let now = chrono::Local::now();
        output.push_str(&format!("{}\n\n", now.format("%-m-%-d-%Y, %-I:%M %p").to_string().to_lowercase()));
        output.push_str("OTSoft 2.7, release date 2/1/2026\n");
        output.push_str("\n\n");

        if self.tie_warning {
            output.push_str("Caution: The BCD algorithm has selected arbitrarily among tied Faithfulness constraint subsets.\n");
            output.push_str("You may wish to try changing the order of the Faithfulness constraints in the input file,\n");
            output.push_str("to see whether this results in a different ranking.\n\n\n");
        }

        // Section 1: Result
        output.push_str("1. Result\n\n");

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

        // Section 2: Tableaux
        output.push_str("2. Tableaux\n\n");

        for form in &tableau.forms {
            output.push('\n');
            output.push_str(&format!("/{}/: \n", form.input));

            // Build constraint header with stratum separators
            // Use broken bar character (U+00A6) for within-stratum, | for between-strata
            let sep_char = '\u{00A6}'; // Broken bar (¦)

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

            for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
                let c_stratum = self.constraint_strata[c_idx];

                // Add separator before this constraint
                if c_idx > 0 {
                    let prev_stratum = self.constraint_strata[c_idx - 1];
                    if c_stratum != prev_stratum {
                        header.push('|');
                    } else {
                        header.push(sep_char);
                    }
                }

                header.push_str(&constraint.abbrev());
            }
            output.push_str(&header);
            output.push('\n');

            // Find winner
            let winner_idx = form.candidates.iter()
                .position(|c| c.frequency > 0);

            // Output each candidate
            for (cand_idx, candidate) in form.candidates.iter().enumerate() {
                let is_winner = Some(cand_idx) == winner_idx;
                let marker = if is_winner { ">" } else { " " };

                // Find the first fatal violation (if any) for this loser
                let first_fatal_idx = if !is_winner {
                    winner_idx.and_then(|wi| {
                        let winner = &form.candidates[wi];
                        candidate.violations.iter().enumerate()
                            .position(|(idx, &viols)| viols > winner.violations[idx])
                    })
                } else {
                    None
                };

                // Candidate surface form (right-aligned to match expected output)
                output.push_str(&format!("{}{:<width$} ", marker, candidate.form, width = max_cand_width));

                // Violations with stratum separators
                for (c_idx, &viols) in candidate.violations.iter().enumerate() {
                    let c_stratum = self.constraint_strata[c_idx];
                    let constraint_abbrev = &tableau.constraints[c_idx].abbrev();
                    let col_width = constraint_abbrev.len();

                    // Add separator
                    if c_idx > 0 {
                        let prev_stratum = self.constraint_strata[c_idx - 1];
                        if c_stratum != prev_stratum {
                            output.push('|');
                        } else {
                            output.push(sep_char);
                        }
                    }

                    // Mark violation as fatal only if it's the FIRST fatal violation
                    let is_fatal = first_fatal_idx == Some(c_idx);
                    output.push_str(&format_violation(col_width, viols, is_fatal));
                }
                output.push('\n');
            }
            output.push('\n');
        }
        output.push('\n');

        // Section 3: Status of Proposed Constraints
        if !self.constraint_necessity.is_empty() {
            output.push_str("3. Status of Proposed Constraints:  Necessary or Unnecessary\n\n");

            let max_abbrev_width = tableau.constraints.iter()
                .map(|c| c.abbrev().len())
                .max()
                .unwrap_or(0);

            for (c_idx, necessity) in self.constraint_necessity.iter().enumerate() {
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

            // Check if mass deletion is possible
            let mass_deletion_possible = self.check_mass_deletion(tableau);
            if mass_deletion_possible {
                output.push_str("\nA check has determined that the grammar will still work even if the \n");
                output.push_str("constraints marked above as unnecessary are removed en masse.\n\n\n");
            } else {
                output.push_str("\n\n");
            }
        }

        // Section 4: Ranking Arguments (FRed)
        if let Some(ref fred) = self.fred_result {
            output.push_str(&fred.format_section4());
        }

        // Section 5: Mini-Tableaux
        if !self.mini_tableaux.is_empty() {
            output.push_str("5. Mini-Tableaux\n\n");
            output.push_str("The following small tableaux may be useful in presenting ranking arguments. \n");
            output.push_str("They include all winner-rival comparisons in which there is just one \n");
            output.push_str("winner-preferring constraint and at least one loser-preferring constraint.  \n");
            output.push_str("Constraints not violated by either candidate are omitted.\n\n");

            for mini in &self.mini_tableaux {
                self.format_mini_tableau(tableau, mini, &mut output);
            }
        }

        output
    }

    /// Check if mass deletion of unnecessary constraints still allows RCD to succeed
    fn check_mass_deletion(&self, tableau: &Tableau) -> bool {
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
        let test_result = modified_tableau.run_rcd_internal(false, &[]);
        test_result.success
    }

    /// Format a mini-tableau
    fn format_mini_tableau(&self, tableau: &Tableau, mini: &MiniTableau, output: &mut String) {
        let form = &tableau.forms[mini.form_index];
        let winner = &form.candidates[mini.winner_index];
        let loser = &form.candidates[mini.loser_index];

        output.push_str(&format!("\n/{}/: \n", form.input));

        // Build header with only included constraints
        let sep_char = '\u{00A6}'; // Broken bar (¦)
        let max_cand_width = winner.form.len().max(loser.form.len()).max(2);

        let mut header = String::new();
        for _ in 0..max_cand_width + 2 {
            header.push(' ');
        }

        for (i, &c_idx) in mini.included_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];

            // Add separator before this constraint
            if i > 0 {
                let prev_c_idx = mini.included_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    header.push('|');
                } else {
                    header.push(sep_char);
                }
            }

            header.push_str(&constraint.abbrev());
        }
        output.push_str(&header);
        output.push('\n');

        // Output winner
        output.push_str(&format!(">{:<width$} ", winner.form, width = max_cand_width));
        for (i, &c_idx) in mini.included_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];
            let col_width = constraint.abbrev().len();

            // Add separator
            if i > 0 {
                let prev_c_idx = mini.included_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    output.push('|');
                } else {
                    output.push(sep_char);
                }
            }

            output.push_str(&format_violation(col_width, winner.violations[c_idx], false));
        }
        output.push('\n');

        // Output loser (without fatal violation markers in mini-tableaux)
        output.push_str(&format!(" {:<width$} ", loser.form, width = max_cand_width));

        for (i, &c_idx) in mini.included_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];
            let col_width = constraint.abbrev().len();

            // Add separator
            if i > 0 {
                let prev_c_idx = mini.included_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    output.push('|');
                } else {
                    output.push(sep_char);
                }
            }

            output.push_str(&format_violation(col_width, loser.violations[c_idx], false));
        }
        output.push_str("\n\n");
    }

    // ── HTML output ──────────────────────────────────────────────────────────

    /// Generate an HTML document containing styled tableaux for this RCD result.
    pub fn format_html_output(&self, tableau: &Tableau, filename: &str) -> String {
        self.format_html_output_with_algorithm(tableau, filename, "Recursive Constraint Demotion")
    }

    /// Generate an HTML document with a configurable algorithm name.
    pub(crate) fn format_html_output_with_algorithm(
        &self,
        tableau: &Tableau,
        filename: &str,
        algorithm_name: &str,
    ) -> String {
        let mut out = String::new();

        // DOCTYPE, head, CSS
        out.push_str("<!DOCTYPE html>\n<html>\n<head>\n");
        out.push_str("<meta charset=\"UTF-8\">\n");
        out.push_str(&format!(
            "<title>OTSoft 2.7 {}</title>\n",
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
        out.push_str("<p>OTSoft 2.7, release date 2/1/2026</p>\n");

        if self.tie_warning {
            out.push_str(
                "<p class=\"warning\">Caution: The BCD algorithm has selected arbitrarily \
                 among tied Faithfulness constraint subsets. You may wish to try changing \
                 the order of the Faithfulness constraints in the input file, to see whether \
                 this results in a different ranking.</p>\n",
            );
        }

        // Section 1: Result
        out.push_str("<h2>1. Result</h2>\n");
        if self.success {
            out.push_str("<p class=\"success\">A ranking was found that generates the correct outputs.</p>\n");
        } else {
            out.push_str("<p class=\"failure\">No ranking was found.</p>\n");
        }
        for stratum in 1..=self.num_strata {
            out.push_str(&format!("<p><strong>Stratum #{stratum}</strong></p>\n<ul>\n"));
            for (c_idx, &c_stratum) in self.constraint_strata.iter().enumerate() {
                if c_stratum == stratum {
                    if let Some(constraint) = tableau.get_constraint(c_idx) {
                        out.push_str(&format!(
                            "  <li>{} ({})</li>\n",
                            html_escape(&constraint.full_name()),
                            html_escape(&constraint.abbrev()),
                        ));
                    }
                }
            }
            out.push_str("</ul>\n");
        }

        // Section 2: Tableaux as HTML tables
        out.push_str("<h2>2. Tableaux</h2>\n");
        for form in &tableau.forms {
            let winner_idx = form.candidates.iter().position(|c| c.frequency > 0);
            out.push_str(&self.format_html_form_table(tableau, form, winner_idx));
        }

        // Section 3: Constraint necessity
        if !self.constraint_necessity.is_empty() {
            out.push_str(
                "<h2>3. Status of Proposed Constraints: Necessary or Unnecessary</h2>\n",
            );
            out.push_str("<table class=\"necessity-table\">\n<tbody>\n");
            for (c_idx, necessity) in self.constraint_necessity.iter().enumerate() {
                let constraint = &tableau.constraints[c_idx];
                let status = match necessity {
                    ConstraintNecessity::Necessary => "Necessary",
                    ConstraintNecessity::UnnecessaryButShownForFaithfulness =>
                        "Not necessary (but included to show Faithfulness violations of a winning candidate)",
                    ConstraintNecessity::CompletelyUnnecessary => "Not necessary",
                };
                out.push_str(&format!(
                    "  <tr><td>{}</td><td>{}</td></tr>\n",
                    html_escape(&constraint.abbrev()),
                    html_escape(status),
                ));
            }
            out.push_str("</tbody>\n</table>\n");
            if self.check_mass_deletion(tableau) {
                out.push_str(
                    "<p>A check has determined that the grammar will still work even if the \
                     constraints marked above as unnecessary are removed en masse.</p>\n",
                );
            }
        }

        // Section 4: Ranking Arguments (FRed)
        if let Some(ref fred) = self.fred_result {
            out.push_str(
                "<h2>4. Ranking Arguments, based on the Fusional Reduction Algorithm</h2>\n",
            );
            out.push_str(&format!(
                "<pre>{}</pre>\n",
                html_escape(&fred.format_section4())
            ));
        }

        // Section 5: Mini-Tableaux
        if !self.mini_tableaux.is_empty() {
            out.push_str("<h2>5. Mini-Tableaux</h2>\n");
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
    fn format_html_form_table(
        &self,
        tableau: &Tableau,
        form: &crate::tableau::InputForm,
        winner_idx: Option<usize>,
    ) -> String {
        let num_constraints = tableau.constraints.len();
        let mut out = String::new();

        out.push_str("<table>\n");

        // Header row: input form + constraint abbreviations
        out.push_str("  <tr>\n");
        out.push_str(&format!("    <th>/{}/</th>\n", html_escape(&form.input)));
        for c_idx in 0..num_constraints {
            let is_last = c_idx + 1 == num_constraints;
            let has_border = !is_last
                && self.constraint_strata[c_idx] != self.constraint_strata[c_idx + 1];
            let class = html_cell_class(false, is_last, has_border);
            out.push_str(&format!(
                "    <th class=\"{}\">{}</th>\n",
                class,
                html_escape(&tableau.constraints[c_idx].abbrev()),
            ));
        }
        out.push_str("  </tr>\n");

        // Compute winner shading point (0-based constraint index where all losers are dead).
        // Cells strictly after this index are shaded in the winner row.
        let winner_shading_point: usize = if let Some(wi) = winner_idx {
            let winner = &form.candidates[wi];
            let losers: Vec<usize> = form.candidates.iter()
                .enumerate()
                .filter(|(i, c)| *i != wi && c.frequency == 0)
                .map(|(i, _)| i)
                .collect();
            if losers.is_empty() {
                usize::MAX
            } else {
                let mut dead = vec![false; losers.len()];
                let mut found = usize::MAX;
                for c_idx in 0..num_constraints {
                    for (l_pos, &l_idx) in losers.iter().enumerate() {
                        if form.candidates[l_idx].violations[c_idx] > winner.violations[c_idx] {
                            dead[l_pos] = true;
                        }
                    }
                    if dead.iter().all(|&d| d) {
                        found = c_idx;
                        break;
                    }
                }
                found
            }
        } else {
            usize::MAX
        };

        // Candidate rows
        for (cand_idx, candidate) in form.candidates.iter().enumerate() {
            let is_winner = Some(cand_idx) == winner_idx;
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
                for c_idx in 0..num_constraints {
                    let is_shaded = c_idx > winner_shading_point;
                    let is_last = c_idx + 1 == num_constraints;
                    let has_border = !is_last
                        && self.constraint_strata[c_idx] != self.constraint_strata[c_idx + 1];
                    let class = html_cell_class(is_shaded, is_last, has_border);
                    let content = format_html_viol(candidate.violations[c_idx], 0, false);
                    out.push_str(&format!("    <td class=\"{class}\">{content}</td>\n"));
                }
            } else {
                // Loser row: find first fatal violation, shade cells after it.
                // The fatal cell itself is NOT shaded (style chosen before flag is set).
                let first_fatal = if let Some(wi) = winner_idx {
                    let winner = &form.candidates[wi];
                    candidate.violations.iter()
                        .enumerate()
                        .position(|(i, &v)| v > winner.violations[i])
                } else {
                    None
                };

                let mut fatal_seen = false;
                for c_idx in 0..num_constraints {
                    let is_fatal = first_fatal == Some(c_idx);
                    // Shade based on fatal_seen BEFORE updating it (matches VB6 behavior)
                    let is_shaded = fatal_seen;
                    let is_last = c_idx + 1 == num_constraints;
                    let has_border = !is_last
                        && self.constraint_strata[c_idx] != self.constraint_strata[c_idx + 1];
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

    /// Render a mini-tableau as an HTML table.
    fn format_html_mini_tableau(&self, tableau: &Tableau, mini: &MiniTableau) -> String {
        let form = &tableau.forms[mini.form_index];
        let winner = &form.candidates[mini.winner_index];
        let loser = &form.candidates[mini.loser_index];
        let included = &mini.included_constraints;
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
                    result.compute_extra_analyses(self);
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
                    result.compute_extra_analyses(self);
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
    pub(crate) fn compute_constraint_necessity(&self, rcd_result: &RCDResult) -> Vec<ConstraintNecessity> {
        let mut necessity = vec![ConstraintNecessity::Necessary; self.constraints.len()];

        // Only analyze if RCD succeeded
        if !rcd_result.success {
            return necessity;
        }

        for (c_idx, nec) in necessity.iter_mut().enumerate() {
            // Test if constraint is necessary
            if !self.is_constraint_necessary(c_idx) {
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

    /// Test if a constraint is necessary by running RCD without it
    fn is_constraint_necessary(&self, constraint_idx: usize) -> bool {
        // Create modified tableau with constraint violations zeroed
        let modified_tableau = self.clone_with_constraint_removed(constraint_idx);

        // Run RCD on modified tableau (without computing extra analyses to avoid recursion)
        let test_result = modified_tableau.run_rcd_internal(false, &[]);

        // Constraint is necessary if RCD fails without it
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
        std::fs::read_to_string("../examples/tiny/input.txt")
            .expect("Failed to load examples/tiny/input.txt")
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
