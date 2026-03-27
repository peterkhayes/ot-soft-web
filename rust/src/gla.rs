//! Gradual Learning Algorithm (GLA)
//!
//! Online error-driven learner in two modes:
//!
//! **Stochastic OT** (Boersma 1997, Boersma & Hayes 2001): Each trial adds
//! Gaussian noise (σ=NoiseMark/NoiseFaith per schedule stage) to ranking values,
//! sorts constraints, evaluates candidates by strict domination, then adjusts
//! ranking values ±PlastMark/PlastFaith.
//!
//! **Online MaxEnt** (Jäger 2004): Each trial samples a candidate via
//! exp(−H)/Z probabilities, then adjusts weights by `plasticity × (gen_viols − obs_viols)`.
//!
//! Reproduces VB6 boersma.frm:GLACore, RankingValueAdjustment, MaxEntSampledCandidate,
//! GLATestGrammar, GenerateMaxEntPredictions.

use wasm_bindgen::prelude::*;
use crate::rng::{Rng, GaussianMode};
use crate::schedule::LearningSchedule;
use crate::tableau::Tableau;

// ─── OT Evaluation (Stochastic OT) ──────────────────────────────────────────
//
// Reproduces VB6 DetermineSelectionPointsAndSort + GenerateAFormStochasticOT.
//
// Gaussian noise is added to each ranking value, with separate sigma for
// Markedness (noise_mark) and Faithfulness (noise_faith) constraints.

fn ot_evaluate(
    candidates: &[crate::tableau::Candidate],
    ranking_values: &[f64],
    is_faith: &[bool],
    noise_mark: f64,
    noise_faith: f64,
    rng: &mut Rng,
) -> usize {
    let nc = ranking_values.len();
    let n_cands = candidates.len();

    // Add Gaussian noise to each ranking value (σ = noise_mark or noise_faith)
    let noisy_rv: Vec<f64> = ranking_values
        .iter()
        .zip(is_faith.iter())
        .map(|(&rv, &faith)| {
            let sigma = if faith { noise_faith } else { noise_mark };
            rv + sigma * rng.gaussian()
        })
        .collect();

    // Sort constraint indices descending by noisy ranking value (highest = most important)
    let mut sorted: Vec<usize> = (0..nc).collect();
    sorted.sort_by(|&a, &b| noisy_rv[b].partial_cmp(&noisy_rv[a]).unwrap_or(std::cmp::Ordering::Equal));

    // King-of-the-hill OT evaluation: start with candidate 0
    let mut best = 0;
    for ci in 1..n_cands {
        for &c in &sorted {
            let best_v = candidates[best].violations[c];
            let ci_v = candidates[ci].violations[c];
            if best_v > ci_v {
                // current champion violates more → ci wins
                best = ci;
                break;
            } else if best_v < ci_v {
                // ci violates more → ci loses
                break;
            }
            // equal: go to next constraint
        }
    }
    best
}

// ─── MaxEnt Sampling ─────────────────────────────────────────────────────────
//
// Reproduces VB6 boersma.frm:MaxEntSampledCandidate.

fn maxent_sample(
    candidates: &[crate::tableau::Candidate],
    weights: &[f64],
    rng: &mut Rng,
) -> usize {
    let nc = weights.len();

    // Compute exp(-harmony) for each candidate
    let e_harmonies: Vec<f64> = candidates.iter().map(|cand| {
        let h: f64 = (0..nc).map(|c| weights[c] * cand.violations[c] as f64).sum();
        (-h).exp()
    }).collect();

    let z: f64 = e_harmonies.iter().sum();

    // Sample proportionally to e_harmonies
    let r = rng.uniform() * z;
    let mut cumsum = 0.0;
    for (ci, &eh) in e_harmonies.iter().enumerate() {
        cumsum += eh;
        if cumsum >= r {
            return ci;
        }
    }
    candidates.len() - 1
}

// ─── MaxEnt Exact Probabilities ──────────────────────────────────────────────
//
// Reproduces VB6 boersma.frm:GenerateMaxEntPredictions.
// Returns predicted probability for each candidate in each form.

fn maxent_predictions(tableau: &Tableau, weights: &[f64]) -> Vec<Vec<f64>> {
    let nc = weights.len();
    tableau.forms.iter().map(|form| {
        let e_harmonies: Vec<f64> = form.candidates.iter().map(|cand| {
            let h: f64 = (0..nc).map(|c| weights[c] * cand.violations[c] as f64).sum();
            (-h).exp()
        }).collect();
        let z: f64 = e_harmonies.iter().sum();
        if z == 0.0 {
            vec![0.0; form.candidates.len()]
        } else {
            e_harmonies.iter().map(|&eh| eh / z).collect()
        }
    }).collect()
}

// ─── Magri Promotion Amount ───────────────────────────────────────────────────
//
// Reproduces VB6 boersma.frm:MagriPromotionAmount.
//
// Counts how many constraints prefer the generated form (would be promoted) and
// how many prefer the observed form (would be demoted), then returns:
//   NumberOfConstraintsDemoted / (NumberOfConstraintsPromoted + 1)
//
// Applied as a scalar multiplier on the promotion plasticity only (demotion is
// unmodified). The +1 in the denominator avoids division by zero.

fn magri_promotion_amount(
    gen_cand: &crate::tableau::Candidate,
    obs_cand: &crate::tableau::Candidate,
    nc: usize,
) -> f64 {
    let mut promoted = 0u32;
    let mut demoted = 0u32;
    for c in 0..nc {
        match gen_cand.violations[c].cmp(&obs_cand.violations[c]) {
            std::cmp::Ordering::Greater => promoted += 1,
            std::cmp::Ordering::Less => demoted += 1,
            std::cmp::Ordering::Equal => {}
        }
    }
    demoted as f64 / (promoted as f64 + 1.0)
}

// ─── A priori ranking enforcement ────────────────────────────────────────────
//
// Reproduces VB6 boersma.frm:AdjustAPrioriRankings_Up and :AdjustAPrioriRankings_Down.
//
// table[i][j] = true means constraint i a priori dominates constraint j (i >> j),
// so ranking_values[i] must be >= ranking_values[j] + gap.

/// Raise dominators until all a priori gaps are satisfied.
/// Called after a constraint is strengthened (raised). Loops until stable.
fn adjust_apriori_up(rv: &mut [f64], apriori: &[Vec<bool>], gap: f64) {
    let nc = rv.len();
    loop {
        let mut changed = false;
        for i in 0..nc {
            for j in 0..nc {
                if apriori[i][j] {
                    let margin = rv[i] - rv[j];
                    // VB6 uses gap - 0.0001 to avoid floating-point equality issues
                    if margin < gap - 0.0001 {
                        rv[i] = rv[j] + gap;
                        changed = true;
                    }
                }
            }
        }
        if !changed { break; }
    }
}

/// Lower dominated constraints until all a priori gaps are satisfied.
/// Called after a constraint is weakened (lowered). Loops until stable.
fn adjust_apriori_down(rv: &mut [f64], apriori: &[Vec<bool>], gap: f64) {
    let nc = rv.len();
    loop {
        let mut changed = false;
        for i in 0..nc {
            for j in 0..nc {
                if apriori[i][j] {
                    let margin = rv[i] - rv[j];
                    if margin < gap {
                        rv[j] = rv[i] - gap;
                        changed = true;
                    }
                }
            }
        }
        if !changed { break; }
    }
}

// ─── Active Constraint Detection ─────────────────────────────────────────────
//
// Reproduces VB6 boersma.frm:GLATestGrammar (active constraint tracking) and
// PrintActiveConstraints.
//
// A constraint is "active" if it causes the winning candidate to defeat a rival
// in at least one input form, using deterministic (no-noise) evaluation over the
// final ranking values.

fn compute_active_constraints(tableau: &Tableau, ranking_values: &[f64]) -> Vec<bool> {
    let nc = ranking_values.len();
    let mut active = vec![false; nc];

    // Sort constraint indices by descending ranking value (same order VB6 uses via SlotFiller)
    let mut sorted: Vec<usize> = (0..nc).collect();
    sorted.sort_by(|&a, &b| {
        ranking_values[b]
            .partial_cmp(&ranking_values[a])
            .unwrap_or(std::cmp::Ordering::Equal)
    });

    for form in &tableau.forms {
        let n_cands = form.candidates.len();
        if n_cands < 2 {
            continue;
        }

        // Find winner by king-of-hill evaluation (deterministic, no noise)
        let mut winner = 0;
        for ci in 1..n_cands {
            for &c in &sorted {
                match form.candidates[winner].violations[c]
                    .cmp(&form.candidates[ci].violations[c])
                {
                    std::cmp::Ordering::Greater => {
                        winner = ci;
                        break;
                    }
                    std::cmp::Ordering::Less => break,
                    std::cmp::Ordering::Equal => {}
                }
            }
        }

        // A constraint is active if it kills any rival: winner violates it less
        // than the rival, and it is the first distinguishing constraint in sorted order.
        for rival in 0..n_cands {
            if rival == winner {
                continue;
            }
            for &c in &sorted {
                match form.candidates[winner].violations[c]
                    .cmp(&form.candidates[rival].violations[c])
                {
                    std::cmp::Ordering::Less => {
                        active[c] = true;
                        break;
                    }
                    std::cmp::Ordering::Greater => break,
                    std::cmp::Ordering::Equal => {}
                }
            }
        }
    }

    active
}

// ─── GLA Result ──────────────────────────────────────────────────────────────

/// Result of running the Gradual Learning Algorithm (GLA)
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct GlaResult {
    /// Final ranking values (StochasticOT) or weights (MaxEnt)
    ranking_values: Vec<f64>,
    /// Predicted probabilities from testing: test_probs[form_idx][cand_idx]
    test_probs: Vec<Vec<f64>>,
    /// Log likelihood of training data under the tested grammar
    log_likelihood: f64,
    /// Whether this is MaxEnt mode (false = StochasticOT)
    maxent_mode: bool,
    /// Pre-formatted schedule description for output
    schedule_description: String,
    /// Parameters stored for output formatting
    test_trials: usize,
    /// Gaussian prior parameters (MaxEnt mode only)
    gaussian_prior: bool,
    sigma: f64,
    /// Whether the Magri update rule was active (StochasticOT only)
    magri_update_rule: bool,
    /// Whether negative weights were permitted
    negative_weights_ok: bool,
    /// Sum of squared error between input and generated proportions
    error_term: f64,
    /// Total number of candidates across all forms
    total_rivals: usize,
    /// Learning time in seconds
    learning_time_secs: f64,
    /// A priori ranking table used during learning (empty if none).
    apriori: Vec<Vec<bool>>,
    /// Minimum gap enforced between a priori ranked constraints.
    apriori_gap: f64,
    /// Optional history of ranking values/weights recorded during learning.
    history: Option<String>,
    /// Optional full (annotated) history: trial, input, generated, heard, values.
    full_history: Option<String>,
    /// Optional history of candidate probabilities at every trial (MaxEnt only).
    candidate_prob_history: Option<String>,
    /// Which constraints are active (cause the winner to defeat a rival).
    active_constraints: Vec<bool>,
}

impl GlaResult {
    /// Create a minimal GlaResult for formatting the pairwise probability table.
    pub(crate) fn for_pairwise_table(ranking_values: Vec<f64>) -> Self {
        Self {
            ranking_values,
            test_probs: Vec::new(),
            log_likelihood: 0.0,
            maxent_mode: false,
            schedule_description: String::new(),
            test_trials: 0,
            gaussian_prior: false,
            sigma: 0.0,
            magri_update_rule: false,
            negative_weights_ok: false,
            error_term: 0.0,
            total_rivals: 0,
            learning_time_secs: 0.0,
            apriori: Vec::new(),
            apriori_gap: 20.0,
            history: None,
            full_history: None,
            candidate_prob_history: None,
            active_constraints: Vec::new(),
        }
    }
}

#[wasm_bindgen]
impl GlaResult {
    pub fn get_ranking_value(&self, constraint_index: usize) -> f64 {
        self.ranking_values.get(constraint_index).copied().unwrap_or(0.0)
    }

    pub fn get_test_prob(&self, form_index: usize, cand_index: usize) -> f64 {
        self.test_probs
            .get(form_index)
            .and_then(|f| f.get(cand_index))
            .copied()
            .unwrap_or(0.0)
    }

    pub fn log_likelihood(&self) -> f64 {
        self.log_likelihood
    }

    pub fn is_maxent_mode(&self) -> bool {
        self.maxent_mode
    }

    pub fn gaussian_prior(&self) -> bool {
        self.gaussian_prior
    }

    pub fn sigma(&self) -> f64 {
        self.sigma
    }

    pub fn history(&self) -> Option<String> {
        self.history.clone()
    }

    pub fn full_history(&self) -> Option<String> {
        self.full_history.clone()
    }

    pub fn candidate_prob_history(&self) -> Option<String> {
        self.candidate_prob_history.clone()
    }

    pub fn get_active_constraint(&self, constraint_index: usize) -> bool {
        self.active_constraints.get(constraint_index).copied().unwrap_or(false)
    }

    /// Format results as text output for download.
    /// Reproduces the structure of OTSoft's GLA text output.
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        let mut out = String::new();
        let mode_name = if self.maxent_mode {
            "GLA-MaxEnt"
        } else {
            "GLA-Stochastic OT"
        };
        let value_label = if self.maxent_mode { "Weights" } else { "Ranking Values" };

        // Header
        out.push_str(&format!(
            "Result of Applying {} to {}\n\n\n",
            mode_name, filename
        ));

        let now = chrono::Local::now();
        out.push_str(&format!(
            "{}\n\n",
            now.format("%-m-%-d-%Y, %-I:%M %p")
                .to_string()
                .to_lowercase()
        ));
        out.push_str(crate::VERSION_STRING);
        out.push_str("\n\n\n");

        // Parameters
        out.push_str("Parameters:\n");
        out.push_str(&self.schedule_description);
        out.push_str(&format!("   Times to test grammar: {}\n", self.test_trials));
        if self.maxent_mode && self.gaussian_prior {
            out.push_str(&format!("   Sigma: {}\n", self.sigma));
            out.push_str("   A Gaussian prior for MaxEnt learning was in effect.\n");
        }
        if !self.maxent_mode && self.magri_update_rule {
            out.push_str("   The Magri update rule was employed.\n");
        }
        if !self.apriori.is_empty() {
            out.push_str(&format!(
                "   A priori rankings in effect (minimum gap: {}).\n",
                self.apriori_gap
            ));
        }
        out.push_str("\n\n");

        // Section 1: Ranking Values / Weights Found (original constraint order)
        // VB6 PrintGLAResults: uses FourDecPlaces (##,##0.0000)
        out.push_str(&format!("1. {} Found\n\n", value_label));
        for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
            out.push_str(&format!(
                "   {:<42}{:.4}\n",
                constraint.full_name(),
                self.ranking_values[c_idx]
            ));
        }
        out.push('\n');
        out.push_str(&format!(
            "   Log likelihood of data: {:.6}\n",
            self.log_likelihood
        ));
        out.push_str("\n\n");

        // Section 2: Matchup to Input Frequencies
        // Reproduces VB6 GLATestGrammar print section.
        // VB6 shows proportions (0–1) with 4 decimal places, plus raw counts.
        out.push_str("2. Matchup to Input Frequencies\n\n");
        for (form_idx, form) in tableau.forms.iter().enumerate() {
            let total_freq: f64 = form.candidates.iter().map(|c| c.frequency as f64).sum();
            out.push_str(&format!("/{}/\n", form.input));

            let max_cand_width = form.candidates.iter()
                .map(|c| c.form.len())
                .max()
                .unwrap_or(0)
                .max(2);

            // Column headers depend on mode
            // VB6 StochasticOT: "Input Fr. Gen Fr.  Input #     Gen. #"
            // VB6 MaxEnt:       "Input Fr.  P     Input #"
            if self.maxent_mode {
                out.push_str(&format!(
                    "  {:<width$}  {:>10}  {:>10}  {:>8}\n",
                    "", "Input Fr.", "Prob", "Input #",
                    width = max_cand_width
                ));
            } else {
                out.push_str(&format!(
                    "  {:<width$}  {:>10}  {:>10}  {:>8}  {:>8}\n",
                    "", "Input Fr.", "Gen Fr.", "Input #", "Gen. #",
                    width = max_cand_width
                ));
            }

            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                let obs_prop = if total_freq > 0.0 {
                    cand.frequency as f64 / total_freq
                } else {
                    0.0
                };
                let gen_prop = self.test_probs
                    .get(form_idx)
                    .and_then(|f| f.get(cand_idx))
                    .copied()
                    .unwrap_or(0.0);
                let marker = if cand.frequency > 0 { ">" } else { " " };

                if self.maxent_mode {
                    out.push_str(&format!(
                        "  {}{:<width$}  {:>10.4}  {:>10.4}  {:>8}\n",
                        marker, cand.form, obs_prop, gen_prop, cand.frequency,
                        width = max_cand_width
                    ));
                } else {
                    // Generated number = proportion × test_trials
                    let gen_count = (gen_prop * self.test_trials as f64).round() as usize;
                    out.push_str(&format!(
                        "  {}{:<width$}  {:>10.4}  {:>10.4}  {:>8}  {:>8}\n",
                        marker, cand.form, obs_prop, gen_prop, cand.frequency, gen_count,
                        width = max_cand_width
                    ));
                }
            }
            out.push('\n');
        }

        let mut section_num = 2;

        // Section 3 (Stochastic OT only): Testing the Grammar: Details
        // Reproduces VB6 boersma.frm:GLATestGrammar (lines 4167-4177).
        if !self.maxent_mode {
            section_num += 1;
            out.push_str(&format!("{}. Testing the Grammar: Details\n\n", section_num));
            out.push_str(&format!(
                "The grammar was tested for {} cycles.\n",
                self.test_trials
            ));
            if self.total_rivals > 0 {
                out.push_str(&format!(
                    "Average error per candidate: {:.3} percent\n",
                    100.0 * self.error_term / self.total_rivals as f64
                ));
            }
            out.push_str(&format!(
                "Learning time: {:.3} minutes\n",
                self.learning_time_secs / 60.0
            ));
            if self.negative_weights_ok {
                out.push_str("Negative weights were permitted.\n");
            } else {
                out.push_str("Negative weights were not permitted.\n");
            }
            out.push_str("\n\n");
        }

        // Sorted ranking values / weights
        // VB6 PrintGLAResults: uses FourDecPlaces (##,##0.0000)
        section_num += 1;
        out.push_str(&format!("{}. {} (sorted)\n\n", section_num, value_label));
        let sorted = crate::tableau::sorted_indices_descending(&self.ranking_values);
        for &c_idx in &sorted {
            out.push_str(&format!(
                "   {:<42}{:.4}\n",
                tableau.constraints[c_idx].full_name(),
                self.ranking_values[c_idx]
            ));
        }

        // Pairwise ranking probabilities (Stochastic OT only)
        if !self.maxent_mode {
            section_num += 1;
            out.push_str("\n\n");
            out.push_str(&self.format_pairwise_probabilities(tableau, section_num));
        }

        // Active Constraints
        // Reproduces VB6 boersma.frm:PrintActiveConstraints.
        // Lists constraints in sorted order with "Active" or "Inactive" status.
        {
            section_num += 1;
            let sorted = crate::tableau::sorted_indices_descending(&self.ranking_values);
            let name_width = tableau.constraints.iter()
                .map(|c| c.full_name().len())
                .max()
                .unwrap_or(10)
                .max(10);
            out.push_str(&format!("\n\n{section_num}. Active Constraints\n\n"));
            out.push_str("A constraint is active if it causes the winning candidate to defeat a rival\nin at least one competition.\n\n");
            out.push_str(&format!("   {:<width$}  Status\n", "Constraint", width = name_width));
            for &c_idx in &sorted {
                let status = if self.active_constraints.get(c_idx).copied().unwrap_or(false) {
                    "Active"
                } else {
                    "Inactive"
                };
                out.push_str(&format!(
                    "   {:<width$}  {}\n",
                    tableau.constraints[c_idx].full_name(),
                    status,
                    width = name_width
                ));
            }
        }

        // A priori rankings (if any)
        // Reproduces VB6 Main.frm:PrintOutTheAprioriRankings
        if !self.apriori.is_empty() {
            let nc = tableau.constraints.len();
            let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();

            section_num += 1;
            out.push_str(&format!("\n\n{section_num}. A Priori Rankings\n\n"));
            out.push_str("In the following table, \"yes\" means that the constraint of the indicated\n");
            out.push_str("row was marked a priori to dominate the constraint in the given column.\n\n");

            // Print the table: rows = dominators, columns = dominated
            // VB6 switches axes: table[i][j]=true → row=j, col=i shows "yes"
            let col_width = abbrevs.iter().map(|a| a.len()).max().unwrap_or(3).max(3);
            let row_label_width = col_width + 2;

            // Header row
            out.push_str(&format!("{:width$}", "", width = row_label_width));
            for abbrev in &abbrevs {
                out.push_str(&format!("  {:>width$}", abbrev, width = col_width));
            }
            out.push('\n');

            // Data rows
            for (row_i, row_abbrev) in abbrevs.iter().enumerate() {
                out.push_str(&format!("{:<width$}", row_abbrev, width = row_label_width));
                for col_j in 0..nc {
                    // VB6 switches axes: table[col_j][row_i] = true means col_j >> row_i
                    // so "yes" appears at (row=row_i, col=col_j) when col_j dominates row_i
                    let cell = if self.apriori[col_j][row_i] { "yes" } else { "" };
                    out.push_str(&format!("  {:>width$}", cell, width = col_width));
                }
                out.push('\n');
            }

            out.push_str(&format!(
                "\n   An a priori ranking was implemented as a minimal difference\n   in ranking values of {}.\n",
                self.apriori_gap
            ));
        }

        out
    }

    /// Format pairwise ranking probability matrix.
    ///
    /// Produces the table from VB6 `boersma.frm:PrintPairwiseRankingProbabilities()`:
    /// rows/columns are constraints sorted by descending ranking value, upper
    /// triangle shows P(row >> col) from the full 481-entry lookup table, lower
    /// triangle is empty (shaded in VB6).
    ///
    /// Only meaningful in Stochastic OT mode (not MaxEnt).
    pub fn pairwise_probabilities_json(&self, tableau: &Tableau) -> String {
        let sorted = crate::tableau::sorted_indices_descending(&self.ranking_values);
        let n = sorted.len();
        let headers: Vec<String> =
            sorted.iter().map(|&i| tableau.constraints[i].abbrev()).collect();

        let matrix: Vec<Vec<String>> = (0..n.saturating_sub(1))
            .map(|i| {
                (1..n)
                    .map(|j| {
                        if j <= i {
                            String::new()
                        } else {
                            let diff =
                                self.ranking_values[sorted[i]] - self.ranking_values[sorted[j]];
                            crate::hasse::lookup_gla_probability_full(diff).to_string()
                        }
                    })
                    .collect()
            })
            .collect();

        // Build JSON manually to avoid adding serde_json as a production dependency.
        let json_str = |s: &str| -> String { format!("\"{}\"", s.replace('\\', "\\\\").replace('"', "\\\"")) };
        let headers_json: Vec<String> = headers.iter().map(|h| json_str(h)).collect();
        let matrix_json: Vec<String> = matrix
            .iter()
            .map(|row| {
                let cells: Vec<String> = row.iter().map(|c| json_str(c)).collect();
                format!("[{}]", cells.join(","))
            })
            .collect();
        format!(
            "{{\"headers\":[{}],\"matrix\":[{}]}}",
            headers_json.join(","),
            matrix_json.join(",")
        )
    }

    /// Only meaningful in Stochastic OT mode (not MaxEnt).
    pub fn format_pairwise_probabilities(&self, tableau: &Tableau, section_num: usize) -> String {
        let sorted = crate::tableau::sorted_indices_descending(&self.ranking_values);
        let n = sorted.len();
        let abbrevs: Vec<String> = sorted.iter().map(|&i| tableau.constraints[i].abbrev()).collect();

        let mut out = String::new();
        out.push_str(&format!("{section_num}. Ranking Value to Ranking Probability Conversion\n\n"));
        out.push_str("The computed ranking values imply the pairwise ranking probabilities given below.\n");
        out.push_str("In the table, the probability given is that of the constraint in the row headings\n");
        out.push_str("outranking the constraint in the column headings.\n\n");

        // Determine column width: max abbreviation length, at least 8 for probability strings
        let col_width = abbrevs.iter().map(|a| a.len()).max().unwrap_or(4).max(8);

        // Column headers (skip first column since it's the row label area)
        let label_width = col_width + 2; // row label column
        out.push_str(&format!("{:width$}", "", width = label_width));
        for abbrev in &abbrevs[1..] {
            out.push_str(&format!("  {:>width$}", abbrev, width = col_width));
        }
        out.push('\n');

        // Data rows
        for (i, row_abbrev) in abbrevs[..n - 1].iter().enumerate() {
            out.push_str(&format!("{:<width$}", row_abbrev, width = label_width));
            for j in 1..n {
                if j <= i {
                    // Lower triangle: shaded/empty
                    out.push_str(&format!("  {:>width$}", "", width = col_width));
                } else {
                    let diff = self.ranking_values[sorted[i]] - self.ranking_values[sorted[j]];
                    let prob = crate::hasse::lookup_gla_probability_full(diff);
                    out.push_str(&format!("  {:>width$}", prob, width = col_width));
                }
            }
            out.push('\n');
        }

        out
    }
}

// ─── Main GLA Algorithm ───────────────────────────────────────────────────────

fn parse_apriori_from_opts(
    tableau: &Tableau,
    opts: &crate::GlaOptions,
) -> Result<Vec<Vec<bool>>, String> {
    let text = opts.apriori_text();
    if text.trim().is_empty() {
        return Ok(vec![]);
    }
    let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
    crate::apriori::parse_apriori(&text, &abbrevs)
}

/// Build a stripped-down GlaOptions for collate runs (no history, no schedule/apriori text).
fn collate_run_opts(opts: &crate::GlaOptions) -> crate::GlaOptions {
    crate::GlaOptions {
        maxent_mode: opts.maxent_mode,
        cycles: opts.cycles,
        initial_plasticity: opts.initial_plasticity,
        final_plasticity: opts.final_plasticity,
        test_trials: opts.test_trials,
        negative_weights_ok: opts.negative_weights_ok,
        gaussian_prior: opts.gaussian_prior,
        sigma: opts.sigma,
        magri_update_rule: opts.magri_update_rule,
        exact_proportions: opts.exact_proportions,
        apriori_gap: opts.apriori_gap,
        generate_history: false,
        generate_full_history: false,
        generate_candidate_prob_history: false,
        learning_schedule: String::new(),
        apriori_text: String::new(),
    }
}

/// Append G and O records for one GLA run to `out`.
///
/// G records: `G\t{run_idx}\t{constraint_abbrev}\t{ranking_value}`
/// O records: `O\t{run_idx}\t{form_idx}\t{input}\t{rival_idx}\t{rival_form}\t{freq}\t{pct_gen}`
/// VB6 skips candidate 0 (the winner), so O records start from rival index 1.
fn append_collate_run(tableau: &Tableau, run_idx: usize, result: &GlaResult, out: &mut String) {
    // G records: constraint ranking values / weights
    for (ci, constraint) in tableau.constraints.iter().enumerate() {
        out.push_str(&format!("G\t{}\t{}\t{:.3}\n", run_idx, constraint.abbrev(), result.ranking_values[ci]));
    }

    // O records: rival candidates with predicted probabilities.
    for (fi, form) in tableau.forms.iter().enumerate() {
        for ri in 1..form.candidates.len() {
            let cand = &form.candidates[ri];
            let pct_gen = result.test_probs.get(fi).and_then(|f| f.get(ri)).copied().unwrap_or(0.0) * 100.0;
            out.push_str(&format!(
                "O\t{}\t{}\t{}\t{}\t{}\t{}\t{:.4}\n",
                run_idx, fi + 1, form.input, ri, cand.form, cand.frequency, pct_gen,
            ));
        }
    }
}

// ─── GLA Learning Helpers ────────────────────────────────────────────────────

/// Grouped optional history buffers for the GLA learning loop.
struct HistoryBuffers {
    /// Ranking values after each mismatch (tab-separated, one row per mismatch).
    history: Option<String>,
    /// Annotated log: trial#, input, generated, heard, per-constraint values.
    full_history: Option<String>,
    /// MaxEnt only: predicted probabilities for every candidate at every trial.
    cand_prob: Option<String>,
}

/// Initialize the three optional history buffers with their header rows.
fn init_history_buffers(
    tableau: &Tableau,
    maxent_mode: bool,
    nc: usize,
    opts: &crate::GlaOptions,
) -> HistoryBuffers {
    use std::fmt::Write;

    let history = if opts.generate_history {
        let mut header = String::from("Trial");
        for c in &tableau.constraints {
            header.push('\t');
            header.push_str(&c.abbrev());
        }
        header.push('\n');
        Some(header)
    } else {
        None
    };

    let full_history = if opts.generate_full_history {
        let mut header = String::from("Trial #\tInput\tGenerated\tHeard");
        for c in &tableau.constraints {
            write!(header, "\t{}", c.abbrev()).unwrap();
        }
        header.push('\n');
        // Initial row: starting values before learning begins
        header.push_str("(Initial)\t\t\t");
        let initial_value = if maxent_mode { 0.0 } else { 100.0 };
        for _ in 0..nc {
            write!(header, "\t{initial_value:.4}").unwrap();
        }
        header.push('\n');
        Some(header)
    } else {
        None
    };

    let cand_prob = if maxent_mode && opts.generate_candidate_prob_history {
        let mut header = String::from("Trial #");
        for form in &tableau.forms {
            for cand in &form.candidates {
                write!(header, "\t/{}/ -> {}", form.input, cand.form).unwrap();
            }
        }
        header.push('\n');
        // Initial row: predictions with zero weights
        let zero_weights = vec![0.0; nc];
        let initial_preds = maxent_predictions(tableau, &zero_weights);
        header.push_str("(initial)");
        for form_preds in &initial_preds {
            for &p in form_preds {
                write!(header, "\t{}", p).unwrap();
            }
        }
        header.push('\n');
        Some(header)
    } else {
        None
    };

    HistoryBuffers { history, full_history, cand_prob }
}

/// Select the next training exemplar from the pool.
///
/// With exact proportions, uses Fisher-Yates shuffle and sequential cursor.
/// Otherwise, selects uniformly at random.
fn select_training_exemplar(
    training_pool: &mut [(usize, usize)],
    pool_size: usize,
    pool_cursor: &mut usize,
    exact_proportions: bool,
    rng: &mut Rng,
) -> (usize, usize) {
    if exact_proportions {
        if *pool_cursor >= pool_size {
            // Fisher-Yates shuffle
            for i in (1..pool_size).rev() {
                let j = (rng.uniform() * (i + 1) as f64) as usize;
                training_pool.swap(i, j);
            }
            *pool_cursor = 0;
        }
        let pair = training_pool[*pool_cursor];
        *pool_cursor += 1;
        pair
    } else {
        let r = rng.uniform();
        let idx = ((r * pool_size as f64) as usize).min(pool_size - 1);
        training_pool[idx]
    }
}

/// Update ranking values after a mismatch (generated != observed).
///
/// MaxEnt: likelihood-based gradient + optional Gaussian prior.
/// StochasticOT: ±plasticity with optional Magri update rule.
#[allow(clippy::too_many_arguments)]
fn update_ranking_values(
    ranking_values: &mut [f64],
    gen_cand: &crate::tableau::Candidate,
    obs_cand: &crate::tableau::Candidate,
    is_faith: &[bool],
    plast_mark: f64,
    plast_faith: f64,
    maxent_mode: bool,
    gaussian_prior: bool,
    sigma: f64,
    negative_weights_ok: bool,
    magri_update_rule: bool,
    nc: usize,
) {
    if maxent_mode {
        // MaxEnt update: reproduces VB6 RankingValueAdjustment (MaxEnt branch).
        // PriorBasedChange = plasticity * (weight - mu) / sigma² / 2  (mu=0)
        // The "/2" is "out of the blue" per the VB6 author's comment; reproduced for fidelity.
        let sigma_sq = sigma * sigma;
        for (c, rv) in ranking_values.iter_mut().enumerate() {
            let plast = if is_faith[c] { plast_faith } else { plast_mark };
            let gen_v = gen_cand.violations[c] as f64;
            let obs_v = obs_cand.violations[c] as f64;
            let likelihood_change = plast * (obs_v - gen_v);
            let prior_change = if gaussian_prior {
                plast * *rv / sigma_sq / 2.0
            } else {
                0.0
            };
            *rv -= likelihood_change + prior_change;
            if !negative_weights_ok && *rv < 0.0 {
                *rv = 0.0;
            }
        }
    } else {
        // StochasticOT update: ±plasticity per constraint.
        // Magri update rule: promotion scaled by demoted/promoted ratio.
        let promo_scale = if magri_update_rule {
            magri_promotion_amount(gen_cand, obs_cand, nc)
        } else {
            1.0
        };
        for (c, rv) in ranking_values.iter_mut().enumerate() {
            let plast = if is_faith[c] { plast_faith } else { plast_mark };
            if gen_cand.violations[c] > obs_cand.violations[c] {
                *rv += plast * promo_scale;
            } else if gen_cand.violations[c] < obs_cand.violations[c] {
                *rv -= plast;
            }
        }
    }
}

/// Record mismatch to history buffer and full history buffer.
#[allow(clippy::too_many_arguments)]
fn record_mismatch_histories(
    bufs: &mut HistoryBuffers,
    tableau: &Tableau,
    ranking_values: &[f64],
    trial_number: usize,
    sel_form: usize,
    generated: usize,
    sel_cand: usize,
    is_faith: &[bool],
    plast_mark: f64,
    plast_faith: f64,
    maxent_mode: bool,
    nc: usize,
) {
    use std::fmt::Write;

    if let Some(ref mut buf) = bufs.history {
        write!(buf, "{}", trial_number).unwrap();
        for rv in ranking_values.iter() {
            write!(buf, "\t{rv:.4}").unwrap();
        }
        buf.push('\n');
    }

    if let Some(ref mut buf) = bufs.full_history {
        // VB6 StochasticOT: per-constraint delta + value written first
        if !maxent_mode {
            let gen_cand = &tableau.forms[sel_form].candidates[generated];
            let obs_cand = &tableau.forms[sel_form].candidates[sel_cand];
            for c in 0..nc {
                let gen_v = gen_cand.violations[c];
                let obs_v = obs_cand.violations[c];
                let plast = if is_faith[c] { plast_faith } else { plast_mark };
                if gen_v > obs_v {
                    write!(buf, "\t{plast}\t{:.4}", ranking_values[c]).unwrap();
                } else if gen_v < obs_v {
                    write!(buf, "\t-{plast}\t{:.4}", ranking_values[c]).unwrap();
                } else {
                    buf.push_str("\t\t");
                }
            }
        }
        // GLACore part: trial#, input, generated, heard, final values
        write!(buf, "{}\t{}\t{}\t{}",
            trial_number,
            tableau.forms[sel_form].input,
            tableau.forms[sel_form].candidates[generated].form,
            tableau.forms[sel_form].candidates[sel_cand].form,
        ).unwrap();
        for rv in ranking_values.iter() {
            write!(buf, "\t{rv:.4}").unwrap();
        }
        buf.push('\n');
    }
}

/// Record candidate probability history for a trial (MaxEnt only, every trial).
fn record_cand_prob_history(
    buf: &mut Option<String>,
    tableau: &Tableau,
    ranking_values: &[f64],
    trial_number: usize,
) {
    if let Some(ref mut buf) = buf {
        use std::fmt::Write;
        let preds = maxent_predictions(tableau, ranking_values);
        write!(buf, "{}", trial_number).unwrap();
        for form_preds in &preds {
            for &p in form_preds {
                write!(buf, "\t{}", p).unwrap();
            }
        }
        buf.push('\n');
    }
}

/// Run stochastic test grammar evaluation (StochasticOT mode).
///
/// Uses noise from the last schedule stage. Returns predicted probabilities.
fn test_grammar_stochastic(
    tableau: &Tableau,
    ranking_values: &[f64],
    is_faith: &[bool],
    schedule: &LearningSchedule,
    test_trials: usize,
    rng: &mut Rng,
) -> Vec<Vec<f64>> {
    let last_stage = schedule.stages.last();
    let noise_mark = last_stage.map(|s| s.noise_mark).unwrap_or(2.0);
    let noise_faith = last_stage.map(|s| s.noise_faith).unwrap_or(2.0);

    let mut counts: Vec<Vec<usize>> = tableau.forms.iter()
        .map(|f| vec![0usize; f.candidates.len()])
        .collect();
    for _ in 0..test_trials {
        for (fi, form) in tableau.forms.iter().enumerate() {
            let winner = ot_evaluate(
                &form.candidates,
                ranking_values,
                is_faith,
                noise_mark,
                noise_faith,
                rng,
            );
            counts[fi][winner] += 1;
        }
    }
    counts.iter().map(|form_counts| {
        let total: usize = form_counts.iter().sum();
        if total == 0 {
            vec![0.0; form_counts.len()]
        } else {
            form_counts.iter().map(|&c| c as f64 / total as f64).collect()
        }
    }).collect()
}

/// Compute the squared error between observed and predicted proportions.
///
/// Returns (error_term, total_rivals).
fn compute_error_term(tableau: &Tableau, test_probs: &[Vec<f64>]) -> (f64, usize) {
    let mut error_term = 0.0f64;
    let mut total_rivals: usize = 0;
    for (fi, form) in tableau.forms.iter().enumerate() {
        let total_freq: f64 = form.candidates.iter().map(|c| c.frequency as f64).sum();
        total_rivals += form.candidates.len();
        for (ci, cand) in form.candidates.iter().enumerate() {
            let obs_prop = if total_freq > 0.0 {
                cand.frequency as f64 / total_freq
            } else {
                0.0
            };
            let gen_prop = test_probs[fi][ci];
            error_term += (obs_prop - gen_prop).powi(2);
        }
    }
    (error_term, total_rivals)
}

/// Compute log likelihood: Σ frequency × log(predicted_proportion).
fn compute_log_likelihood(tableau: &Tableau, test_probs: &[Vec<f64>]) -> f64 {
    let mut log_likelihood = 0.0f64;
    for (fi, form) in tableau.forms.iter().enumerate() {
        for (ci, cand) in form.candidates.iter().enumerate() {
            if cand.frequency > 0 {
                let pred = test_probs[fi][ci];
                if pred > 0.0 {
                    log_likelihood += cand.frequency as f64 * pred.ln();
                }
            }
        }
    }
    log_likelihood
}

impl Tableau {
    /// Run the Gradual Learning Algorithm using the default 4-stage schedule.
    ///
    /// This is a convenience wrapper around `run_gla_with_schedule`.
    /// Reproduces VB6 boersma.frm:GLACore and related subroutines.
    pub fn run_gla(&self, opts: &crate::GlaOptions) -> GlaResult {
        let schedule = LearningSchedule::default_4stage(opts.cycles, opts.initial_plasticity, opts.final_plasticity);
        let apriori = parse_apriori_from_opts(self, opts)
            .expect("Invalid a priori rankings text in GlaOptions");
        self.run_gla_with_schedule(&schedule, &apriori, opts)
    }

    /// Run GLA `run_count` times and format results as `CollateRuns.txt` content.
    ///
    /// Reproduces VB6 `boersma.frm:MultipleRuns`. Each run appends:
    ///   - **G records**: `G\t{run}\t{constraint_abbrev}\t{ranking_value}` (one per constraint)
    ///   - **O records**: `O\t{run}\t{form_idx}\t{input}\t{rival_idx}\t{rival_form}\t{freq}\t{pct_gen}`
    ///     (one per non-first candidate per form; VB6 skips candidate 0, the winner)
    pub fn format_collate_runs_output(
        &self,
        run_count: usize,
        schedule: &LearningSchedule,
        apriori: &[Vec<bool>],
        opts: &crate::GlaOptions,
    ) -> String {
        let run_opts = collate_run_opts(opts);
        let mut out = String::new();
        for run_idx in 1..=run_count {
            let result = self.run_gla_with_schedule(schedule, apriori, &run_opts);
            append_collate_run(self, run_idx, &result, &mut out);
        }
        out
    }

    /// Run the Gradual Learning Algorithm with an explicit learning schedule.
    ///
    /// Supports separate plasticity for Markedness vs Faithfulness constraints
    /// (PlastMark / PlastFaith per stage), matching VB6 behavior.
    ///
    /// `apriori`: parsed a priori rankings table (empty = no a priori constraints).
    /// Reproduces VB6 boersma.frm:AdjustAPrioriRankings_Up / _Down.
    pub fn run_gla_with_schedule(
        &self,
        schedule: &LearningSchedule,
        apriori: &[Vec<bool>],
        opts: &crate::GlaOptions,
    ) -> GlaResult {
        let maxent_mode = opts.maxent_mode;
        let nc = self.constraints.len();

        // Precomputed faithfulness flags for PlastMark vs PlastFaith distinction
        let is_faith: Vec<bool> = self.constraints.iter().map(|c| c.is_faithfulness()).collect();

        // ── Build training pool (weighted by frequency) ──────────────────────
        let mut training_pool: Vec<(usize, usize)> = Vec::new();
        for (fi, form) in self.forms.iter().enumerate() {
            for (ci, cand) in form.candidates.iter().enumerate() {
                for _ in 0..cand.frequency {
                    training_pool.push((fi, ci));
                }
            }
        }
        let pool_size = training_pool.len();

        // Round up schedule for exact proportions
        let schedule = if opts.exact_proportions && pool_size > 0 {
            let mut s = schedule.clone();
            s.round_stages_to_multiple(pool_size);
            s
        } else {
            schedule.clone()
        };

        // ── Initialize ranking values / weights ──────────────────────────────
        let initial_value = if maxent_mode { 0.0 } else { 100.0 };
        let mut ranking_values = vec![initial_value; nc];

        // Apply initial a priori enforcement
        if !apriori.is_empty() && !maxent_mode {
            adjust_apriori_up(&mut ranking_values, apriori, opts.apriori_gap);
            adjust_apriori_down(&mut ranking_values, apriori, opts.apriori_gap);
        }

        // ── History buffers ──────────────────────────────────────────────────
        let mut bufs = init_history_buffers(self, maxent_mode, nc, opts);
        let mut trial_number: usize = 0;

        let mut rng = Rng::new(GaussianMode::Standard);
        let learning_start = chrono::Utc::now();

        let mode_name = if maxent_mode { "MaxEnt" } else { "StochasticOT" };
        crate::ot_log!("Starting GLA ({}) with {} constraints, {} training exemplars, {} cycles",
            mode_name, nc, pool_size, schedule.total_cycles());

        // Start at pool_size so the first use triggers a shuffle
        let mut pool_cursor: usize = pool_size;

        // ── Main learning loop ────────────────────────────────────────────────
        if pool_size > 0 {
            for (stage_idx, stage) in schedule.stages.iter().enumerate() {
                for _ in 0..stage.trials {
                    trial_number += 1;

                    let (sel_form, sel_cand) = select_training_exemplar(
                        &mut training_pool, pool_size, &mut pool_cursor,
                        opts.exact_proportions, &mut rng,
                    );

                    // Generate a form using the current grammar
                    let generated = if maxent_mode {
                        maxent_sample(&self.forms[sel_form].candidates, &ranking_values, &mut rng)
                    } else {
                        ot_evaluate(
                            &self.forms[sel_form].candidates,
                            &ranking_values,
                            &is_faith,
                            stage.noise_mark,
                            stage.noise_faith,
                            &mut rng,
                        )
                    };

                    // Update weights and record histories on mismatch
                    if generated != sel_cand {
                        update_ranking_values(
                            &mut ranking_values,
                            &self.forms[sel_form].candidates[generated],
                            &self.forms[sel_form].candidates[sel_cand],
                            &is_faith,
                            stage.plast_mark,
                            stage.plast_faith,
                            maxent_mode,
                            opts.gaussian_prior,
                            opts.sigma,
                            opts.negative_weights_ok,
                            opts.magri_update_rule,
                            nc,
                        );
                        record_mismatch_histories(
                            &mut bufs, self, &ranking_values, trial_number,
                            sel_form, generated, sel_cand, &is_faith,
                            stage.plast_mark, stage.plast_faith, maxent_mode, nc,
                        );
                    }

                    record_cand_prob_history(
                        &mut bufs.cand_prob, self, &ranking_values, trial_number,
                    );
                }
                crate::ot_log!("GLA stage {}/{} complete (plast_mark={:.4}, plast_faith={:.4})",
                    stage_idx + 1, schedule.stages.len(), stage.plast_mark, stage.plast_faith);
            }
        }

        let learning_time_secs = {
            let elapsed = chrono::Utc::now() - learning_start;
            elapsed.num_milliseconds() as f64 / 1000.0
        };

        // ── Test grammar ──────────────────────────────────────────────────────
        let test_probs = if maxent_mode {
            maxent_predictions(self, &ranking_values)
        } else {
            test_grammar_stochastic(
                self, &ranking_values, &is_faith, &schedule, opts.test_trials, &mut rng,
            )
        };

        let (error_term, total_rivals) = compute_error_term(self, &test_probs);
        let log_likelihood = compute_log_likelihood(self, &test_probs);

        crate::ot_log!("GLA ({}) DONE: log_likelihood = {:.6}", mode_name, log_likelihood);

        let active_constraints = compute_active_constraints(self, &ranking_values);

        GlaResult {
            ranking_values,
            test_probs,
            log_likelihood,
            maxent_mode,
            schedule_description: schedule.format_description(),
            test_trials: opts.test_trials,
            gaussian_prior: opts.gaussian_prior,
            sigma: opts.sigma,
            magri_update_rule: opts.magri_update_rule,
            negative_weights_ok: opts.negative_weights_ok,
            error_term,
            total_rivals,
            learning_time_secs,
            apriori: apriori.to_vec(),
            apriori_gap: opts.apriori_gap,
            history: bufs.history,
            full_history: bufs.full_history,
            candidate_prob_history: bufs.cand_prob,
            active_constraints,
        }
    }
}

// ─── Chunked Runner ────────────────────────────────────────────────────────────
//
// Stateful GLA runner for chunked execution. Holds all algorithm state between
// `run_chunk` calls so the JS main thread can yield for UI updates.

/// Trait documenting the chunked runner contract.
/// Not exported via wasm_bindgen, but mirrored by a TypeScript interface.
/// Implemented by each algorithm's runner struct for consistency.
#[allow(dead_code)]
pub(crate) trait ChunkedRunner {
    /// Advance up to `max_work` units of work. Returns true when done.
    fn run_chunk(&mut self, max_work: usize) -> bool;
    /// Progress as [completed, total].
    fn progress(&self) -> [f64; 2];
}

/// Execution phase of the GLA runner.
enum GlaPhase {
    /// Learning phase: processing training trials.
    Learning,
    /// Testing phase: evaluating grammar with test trials (StochasticOT only).
    Testing,
    /// All computation complete.
    Done,
}

/// Chunked GLA runner for interactive progress reporting.
///
/// Holds all algorithm state so learning can be split across multiple
/// `run_chunk` calls, yielding control to the browser between chunks.
#[wasm_bindgen]
pub struct GlaRunner {
    // ── Immutable config ────────────────────────────────────────────────────
    tableau: Tableau,
    schedule: LearningSchedule,
    apriori: Vec<Vec<bool>>,
    maxent_mode: bool,
    test_trials: usize,
    negative_weights_ok: bool,
    gaussian_prior: bool,
    sigma: f64,
    magri_update_rule: bool,
    exact_proportions: bool,
    apriori_gap: f64,
    is_faith: Vec<bool>,
    total_learning_trials: usize,

    // ── Mutable algorithm state ─────────────────────────────────────────────
    ranking_values: Vec<f64>,
    training_pool: Vec<(usize, usize)>,
    pool_cursor: usize,
    rng: Rng,
    phase: GlaPhase,
    /// Current position: (stage_index, trial_within_stage)
    current_stage: usize,
    current_trial_in_stage: usize,
    trials_completed: usize,
    trial_number: usize,
    learning_start: chrono::DateTime<chrono::Utc>,

    // ── History buffers ─────────────────────────────────────────────────────
    history_buf: Option<String>,
    full_history_buf: Option<String>,
    cand_prob_history_buf: Option<String>,

    // ── Testing state (StochasticOT) ────────────────────────────────────────
    test_counts: Vec<Vec<usize>>,
    test_trials_completed: usize,
    test_noise_mark: f64,
    test_noise_faith: f64,

    // ── Final result (populated during Done phase) ──────────────────────────
    result: Option<GlaResult>,
}

#[wasm_bindgen]
impl GlaRunner {
    #[wasm_bindgen(constructor)]
    pub fn new(text: &str, opts: &crate::GlaOptions) -> Result<GlaRunner, String> {
        let tableau = Tableau::parse(text)?;
        let schedule = crate::build_gla_schedule(opts)?;
        let apriori = crate::parse_gla_apriori(&tableau, opts)?;

        let maxent_mode = opts.maxent_mode;
        let nc = tableau.constraints.len();
        let is_faith: Vec<bool> = tableau.constraints.iter().map(|c| c.is_faithfulness()).collect();

        // Build training pool
        let mut training_pool: Vec<(usize, usize)> = Vec::new();
        for (fi, form) in tableau.forms.iter().enumerate() {
            for (ci, cand) in form.candidates.iter().enumerate() {
                for _ in 0..cand.frequency {
                    training_pool.push((fi, ci));
                }
            }
        }
        let pool_size = training_pool.len();

        // Round schedule for exact proportions
        let schedule = if opts.exact_proportions && pool_size > 0 {
            let mut s = schedule;
            s.round_stages_to_multiple(pool_size);
            s
        } else {
            schedule
        };

        let total_learning_trials = schedule.total_cycles();

        // Initialize ranking values
        let initial_value = if maxent_mode { 0.0 } else { 100.0 };
        let mut ranking_values = vec![initial_value; nc];

        // Apply initial a priori enforcement
        if !apriori.is_empty() && !maxent_mode {
            adjust_apriori_up(&mut ranking_values, &apriori, opts.apriori_gap);
            adjust_apriori_down(&mut ranking_values, &apriori, opts.apriori_gap);
        }

        // History buffers
        let history_buf = if opts.generate_history {
            let mut header = String::from("Trial");
            for c in &tableau.constraints {
                header.push('\t');
                header.push_str(&c.abbrev());
            }
            header.push('\n');
            Some(header)
        } else {
            None
        };

        let full_history_buf = if opts.generate_full_history {
            use std::fmt::Write;
            let mut header = String::from("Trial #\tInput\tGenerated\tHeard");
            for c in &tableau.constraints {
                write!(header, "\t{}", c.abbrev()).unwrap();
            }
            header.push('\n');
            header.push_str("(Initial)\t\t\t");
            for _ in 0..nc {
                write!(header, "\t{initial_value:.4}").unwrap();
            }
            header.push('\n');
            Some(header)
        } else {
            None
        };

        let cand_prob_history_buf = if maxent_mode && opts.generate_candidate_prob_history {
            use std::fmt::Write;
            let mut header = String::from("Trial #");
            for form in &tableau.forms {
                for cand in &form.candidates {
                    write!(header, "\t/{}/ -> {}", form.input, cand.form).unwrap();
                }
            }
            header.push('\n');
            let zero_weights = vec![0.0; nc];
            let initial_preds = maxent_predictions(&tableau, &zero_weights);
            header.push_str("(initial)");
            for form_preds in &initial_preds {
                for &p in form_preds {
                    write!(header, "\t{}", p).unwrap();
                }
            }
            header.push('\n');
            Some(header)
        } else {
            None
        };

        Ok(GlaRunner {
            schedule,
            apriori,
            maxent_mode,
            test_trials: if maxent_mode { 0 } else { opts.test_trials },
            negative_weights_ok: opts.negative_weights_ok,
            gaussian_prior: opts.gaussian_prior,
            sigma: opts.sigma,
            magri_update_rule: opts.magri_update_rule,
            exact_proportions: opts.exact_proportions,
            apriori_gap: opts.apriori_gap,
            is_faith,
            total_learning_trials,
            ranking_values,
            pool_cursor: pool_size, // triggers shuffle on first use
            rng: Rng::new(GaussianMode::Standard),
            phase: GlaPhase::Learning,
            current_stage: 0,
            current_trial_in_stage: 0,
            trials_completed: 0,
            trial_number: 0,
            learning_start: chrono::Utc::now(),
            history_buf,
            full_history_buf,
            cand_prob_history_buf,
            test_counts: tableau.forms.iter().map(|f| vec![0usize; f.candidates.len()]).collect(),
            test_trials_completed: 0,
            test_noise_mark: 2.0,
            test_noise_faith: 2.0,
            training_pool,
            tableau,
            result: None,
        })
    }

    /// Advance up to `max_trials` trials. Returns true when all computation
    /// (learning + testing) is complete.
    pub fn run_chunk(&mut self, max_trials: usize) -> bool {
        match self.phase {
            GlaPhase::Learning => self.run_learning_chunk(max_trials),
            GlaPhase::Testing => self.run_testing_chunk(max_trials),
            GlaPhase::Done => true,
        }
    }

    /// Progress as [completed, total]. During learning, counts learning trials.
    /// During testing, adds test trials to the learning total.
    pub fn progress(&self) -> Vec<f64> {
        let total = (self.total_learning_trials + self.test_trials) as f64;
        let completed = match self.phase {
            GlaPhase::Learning => self.trials_completed as f64,
            GlaPhase::Testing => {
                (self.total_learning_trials + self.test_trials_completed) as f64
            }
            GlaPhase::Done => total,
        };
        vec![completed, total]
    }

    /// Extract the final GlaResult. Only valid after `run_chunk` returns true.
    /// Panics if called before completion.
    pub fn take_result(&mut self) -> GlaResult {
        self.result.take().expect("GlaRunner: take_result called before completion")
    }
}

impl GlaRunner {
    fn run_learning_chunk(&mut self, max_trials: usize) -> bool {
        let pool_size = self.training_pool.len();
        if pool_size == 0 {
            self.finish_learning();
            return self.maybe_start_testing();
        }

        let nc = self.ranking_values.len();
        let mut work_done = 0;

        while work_done < max_trials {
            // Check if current stage is exhausted
            if self.current_stage >= self.schedule.stages.len() {
                self.finish_learning();
                return self.maybe_start_testing();
            }

            let stage = &self.schedule.stages[self.current_stage];
            if self.current_trial_in_stage >= stage.trials {
                crate::ot_log!("GLA stage {}/{} complete (plast_mark={:.4}, plast_faith={:.4})",
                    self.current_stage + 1, self.schedule.stages.len(),
                    stage.plast_mark, stage.plast_faith);
                self.current_stage += 1;
                self.current_trial_in_stage = 0;
                continue;
            }

            // ── Single trial ────────────────────────────────────────────────
            self.trial_number += 1;
            self.current_trial_in_stage += 1;
            self.trials_completed += 1;
            work_done += 1;

            // Select observed exemplar
            let (sel_form, sel_cand) = if self.exact_proportions {
                if self.pool_cursor >= pool_size {
                    for i in (1..pool_size).rev() {
                        let j = (self.rng.uniform() * (i + 1) as f64) as usize;
                        self.training_pool.swap(i, j);
                    }
                    self.pool_cursor = 0;
                }
                let pair = self.training_pool[self.pool_cursor];
                self.pool_cursor += 1;
                pair
            } else {
                let r = self.rng.uniform();
                let idx = ((r * pool_size as f64) as usize).min(pool_size - 1);
                self.training_pool[idx]
            };

            // Generate a form
            let generated = if self.maxent_mode {
                maxent_sample(
                    &self.tableau.forms[sel_form].candidates,
                    &self.ranking_values,
                    &mut self.rng,
                )
            } else {
                ot_evaluate(
                    &self.tableau.forms[sel_form].candidates,
                    &self.ranking_values,
                    &self.is_faith,
                    stage.noise_mark,
                    stage.noise_faith,
                    &mut self.rng,
                )
            };

            // Update weights on mismatch
            if generated != sel_cand {
                let gen_cand = &self.tableau.forms[sel_form].candidates[generated];
                let obs_cand = &self.tableau.forms[sel_form].candidates[sel_cand];

                if self.maxent_mode {
                    let sigma_sq = self.sigma * self.sigma;
                    for (c, rv) in self.ranking_values.iter_mut().enumerate() {
                        let plast = if self.is_faith[c] { stage.plast_faith } else { stage.plast_mark };
                        let gen_v = gen_cand.violations[c] as f64;
                        let obs_v = obs_cand.violations[c] as f64;
                        let likelihood_change = plast * (gen_v - obs_v);
                        let prior_change = if self.gaussian_prior {
                            plast * *rv / sigma_sq / 2.0
                        } else {
                            0.0
                        };
                        *rv -= likelihood_change + prior_change;
                        if !self.negative_weights_ok && *rv < 0.0 {
                            *rv = 0.0;
                        }
                    }
                } else {
                    let promo_scale = if self.magri_update_rule {
                        magri_promotion_amount(gen_cand, obs_cand, nc)
                    } else {
                        1.0
                    };
                    for (c, rv) in self.ranking_values.iter_mut().enumerate() {
                        let plast = if self.is_faith[c] { stage.plast_faith } else { stage.plast_mark };
                        let gen_v = gen_cand.violations[c];
                        let obs_v = obs_cand.violations[c];
                        if gen_v > obs_v {
                            *rv += plast * promo_scale;
                        } else if gen_v < obs_v {
                            *rv -= plast;
                        }
                    }
                }

                // Record history
                if let Some(ref mut buf) = self.history_buf {
                    use std::fmt::Write;
                    write!(buf, "{}", self.trial_number).unwrap();
                    for rv in self.ranking_values.iter() {
                        write!(buf, "\t{rv:.4}").unwrap();
                    }
                    buf.push('\n');
                }

                // Record full history
                if let Some(ref mut buf) = self.full_history_buf {
                    use std::fmt::Write;
                    if !self.maxent_mode {
                        let gen_cand = &self.tableau.forms[sel_form].candidates[generated];
                        let obs_cand = &self.tableau.forms[sel_form].candidates[sel_cand];
                        for c in 0..nc {
                            let gen_v = gen_cand.violations[c];
                            let obs_v = obs_cand.violations[c];
                            let plast = if self.is_faith[c] { stage.plast_faith } else { stage.plast_mark };
                            if gen_v > obs_v {
                                write!(buf, "\t{plast}\t{:.4}", self.ranking_values[c]).unwrap();
                            } else if gen_v < obs_v {
                                write!(buf, "\t-{plast}\t{:.4}", self.ranking_values[c]).unwrap();
                            } else {
                                buf.push_str("\t\t");
                            }
                        }
                    }
                    write!(buf, "{}\t{}\t{}\t{}",
                        self.trial_number,
                        self.tableau.forms[sel_form].input,
                        self.tableau.forms[sel_form].candidates[generated].form,
                        self.tableau.forms[sel_form].candidates[sel_cand].form,
                    ).unwrap();
                    for rv in self.ranking_values.iter() {
                        write!(buf, "\t{rv:.4}").unwrap();
                    }
                    buf.push('\n');
                }
            }

            // Candidate probability history (every trial, not just mismatches)
            if let Some(ref mut buf) = self.cand_prob_history_buf {
                use std::fmt::Write;
                let preds = maxent_predictions(&self.tableau, &self.ranking_values);
                write!(buf, "{}", self.trial_number).unwrap();
                for form_preds in &preds {
                    for &p in form_preds {
                        write!(buf, "\t{}", p).unwrap();
                    }
                }
                buf.push('\n');
            }
        }

        false // not done yet
    }

    fn finish_learning(&mut self) {
        // Store noise values from last stage for testing
        if let Some(last) = self.schedule.stages.last() {
            self.test_noise_mark = last.noise_mark;
            self.test_noise_faith = last.noise_faith;
        }
    }

    fn maybe_start_testing(&mut self) -> bool {
        if self.maxent_mode || self.test_trials == 0 {
            // MaxEnt: exact predictions, no stochastic testing needed
            self.finalize();
            true
        } else {
            self.phase = GlaPhase::Testing;
            false
        }
    }

    fn run_testing_chunk(&mut self, max_trials: usize) -> bool {
        let mut work_done = 0;
        while work_done < max_trials && self.test_trials_completed < self.test_trials {
            for (fi, form) in self.tableau.forms.iter().enumerate() {
                let winner = ot_evaluate(
                    &form.candidates,
                    &self.ranking_values,
                    &self.is_faith,
                    self.test_noise_mark,
                    self.test_noise_faith,
                    &mut self.rng,
                );
                self.test_counts[fi][winner] += 1;
            }
            self.test_trials_completed += 1;
            work_done += 1;
        }

        if self.test_trials_completed >= self.test_trials {
            self.finalize();
            true
        } else {
            false
        }
    }

    fn finalize(&mut self) {
        let learning_time_secs = {
            let elapsed = chrono::Utc::now() - self.learning_start;
            elapsed.num_milliseconds() as f64 / 1000.0
        };

        // Compute test probabilities
        let test_probs = if self.maxent_mode {
            maxent_predictions(&self.tableau, &self.ranking_values)
        } else {
            self.test_counts.iter().map(|form_counts| {
                let total: usize = form_counts.iter().sum();
                if total == 0 {
                    vec![0.0; form_counts.len()]
                } else {
                    form_counts.iter().map(|&c| c as f64 / total as f64).collect()
                }
            }).collect()
        };

        // Error term
        let mut error_term = 0.0f64;
        let mut total_rivals: usize = 0;
        for (fi, form) in self.tableau.forms.iter().enumerate() {
            let total_freq: f64 = form.candidates.iter().map(|c| c.frequency as f64).sum();
            total_rivals += form.candidates.len();
            for (ci, cand) in form.candidates.iter().enumerate() {
                let obs_prop = if total_freq > 0.0 { cand.frequency as f64 / total_freq } else { 0.0 };
                let gen_prop = test_probs[fi][ci];
                error_term += (obs_prop - gen_prop).powi(2);
            }
        }

        // Log likelihood
        let mut log_likelihood = 0.0f64;
        for (fi, form) in self.tableau.forms.iter().enumerate() {
            for (ci, cand) in form.candidates.iter().enumerate() {
                if cand.frequency > 0 {
                    let pred = test_probs[fi][ci];
                    if pred > 0.0 {
                        log_likelihood += cand.frequency as f64 * pred.ln();
                    }
                }
            }
        }

        self.phase = GlaPhase::Done;
        self.result = Some(GlaResult {
            ranking_values: self.ranking_values.clone(),
            test_probs,
            log_likelihood,
            maxent_mode: self.maxent_mode,
            schedule_description: self.schedule.format_description(),
            test_trials: self.test_trials,
            gaussian_prior: self.gaussian_prior,
            sigma: self.sigma,
            magri_update_rule: self.magri_update_rule,
            negative_weights_ok: self.negative_weights_ok,
            error_term,
            total_rivals,
            learning_time_secs,
            apriori: self.apriori.clone(),
            apriori_gap: self.apriori_gap,
            history: self.history_buf.take(),
            full_history: self.full_history_buf.take(),
            candidate_prob_history: self.cand_prob_history_buf.take(),
            active_constraints: compute_active_constraints(&self.tableau, &self.ranking_values),
        });
    }
}

impl ChunkedRunner for GlaRunner {
    fn run_chunk(&mut self, max_work: usize) -> bool {
        GlaRunner::run_chunk(self, max_work)
    }

    fn progress(&self) -> [f64; 2] {
        let p = GlaRunner::progress(self);
        [p[0], p[1]]
    }
}

// ─── GlaMultipleRunsRunner ────────────────────────────────────────────────────

/// Chunked runner for the "Run N times & Download" (CollateRuns) workflow.
///
/// Each `run_chunk(max_runs)` call completes up to `max_runs` full GLA runs,
/// yielding control to the browser between calls for progress updates.
/// Progress is reported as [runs_completed, run_count].
#[wasm_bindgen]
pub struct GlaMultipleRunsRunner {
    tableau: Tableau,
    schedule: LearningSchedule,
    apriori: Vec<Vec<bool>>,
    run_opts: crate::GlaOptions,
    run_count: usize,
    runs_completed: usize,
    output: String,
}

#[wasm_bindgen]
impl GlaMultipleRunsRunner {
    #[wasm_bindgen(constructor)]
    pub fn new(text: &str, run_count: u32, opts: &crate::GlaOptions) -> Result<GlaMultipleRunsRunner, String> {
        let tableau = Tableau::parse(text)?;
        let schedule = crate::build_gla_schedule(opts)?;
        let apriori = crate::parse_gla_apriori(&tableau, opts)?;
        let run_opts = collate_run_opts(opts);
        let run_count = run_count as usize;

        Ok(GlaMultipleRunsRunner {
            tableau,
            schedule,
            apriori,
            run_opts,
            run_count,
            runs_completed: 0,
            output: String::new(),
        })
    }

    /// Run up to `max_runs` complete GLA runs. Returns true when all runs are done.
    pub fn run_chunk(&mut self, max_runs: usize) -> bool {
        let end = (self.runs_completed + max_runs).min(self.run_count);
        for run_idx in (self.runs_completed + 1)..=end {
            let result = self.tableau.run_gla_with_schedule(&self.schedule, &self.apriori, &self.run_opts);
            append_collate_run(&self.tableau, run_idx, &result, &mut self.output);
        }
        self.runs_completed = end;
        self.runs_completed >= self.run_count
    }

    /// Progress as [completed_runs, total_runs].
    pub fn progress(&self) -> Vec<f64> {
        vec![self.runs_completed as f64, self.run_count as f64]
    }

    /// Extract the final collated output. Only valid after `run_chunk` returns true.
    pub fn take_result(&mut self) -> String {
        std::mem::take(&mut self.output)
    }
}

impl ChunkedRunner for GlaMultipleRunsRunner {
    fn run_chunk(&mut self, max_work: usize) -> bool {
        GlaMultipleRunsRunner::run_chunk(self, max_work)
    }

    fn progress(&self) -> [f64; 2] {
        let p = GlaMultipleRunsRunner::progress(self);
        [p[0], p[1]]
    }
}

// ─── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use crate::schedule::LearningSchedule;
    use crate::tableau::Tableau;

    fn load_tiny() -> String {
        std::fs::read_to_string("../examples/TinyIllustrativeFile.txt")
            .expect("Failed to load examples/TinyIllustrativeFile.txt")
    }

    #[test]
    fn test_gla_sot_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(), "ranking value {} should be finite", c);
        }
        assert!(result.log_likelihood().is_finite());
    }

    #[test]
    fn test_gla_maxent_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(), "weight {} should be finite", c);
            assert!(result.get_ranking_value(c) >= 0.0, "weight {} should be >= 0 by default", c);
        }
        assert!(result.log_likelihood().is_finite());
    }

    #[test]
    fn test_gla_maxent_probs_sum_to_one() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, ..Default::default()
        });

        for fi in 0..tableau.form_count() {
            let form = tableau.get_form(fi).unwrap();
            let total: f64 = (0..form.candidate_count())
                .map(|ci| result.get_test_prob(fi, ci))
                .sum();
            assert!((total - 1.0).abs() < 1e-9, "MaxEnt probs for form {} should sum to 1, got {}", fi, total);
        }
    }

    #[test]
    fn test_gla_sot_initial_values_100() {
        // StochasticOT should start at 100 (Boersma's canonical value)
        // After 0 cycles, values should still be 100
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 0, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 100,
            ..Default::default()
        });
        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert_eq!(result.get_ranking_value(c), 100.0,
                "StochasticOT should start at 100, got {}", result.get_ranking_value(c));
        }
    }

    #[test]
    fn test_gla_maxent_initial_values_zero() {
        // MaxEnt should start at 0
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 0, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, ..Default::default()
        });
        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert_eq!(result.get_ranking_value(c), 0.0,
                "MaxEnt should start at 0, got {}", result.get_ranking_value(c));
        }
    }

    #[test]
    fn test_gla_maxent_gaussian_prior_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, gaussian_prior: true, ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(), "weight {} should be finite with prior", c);
        }
        assert!(result.log_likelihood().is_finite());
        assert!(result.gaussian_prior(), "gaussian_prior() should be true");
        assert_eq!(result.sigma(), 1.0, "sigma() should be 1.0");
    }

    #[test]
    fn test_gla_sot_magri_update_rule_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            magri_update_rule: true, ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(),
                "ranking value {} should be finite with Magri rule", c);
        }
        assert!(result.log_likelihood().is_finite());

        let output = result.format_output(&tableau, "test.txt");
        assert!(output.contains("The Magri update rule was employed."),
            "output should mention Magri update rule");
    }

    #[test]
    fn test_gla_custom_schedule_runs() {
        // A custom schedule with separate M/F plasticity should run without errors
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule_text = "Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n\
                             250\t2\t0.5\t2\t2\n\
                             250\t0.2\t0.05\t2\t2\n";
        let schedule = LearningSchedule::parse(schedule_text).unwrap();
        let result = tableau.run_gla_with_schedule(
            &schedule,
            &[],
            &crate::GlaOptions { test_trials: 200, ..Default::default() },
        );

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite());
        }
    }

    #[test]
    fn test_gla_multiple_runs_format() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule = LearningSchedule::default_4stage(500, 2.0, 0.001);
        let output = tableau.format_collate_runs_output(
            3,
            &schedule,
            &[],
            &crate::GlaOptions { test_trials: 200, ..Default::default() },
        );

        let nc = tableau.constraint_count();
        let lines: Vec<&str> = output.lines().collect();

        // Count G records
        let g_lines: Vec<&&str> = lines.iter().filter(|l| l.starts_with("G\t")).collect();
        assert_eq!(g_lines.len(), 3 * nc, "should have run_count * nc G records");

        // Count O records (one per rival per form across all runs)
        let o_lines: Vec<&&str> = lines.iter().filter(|l| l.starts_with("O\t")).collect();
        assert!(!o_lines.is_empty(), "should have O records");

        // G records have 4 tab-separated fields
        for line in &g_lines {
            let fields: Vec<&str> = line.split('\t').collect();
            assert_eq!(fields.len(), 4, "G record should have 4 fields: {}", line);
        }

        // O records have 8 tab-separated fields
        for line in &o_lines {
            let fields: Vec<&str> = line.split('\t').collect();
            assert_eq!(fields.len(), 8, "O record should have 8 fields: {}", line);
        }

        // Run indices go from 1 to 3
        for run_idx in 1..=3 {
            let run_str = run_idx.to_string();
            let has_run = lines
                .iter()
                .any(|l| l.starts_with("G\t") && l.split('\t').nth(1) == Some(&run_str));
            assert!(has_run, "should have G records for run {}", run_idx);
        }
    }

    #[test]
    fn test_gla_sot_pairwise_probabilities_structure() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        });

        let table = result.format_pairwise_probabilities(&tableau, 5);
        assert!(table.contains("5. Ranking Value to Ranking Probability Conversion"));
        assert!(table.contains("outranking the constraint in the column headings"));

        // Should contain constraint abbreviations
        for c in &tableau.constraints {
            assert!(table.contains(&c.abbrev()), "table should contain abbrev {}", c.abbrev());
        }

        // Should contain probability values (at least "0.5" somewhere)
        assert!(table.contains("0.5") || table.contains("0.9") || table.contains(">.999999"),
            "table should contain probability values");
    }

    #[test]
    fn test_gla_sot_pairwise_probabilities_json() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        });

        let json_str = result.pairwise_probabilities_json(&tableau);
        let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();

        let headers = parsed["headers"].as_array().unwrap();
        assert_eq!(headers.len(), tableau.constraints.len());

        let matrix = parsed["matrix"].as_array().unwrap();
        // Matrix has n-1 rows (all but last constraint)
        assert_eq!(matrix.len(), tableau.constraints.len() - 1);
        // Each row has n-1 columns (all but first constraint)
        for row in matrix {
            assert_eq!(row.as_array().unwrap().len(), tableau.constraints.len() - 1);
        }

        // Upper triangle should have probability values, lower triangle empty strings
        let first_row = matrix[0].as_array().unwrap();
        assert!(!first_row[0].as_str().unwrap().is_empty(), "first cell should be a probability");

        // Every header should be a constraint abbreviation
        let abbrevs: Vec<String> = (0..tableau.constraints.len())
            .map(|i| tableau.constraints[i].abbrev())
            .collect();
        for h in headers {
            assert!(abbrevs.contains(&h.as_str().unwrap().to_string()));
        }
    }

    #[test]
    fn test_gla_sot_format_output_includes_pairwise() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        });

        let output = result.format_output(&tableau, "test.txt");
        assert!(output.contains("5. Ranking Value to Ranking Probability Conversion"),
            "Stochastic OT output should include pairwise probability section");
    }

    #[test]
    fn test_gla_sot_format_output_includes_testing_details() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        });

        let output = result.format_output(&tableau, "test.txt");
        assert!(output.contains("3. Testing the Grammar: Details"),
            "Stochastic OT output should include testing details section");
        assert!(output.contains("The grammar was tested for 200 cycles."),
            "should show test cycle count");
        assert!(output.contains("Average error per candidate:"),
            "should show average error");
        assert!(output.contains("Learning time:"),
            "should show learning time");
        assert!(output.contains("Negative weights were not permitted."),
            "should show negative weights status");
    }

    #[test]
    fn test_gla_maxent_format_output_excludes_testing_details() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, ..Default::default()
        });

        let output = result.format_output(&tableau, "test.txt");
        assert!(!output.contains("Testing the Grammar: Details"),
            "MaxEnt output should NOT include testing details section");
    }

    #[test]
    fn test_gla_maxent_format_output_excludes_pairwise() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, ..Default::default()
        });

        let output = result.format_output(&tableau, "test.txt");
        assert!(!output.contains("Ranking Value to Ranking Probability Conversion"),
            "MaxEnt output should NOT include pairwise probability section");
    }

    #[test]
    fn test_gla_custom_schedule_output_contains_stage_info() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule_text = "Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n\
                             250\t2\t0.5\t2\t2\n\
                             250\t0.2\t0.05\t2\t2\n";
        let schedule = LearningSchedule::parse(schedule_text).unwrap();
        let result = tableau.run_gla_with_schedule(
            &schedule,
            &[],
            &crate::GlaOptions { test_trials: 200, ..Default::default() },
        );
        let output = result.format_output(&tableau, "test.txt");

        // Custom schedule should mention stages in the output
        assert!(output.contains("Custom learning schedule"), "output should mention custom schedule");
        assert!(output.contains("PlastMark"), "output should show PlastMark column");
    }

    #[test]
    fn test_gla_sot_exact_proportions_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            exact_proportions: true, ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(),
                "ranking value {} should be finite with exact proportions", c);
        }
        assert!(result.log_likelihood().is_finite());
    }

    #[test]
    fn test_gla_maxent_exact_proportions_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, exact_proportions: true, ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(),
                "weight {} should be finite with exact proportions", c);
        }
        assert!(result.log_likelihood().is_finite());
    }

    #[test]
    fn test_gla_sot_history_generation() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule = LearningSchedule::default_4stage(500, 2.0, 0.001);
        let result = tableau.run_gla_with_schedule(
            &schedule,
            &[],
            &crate::GlaOptions { test_trials: 200, generate_history: true, ..Default::default() },
        );

        let history = result.history().expect("history should be Some when generate_history=true");
        let lines: Vec<&str> = history.lines().collect();

        // Header should start with "Trial" and contain constraint abbreviations
        assert!(lines[0].starts_with("Trial\t"), "header should start with Trial");
        let nc = tableau.constraint_count();
        let header_cols: Vec<&str> = lines[0].split('\t').collect();
        assert_eq!(header_cols.len(), nc + 1, "header should have Trial + {} constraint columns", nc);

        // Should have at least some data rows (mismatches)
        assert!(lines.len() > 1, "history should have data rows");

        // Each data row should have trial number + nc values
        for line in &lines[1..] {
            let cols: Vec<&str> = line.split('\t').collect();
            assert_eq!(cols.len(), nc + 1, "each row should have {} columns, got {}", nc + 1, cols.len());
            // Trial number should parse as integer
            cols[0].parse::<usize>().expect("trial number should be an integer");
        }
    }

    #[test]
    fn test_gla_history_none_when_disabled() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        });
        assert!(result.history().is_none(), "history should be None when generate_history=false");
    }

    #[test]
    fn test_gla_full_history_generation() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule = LearningSchedule::default_4stage(500, 2.0, 0.001);
        let result = tableau.run_gla_with_schedule(
            &schedule,
            &[],
            &crate::GlaOptions { test_trials: 200, generate_full_history: true, ..Default::default() },
        );

        let fh = result.full_history().expect("full_history should be Some when enabled");
        let lines: Vec<&str> = fh.lines().collect();
        let nc = tableau.constraint_count();

        // Header: Trial # \t Input \t Generated \t Heard \t [constraints...]
        let header_cols: Vec<&str> = lines[0].split('\t').collect();
        assert_eq!(header_cols.len(), nc + 4, "header should have 4 leading + {} constraint columns", nc);
        assert_eq!(header_cols[0], "Trial #");
        assert_eq!(header_cols[1], "Input");
        assert_eq!(header_cols[2], "Generated");
        assert_eq!(header_cols[3], "Heard");

        // Initial row
        assert!(lines[1].starts_with("(Initial)\t\t\t"), "second line should be initial row");
        let init_cols: Vec<&str> = lines[1].split('\t').collect();
        assert_eq!(init_cols.len(), nc + 4, "initial row should have same column count as header");

        // Data rows (at least some mismatches should occur)
        // StochasticOT: per-constraint delta+value (2 cols each) written first by
        // RankingValueAdjustment, then trial/input/gen/heard + final values by GLACore.
        // Total columns = nc*2 (deltas) + 4 (trial info) + nc (final values) = nc*3 + 4.
        let expected_data_cols = nc * 3 + 4;
        assert!(lines.len() > 2, "should have data rows after header + initial");
        for line in &lines[2..] {
            let cols: Vec<&str> = line.split('\t').collect();
            assert_eq!(cols.len(), expected_data_cols,
                "each data row should have {} columns (nc*3+4), got {}", expected_data_cols, cols.len());
        }
    }

    #[test]
    fn test_gla_full_history_none_when_disabled() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        });
        assert!(result.full_history().is_none(), "full_history should be None when disabled");
    }

    #[test]
    fn test_gla_candidate_prob_history_maxent() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let total_cycles = 10;
        let schedule = LearningSchedule::default_4stage(total_cycles, 2.0, 0.001);
        let total_trials = schedule.total_cycles();
        let result = tableau.run_gla_with_schedule(
            &schedule,
            &[],
            &crate::GlaOptions {
                maxent_mode: true, test_trials: 0, generate_candidate_prob_history: true,
                ..Default::default()
            },
        );

        let cph = result.candidate_prob_history()
            .expect("candidate_prob_history should be Some when enabled in MaxEnt mode");
        let lines: Vec<&str> = cph.lines().collect();

        // Header: Trial # + one column per candidate across all forms
        assert!(lines[0].starts_with("Trial #\t"), "header should start with Trial #");
        let total_cands: usize = (0..tableau.form_count())
            .map(|fi| tableau.get_form(fi).unwrap().candidate_count())
            .sum();
        let header_cols: Vec<&str> = lines[0].split('\t').collect();
        assert_eq!(header_cols.len(), total_cands + 1,
            "header should have Trial # + {} candidate columns", total_cands);

        // Each header column (except first) should contain " -> "
        for col in &header_cols[1..] {
            assert!(col.contains(" -> "), "column '{}' should contain ' -> '", col);
        }

        // Initial row
        assert!(lines[1].starts_with("(initial)"), "second line should be initial row");

        // Data rows: one per trial (not just mismatches)
        let data_lines = lines.len() - 2; // minus header and initial
        assert_eq!(data_lines, total_trials,
            "should have one data row per trial ({}), got {}", total_trials, data_lines);

        // Each data row should have correct column count
        for line in &lines[2..] {
            let cols: Vec<&str> = line.split('\t').collect();
            assert_eq!(cols.len(), total_cands + 1);
            cols[0].parse::<usize>().expect("trial number should be an integer");
        }
    }

    #[test]
    fn test_gla_candidate_prob_history_none_when_disabled() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(&crate::GlaOptions {
            maxent_mode: true, cycles: 50, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, ..Default::default()
        });
        assert!(result.candidate_prob_history().is_none(),
            "candidate_prob_history should be None when not requested");
    }

    #[test]
    fn test_gla_candidate_prob_history_none_for_stochastic_ot() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule = LearningSchedule::default_4stage(10, 2.0, 0.001);
        // Even with generate_candidate_prob_history=true, SOT should not produce it
        let result = tableau.run_gla_with_schedule(
            &schedule,
            &[],
            &crate::GlaOptions {
                test_trials: 200, generate_candidate_prob_history: true,
                ..Default::default()
            },
        );
        assert!(result.candidate_prob_history().is_none(),
            "candidate_prob_history should be None in Stochastic OT mode");
    }

    #[test]
    fn test_gla_apriori_enforces_initial_gap() {
        // VB6 v2.7: a priori enforcement only happens during initialization, not
        // during learning. Verify the initial ranking values satisfy the gap.
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let nc = tableau.constraint_count();
        if nc < 2 {
            return; // Need at least 2 constraints
        }

        // table[0][1] = true means constraint 0 a priori dominates constraint 1
        let mut apriori = vec![vec![false; nc]; nc];
        apriori[0][1] = true;

        let gap = 20.0;
        // Use 0 cycles so we only get initialization, not learning
        let result = tableau.run_gla_with_schedule(
            &crate::schedule::LearningSchedule::default_4stage(0, 2.0, 0.001),
            &apriori,
            &crate::GlaOptions { test_trials: 0, apriori_gap: gap, ..Default::default() },
        );

        let rv0 = result.get_ranking_value(0);
        let rv1 = result.get_ranking_value(1);
        assert!(
            rv0 - rv1 >= gap - 0.001,
            "initial a priori gap not satisfied: rv[0]={rv0:.3}, rv[1]={rv1:.3}, required gap={gap}"
        );
    }

    #[test]
    fn test_gla_runner_completes() {
        let text = load_tiny();
        let opts = crate::GlaOptions {
            cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001, test_trials: 200,
            ..Default::default()
        };
        let mut runner = super::GlaRunner::new(&text, &opts).unwrap();

        // Run in small chunks until done
        let mut iterations = 0;
        while !runner.run_chunk(100) {
            iterations += 1;
            let p = runner.progress();
            assert!(p[0] <= p[1], "completed should not exceed total");
            assert!(iterations < 10_000, "should complete in bounded iterations");
        }
        assert!(iterations > 0, "should take multiple chunks");

        let result = runner.take_result();
        let nc = Tableau::parse(&text).unwrap().constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite());
        }
        assert!(result.log_likelihood().is_finite());
    }

    #[test]
    fn test_gla_runner_maxent_completes() {
        let text = load_tiny();
        let opts = crate::GlaOptions {
            maxent_mode: true, cycles: 500, initial_plasticity: 2.0, final_plasticity: 0.001,
            test_trials: 0, ..Default::default()
        };
        let mut runner = super::GlaRunner::new(&text, &opts).unwrap();

        while !runner.run_chunk(100) {}

        let result = runner.take_result();
        assert!(result.log_likelihood().is_finite());
        assert!(result.is_maxent_mode());
    }
}
