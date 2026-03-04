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
    /// Optional history of ranking values/weights recorded during learning.
    history: Option<String>,
    /// Optional full (annotated) history: trial, input, generated, heard, values.
    full_history: Option<String>,
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
            history: None,
            full_history: None,
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
        out.push_str("OTSoft 2.7, release date 2/1/2026\n\n\n");

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
        out.push_str("\n\n");

        // Section 1: Ranking Values / Weights Found (original constraint order)
        out.push_str(&format!("1. {} Found\n\n", value_label));
        for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
            out.push_str(&format!(
                "   {:<42}{:.3}\n",
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
            if self.maxent_mode {
                out.push_str(&format!(
                    "  {:<width$}  {:>9}  {:>9}\n",
                    "", "Input%", "Prob",
                    width = max_cand_width
                ));
            } else {
                out.push_str(&format!(
                    "  {:<width$}  {:>9}  {:>9}\n",
                    "", "Input%", "Gen%",
                    width = max_cand_width
                ));
            }

            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                let obs_pct = if total_freq > 0.0 {
                    cand.frequency as f64 / total_freq * 100.0
                } else {
                    0.0
                };
                let gen_pct = self.test_probs
                    .get(form_idx)
                    .and_then(|f| f.get(cand_idx))
                    .copied()
                    .unwrap_or(0.0)
                    * 100.0;
                let marker = if cand.frequency > 0 { ">" } else { " " };
                out.push_str(&format!(
                    "  {}{:<width$}  {:>8.1}%  {:>8.1}%\n",
                    marker, cand.form, obs_pct, gen_pct,
                    width = max_cand_width
                ));
            }
            out.push('\n');
        }

        // Section 3: Sorted ranking values / weights
        out.push_str(&format!("3. {} (sorted)\n\n", value_label));
        let sorted = crate::tableau::sorted_indices_descending(&self.ranking_values);
        for &c_idx in &sorted {
            out.push_str(&format!(
                "   {:<42}{:.3}\n",
                tableau.constraints[c_idx].full_name(),
                self.ranking_values[c_idx]
            ));
        }

        // Section 4: Pairwise ranking probabilities (Stochastic OT only)
        if !self.maxent_mode {
            out.push_str("\n\n");
            out.push_str(&self.format_pairwise_probabilities(tableau));
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
    pub fn format_pairwise_probabilities(&self, tableau: &Tableau) -> String {
        let sorted = crate::tableau::sorted_indices_descending(&self.ranking_values);
        let n = sorted.len();
        let abbrevs: Vec<String> = sorted.iter().map(|&i| tableau.constraints[i].abbrev()).collect();

        let mut out = String::new();
        out.push_str("4. Ranking Value to Ranking Probability Conversion\n\n");
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

impl Tableau {
    /// Run the Gradual Learning Algorithm using the default 4-stage schedule.
    ///
    /// This is a convenience wrapper around `run_gla_with_schedule`.
    /// Reproduces VB6 boersma.frm:GLACore and related subroutines.
    #[allow(clippy::too_many_arguments)]
    pub fn run_gla(
        &self,
        maxent_mode: bool,
        cycles: usize,
        initial_plasticity: f64,
        final_plasticity: f64,
        test_trials: usize,
        negative_weights_ok: bool,
        gaussian_prior: bool,
        sigma: f64,
        magri_update_rule: bool,
        exact_proportions: bool,
    ) -> GlaResult {
        let schedule = LearningSchedule::default_4stage(cycles, initial_plasticity, final_plasticity);
        self.run_gla_with_schedule(
            maxent_mode,
            &schedule,
            test_trials,
            negative_weights_ok,
            gaussian_prior,
            sigma,
            magri_update_rule,
            exact_proportions,
            false,
            false,
        )
    }

    /// Run GLA `run_count` times and format results as `CollateRuns.txt` content.
    ///
    /// Reproduces VB6 `boersma.frm:MultipleRuns`. Each run appends:
    ///   - **G records**: `G\t{run}\t{constraint_abbrev}\t{ranking_value}` (one per constraint)
    ///   - **O records**: `O\t{run}\t{form_idx}\t{input}\t{rival_idx}\t{rival_form}\t{freq}\t{pct_gen}`
    ///     (one per non-first candidate per form; VB6 skips candidate 0, the winner)
    #[allow(clippy::too_many_arguments)]
    pub fn format_collate_runs_output(
        &self,
        run_count: usize,
        maxent_mode: bool,
        schedule: &LearningSchedule,
        test_trials: usize,
        negative_weights_ok: bool,
        gaussian_prior: bool,
        sigma: f64,
        magri_update_rule: bool,
        exact_proportions: bool,
    ) -> String {
        let nc = self.constraints.len();
        let mut out = String::new();

        for run_idx in 1..=run_count {
            let result = self.run_gla_with_schedule(
                maxent_mode,
                schedule,
                test_trials,
                negative_weights_ok,
                gaussian_prior,
                sigma,
                magri_update_rule,
                exact_proportions,
                false,
                false,
            );

            // G records: constraint ranking values / weights
            for ci in 0..nc {
                out.push_str(&format!(
                    "G\t{}\t{}\t{:.3}\n",
                    run_idx,
                    self.constraints[ci].abbrev(),
                    result.ranking_values[ci],
                ));
            }

            // O records: rival candidates with predicted probabilities.
            // VB6 loops from RivalIndex=1, skipping candidate 0 (the winner).
            for (fi, form) in self.forms.iter().enumerate() {
                for ri in 1..form.candidates.len() {
                    let cand = &form.candidates[ri];
                    let pct_gen = result
                        .test_probs
                        .get(fi)
                        .and_then(|f| f.get(ri))
                        .copied()
                        .unwrap_or(0.0)
                        * 100.0;
                    out.push_str(&format!(
                        "O\t{}\t{}\t{}\t{}\t{}\t{}\t{:.4}\n",
                        run_idx,
                        fi + 1,
                        form.input,
                        ri,
                        cand.form,
                        cand.frequency,
                        pct_gen,
                    ));
                }
            }
        }

        out
    }

    /// Run the Gradual Learning Algorithm with an explicit learning schedule.
    ///
    /// Supports separate plasticity for Markedness vs Faithfulness constraints
    /// (PlastMark / PlastFaith per stage), matching VB6 behavior.
    ///
    /// # Arguments
    /// * `maxent_mode` — if true, run online MaxEnt; if false, run Stochastic OT
    /// * `schedule` — multi-stage plasticity schedule
    /// * `test_trials` — evaluation trials after learning; only used in StochasticOT mode
    /// * `negative_weights_ok` — allow weights below 0 (MaxEnt mode only)
    /// * `gaussian_prior` — apply Gaussian L2 prior each update (MaxEnt mode only)
    /// * `sigma` — standard deviation of the Gaussian prior (mu=0 for all constraints)
    /// * `magri_update_rule` — scale promotion plasticity by Magri's factor (StochasticOT only)
    #[allow(clippy::too_many_arguments)]
    pub fn run_gla_with_schedule(
        &self,
        maxent_mode: bool,
        schedule: &LearningSchedule,
        test_trials: usize,
        negative_weights_ok: bool,
        gaussian_prior: bool,
        sigma: f64,
        magri_update_rule: bool,
        exact_proportions: bool,
        generate_history: bool,
        generate_full_history: bool,
    ) -> GlaResult {
        let nc = self.constraints.len();

        // ── Faithfulness flags ────────────────────────────────────────────────
        // Precomputed to distinguish PlastMark vs PlastFaith in the update loop.
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

        // ── Round up schedule for exact proportions ──────────────────────────
        let schedule = if exact_proportions && pool_size > 0 {
            let mut s = schedule.clone();
            s.round_stages_to_multiple(pool_size);
            s
        } else {
            schedule.clone()
        };

        // ── Initialize ranking values / weights ──────────────────────────────
        // StochasticOT starts at 100 (Boersma's canonical value)
        // MaxEnt starts at 0
        let initial_value = if maxent_mode { 0.0 } else { 100.0 };
        let mut ranking_values = vec![initial_value; nc];

        // ── History buffer ──────────────────────────────────────────────────
        let mut history_buf = if generate_history {
            let mut header = String::from("Trial");
            for c in &self.constraints {
                header.push('\t');
                header.push_str(&c.abbrev());
            }
            header.push('\n');
            Some(header)
        } else {
            None
        };
        let mut trial_number: usize = 0;

        // ── Full history buffer ─────────────────────────────────────────
        // Reproduces VB6 boersma.frm FullHistory output: annotated log with
        // trial, input, generated, heard, then one column per constraint.
        let mut full_history_buf = if generate_full_history {
            use std::fmt::Write;
            let mut header = String::from("Trial #\tInput\tGenerated\tHeard");
            for c in &self.constraints {
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

        let mut rng = Rng::new(GaussianMode::Standard);

        let mode_name = if maxent_mode { "MaxEnt" } else { "StochasticOT" };
        let total_cycles = schedule.total_cycles();
        crate::ot_log!("Starting GLA ({}) with {} constraints, {} training exemplars, {} cycles",
            mode_name, nc, pool_size, total_cycles);

        // ── Exact proportions cursor ─────────────────────────────────────────
        // Start at pool_size so the first use triggers a shuffle.
        let mut pool_cursor: usize = pool_size;

        // ── Main learning loop ────────────────────────────────────────────────
        if pool_size > 0 {
            for (stage_idx, stage) in schedule.stages.iter().enumerate() {
                for _ in 0..stage.trials {
                    trial_number += 1;

                    // Select observed exemplar
                    let (sel_form, sel_cand) = if exact_proportions {
                        if pool_cursor >= pool_size {
                            // Fisher-Yates shuffle
                            for i in (1..pool_size).rev() {
                                let j = (rng.uniform() * (i + 1) as f64) as usize;
                                training_pool.swap(i, j);
                            }
                            pool_cursor = 0;
                        }
                        let pair = training_pool[pool_cursor];
                        pool_cursor += 1;
                        pair
                    } else {
                        let r = rng.uniform();
                        let idx = ((r * pool_size as f64) as usize).min(pool_size - 1);
                        training_pool[idx]
                    };

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

                    // Skip if the grammar already produced the correct form
                    if generated == sel_cand {
                        continue;
                    }

                    let gen_cand = &self.forms[sel_form].candidates[generated];
                    let obs_cand = &self.forms[sel_form].candidates[sel_cand];

                    // Update ranking values / weights
                    if maxent_mode {
                        // MaxEnt update: reproduces VB6 RankingValueAdjustment (MaxEnt branch).
                        // PriorBasedChange = plasticity * (weight - mu) / sigma² / 2  (mu=0)
                        // The "/2" is "out of the blue" per the VB6 author's comment; reproduced for fidelity.
                        let sigma_sq = sigma * sigma;
                        for (c, rv) in ranking_values.iter_mut().enumerate() {
                            let plast = if is_faith[c] { stage.plast_faith } else { stage.plast_mark };
                            let gen_v = gen_cand.violations[c] as f64;
                            let obs_v = obs_cand.violations[c] as f64;
                            let likelihood_change = plast * (gen_v - obs_v);
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
                        // Reproduces VB6 RankingValueAdjustment (StochasticOT branch).
                        //
                        // Magri update rule: promotion is scaled by
                        //   NumberOfConstraintsDemoted / (NumberOfConstraintsPromoted + 1)
                        // Demotion is unaffected (VB6 comment: "no special regime for demotion").
                        let promo_scale = if magri_update_rule {
                            magri_promotion_amount(gen_cand, obs_cand, nc)
                        } else {
                            1.0
                        };
                        for (c, rv) in ranking_values.iter_mut().enumerate() {
                            let plast = if is_faith[c] { stage.plast_faith } else { stage.plast_mark };
                            let gen_v = gen_cand.violations[c];
                            let obs_v = obs_cand.violations[c];
                            if gen_v > obs_v {
                                // Generated violates more: strengthen (raise)
                                *rv += plast * promo_scale;
                            } else if gen_v < obs_v {
                                // Generated violates less: weaken (lower)
                                *rv -= plast;
                            }
                        }
                    }

                    // Record history after each mismatch (VB6: writes after RankingValueAdjustment)
                    if let Some(ref mut buf) = history_buf {
                        use std::fmt::Write;
                        write!(buf, "{}", trial_number).unwrap();
                        for rv in ranking_values.iter() {
                            write!(buf, "\t{rv:.4}").unwrap();
                        }
                        buf.push('\n');
                    }

                    // Record full history (annotated log with input/generated/heard)
                    if let Some(ref mut buf) = full_history_buf {
                        use std::fmt::Write;
                        write!(buf, "{}\t{}\t{}\t{}",
                            trial_number,
                            self.forms[sel_form].input,
                            gen_cand.form,
                            obs_cand.form,
                        ).unwrap();
                        for rv in ranking_values.iter() {
                            write!(buf, "\t{rv:.4}").unwrap();
                        }
                        buf.push('\n');
                    }
                }
                crate::ot_log!("GLA stage {}/{} complete (plast_mark={:.4}, plast_faith={:.4})",
                    stage_idx + 1, schedule.stages.len(), stage.plast_mark, stage.plast_faith);
            }
        }

        // ── Test grammar ──────────────────────────────────────────────────────
        let test_probs = if maxent_mode {
            // MaxEnt: exact predicted probabilities (reproduces GenerateMaxEntPredictions)
            maxent_predictions(self, &ranking_values)
        } else {
            // StochasticOT: stochastic test (reproduces GLATestGrammar).
            // Use noise from the last stage (finest plasticity, lowest noise stage).
            let last_stage = schedule.stages.last();
            let noise_mark = last_stage.map(|s| s.noise_mark).unwrap_or(2.0);
            let noise_faith = last_stage.map(|s| s.noise_faith).unwrap_or(2.0);

            let mut counts: Vec<Vec<usize>> = self.forms.iter()
                .map(|f| vec![0usize; f.candidates.len()])
                .collect();
            for _ in 0..test_trials {
                for (fi, form) in self.forms.iter().enumerate() {
                    let winner = ot_evaluate(
                        &form.candidates,
                        &ranking_values,
                        &is_faith,
                        noise_mark,
                        noise_faith,
                        &mut rng,
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
        };

        // ── Log likelihood ────────────────────────────────────────────────────
        // Σ frequency × log(predicted_proportion)
        let mut log_likelihood = 0.0f64;
        for (fi, form) in self.forms.iter().enumerate() {
            for (ci, cand) in form.candidates.iter().enumerate() {
                if cand.frequency > 0 {
                    let pred = test_probs[fi][ci];
                    if pred > 0.0 {
                        log_likelihood += cand.frequency as f64 * pred.ln();
                    }
                }
            }
        }

        crate::ot_log!("GLA ({}) DONE: log_likelihood = {:.6}", mode_name, log_likelihood);

        GlaResult {
            ranking_values,
            test_probs,
            log_likelihood,
            maxent_mode,
            schedule_description: schedule.format_description(),
            test_trials,
            gaussian_prior,
            sigma,
            magri_update_rule,
            history: history_buf,
            full_history: full_history_buf,
        }
    }
}

// ─── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use crate::schedule::LearningSchedule;
    use crate::tableau::Tableau;

    fn load_tiny() -> String {
        std::fs::read_to_string("../examples/TinyIllustrativeFile/input.txt")
            .expect("Failed to load examples/TinyIllustrativeFile/input.txt")
    }

    #[test]
    fn test_gla_sot_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0, false, false);

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(), "ranking value {} should be finite", c);
        }
        assert!(result.log_likelihood().is_finite());
    }

    #[test]
    fn test_gla_maxent_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, false, 1.0, false, false);

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
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, false, 1.0, false, false);

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
        let result = tableau.run_gla(false, 0, 2.0, 0.001, 100, false, false, 1.0, false, false);
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
        let result = tableau.run_gla(true, 0, 2.0, 0.001, 0, false, false, 1.0, false, false);
        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert_eq!(result.get_ranking_value(c), 0.0,
                "MaxEnt should start at 0, got {}", result.get_ranking_value(c));
        }
    }

    #[test]
    fn test_gla_maxent_gaussian_prior_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, true, 1.0, false, false);

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
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0, true, false);

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
        let result = tableau.run_gla_with_schedule(false, &schedule, 200, false, false, 1.0, false, false, false, false);

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
            3,     // run_count
            false, // maxent_mode
            &schedule,
            200,   // test_trials
            false, // negative_weights_ok
            false, // gaussian_prior
            1.0,   // sigma
            false, // magri_update_rule
            false, // exact_proportions
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
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0, false, false);

        let table = result.format_pairwise_probabilities(&tableau);
        assert!(table.contains("4. Ranking Value to Ranking Probability Conversion"));
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
    fn test_gla_sot_format_output_includes_pairwise() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0, false, false);

        let output = result.format_output(&tableau, "test.txt");
        assert!(output.contains("4. Ranking Value to Ranking Probability Conversion"),
            "Stochastic OT output should include pairwise probability section");
    }

    #[test]
    fn test_gla_maxent_format_output_excludes_pairwise() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, false, 1.0, false, false);

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
        let result = tableau.run_gla_with_schedule(false, &schedule, 200, false, false, 1.0, false, false, false, false);
        let output = result.format_output(&tableau, "test.txt");

        // Custom schedule should mention stages in the output
        assert!(output.contains("Custom learning schedule"), "output should mention custom schedule");
        assert!(output.contains("PlastMark"), "output should show PlastMark column");
    }

    #[test]
    fn test_gla_sot_exact_proportions_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0, false, true);

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
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, false, 1.0, false, true);

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
        let result = tableau.run_gla_with_schedule(false, &schedule, 200, false, false, 1.0, false, false, true, false);

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
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0, false, false);
        assert!(result.history().is_none(), "history should be None when generate_history=false");
    }

    #[test]
    fn test_gla_full_history_generation() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule = LearningSchedule::default_4stage(500, 2.0, 0.001);
        let result = tableau.run_gla_with_schedule(false, &schedule, 200, false, false, 1.0, false, false, false, true);

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
        assert!(lines.len() > 2, "should have data rows after header + initial");
        for line in &lines[2..] {
            let cols: Vec<&str> = line.split('\t').collect();
            assert_eq!(cols.len(), nc + 4, "each data row should have {} columns", nc + 4);
            // Trial number
            cols[0].parse::<usize>().expect("trial number should be an integer");
            // Input, Generated, Heard should be non-empty
            assert!(!cols[1].is_empty(), "input should be non-empty");
            assert!(!cols[2].is_empty(), "generated should be non-empty");
            assert!(!cols[3].is_empty(), "heard should be non-empty");
        }
    }

    #[test]
    fn test_gla_full_history_none_when_disabled() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0, false, false);
        assert!(result.full_history().is_none(), "full_history should be None when disabled");
    }
}
