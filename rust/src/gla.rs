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
use crate::schedule::LearningSchedule;
use crate::tableau::Tableau;

// ─── Gaussian RNG ─────────────────────────────────────────────────────────────
// Box-Muller transform (matches the VB6 boersma.frm:Gaussian function).
// Note: the VB6 function multiplies by 2 (giving σ=2), which is the noise level
// used directly for StochasticOT ranking value perturbation.

fn getrandom_uniform() -> f64 {
    let mut bytes = [0u8; 8];
    getrandom::getrandom(&mut bytes).expect("getrandom failed");
    let n = u64::from_le_bytes(bytes);
    (n as f64) * (1.0 / 18_446_744_073_709_551_616.0_f64)
}

struct Rng {
    stored: Option<f64>,
}

impl Rng {
    fn new() -> Self {
        Rng { stored: None }
    }

    fn uniform(&mut self) -> f64 {
        getrandom_uniform()
    }

    /// Standard normal deviate (σ=1). Scale by noise_sigma at call site.
    fn gaussian(&mut self) -> f64 {
        if let Some(stored) = self.stored.take() {
            return stored;
        }
        loop {
            let v1 = 2.0 * getrandom_uniform() - 1.0;
            let v2 = 2.0 * getrandom_uniform() - 1.0;
            let r = v1 * v1 + v2 * v2;
            if r > 0.0 && r < 1.0 {
                let fac = (-2.0 * r.ln() / r).sqrt();
                self.stored = Some(v1 * fac);
                return v2 * fac;
            }
        }
    }
}

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

    /// Format results as text output for download.
    /// Reproduces the structure of OTSoft's GLA text output.
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        let nc = tableau.constraints.len();
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
        let mut sorted: Vec<usize> = (0..nc).collect();
        sorted.sort_by(|&a, &b| {
            self.ranking_values[b]
                .partial_cmp(&self.ranking_values[a])
                .unwrap_or(std::cmp::Ordering::Equal)
        });
        for &c_idx in &sorted {
            out.push_str(&format!(
                "   {:<42}{:.3}\n",
                tableau.constraints[c_idx].full_name(),
                self.ranking_values[c_idx]
            ));
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
    ) -> GlaResult {
        let schedule = LearningSchedule::default_4stage(cycles, initial_plasticity, final_plasticity);
        self.run_gla_with_schedule(
            maxent_mode,
            &schedule,
            test_trials,
            negative_weights_ok,
            gaussian_prior,
            sigma,
        )
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
    #[allow(clippy::too_many_arguments)]
    pub fn run_gla_with_schedule(
        &self,
        maxent_mode: bool,
        schedule: &LearningSchedule,
        test_trials: usize,
        negative_weights_ok: bool,
        gaussian_prior: bool,
        sigma: f64,
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

        // ── Initialize ranking values / weights ──────────────────────────────
        // StochasticOT starts at 100 (Boersma's canonical value)
        // MaxEnt starts at 0
        let initial_value = if maxent_mode { 0.0 } else { 100.0 };
        let mut ranking_values = vec![initial_value; nc];

        let mut rng = Rng::new();

        let mode_name = if maxent_mode { "MaxEnt" } else { "StochasticOT" };
        let total_cycles = schedule.total_cycles();
        crate::ot_log!("Starting GLA ({}) with {} constraints, {} training exemplars, {} cycles",
            mode_name, nc, pool_size, total_cycles);

        // ── Main learning loop ────────────────────────────────────────────────
        if pool_size > 0 {
            for (stage_idx, stage) in schedule.stages.iter().enumerate() {
                for _ in 0..stage.trials {
                    // Select observed exemplar weighted by frequency
                    let r = rng.uniform();
                    let idx = ((r * pool_size as f64) as usize).min(pool_size - 1);
                    let (sel_form, sel_cand) = training_pool[idx];

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
                        // StochasticOT update: ±plasticity per constraint
                        // Reproduces VB6 RankingValueAdjustment (StochasticOT branch)
                        for (c, rv) in ranking_values.iter_mut().enumerate() {
                            let plast = if is_faith[c] { stage.plast_faith } else { stage.plast_mark };
                            let gen_v = gen_cand.violations[c];
                            let obs_v = obs_cand.violations[c];
                            if gen_v > obs_v {
                                // Generated violates more: strengthen (raise)
                                *rv += plast;
                            } else if gen_v < obs_v {
                                // Generated violates less: weaken (lower)
                                *rv -= plast;
                            }
                        }
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
        }
    }
}

// ─── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use crate::schedule::LearningSchedule;
    use crate::tableau::Tableau;

    fn load_tiny() -> String {
        std::fs::read_to_string("../examples/tiny/input.txt")
            .expect("Failed to load examples/tiny/input.txt")
    }

    #[test]
    fn test_gla_sot_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(false, 500, 2.0, 0.001, 200, false, false, 1.0);

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(), "ranking value {} should be finite", c);
        }
        assert!(result.log_likelihood().is_finite());
    }

    #[test]
    fn test_gla_maxent_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, false, 1.0);

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
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, false, 1.0);

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
        let result = tableau.run_gla(false, 0, 2.0, 0.001, 100, false, false, 1.0);
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
        let result = tableau.run_gla(true, 0, 2.0, 0.001, 0, false, false, 1.0);
        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert_eq!(result.get_ranking_value(c), 0.0,
                "MaxEnt should start at 0, got {}", result.get_ranking_value(c));
        }
    }

    #[test]
    fn test_gla_maxent_gaussian_prior_runs() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let result = tableau.run_gla(true, 500, 2.0, 0.001, 0, false, true, 1.0);

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite(), "weight {} should be finite with prior", c);
        }
        assert!(result.log_likelihood().is_finite());
        assert!(result.gaussian_prior(), "gaussian_prior() should be true");
        assert_eq!(result.sigma(), 1.0, "sigma() should be 1.0");
    }

    #[test]
    fn test_gla_custom_schedule_runs() {
        // A custom schedule with separate M/F plasticity should run without errors
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule_text = "Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n\
                             250\t2\t0.5\t2\t2\n\
                             250\t0.2\t0.05\t2\t2\n";
        let schedule = LearningSchedule::parse(schedule_text).unwrap();
        let result = tableau.run_gla_with_schedule(false, &schedule, 200, false, false, 1.0);

        let nc = tableau.constraint_count();
        for c in 0..nc {
            assert!(result.get_ranking_value(c).is_finite());
        }
    }

    #[test]
    fn test_gla_custom_schedule_output_contains_stage_info() {
        let tableau = Tableau::parse(&load_tiny()).unwrap();
        let schedule_text = "Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n\
                             250\t2\t0.5\t2\t2\n\
                             250\t0.2\t0.05\t2\t2\n";
        let schedule = LearningSchedule::parse(schedule_text).unwrap();
        let result = tableau.run_gla_with_schedule(false, &schedule, 200, false, false, 1.0);
        let output = result.format_output(&tableau, "test.txt");

        // Custom schedule should mention stages in the output
        assert!(output.contains("Custom learning schedule"), "output should mention custom schedule");
        assert!(output.contains("PlastMark"), "output should show PlastMark column");
    }
}
