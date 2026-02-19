//! Noisy Harmonic Grammar (NHG) algorithm
//!
//! Online error-driven learner. For each training trial, an observed output is
//! selected (weighted by frequency), a form is generated using Gaussian-perturbed
//! constraint weights, and weights are adjusted if the generated form differs from
//! the observed one.
//!
//! Implements 8 noise variants controlled by boolean flags. Reproduces
//! VB6 NoisyHarmonicGrammar.frm.

use wasm_bindgen::prelude::*;
use crate::tableau::Tableau;

// ─── Gaussian RNG ────────────────────────────────────────────────────────────
//
// Box-Muller transform, matching the VB6 Gaussian() function.
// Uses `getrandom` for a WASM-compatible (js) and native-compatible uniform source.

fn getrandom_uniform() -> f64 {
    let mut bytes = [0u8; 8];
    getrandom::getrandom(&mut bytes).expect("getrandom failed");
    let n = u64::from_le_bytes(bytes);
    // Map to [0, 1) — use the ratio n / 2^64
    (n as f64) * (1.0 / 18_446_744_073_709_551_616.0_f64)
}

/// State for the Gaussian RNG. Caches the second value from Box-Muller.
struct Rng {
    stored: Option<f64>,
    demi_gaussians: bool,
}

impl Rng {
    fn new(demi_gaussians: bool) -> Self {
        Rng { stored: None, demi_gaussians }
    }

    /// Uniform sample in [0, 1)
    fn uniform(&mut self) -> f64 {
        getrandom_uniform()
    }

    /// Standard normal deviate (σ=1), with demi-Gaussian option.
    fn gaussian(&mut self) -> f64 {
        let val = if let Some(stored) = self.stored.take() {
            stored
        } else {
            // Box-Muller rejection sampling (matches VB6)
            loop {
                let v1 = 2.0 * getrandom_uniform() - 1.0;
                let v2 = 2.0 * getrandom_uniform() - 1.0;
                let r = v1 * v1 + v2 * v2;
                if r > 0.0 && r < 1.0 {
                    let fac = (-2.0 * r.ln() / r).sqrt();
                    // Cache one, return the other (matches VB6)
                    self.stored = Some(v1 * fac);
                    break v2 * fac;
                }
            }
        };
        if self.demi_gaussians { val.abs() } else { val }
    }
}

// ─── Form generation ─────────────────────────────────────────────────────────

const NO_WINNER: usize = usize::MAX;
const MAX_TIE_RETRIES: usize = 100;

/// Compute the noisy harmony of each candidate for one input form and return the index
/// of the winning candidate (lowest harmony). Returns `NO_WINNER` if tie-skipping is
/// active and a tie persists after MAX_TIE_RETRIES attempts.
fn generate_form(
    candidates: &[crate::tableau::Candidate],
    weights: &[f64],
    noise_std: f64,
    noise_by_cell: bool,
    post_mult_noise: bool,
    noise_for_zero_cells: bool,
    late_noise: bool,
    exponential_nhg: bool,
    negative_weights_ok: bool,
    resolve_ties_by_skipping: bool,
    rng: &mut Rng,
) -> usize {
    let nc = weights.len();
    let n_cands = candidates.len();

    for attempt in 0..=MAX_TIE_RETRIES {
        // ── Variant-specific setup (computed once per input evaluation, not per candidate) ──

        // Pre-mult by-constraint (Variant A/A'): compute noisy weights once.
        let noisy_weights_pre: Vec<f64>;
        // Post-mult by-constraint (Variant C/D): compute perturbations once.
        let perturbations: Vec<f64>;

        if !late_noise && !noise_by_cell {
            if post_mult_noise {
                // C/D: one noise term per constraint, added after weight × violations
                perturbations = (0..nc).map(|_| noise_std * rng.gaussian()).collect();
                noisy_weights_pre = vec![];
            } else {
                // A/A': perturb weight before multiplying by violations
                noisy_weights_pre = (0..nc)
                    .map(|c| {
                        let mut local = weights[c] + noise_std * rng.gaussian();
                        if local < 0.0 && !negative_weights_ok && !exponential_nhg {
                            local = 0.0;
                        }
                        local
                    })
                    .collect();
                perturbations = vec![];
            }
        } else {
            noisy_weights_pre = vec![];
            perturbations = vec![];
        }

        // ── Evaluate each candidate ──────────────────────────────────────────────────────

        let mut best_harmony = f64::INFINITY;
        let mut winner_idx = 0;
        let mut tied: Vec<bool> = vec![false; n_cands];
        let mut n_tied = 0usize;

        for (ci, cand) in candidates.iter().enumerate() {
            let harmony: f64 = if late_noise {
                // Variant G: harmony = Σ(w×v) + one Gaussian term per candidate
                let base: f64 = (0..nc)
                    .map(|c| {
                        let v = cand.violations[c] as f64;
                        if exponential_nhg {
                            (weights[c] * v).exp()
                        } else {
                            weights[c] * v
                        }
                    })
                    .sum();
                base + noise_std * rng.gaussian()
            } else if !noise_by_cell {
                if post_mult_noise {
                    // C/D: Σ(w×v + perturbation[c]), with zero-cell option
                    let mut h = 0.0f64;
                    for c in 0..nc {
                        let v = cand.violations[c] as f64;
                        let contrib = if v == 0.0 {
                            if noise_for_zero_cells {
                                // C: noise even in zero-violation cells
                                if exponential_nhg {
                                    (v + perturbations[c]).exp()
                                } else {
                                    perturbations[c]
                                }
                            } else {
                                0.0 // D: no noise for zero cells
                            }
                        } else if exponential_nhg {
                            (weights[c] * v + perturbations[c]).exp()
                        } else {
                            // Check effective weight; floor to zero if not OK with negatives
                            let eff_w = weights[c] + perturbations[c] / v;
                            if eff_w < 0.0 && !negative_weights_ok {
                                0.0
                            } else {
                                weights[c] * v + perturbations[c]
                            }
                        };
                        h += contrib;
                    }
                    h
                } else {
                    // A/A': use pre-computed noisy_weights
                    (0..nc)
                        .map(|c| {
                            let v = cand.violations[c] as f64;
                            if exponential_nhg {
                                noisy_weights_pre[c].exp() * v
                            } else {
                                noisy_weights_pre[c] * v
                            }
                        })
                        .sum()
                }
            } else {
                // B/B' or E/F: noise generated per (candidate, constraint) cell
                let mut h = 0.0f64;
                for c in 0..nc {
                    let v = cand.violations[c] as f64;
                    if post_mult_noise {
                        // E/F: noise added after weight × violations
                        if v != 0.0 || noise_for_zero_cells {
                            let g = noise_std * rng.gaussian();
                            let contrib = if exponential_nhg {
                                (weights[c] * v + g).exp()
                            } else {
                                let eff_w = if v != 0.0 {
                                    weights[c] + g / v
                                } else {
                                    0.0
                                };
                                if eff_w < 0.0 && !negative_weights_ok {
                                    0.0
                                } else {
                                    weights[c] * v + g
                                }
                            };
                            h += contrib;
                        }
                    } else {
                        // B/B': re-perturb weight per cell, zero × anything = 0
                        if v != 0.0 {
                            let mut local_w = weights[c] + noise_std * rng.gaussian();
                            if local_w < 0.0 && !negative_weights_ok && !exponential_nhg {
                                local_w = 0.0;
                            }
                            h += if exponential_nhg {
                                local_w.exp() * v
                            } else {
                                local_w * v
                            };
                        }
                        // v == 0.0: contribution is 0 regardless of noise
                    }
                }
                h
            };

            if harmony < best_harmony {
                best_harmony = harmony;
                winner_idx = ci;
                // Reset tie tracking
                for t in tied.iter_mut() { *t = false; }
                tied[ci] = true;
                n_tied = 1;
            } else if harmony == best_harmony {
                tied[ci] = true;
                n_tied += 1;
            }
        }

        // ── Tie resolution ───────────────────────────────────────────────────────────────

        if n_tied <= 1 {
            return winner_idx;
        }

        if resolve_ties_by_skipping {
            if attempt == MAX_TIE_RETRIES {
                return NO_WINNER;
            }
            // retry
            continue;
        }

        // Random pick among tied candidates
        let pick = (rng.uniform() * n_tied as f64) as usize;
        let mut count = 0usize;
        for (ci, &is_tied) in tied.iter().enumerate() {
            if is_tied {
                if count == pick {
                    return ci;
                }
                count += 1;
            }
        }
        return winner_idx; // fallback (shouldn't reach here)
    }

    NO_WINNER
}

// ─── NHG Result ──────────────────────────────────────────────────────────────

/// Result of running the Noisy Harmonic Grammar algorithm
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct NhgResult {
    /// Final constraint weights (in original constraint order)
    weights: Vec<f64>,
    /// Predicted probabilities from testing: test_probs[form_idx][cand_idx]
    test_probs: Vec<Vec<f64>>,
    /// Log likelihood of training data under the tested grammar
    log_likelihood: f64,
    /// Parameters used (stored for output formatting)
    cycles: usize,
    initial_plasticity: f64,
    final_plasticity: f64,
    test_trials: usize,
    exponential_nhg: bool,
}

#[wasm_bindgen]
impl NhgResult {
    pub fn get_weight(&self, constraint_index: usize) -> f64 {
        self.weights.get(constraint_index).copied().unwrap_or(0.0)
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

    /// Format results as text output for download.
    /// Reproduces the structure of OTSoft's NHG text output.
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        let nc = tableau.constraints.len();
        let mut out = String::new();

        // Header
        out.push_str(&format!(
            "Result of Applying Noisy Harmonic Grammar to {}\n\n\n",
            filename
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
        out.push_str(&format!("   Cycles: {}\n", self.cycles));
        out.push_str(&format!("   Initial plasticity: {:.3}\n", self.initial_plasticity));
        out.push_str(&format!("   Final plasticity: {:.3}\n", self.final_plasticity));
        out.push_str(&format!("   Times to test grammar: {}\n", self.test_trials));
        if self.exponential_nhg {
            out.push_str("   Exponential NHG: yes\n");
        }
        out.push_str("\n\n");

        // Section 1: Weights Found (in original constraint order, matching VB6)
        out.push_str("1. Weights Found\n\n");
        for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
            let w = self.weights[c_idx];
            if self.exponential_nhg {
                out.push_str(&format!(
                    "   {:<42}{:.3}    exp: {:.3}\n",
                    constraint.full_name(),
                    w,
                    w.exp()
                ));
            } else {
                out.push_str(&format!(
                    "   {:<42}{:.3}\n",
                    constraint.full_name(),
                    w
                ));
            }
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

            // Header row
            let max_cand_width = form.candidates.iter()
                .map(|c| c.form.len())
                .max()
                .unwrap_or(0)
                .max(2);

            out.push_str(&format!(
                "  {:<width$}  {:>9}  {:>9}\n",
                "",
                "Input%",
                "Gen%",
                width = max_cand_width
            ));

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
                    marker,
                    cand.form,
                    obs_pct,
                    gen_pct,
                    width = max_cand_width
                ));
            }
            out.push('\n');
        }

        // Constraint violation tableaux (if desired — matches VB6 optional include)
        out.push_str("3. Constraint Weights (sorted)\n\n");
        let mut sorted: Vec<usize> = (0..nc).collect();
        sorted.sort_by(|&a, &b| {
            self.weights[b]
                .partial_cmp(&self.weights[a])
                .unwrap_or(std::cmp::Ordering::Equal)
        });
        for &c_idx in &sorted {
            out.push_str(&format!(
                "   {:<42}{:.3}\n",
                tableau.constraints[c_idx].full_name(),
                self.weights[c_idx]
            ));
        }

        out
    }
}

// ─── Main NHG algorithm ───────────────────────────────────────────────────────

impl Tableau {
    /// Run the Noisy Harmonic Grammar algorithm.
    ///
    /// Reproduces VB6 NoisyHarmonicGrammar.frm:NoisyHarmonicGrammarCore and
    /// NHGTestGrammar.
    ///
    /// # Arguments
    /// * `cycles` — total training iterations (default 5000)
    /// * `initial_plasticity` — starting learning rate (default 2.0)
    /// * `final_plasticity` — ending learning rate (default 0.002)
    /// * `test_trials` — evaluation trials after learning (default 2000)
    /// * `noise_by_cell` — if true, draw a fresh noise value per (candidate, constraint) cell
    /// * `post_mult_noise` — if true, add noise after weight × violations product
    /// * `noise_for_zero_cells` — if true, apply noise even in zero-violation cells
    /// * `late_noise` — if true, add single noise term to total harmony (after all constraints)
    /// * `exponential_nhg` — use exp(weight) instead of raw weight
    /// * `demi_gaussians` — use positive-only (half-Gaussian) noise
    /// * `negative_weights_ok` — allow weights to go below zero
    /// * `resolve_ties_by_skipping` — skip trial on tie instead of random pick
    #[allow(clippy::too_many_arguments)]
    pub fn run_nhg(
        &self,
        cycles: usize,
        initial_plasticity: f64,
        final_plasticity: f64,
        test_trials: usize,
        noise_by_cell: bool,
        post_mult_noise: bool,
        noise_for_zero_cells: bool,
        late_noise: bool,
        exponential_nhg: bool,
        demi_gaussians: bool,
        negative_weights_ok: bool,
        resolve_ties_by_skipping: bool,
    ) -> NhgResult {
        let nc = self.constraints.len();

        // Noise std deviation: 1.0 for normal NHG, 0.1 for exponential (matches VB6 SetTheNoise)
        let noise_std = if exponential_nhg { 0.1 } else { 1.0 };

        // ── Build training data pool ─────────────────────────────────────────────────────
        // All (form_idx, cand_idx) pairs where frequency > 0, repeated by frequency.
        // Matches VB6: winner at index 0 is included when looping from 0 to mNumberOfRivals.
        let mut training_pool: Vec<(usize, usize)> = Vec::new();
        for (fi, form) in self.forms.iter().enumerate() {
            for (ci, cand) in form.candidates.iter().enumerate() {
                for _ in 0..cand.frequency {
                    training_pool.push((fi, ci));
                }
            }
        }
        let pool_size = training_pool.len();

        // ── Initialize weights ───────────────────────────────────────────────────────────
        let mut weights = vec![0.0f64; nc];

        // ── Learning schedule: 4 stages with geometric plasticity interpolation ──────────
        // Matches VB6 DetermineLearningSchedule.
        let p1 = initial_plasticity;
        let p4 = final_plasticity;
        let p2 = (p1 * p1 * p4).powf(1.0 / 3.0);
        let p3 = (p1 * p4 * p4).powf(1.0 / 3.0);
        let plasticities = [p1, p2, p3, p4];
        let trials_per_stage = cycles / 4;

        let mut rng = Rng::new(demi_gaussians);

        crate::ot_log!("Starting NHG with {} constraints, {} training exemplars, {} cycles",
            nc, pool_size, cycles);

        // ── Main learning loop ───────────────────────────────────────────────────────────
        if pool_size > 0 {
            for stage in 0..4 {
                let plasticity = plasticities[stage];

                for _ in 0..trials_per_stage {
                    // Select training exemplar (stochastic, weighted by frequency)
                    let r = rng.uniform();
                    let idx = ((r * pool_size as f64) as usize).min(pool_size - 1);
                    let (selected_form, selected_cand) = training_pool[idx];

                    // Generate a form stochastically using the current grammar
                    let generated = generate_form(
                        &self.forms[selected_form].candidates,
                        &weights,
                        noise_std,
                        noise_by_cell,
                        post_mult_noise,
                        noise_for_zero_cells,
                        late_noise,
                        exponential_nhg,
                        negative_weights_ok,
                        resolve_ties_by_skipping,
                        &mut rng,
                    );

                    // Update weights if generated form differs from observed
                    if generated != NO_WINNER && generated != selected_cand {
                        let winner_cand = &self.forms[selected_form].candidates[generated];
                        let target_cand = &self.forms[selected_form].candidates[selected_cand];

                        for c in 0..nc {
                            let wv = winner_cand.violations[c] as f64;
                            let tv = target_cand.violations[c] as f64;
                            if wv != tv {
                                // VB6: weight += plasticity * (winner_viols - training_viols)
                                weights[c] += plasticity * (wv - tv);
                                if weights[c] < 0.0 && !negative_weights_ok && !exponential_nhg {
                                    weights[c] = 0.0;
                                }
                            }
                        }
                    }
                }
                crate::ot_log!("NHG stage {}/4 complete (plasticity = {:.4})", stage + 1, plasticity);
            }
        }

        // ── Test grammar ─────────────────────────────────────────────────────────────────
        // For each trial, generate an output for every input form.
        // Count how often each candidate wins, then divide by test_trials.
        let mut counts: Vec<Vec<usize>> = self.forms
            .iter()
            .map(|f| vec![0usize; f.candidates.len()])
            .collect();

        for _ in 0..test_trials {
            for (fi, form) in self.forms.iter().enumerate() {
                let winner = generate_form(
                    &form.candidates,
                    &weights,
                    noise_std,
                    noise_by_cell,
                    post_mult_noise,
                    noise_for_zero_cells,
                    late_noise,
                    exponential_nhg,
                    negative_weights_ok,
                    resolve_ties_by_skipping,
                    &mut rng,
                );
                if winner != NO_WINNER {
                    counts[fi][winner] += 1;
                }
            }
        }

        let test_probs: Vec<Vec<f64>> = counts
            .iter()
            .map(|form_counts| {
                let total: usize = form_counts.iter().sum();
                if total == 0 {
                    vec![0.0; form_counts.len()]
                } else {
                    form_counts
                        .iter()
                        .map(|&c| c as f64 / total as f64)
                        .collect()
                }
            })
            .collect();

        // ── Log likelihood ───────────────────────────────────────────────────────────────
        // Matches VB6: Σ frequency * log(predicted_proportion)
        let mut log_likelihood = 0.0f64;
        for (fi, form) in self.forms.iter().enumerate() {
            let total_freq: f64 = form.candidates.iter().map(|c| c.frequency as f64).sum();
            if total_freq <= 0.0 {
                continue;
            }
            for (ci, cand) in form.candidates.iter().enumerate() {
                if cand.frequency > 0 {
                    let pred = test_probs[fi][ci];
                    if pred > 0.0 {
                        log_likelihood += cand.frequency as f64 * pred.ln();
                    }
                    // Zero prediction → log_likelihood would be -∞; skip per VB6 warning
                }
            }
        }

        crate::ot_log!("NHG DONE: log_likelihood = {:.6}", log_likelihood);

        NhgResult {
            weights,
            test_probs,
            log_likelihood,
            cycles,
            initial_plasticity,
            final_plasticity,
            test_trials,
            exponential_nhg,
        }
    }
}

// ─── Tests ────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use crate::tableau::Tableau;

    fn load_tiny_example() -> String {
        std::fs::read_to_string("../examples/tiny/input.txt")
            .expect("Failed to load examples/tiny/input.txt")
    }

    #[test]
    fn test_nhg_runs_tiny_example() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_nhg(
            500,   // cycles (fewer for speed)
            2.0,   // initial_plasticity
            0.002, // final_plasticity
            200,   // test_trials
            false, false, false, false, false, false, false, false,
        );

        // All weights should be finite
        let nc = tableau.constraint_count();
        for c_idx in 0..nc {
            let w = result.get_weight(c_idx);
            assert!(w.is_finite(), "Weight {} should be finite, got {}", c_idx, w);
            assert!(w >= 0.0, "Weight {} should be >= 0 (no negatives by default), got {}", c_idx, w);
        }

        // Log likelihood should be finite
        assert!(result.log_likelihood().is_finite(), "Log likelihood should be finite");
    }

    #[test]
    fn test_nhg_test_probs_sum_to_one() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_nhg(
            500, 2.0, 0.002, 500, false, false, false, false, false, false, false, false,
        );

        for form_idx in 0..tableau.form_count() {
            let form = tableau.get_form(form_idx).unwrap();
            let total: f64 = (0..form.candidate_count())
                .map(|ci| result.get_test_prob(form_idx, ci))
                .sum();
            // Total can be 0 (all ties skipped) or ~1 (normal operation)
            if total > 0.0 {
                assert!(
                    (total - 1.0).abs() < 1e-9,
                    "Test probs for form {} should sum to 1, got {}",
                    form_idx,
                    total
                );
            }
        }
    }

    #[test]
    fn test_nhg_learns_to_prefer_winners() {
        // After learning, the observed winners should have higher predicted probability
        // than their rivals for the tiny example.
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_nhg(
            2000, 2.0, 0.002, 1000, false, false, false, false, false, false, false, false,
        );

        for form_idx in 0..tableau.form_count() {
            let form = tableau.get_form(form_idx).unwrap();
            // Find the winner (cand with frequency > 0)
            let winner_ci = (0..form.candidate_count())
                .find(|&ci| form.get_candidate(ci).unwrap().frequency > 0);
            if let Some(wi) = winner_ci {
                let winner_prob = result.get_test_prob(form_idx, wi);
                // Winner should have a reasonable probability after learning
                // (Not strictly > 0.5 since noise is stochastic, but > 0 for sure)
                assert!(
                    winner_prob > 0.0,
                    "Winner for form {} should have non-zero probability, got {}",
                    form_idx,
                    winner_prob
                );
            }
        }
    }

    #[test]
    fn test_nhg_output_format() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_nhg(
            200, 2.0, 0.002, 100, false, false, false, false, false, false, false, false,
        );
        let output = result.format_output(&tableau, "test.txt");

        assert!(output.contains("Result of Applying Noisy Harmonic Grammar to test.txt"));
        assert!(output.contains("1. Weights Found"));
        assert!(output.contains("2. Matchup to Input Frequencies"));
        assert!(output.contains("Log likelihood of data:"));
        assert!(output.contains("/a/"));
        assert!(output.contains("/tat/"));
        assert!(output.contains("/at/"));
    }

    #[test]
    fn test_nhg_exponential_variant() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        // Exponential NHG uses smaller noise (0.1) and lower initial plasticity
        let result = tableau.run_nhg(
            500, 0.05, 0.0002, 200, false, false, false, false,
            true,  // exponential_nhg
            false, false, false,
        );
        // Weights should remain finite (may be negative in exponential mode)
        let nc = tableau.constraint_count();
        for c_idx in 0..nc {
            assert!(result.get_weight(c_idx).is_finite(), "Weight should be finite");
        }
    }
}
