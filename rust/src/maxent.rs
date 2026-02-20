//! Maximum Entropy (Batch) algorithm
//!
//! Implements the Generalized Iterative Scaling (GIS) optimizer for MaxEnt
//! grammars, following Goodman (2002) as implemented in OTSoft's MyMaxEnt.frm.
//!
//! All candidates (winners and losers alike) participate in the softmax.
//! The algorithm iteratively adjusts constraint weights so that the expected
//! violation counts match the observed counts.

use wasm_bindgen::prelude::*;
use crate::tableau::Tableau;

/// Result of running the MaxEnt algorithm
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct MaxEntResult {
    /// Final constraint weights
    weights: Vec<f64>,
    /// Predicted probabilities: predicted_probs[form_idx][cand_idx]
    predicted_probs: Vec<Vec<f64>>,
    /// Log probability of the training data under the learned grammar
    log_prob: f64,
    /// Number of GIS iterations actually run
    iterations: usize,
    /// Parameters used
    weight_min: f64,
    weight_max: f64,
}

#[wasm_bindgen]
impl MaxEntResult {
    pub fn get_weight(&self, constraint_index: usize) -> f64 {
        self.weights.get(constraint_index).copied().unwrap_or(0.0)
    }

    pub fn get_predicted_prob(&self, form_index: usize, cand_index: usize) -> f64 {
        self.predicted_probs
            .get(form_index)
            .and_then(|f| f.get(cand_index))
            .copied()
            .unwrap_or(0.0)
    }

    pub fn log_prob(&self) -> f64 {
        self.log_prob
    }

    pub fn iterations(&self) -> usize {
        self.iterations
    }

    /// Format results as text output for download.
    /// Reproduces the structure of OTSoft's MaxEnt text output.
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        let nc = tableau.constraints.len();
        let mut out = String::new();

        // Header
        out.push_str(&format!(
            "Results of Applying Maximum Entropy to {}\n\n\n",
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
        out.push_str(&format!("   Iterations: {}\n", self.iterations));
        out.push_str(&format!("   Weight minimum: {}\n", self.weight_min));
        out.push_str(&format!("   Weight maximum: {}\n", self.weight_max));
        out.push_str("\n\n");

        // Section 1: Constraint weights
        out.push_str("1. Constraint Weights\n\n");

        // Sort by weight descending for display
        let mut sorted_constraints: Vec<usize> = (0..nc).collect();
        sorted_constraints.sort_by(|&a, &b| {
            self.weights[b].partial_cmp(&self.weights[a]).unwrap_or(std::cmp::Ordering::Equal)
        });

        for &c_idx in &sorted_constraints {
            let constraint = &tableau.constraints[c_idx];
            out.push_str(&format!(
                "   {:<42}{:.3}\n",
                constraint.full_name(),
                self.weights[c_idx],
            ));
        }
        out.push('\n');
        out.push_str(&format!("   Log probability of data: {:.6}\n", self.log_prob));
        out.push_str("\n\n");

        // Section 2: Tableaux with predicted probabilities
        out.push_str("2. Tableaux\n\n");

        for (form_idx, form) in tableau.forms.iter().enumerate() {
            out.push_str(&format!("\n/{}/:\n", form.input));

            // Calculate total frequency and observed proportions
            let total_freq: f64 = form.candidates.iter().map(|c| c.frequency as f64).sum();

            // Header row: constraint abbreviations
            let max_cand_width = form.candidates.iter()
                .map(|c| c.form.len())
                .max()
                .unwrap_or(0)
                .max(2);

            // Build header
            let mut header = format!("{:<width$}  {:>7}  {:>7}", "", "Obs%", "Pred%", width = max_cand_width + 2);
            for c_idx in 0..nc {
                header.push_str(&format!("  {:>width$}", tableau.constraints[c_idx].abbrev(), width = tableau.constraints[c_idx].abbrev().len().max(3)));
            }
            out.push_str(&header);
            out.push('\n');

            // Candidate rows
            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                let obs_pct = if total_freq > 0.0 {
                    cand.frequency as f64 / total_freq * 100.0
                } else {
                    0.0
                };
                let pred_pct = self.predicted_probs
                    .get(form_idx)
                    .and_then(|f| f.get(cand_idx))
                    .copied()
                    .unwrap_or(0.0)
                    * 100.0;

                let marker = if cand.frequency > 0 { ">" } else { " " };
                let mut row = format!(
                    "{}{:<width$}  {:>6.1}%  {:>6.1}%",
                    marker,
                    cand.form,
                    obs_pct,
                    pred_pct,
                    width = max_cand_width
                );
                for c_idx in 0..nc {
                    let col_width = tableau.constraints[c_idx].abbrev().len().max(3);
                    let v = cand.violations[c_idx];
                    if v == 0 {
                        row.push_str(&format!("  {:>width$}", "", width = col_width));
                    } else {
                        row.push_str(&format!("  {:>width$}", v, width = col_width));
                    }
                }
                out.push_str(&row);
                out.push('\n');
            }
            out.push('\n');
        }

        out
    }
}

impl Tableau {
    /// Run the MaxEnt algorithm with Generalized Iterative Scaling.
    ///
    /// All candidates (winners and losers) participate in the softmax.
    /// Faithfulness/Markedness distinction is not used — MaxEnt treats all
    /// constraints uniformly.
    ///
    /// Reproduces VB6 MyMaxEnt.frm:MaxEntCore / CalculateExpectedViolations.
    pub fn run_maxent(&self, iterations: usize, weight_min: f64, weight_max: f64) -> MaxEntResult {
        let nc = self.constraints.len();

        // Initialize all weights to 0
        let mut weights = vec![0.0f64; nc];

        // Pre-calculate total frequency per input form
        let total_freq_per_form: Vec<f64> = self.forms.iter()
            .map(|form| form.candidates.iter().map(|c| c.frequency as f64).sum())
            .collect();

        // Calculate observed violations: sum over all candidates of (freq × violations)
        let mut observed = vec![0.0f64; nc];
        for form in &self.forms {
            for cand in &form.candidates {
                let freq = cand.frequency as f64;
                for (c_idx, obs) in observed.iter_mut().enumerate() {
                    *obs += freq * cand.violations[c_idx] as f64;
                }
            }
        }
        // Replace zero observed violations with epsilon to avoid log(0)
        for obs in &mut observed {
            if *obs == 0.0 {
                *obs = 1e-9;
            }
        }

        // Calculate slowing factor = max total violations across any single candidate
        // (Required for GIS convergence; Goodman 2002)
        let mut slowing_factor = 1.0f64;
        for form in &self.forms {
            for cand in &form.candidates {
                let total: f64 = cand.violations.iter().map(|&v| v as f64).sum();
                if total > slowing_factor {
                    slowing_factor = total;
                }
            }
        }

        crate::ot_log!("Starting MaxEnt with {} constraints, {} forms, {} iterations (weights: {}..{})",
            nc, self.forms.len(), iterations, weight_min, weight_max);

        // GIS main loop
        for _ in 0..iterations {
            // Step 1: Calculate predicted proportions for each form
            let predicted = self.calculate_predicted_probs(&weights);

            // Step 2: Calculate expected violations
            let mut expected = vec![0.0f64; nc];
            for (form_idx, form) in self.forms.iter().enumerate() {
                let total_freq = total_freq_per_form[form_idx];
                for (cand_idx, cand) in form.candidates.iter().enumerate() {
                    let pred_prob = predicted[form_idx][cand_idx];
                    for (c_idx, exp) in expected.iter_mut().enumerate() {
                        *exp += total_freq * pred_prob * cand.violations[c_idx] as f64;
                    }
                }
            }

            // Step 3: Update weights (GIS without prior)
            for c_idx in 0..nc {
                if expected[c_idx] <= 0.0 {
                    // Skip — cannot compute log ratio
                    continue;
                }
                let delta = (observed[c_idx] / expected[c_idx]).ln() / slowing_factor;
                weights[c_idx] = (weights[c_idx] - delta).clamp(weight_min, weight_max);
            }
        }

        // Calculate final predicted probabilities
        let predicted_probs = self.calculate_predicted_probs(&weights);

        // Calculate log probability of data
        let log_prob = self.calculate_log_prob(&weights, &predicted_probs);
        crate::ot_log!("MaxEnt DONE: log_prob = {:.6} after {} iterations", log_prob, iterations);

        MaxEntResult {
            weights,
            predicted_probs,
            log_prob,
            iterations,
            weight_min,
            weight_max,
        }
    }

    /// Calculate softmax predicted probabilities for all forms/candidates.
    fn calculate_predicted_probs(&self, weights: &[f64]) -> Vec<Vec<f64>> {
        let nc = weights.len();
        self.forms.iter().map(|form| {
            let e_harmonies: Vec<f64> = form.candidates.iter().map(|cand| {
                let harmony: f64 = (0..nc)
                    .map(|c| weights[c] * cand.violations[c] as f64)
                    .sum();
                let eh = (-harmony).exp();
                // Treat NaN/overflow as 0
                if eh.is_finite() { eh } else { 0.0 }
            }).collect();

            let z: f64 = e_harmonies.iter().sum();
            let z = if z > 0.0 { z } else { 1e-300 };
            e_harmonies.iter().map(|&eh| eh / z).collect()
        }).collect()
    }

    /// Calculate log probability of the training data.
    fn calculate_log_prob(&self, _weights: &[f64], predicted_probs: &[Vec<f64>]) -> f64 {
        let mut log_prob = 0.0f64;
        for (form_idx, form) in self.forms.iter().enumerate() {
            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                if cand.frequency > 0 {
                    let prob = predicted_probs[form_idx][cand_idx];
                    if prob > 0.0 {
                        log_prob += cand.frequency as f64 * prob.ln();
                    }
                }
            }
        }
        log_prob
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
    fn test_maxent_tiny_example() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_maxent(100, 0.0, 50.0);

        // After 100 iterations, markedness weights should be higher than faithfulness
        // (the data has only markedness violations causing losers)
        let w_noons = result.get_weight(0); // *No Onset
        let w_coda = result.get_weight(1);  // *Coda
        let w_max = result.get_weight(2);   // Max
        let w_dep = result.get_weight(3);   // Dep

        println!("*NoOns: {:.4}, *Coda: {:.4}, Max: {:.4}, Dep: {:.4}", w_noons, w_coda, w_max, w_dep);
        println!("Log prob: {:.4}", result.log_prob());

        // Markedness constraints should have nonzero weight (they're violated by losers)
        assert!(w_noons > 0.0 || w_coda > 0.0, "At least one markedness constraint should have weight > 0");

        // Log prob should be finite and negative
        assert!(result.log_prob().is_finite(), "Log prob should be finite");
        assert!(result.log_prob() < 0.0, "Log prob should be negative");
    }

    #[test]
    fn test_maxent_predicted_probs_sum_to_one() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_maxent(10, 0.0, 50.0);

        // For each input form, predicted probabilities should sum to 1
        let form_count = tableau.form_count();
        for form_idx in 0..form_count {
            let form = tableau.get_form(form_idx).unwrap();
            let total: f64 = (0..form.candidate_count())
                .map(|c_idx| result.get_predicted_prob(form_idx, c_idx))
                .sum();
            assert!(
                (total - 1.0).abs() < 1e-9,
                "Predicted probs for form {} should sum to 1, got {}",
                form_idx,
                total
            );
        }
    }

    #[test]
    fn test_maxent_output_format() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_maxent(5, 0.0, 50.0);
        let output = result.format_output(&tableau, "test.txt");

        assert!(output.contains("Results of Applying Maximum Entropy to test.txt"));
        assert!(output.contains("1. Constraint Weights"));
        assert!(output.contains("2. Tableaux"));
        assert!(output.contains("Log probability of data:"));
        assert!(output.contains("/a/:"));
        assert!(output.contains("/tat/:"));
        assert!(output.contains("/at/:"));
    }

    #[test]
    fn test_maxent_weight_bounds() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_maxent(100, 1.0, 10.0);

        let nc = tableau.constraint_count();
        for c_idx in 0..nc {
            let w = result.get_weight(c_idx);
            assert!(w >= 1.0, "Weight {} should be >= min 1.0, got {}", c_idx, w);
            assert!(w <= 10.0, "Weight {} should be <= max 10.0, got {}", c_idx, w);
        }
    }
}
