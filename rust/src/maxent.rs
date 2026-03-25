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
    /// Gaussian prior parameters (None if no prior)
    use_prior: bool,
    sigma_squared: f64,
    /// Optional history of weights recorded at each GIS iteration.
    history: Option<String>,
    /// Optional history of output probabilities recorded at each GIS iteration.
    output_prob_history: Option<String>,
    /// Learning time in seconds.
    learning_time_secs: f64,
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

    pub fn use_prior(&self) -> bool {
        self.use_prior
    }

    pub fn sigma_squared(&self) -> f64 {
        self.sigma_squared
    }

    pub fn history(&self) -> Option<String> {
        self.history.clone()
    }

    pub fn output_prob_history(&self) -> Option<String> {
        self.output_prob_history.clone()
    }

    /// Format results as text output for download.
    ///
    /// Reproduces the structure of VB6 OTSoft's MaxEnt text output
    /// (DraftOutput.txt), matching the section ordering and formatting of
    /// MyMaxEnt.frm: PrintAHeader, MaxEntCore (sections 1–2), PrintMaxentResults
    /// (section 3), PrintTableaux (section 4), PrintFinalDetails.
    pub fn format_output(&self, tableau: &Tableau, filename: &str, sort_by_weight: bool) -> String {
        let nc = tableau.constraints.len();
        let mut out = String::new();

        // ── Header ──────────────────────────────────────────────────────────
        // Reproduces MyMaxEnt.frm:PrintAHeader → mTmpFile output.
        out.push_str(&format!(
            "Result of Applying Maximum Entropy to {}\n\n\n",
            filename
        ));

        out.push_str(crate::VERSION_STRING);
        out.push_str("\n\n");
        let now = chrono::Local::now();
        out.push_str(&format!(
            "{}\n\n\n",
            now.format("%-m-%-d-%Y, %-I:%M %p")
                .to_string()
                .to_lowercase()
        ));

        // VB6: "For more detailed examination..." note referencing TabbedOutput.txt.
        let file_stem = filename.rsplit_once('.').map_or(filename, |(s, _)| s);
        out.push_str("For more detailed examination of results, please use a spreadsheet program to open the file \n");
        out.push_str(&format!(
            "TabbedOutput.txt, located in the folder FilesFor{}.\n\n\n",
            file_stem
        ));

        // ── Section 1: Constraints and weights ──────────────────────────────
        // Reproduces MaxEntCore → PrintLevel1Header("Constraints and weights").
        // Format: weight right-justified in 8 chars, tab, constraint full name.
        out.push_str("\n1. Constraints and weights\n\n\n");
        for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
            out.push_str(&format!(
                "{:>8}\t{}\n",
                format!("{:.3}", self.weights[c_idx]),
                constraint.full_name()
            ));
        }
        out.push('\n');

        // ── Section 2: Inputs, candidates, frequencies, proportions, probs ──
        // Reproduces MaxEntCore frequency/proportion table via s.PrintTable.
        // Each form produces a sub-table with a header row, a dummy "input" row
        // (VB6 RivalIndex=0), and one row per candidate.
        out.push_str("\n2. Inputs, candidates, input frequencies, input proportions, predicted probabilities\n\n");
        for (form_idx, form) in tableau.forms.iter().enumerate() {
            // Header row
            out.push_str("Inputs  Candidates  Input frequencies  Input proportions  Predicted probabilities\n");

            // VB6 RivalIndex=0 row: input form name, blank candidate, zeros
            out.push_str(&format!(
                "{}  {}  0  0.000  0.000\n",
                form.input, ""
            ));

            // Candidate rows (VB6 RivalIndex=1..N)
            let total_freq: f64 = form.candidates.iter().map(|c| c.frequency as f64).sum();
            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                let obs_prop = if total_freq > 0.0 {
                    cand.frequency as f64 / total_freq
                } else {
                    0.0
                };
                let pred_prob = self.predicted_probs
                    .get(form_idx)
                    .and_then(|f| f.get(cand_idx))
                    .copied()
                    .unwrap_or(0.0);
                out.push_str(&format!(
                    "  {}  {}  {:.3}  {:.3}\n",
                    cand.form, cand.frequency, obs_prop, pred_prob
                ));
            }
            out.push_str("\n\n");
        }

        // Probability of data (VB6: PrintPara with Str(LogProbabilityOfData()))
        out.push_str(&format!("Probability of data = {}\n\n\n", self.log_prob));

        // ── Section 3: Weights Found ─────────────────────────────────────────
        // Reproduces PrintMaxentResults with ThingFound="Weights".
        // Format: weight padded to 10 chars, weight again, 3 spaces, constraint name.
        let constraint_order: Vec<usize> = if sort_by_weight {
            crate::tableau::sorted_indices_descending(&self.weights)
        } else {
            (0..self.weights.len()).collect()
        };

        out.push_str("\n3. Weights Found\n\n");
        for &c_idx in &constraint_order {
            let w = format!("{:.3}", self.weights[c_idx]);
            out.push_str(&format!(
                "{:<10}{w}   {}\n",
                w,
                tableau.constraints[c_idx].full_name()
            ));
        }

        // ── Section 4: Tableaux ──────────────────────────────────────────────
        // Reproduces PrintTableaux: per-form tableaux with harmony, exp(-H),
        // predicted, observed, and constraint violation columns.
        out.push_str("\n4. Tableaux\n\n");
        for (form_idx, form) in tableau.forms.iter().enumerate() {
            let total_freq: f64 = form.candidates.iter().map(|c| c.frequency as f64).sum();

            // Header row
            out.push_str("Input  Candidate  Harmony  exp(-H)  Predicted  Observed");
            for c_idx in 0..nc {
                out.push_str(&format!("  {}", tableau.constraints[c_idx].abbrev()));
            }
            out.push('\n');

            // Weights row (columns 1–6 empty, then weights in constraint columns)
            out.push_str("                                                        ");
            for c_idx in 0..nc {
                let col_w = tableau.constraints[c_idx].abbrev().len().max(5);
                out.push_str(&format!("  {:>width$}", format!("{:.3}", self.weights[c_idx]), width = col_w));
            }
            out.push('\n');

            // Candidate rows
            for (cand_idx, cand) in form.candidates.iter().enumerate() {
                let obs_prop = if total_freq > 0.0 {
                    cand.frequency as f64 / total_freq
                } else {
                    0.0
                };
                let pred_prob = self.predicted_probs
                    .get(form_idx)
                    .and_then(|f| f.get(cand_idx))
                    .copied()
                    .unwrap_or(0.0);

                // Compute harmony = dot product of weights and violations
                let harmony: f64 = (0..nc)
                    .map(|c| self.weights[c] * cand.violations[c] as f64)
                    .sum();
                let e_harmony = (-harmony).exp();

                // VB6: Input column only filled for first candidate (RivalIndex=1).
                // In VB6 v2.7 the loop starts at 1 and the If RivalIndex=0 check
                // never fires, so the Input column is always empty.
                out.push_str(&format!(
                    "  {}  {:.3}  {:.3}  {:.3}  {:.3}",
                    cand.form, harmony, e_harmony, pred_prob, obs_prop
                ));

                // Violation columns as asterisks
                for c_idx in 0..nc {
                    let col_w = tableau.constraints[c_idx].abbrev().len().max(5);
                    let v = cand.violations[c_idx];
                    let stars = if v > 0 {
                        "*".repeat(v as usize)
                    } else {
                        String::new()
                    };
                    out.push_str(&format!("  {:>width$}", stars, width = col_w));
                }
                out.push('\n');
            }
            out.push('\n');
        }

        // ── Learning time ────────────────────────────────────────────────────
        // Reproduces PrintFinalDetails.
        out.push_str(&format!(
            "Learning time:  {:.3} minutes\n\n\n",
            self.learning_time_secs / 60.0
        ));

        out
    }
}

/// Newton's method for the Gaussian prior GIS update.
///
/// Reproduces VB6 MyMaxEnt.frm:DeltaUsingPrior.
///
/// Solves (Goodman 2002, p. 12):
///   0 = expected * exp(delta * slowing_factor) + (weight + delta) / sigma_squared - observed
///
/// Iterates until |function_value| < 1e-5 or up to 100,000 iterations.
fn delta_using_prior(expected: f64, observed: f64, weight: f64, slowing_factor: f64, sigma_squared: f64) -> f64 {
    let mut delta = 0.0f64;
    for _ in 0..100_000 {
        let f = expected * (delta * slowing_factor).exp() + (weight + delta) / sigma_squared - observed;
        let df = expected * slowing_factor * (delta * slowing_factor).exp() + 1.0 / sigma_squared;
        if df == 0.0 {
            break;
        }
        delta -= f / df;
        if f.abs() < 1e-5 {
            break;
        }
    }
    delta
}

impl Tableau {
    /// Run the MaxEnt algorithm with Generalized Iterative Scaling.
    ///
    /// All candidates (winners and losers) participate in the softmax.
    /// Faithfulness/Markedness distinction is not used — MaxEnt treats all
    /// constraints uniformly.
    ///
    /// Reproduces VB6 MyMaxEnt.frm:MaxEntCore / CalculateExpectedViolations.
    ///
    /// When `use_prior` is true, applies Gaussian L2 regularization with
    /// variance `sigma_squared` (Goodman 2002). The Newton's method update
    /// divides by 1000, reproducing the VB6 behavior (including the `/ 1000`
    /// noted as unexplained in the VB6 source).
    pub fn run_maxent(&self, opts: &crate::MaxEntOptions) -> MaxEntResult {
        let iterations = opts.iterations;
        let weight_min = opts.weight_min;
        let weight_max = opts.weight_max;
        let use_prior = opts.use_prior;
        let sigma_squared = opts.sigma_squared;
        let generate_history = opts.generate_history;
        let generate_output_prob_history = opts.generate_output_prob_history;
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

        crate::ot_log!("Starting MaxEnt with {} constraints, {} forms, {} iterations (weights: {}..{}, prior: {})",
            nc, self.forms.len(), iterations, weight_min, weight_max, use_prior);

        let learning_start = chrono::Utc::now();

        // ── History buffer ──────────────────────────────────────────────────
        let mut history_buf = if generate_history {
            use std::fmt::Write;
            // Header: leading tab + constraint abbreviations (matching VB6 HistoryOfWeights.txt)
            let mut header = String::new();
            for c in &self.constraints {
                header.push('\t');
                header.push_str(&c.abbrev());
            }
            header.push('\n');
            // Row 0: initial weights (all zeros)
            write!(header, "0").unwrap();
            for w in &weights {
                write!(header, "\t{w:.4}").unwrap();
            }
            header.push('\n');
            Some(header)
        } else {
            None
        };

        // ── Output probability history buffer ─────────────────────────────
        let mut output_prob_history_buf = if generate_output_prob_history {
            // Header: per-form groups separated by tabs (tab before form if not first)
            // Each group: {input}\t{cand1}\t{cand2}...
            let mut header = String::new();
            for (form_idx, form) in self.forms.iter().enumerate() {
                if form_idx > 0 {
                    header.push('\t');
                }
                header.push_str(&form.input);
                for cand in &form.candidates {
                    header.push('\t');
                    header.push_str(&cand.form);
                }
            }
            header.push('\n');
            Some(header)
        } else {
            None
        };

        // GIS main loop
        for iter_num in 1..=iterations {
            // Step 1: Calculate predicted proportions for each form
            let predicted = self.calculate_predicted_probs(&weights);

            // Record output probability history (before weight update, matching VB6 timing)
            if let Some(ref mut buf) = output_prob_history_buf {
                use std::fmt::Write;
                for (form_idx, form) in self.forms.iter().enumerate() {
                    // VB6 bug: no separator between forms in data rows (header has tab separator)
                    write!(buf, "{iter_num}\t").unwrap();
                    for (cand_idx, _cand) in form.candidates.iter().enumerate() {
                        write!(buf, "\t{}", predicted[form_idx][cand_idx]).unwrap();
                    }
                }
                buf.push('\n');
            }

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

            // Step 3: Update weights
            for c_idx in 0..nc {
                if use_prior {
                    // Gaussian prior: Newton's method (Goodman 2002, p. 12)
                    // Reproduces VB6 DeltaUsingPrior + the unexplained / 1000 scaling.
                    let delta = delta_using_prior(
                        expected[c_idx], observed[c_idx], weights[c_idx],
                        slowing_factor, sigma_squared,
                    );
                    weights[c_idx] = (weights[c_idx] - delta / 1000.0).clamp(weight_min, weight_max);
                } else {
                    // Standard GIS (no prior)
                    if expected[c_idx] <= 0.0 {
                        continue;
                    }
                    let delta = (observed[c_idx] / expected[c_idx]).ln() / slowing_factor;
                    weights[c_idx] = (weights[c_idx] - delta).clamp(weight_min, weight_max);
                }
            }

            // Record history after each iteration
            if let Some(ref mut buf) = history_buf {
                use std::fmt::Write;
                write!(buf, "{iter_num}").unwrap();
                for w in &weights {
                    write!(buf, "\t{w:.4}").unwrap();
                }
                buf.push('\n');
            }
        }

        let learning_time_secs = {
            let elapsed = chrono::Utc::now() - learning_start;
            elapsed.num_milliseconds() as f64 / 1000.0
        };

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
            use_prior,
            sigma_squared,
            history: history_buf,
            output_prob_history: output_prob_history_buf,
            learning_time_secs,
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

// ─── Chunked MaxEnt Runner ────────────────────────────────────────────────────

/// Chunked MaxEnt runner for interactive progress reporting.
///
/// Holds all GIS state so iterations can be split across multiple
/// `run_chunk` calls, yielding control to the browser between chunks.
#[wasm_bindgen]
pub struct MaxEntRunner {
    // ── Immutable config ──────────────────────────────────────────────────────
    tableau: Tableau,
    total_iterations: usize,
    weight_min: f64,
    weight_max: f64,
    use_prior: bool,
    sigma_squared: f64,

    // ── Pre-computed ──────────────────────────────────────────────────────────
    total_freq_per_form: Vec<f64>,
    observed: Vec<f64>,
    slowing_factor: f64,

    // ── Mutable state ─────────────────────────────────────────────────────────
    weights: Vec<f64>,
    current_iter: usize,

    // ── History buffers ───────────────────────────────────────────────────────
    history_buf: Option<String>,
    output_prob_history_buf: Option<String>,

    // ── Final result ───────────────────────────────────────────────────────────
    result: Option<MaxEntResult>,
}

#[wasm_bindgen]
impl MaxEntRunner {
    #[wasm_bindgen(constructor)]
    pub fn new(text: &str, opts: &crate::MaxEntOptions) -> Result<MaxEntRunner, String> {
        let tableau = Tableau::parse(text)?;
        let nc = tableau.constraints.len();

        let total_freq_per_form: Vec<f64> = tableau.forms.iter()
            .map(|form| form.candidates.iter().map(|c| c.frequency as f64).sum())
            .collect();

        let mut observed = vec![0.0f64; nc];
        for form in &tableau.forms {
            for cand in &form.candidates {
                let freq = cand.frequency as f64;
                for (c_idx, obs) in observed.iter_mut().enumerate() {
                    *obs += freq * cand.violations[c_idx] as f64;
                }
            }
        }
        for obs in &mut observed {
            if *obs == 0.0 { *obs = 1e-9; }
        }

        let mut slowing_factor = 1.0f64;
        for form in &tableau.forms {
            for cand in &form.candidates {
                let total: f64 = cand.violations.iter().map(|&v| v as f64).sum();
                if total > slowing_factor { slowing_factor = total; }
            }
        }

        let weights = vec![0.0f64; nc];

        let history_buf = if opts.generate_history {
            use std::fmt::Write;
            let mut header = String::new();
            for c in &tableau.constraints {
                header.push('\t');
                header.push_str(&c.abbrev());
            }
            header.push('\n');
            // Row 0: initial weights (all zeros)
            let mut row = String::from("0");
            for w in &weights {
                write!(row, "\t{w:.4}").unwrap();
            }
            row.push('\n');
            Some(header + &row)
        } else {
            None
        };

        let output_prob_history_buf = if opts.generate_output_prob_history {
            let mut header = String::new();
            for (form_idx, form) in tableau.forms.iter().enumerate() {
                if form_idx > 0 { header.push('\t'); }
                header.push_str(&form.input);
                for cand in &form.candidates {
                    header.push('\t');
                    header.push_str(&cand.form);
                }
            }
            header.push('\n');
            Some(header)
        } else {
            None
        };

        Ok(MaxEntRunner {
            total_iterations: opts.iterations,
            weight_min: opts.weight_min,
            weight_max: opts.weight_max,
            use_prior: opts.use_prior,
            sigma_squared: opts.sigma_squared,
            total_freq_per_form,
            observed,
            slowing_factor,
            weights,
            current_iter: 0,
            history_buf,
            output_prob_history_buf,
            result: None,
            tableau,
        })
    }

    /// Advance up to `max_iters` GIS iterations. Returns true when complete.
    pub fn run_chunk(&mut self, max_iters: usize) -> bool {
        let nc = self.weights.len();
        let mut work_done = 0;

        while work_done < max_iters && self.current_iter < self.total_iterations {
            self.current_iter += 1;
            let iter_num = self.current_iter;

            let predicted = self.tableau.calculate_predicted_probs(&self.weights);

            if let Some(ref mut buf) = self.output_prob_history_buf {
                use std::fmt::Write;
                for (form_idx, form) in self.tableau.forms.iter().enumerate() {
                    write!(buf, "{iter_num}\t").unwrap();
                    for (cand_idx, _) in form.candidates.iter().enumerate() {
                        write!(buf, "\t{}", predicted[form_idx][cand_idx]).unwrap();
                    }
                }
                buf.push('\n');
            }

            let mut expected = vec![0.0f64; nc];
            for (form_idx, form) in self.tableau.forms.iter().enumerate() {
                let total_freq = self.total_freq_per_form[form_idx];
                for (cand_idx, cand) in form.candidates.iter().enumerate() {
                    let pred_prob = predicted[form_idx][cand_idx];
                    for (c_idx, exp) in expected.iter_mut().enumerate() {
                        *exp += total_freq * pred_prob * cand.violations[c_idx] as f64;
                    }
                }
            }

            for (c_idx, w) in self.weights.iter_mut().enumerate() {
                if self.use_prior {
                    let delta = delta_using_prior(
                        expected[c_idx], self.observed[c_idx], *w,
                        self.slowing_factor, self.sigma_squared,
                    );
                    *w = (*w - delta / 1000.0).clamp(self.weight_min, self.weight_max);
                } else {
                    if expected[c_idx] <= 0.0 { continue; }
                    let delta = (self.observed[c_idx] / expected[c_idx]).ln() / self.slowing_factor;
                    *w = (*w - delta).clamp(self.weight_min, self.weight_max);
                }
            }

            if let Some(ref mut buf) = self.history_buf {
                use std::fmt::Write;
                write!(buf, "{iter_num}").unwrap();
                for w in &self.weights {
                    write!(buf, "\t{w:.4}").unwrap();
                }
                buf.push('\n');
            }

            work_done += 1;
        }

        if self.current_iter >= self.total_iterations {
            self.finalize();
            true
        } else {
            false
        }
    }

    /// Progress as [completed, total].
    pub fn progress(&self) -> Vec<f64> {
        vec![self.current_iter as f64, self.total_iterations as f64]
    }

    /// Extract the final MaxEntResult. Only valid after `run_chunk` returns true.
    pub fn take_result(&mut self) -> MaxEntResult {
        self.result.take().expect("MaxEntRunner: take_result called before completion")
    }
}

impl crate::gla::ChunkedRunner for MaxEntRunner {
    fn run_chunk(&mut self, max_work: usize) -> bool {
        MaxEntRunner::run_chunk(self, max_work)
    }
    fn progress(&self) -> [f64; 2] {
        let p = MaxEntRunner::progress(self);
        [p[0], p[1]]
    }
}

impl MaxEntRunner {
    fn finalize(&mut self) {
        let predicted_probs = self.tableau.calculate_predicted_probs(&self.weights);
        let log_prob = self.tableau.calculate_log_prob(&self.weights, &predicted_probs);
        crate::ot_log!("MaxEnt DONE: log_prob = {:.6} after {} iterations", log_prob, self.total_iterations);
        self.result = Some(MaxEntResult {
            weights: self.weights.clone(),
            predicted_probs,
            log_prob,
            iterations: self.total_iterations,
            use_prior: self.use_prior,
            sigma_squared: self.sigma_squared,
            history: self.history_buf.take(),
            output_prob_history: self.output_prob_history_buf.take(),
            learning_time_secs: 0.0, // Not tracked in chunked runner
        });
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
    fn test_maxent_tiny_example() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_maxent(&crate::MaxEntOptions { iterations: 100, ..Default::default() });

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
        let result = tableau.run_maxent(&crate::MaxEntOptions { iterations: 10, ..Default::default() });

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
    fn test_maxent_weight_bounds() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_maxent(&crate::MaxEntOptions {
            iterations: 100, weight_min: 1.0, weight_max: 10.0, ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c_idx in 0..nc {
            let w = result.get_weight(c_idx);
            assert!(w >= 1.0, "Weight {} should be >= min 1.0, got {}", c_idx, w);
            assert!(w <= 10.0, "Weight {} should be <= max 10.0, got {}", c_idx, w);
        }
    }

    #[test]
    fn test_maxent_gaussian_prior_runs() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();
        // With sigma_squared=1.0, the prior term is enabled; results should still be finite
        let result = tableau.run_maxent(&crate::MaxEntOptions {
            iterations: 10, use_prior: true, ..Default::default()
        });

        let nc = tableau.constraint_count();
        for c_idx in 0..nc {
            let w = result.get_weight(c_idx);
            assert!(w.is_finite(), "Weight {} should be finite with prior", c_idx);
            assert!(w >= 0.0, "Weight {} should be >= 0", c_idx);
        }
        assert!(result.log_prob().is_finite(), "Log prob should be finite with prior");
        assert!(result.use_prior(), "use_prior() should return true");
        assert_eq!(result.sigma_squared(), 1.0, "sigma_squared() should return 1.0");
    }

    #[test]
    fn test_maxent_history_generation() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();

        // Without history
        let result = tableau.run_maxent(&crate::MaxEntOptions { iterations: 10, ..Default::default() });
        assert!(result.history().is_none(), "history should be None when not requested");

        // With history
        let result = tableau.run_maxent(&crate::MaxEntOptions {
            iterations: 10, generate_history: true, ..Default::default()
        });
        let history = result.history().expect("history should be Some when generate_history=true");
        let lines: Vec<&str> = history.lines().collect();

        // Header: leading tab + constraint abbreviations
        assert!(lines[0].starts_with('\t'), "header should start with a tab");
        let nc = tableau.constraint_count();
        let header_cols: Vec<&str> = lines[0].split('\t').collect();
        // First col is empty (leading tab), then nc constraint names
        assert_eq!(header_cols.len(), nc + 1, "header should have leading-tab + {} constraints", nc);

        // iterations=10 → row 0 (initial) + 10 iteration rows = 11 data rows + 1 header = 12 lines
        assert_eq!(lines.len(), 12, "should have header + initial + 10 iteration rows");

        // Row 0 should be initial weights (all zeros)
        assert!(lines[1].starts_with("0\t"), "first data row should be iteration 0");

        // Last row should be iteration 10
        assert!(lines[11].starts_with("10\t"), "last data row should be iteration 10");
    }

    #[test]
    fn test_maxent_output_prob_history() {
        let text = load_tiny_example();
        let tableau = Tableau::parse(&text).unwrap();

        // Disabled: should be None
        let result = tableau.run_maxent(&crate::MaxEntOptions { iterations: 10, ..Default::default() });
        assert!(result.output_prob_history().is_none(), "should be None when disabled");

        // Enabled: should have header + 10 data rows
        let result = tableau.run_maxent(&crate::MaxEntOptions {
            iterations: 10, generate_output_prob_history: true, ..Default::default()
        });
        let history = result.output_prob_history().expect("should be Some when enabled");
        let lines: Vec<&str> = history.lines().collect();

        // Header should contain input form names
        let header = lines[0];
        for form in &tableau.forms {
            assert!(header.contains(&form.input), "header should contain input '{}'", form.input);
        }

        // Data row count == iterations
        assert_eq!(lines.len() - 1, 10, "should have exactly 10 data rows (one per iteration)");

        // First data row should start with "1\t" (no initial row, unlike weight history)
        assert!(lines[1].starts_with("1\t"), "first data row should be iteration 1");

        // Last data row should start with "10\t"
        assert!(lines[10].starts_with("10\t"), "last data row should be iteration 10");
    }
}
