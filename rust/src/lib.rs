//! OT-Soft: Optimality Theory constraint ranking
//!
//! This library implements parsing of OT tableaux and constraint ranking
//! algorithms: Recursive Constraint Demotion (RCD), Biased Constraint
//! Demotion (BCD), Low Faithfulness Constraint Demotion (LFCD),
//! Maximum Entropy, and Noisy Harmonic Grammar (NHG).
//!
//! ## Modules
//!
//! - `tableau`: Data structures and parsing for OT tableaux
//! - `rcd`: Recursive Constraint Demotion algorithm
//! - `bcd`: Biased Constraint Demotion algorithm
//! - `lfcd`: Low Faithfulness Constraint Demotion algorithm
//! - `maxent`: Batch Maximum Entropy (GIS optimizer)
//! - `nhg`: Noisy Harmonic Grammar (online learner)

use wasm_bindgen::prelude::*;

mod tableau;
mod rcd;
mod bcd;
mod lfcd;
mod maxent;
mod nhg;

// Re-export public types
pub use tableau::{Tableau, Constraint, Candidate, InputForm};
pub use rcd::RCDResult;
pub use maxent::MaxEntResult;
pub use nhg::NhgResult;

/// Initialize the module
#[wasm_bindgen(start)]
pub fn init() {
    // Set up panic hook for better error messages in the browser
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}

/// Parse a tableau file and return the Tableau object for JavaScript to use
#[wasm_bindgen]
pub fn parse_tableau(text: &str) -> Result<Tableau, String> {
    Tableau::parse(text)
}

/// Run RCD on a parsed tableau
#[wasm_bindgen]
pub fn run_rcd(text: &str) -> Result<RCDResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_rcd())
}

/// Format RCD results as text for download
#[wasm_bindgen]
pub fn format_rcd_output(text: &str, filename: &str) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_rcd();
    Ok(result.format_output(&tableau, filename))
}

/// Run BCD on a parsed tableau
#[wasm_bindgen]
pub fn run_bcd(text: &str, specific: bool) -> Result<RCDResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_bcd(specific))
}

/// Run Maximum Entropy on a tableau
#[wasm_bindgen]
pub fn run_maxent(text: &str, iterations: usize, weight_min: f64, weight_max: f64) -> Result<MaxEntResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_maxent(iterations, weight_min, weight_max))
}

/// Format MaxEnt results as text for download
#[wasm_bindgen]
pub fn format_maxent_output(text: &str, filename: &str, iterations: usize, weight_min: f64, weight_max: f64) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_maxent(iterations, weight_min, weight_max);
    Ok(result.format_output(&tableau, filename))
}

/// Run Noisy Harmonic Grammar on a tableau
#[wasm_bindgen]
#[allow(clippy::too_many_arguments)]
pub fn run_nhg(
    text: &str,
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
) -> Result<NhgResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_nhg(
        cycles, initial_plasticity, final_plasticity, test_trials,
        noise_by_cell, post_mult_noise, noise_for_zero_cells, late_noise,
        exponential_nhg, demi_gaussians, negative_weights_ok, resolve_ties_by_skipping,
    ))
}

/// Format NHG results as text for download
#[wasm_bindgen]
#[allow(clippy::too_many_arguments)]
pub fn format_nhg_output(
    text: &str,
    filename: &str,
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
) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_nhg(
        cycles, initial_plasticity, final_plasticity, test_trials,
        noise_by_cell, post_mult_noise, noise_for_zero_cells, late_noise,
        exponential_nhg, demi_gaussians, negative_weights_ok, resolve_ties_by_skipping,
    );
    Ok(result.format_output(&tableau, filename))
}

/// Format BCD results as text for download
#[wasm_bindgen]
pub fn format_bcd_output(text: &str, filename: &str, specific: bool) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_bcd(specific);
    let algorithm_name = if specific {
        "Biased Constraint Demotion (Specific)"
    } else {
        "Biased Constraint Demotion"
    };
    Ok(result.format_output_with_algorithm(&tableau, filename, algorithm_name))
}

/// Run Low Faithfulness Constraint Demotion on a parsed tableau
#[wasm_bindgen]
pub fn run_lfcd(text: &str) -> Result<RCDResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_lfcd())
}

/// Format LFCD results as text for download
#[wasm_bindgen]
pub fn format_lfcd_output(text: &str, filename: &str) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_lfcd();
    Ok(result.format_output_with_algorithm(
        &tableau,
        filename,
        "Low Faithfulness Constraint Demotion",
    ))
}
