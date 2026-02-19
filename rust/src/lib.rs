//! OT-Soft: Optimality Theory constraint ranking
//!
//! This library implements parsing of OT tableaux and constraint ranking
//! algorithms: Recursive Constraint Demotion (RCD), Biased Constraint
//! Demotion (BCD), Low Faithfulness Constraint Demotion (LFCD),
//! Maximum Entropy, Noisy Harmonic Grammar (NHG), and the Gradual
//! Learning Algorithm (GLA) in Stochastic OT and online MaxEnt modes.
//!
//! ## Modules
//!
//! - `tableau`: Data structures and parsing for OT tableaux
//! - `rcd`: Recursive Constraint Demotion algorithm
//! - `bcd`: Biased Constraint Demotion algorithm
//! - `lfcd`: Low Faithfulness Constraint Demotion algorithm
//! - `maxent`: Batch Maximum Entropy (GIS optimizer)
//! - `nhg`: Noisy Harmonic Grammar (online learner)
//! - `gla`: Gradual Learning Algorithm (StochasticOT and online MaxEnt)

use wasm_bindgen::prelude::*;

/// Log a message to the browser console (WASM) or stderr (native).
#[macro_export]
macro_rules! ot_log {
    ($($arg:tt)*) => {
        #[cfg(target_arch = "wasm32")]
        {
            use wasm_bindgen::prelude::*;
            #[wasm_bindgen]
            extern "C" {
                #[wasm_bindgen(js_namespace = console)]
                fn log(s: &str);
            }
            log(&format!($($arg)*));
        }
        #[cfg(not(target_arch = "wasm32"))]
        {
            eprintln!($($arg)*);
        }
    };
}

mod tableau;
mod rcd;
mod bcd;
mod lfcd;
mod fred;
mod apriori;
mod maxent;
mod nhg;
mod gla;

// Re-export public types
pub use tableau::{Tableau, Constraint, Candidate, InputForm};
pub use rcd::RCDResult;
pub use fred::FRedResult;
pub use maxent::MaxEntResult;
pub use nhg::NhgResult;
pub use gla::GlaResult;

/// Initialize the module
#[wasm_bindgen(start)]
pub fn init() {
    // Set up panic hook for better error messages in the browser
    console_error_panic_hook::set_once();
}

/// Parse a tableau file and return the Tableau object for JavaScript to use
#[wasm_bindgen]
pub fn parse_tableau(text: &str) -> Result<Tableau, String> {
    Tableau::parse(text)
}

/// Run RCD on a parsed tableau.
/// `apriori_text`: contents of an a priori rankings file, or empty string for none.
#[wasm_bindgen]
pub fn run_rcd(text: &str, apriori_text: &str) -> Result<RCDResult, String> {
    let tableau = Tableau::parse(text)?;
    if apriori_text.trim().is_empty() {
        Ok(tableau.run_rcd())
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        let apriori = apriori::parse_apriori(apriori_text, &abbrevs)?;
        Ok(tableau.run_rcd_with_apriori(&apriori))
    }
}

/// Format RCD results as text for download.
/// `apriori_text`: contents of an a priori rankings file, or empty string for none.
#[wasm_bindgen]
pub fn format_rcd_output(text: &str, filename: &str, apriori_text: &str) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = if apriori_text.trim().is_empty() {
        tableau.run_rcd()
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        let apriori = apriori::parse_apriori(apriori_text, &abbrevs)?;
        tableau.run_rcd_with_apriori(&apriori)
    };
    Ok(result.format_output(&tableau, filename))
}

/// Run FRed (Fusional Reduction Algorithm) on a tableau.
///
/// `apriori_text`: contents of an a priori rankings file, or empty string for none.
/// `use_mib`: if true, use Most Informative Basis; if false (default), use Skeletal Basis.
#[wasm_bindgen]
pub fn run_fred(text: &str, apriori_text: &str, use_mib: bool) -> Result<FRedResult, String> {
    let tableau = Tableau::parse(text)?;
    if apriori_text.trim().is_empty() {
        Ok(tableau.run_fred(use_mib))
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        let apriori = apriori::parse_apriori(apriori_text, &abbrevs)?;
        Ok(tableau.run_fred_with_apriori(use_mib, &apriori))
    }
}

/// Format FRed results as a standalone text output for download.
#[wasm_bindgen]
pub fn format_fred_output(text: &str, _filename: &str, apriori_text: &str, use_mib: bool) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = if apriori_text.trim().is_empty() {
        tableau.run_fred(use_mib)
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        let apriori = apriori::parse_apriori(apriori_text, &abbrevs)?;
        tableau.run_fred_with_apriori(use_mib, &apriori)
    };
    Ok(result.format_section4())
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

/// Run Low Faithfulness Constraint Demotion on a parsed tableau.
/// `apriori_text`: contents of an a priori rankings file, or empty string for none.
#[wasm_bindgen]
pub fn run_lfcd(text: &str, apriori_text: &str) -> Result<RCDResult, String> {
    let tableau = Tableau::parse(text)?;
    if apriori_text.trim().is_empty() {
        Ok(tableau.run_lfcd())
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        let apriori = apriori::parse_apriori(apriori_text, &abbrevs)?;
        Ok(tableau.run_lfcd_with_apriori(&apriori))
    }
}

/// Format LFCD results as text for download.
/// `apriori_text`: contents of an a priori rankings file, or empty string for none.
#[wasm_bindgen]
pub fn format_lfcd_output(text: &str, filename: &str, apriori_text: &str) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = if apriori_text.trim().is_empty() {
        tableau.run_lfcd()
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        let apriori = apriori::parse_apriori(apriori_text, &abbrevs)?;
        tableau.run_lfcd_with_apriori(&apriori)
    };
    Ok(result.format_output_with_algorithm(
        &tableau,
        filename,
        "Low Faithfulness Constraint Demotion",
    ))
}

/// Run the Gradual Learning Algorithm (GLA) on a tableau.
///
/// `maxent_mode`: if true, run online MaxEnt; if false, run Stochastic OT.
#[wasm_bindgen]
pub fn run_gla(
    text: &str,
    maxent_mode: bool,
    cycles: usize,
    initial_plasticity: f64,
    final_plasticity: f64,
    test_trials: usize,
    negative_weights_ok: bool,
) -> Result<GlaResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_gla(
        maxent_mode, cycles, initial_plasticity, final_plasticity,
        test_trials, negative_weights_ok,
    ))
}

/// Format GLA results as text for download.
#[wasm_bindgen]
pub fn format_gla_output(
    text: &str,
    filename: &str,
    maxent_mode: bool,
    cycles: usize,
    initial_plasticity: f64,
    final_plasticity: f64,
    test_trials: usize,
    negative_weights_ok: bool,
) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_gla(
        maxent_mode, cycles, initial_plasticity, final_plasticity,
        test_trials, negative_weights_ok,
    );
    Ok(result.format_output(&tableau, filename))
}
