//! OT-Soft: Optimality Theory constraint ranking
//!
//! This library implements parsing of OT tableaux and constraint ranking
//! algorithms: Recursive Constraint Demotion (RCD) and Biased Constraint
//! Demotion (BCD).
//!
//! ## Modules
//!
//! - `tableau`: Data structures and parsing for OT tableaux
//! - `rcd`: Recursive Constraint Demotion algorithm
//! - `bcd`: Biased Constraint Demotion algorithm

use wasm_bindgen::prelude::*;

mod tableau;
mod rcd;
mod bcd;
mod maxent;

// Re-export public types
pub use tableau::{Tableau, Constraint, Candidate, InputForm};
pub use rcd::RCDResult;
pub use maxent::MaxEntResult;

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
