//! OT-Soft: Optimality Theory constraint ranking
//!
//! This library implements parsing of OT tableaux and the Recursive Constraint
//! Demotion (RCD) algorithm for finding stratified constraint rankings.
//!
//! ## Modules
//!
//! - `tableau`: Data structures and parsing for OT tableaux
//! - `rcd`: Recursive Constraint Demotion algorithm

use wasm_bindgen::prelude::*;

mod tableau;
mod rcd;

// Re-export public types
pub use tableau::{Tableau, Constraint, Candidate, InputForm};
pub use rcd::RCDResult;

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
