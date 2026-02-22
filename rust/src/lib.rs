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

/// Options for the Noisy Harmonic Grammar algorithm.
#[wasm_bindgen]
pub struct NhgOptions {
    pub cycles: usize,
    pub initial_plasticity: f64,
    pub final_plasticity: f64,
    pub test_trials: usize,
    pub noise_by_cell: bool,
    pub post_mult_noise: bool,
    pub noise_for_zero_cells: bool,
    pub late_noise: bool,
    pub exponential_nhg: bool,
    pub demi_gaussians: bool,
    pub negative_weights_ok: bool,
    pub resolve_ties_by_skipping: bool,
}

impl Default for NhgOptions {
    fn default() -> Self { Self::new() }
}

#[wasm_bindgen]
impl NhgOptions {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            cycles: 5000,
            initial_plasticity: 2.0,
            final_plasticity: 0.002,
            test_trials: 2000,
            noise_by_cell: false,
            post_mult_noise: false,
            noise_for_zero_cells: false,
            late_noise: false,
            exponential_nhg: false,
            demi_gaussians: false,
            negative_weights_ok: false,
            resolve_ties_by_skipping: false,
        }
    }
}

/// Options for the Gradual Learning Algorithm.
#[wasm_bindgen]
pub struct GlaOptions {
    pub maxent_mode: bool,
    pub cycles: usize,
    pub initial_plasticity: f64,
    pub final_plasticity: f64,
    pub test_trials: usize,
    pub negative_weights_ok: bool,
}

impl Default for GlaOptions {
    fn default() -> Self { Self::new() }
}

#[wasm_bindgen]
impl GlaOptions {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            maxent_mode: false,
            cycles: 1_000_000,
            initial_plasticity: 2.0,
            final_plasticity: 0.001,
            test_trials: 10000,
            negative_weights_ok: false,
        }
    }
}

/// Options for the Maximum Entropy algorithm.
#[wasm_bindgen]
pub struct MaxEntOptions {
    pub iterations: usize,
    pub weight_min: f64,
    pub weight_max: f64,
}

impl Default for MaxEntOptions {
    fn default() -> Self { Self::new() }
}

#[wasm_bindgen]
impl MaxEntOptions {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            iterations: 100,
            weight_min: 0.0,
            weight_max: 50.0,
        }
    }
}

/// Options controlling factorial typology output.
#[wasm_bindgen]
pub struct FtOptions {
    pub include_full_listing: bool,
}

impl Default for FtOptions {
    fn default() -> Self { Self::new() }
}

#[wasm_bindgen]
impl FtOptions {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self { include_full_listing: false }
    }
}

/// Options controlling the FRed (ranking argumentation) section of output.
/// Used by RCD, BCD, and LFCD format functions.
#[wasm_bindgen]
pub struct FredOptions {
    pub include_fred: bool,
    pub use_mib: bool,
    pub show_details: bool,
    pub include_mini_tableaux: bool,
}

impl Default for FredOptions {
    fn default() -> Self { Self::new() }
}

#[wasm_bindgen]
impl FredOptions {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            include_fred: true,
            use_mib: false,
            show_details: true,
            include_mini_tableaux: true,
        }
    }
}

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
mod factorial_typology;

// Re-export public types
pub use tableau::{Tableau, Constraint, Candidate, InputForm};
pub use rcd::RCDResult;
pub use fred::FRedResult;
pub use maxent::MaxEntResult;
pub use nhg::NhgResult;
pub use gla::GlaResult;
pub use factorial_typology::FactorialTypologyResult;

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
pub fn format_rcd_output(
    text: &str,
    filename: &str,
    apriori_text: &str,
    fred_opts: &FredOptions,
) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let apriori = if apriori_text.trim().is_empty() {
        Vec::new()
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        apriori::parse_apriori(apriori_text, &abbrevs)?
    };
    let mut result = if apriori.is_empty() {
        tableau.run_rcd()
    } else {
        tableau.run_rcd_with_apriori(&apriori)
    };
    result.apply_fred_options(&tableau, &apriori, fred_opts.include_fred, fred_opts.use_mib, fred_opts.show_details, fred_opts.include_mini_tableaux);
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
pub fn run_maxent(text: &str, opts: &MaxEntOptions) -> Result<MaxEntResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_maxent(opts.iterations, opts.weight_min, opts.weight_max))
}

/// Format MaxEnt results as text for download
#[wasm_bindgen]
pub fn format_maxent_output(text: &str, filename: &str, opts: &MaxEntOptions) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_maxent(opts.iterations, opts.weight_min, opts.weight_max);
    Ok(result.format_output(&tableau, filename))
}

/// Run Noisy Harmonic Grammar on a tableau
#[wasm_bindgen]
pub fn run_nhg(text: &str, opts: &NhgOptions) -> Result<NhgResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_nhg(
        opts.cycles, opts.initial_plasticity, opts.final_plasticity, opts.test_trials,
        opts.noise_by_cell, opts.post_mult_noise, opts.noise_for_zero_cells, opts.late_noise,
        opts.exponential_nhg, opts.demi_gaussians, opts.negative_weights_ok, opts.resolve_ties_by_skipping,
    ))
}

/// Format NHG results as text for download
#[wasm_bindgen]
pub fn format_nhg_output(text: &str, filename: &str, opts: &NhgOptions) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_nhg(
        opts.cycles, opts.initial_plasticity, opts.final_plasticity, opts.test_trials,
        opts.noise_by_cell, opts.post_mult_noise, opts.noise_for_zero_cells, opts.late_noise,
        opts.exponential_nhg, opts.demi_gaussians, opts.negative_weights_ok, opts.resolve_ties_by_skipping,
    );
    Ok(result.format_output(&tableau, filename))
}

/// Format BCD results as text for download.
#[wasm_bindgen]
pub fn format_bcd_output(
    text: &str,
    filename: &str,
    specific: bool,
    fred_opts: &FredOptions,
) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let mut result = tableau.run_bcd(specific);
    let algorithm_name = if specific {
        "Biased Constraint Demotion (Specific)"
    } else {
        "Biased Constraint Demotion"
    };
    result.apply_fred_options(&tableau, &[], fred_opts.include_fred, fred_opts.use_mib, fred_opts.show_details, fred_opts.include_mini_tableaux);
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
pub fn format_lfcd_output(
    text: &str,
    filename: &str,
    apriori_text: &str,
    fred_opts: &FredOptions,
) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let apriori = if apriori_text.trim().is_empty() {
        Vec::new()
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        apriori::parse_apriori(apriori_text, &abbrevs)?
    };
    let mut result = if apriori.is_empty() {
        tableau.run_lfcd()
    } else {
        tableau.run_lfcd_with_apriori(&apriori)
    };
    result.apply_fred_options(&tableau, &apriori, fred_opts.include_fred, fred_opts.use_mib, fred_opts.show_details, fred_opts.include_mini_tableaux);
    Ok(result.format_output_with_algorithm(
        &tableau,
        filename,
        "Low Faithfulness Constraint Demotion",
    ))
}

/// Run the Gradual Learning Algorithm (GLA) on a tableau.
#[wasm_bindgen]
pub fn run_gla(text: &str, opts: &GlaOptions) -> Result<GlaResult, String> {
    let tableau = Tableau::parse(text)?;
    Ok(tableau.run_gla(
        opts.maxent_mode, opts.cycles, opts.initial_plasticity, opts.final_plasticity,
        opts.test_trials, opts.negative_weights_ok,
    ))
}

/// Format GLA results as text for download.
#[wasm_bindgen]
pub fn format_gla_output(text: &str, filename: &str, opts: &GlaOptions) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let result = tableau.run_gla(
        opts.maxent_mode, opts.cycles, opts.initial_plasticity, opts.final_plasticity,
        opts.test_trials, opts.negative_weights_ok,
    );
    Ok(result.format_output(&tableau, filename))
}

/// Run factorial typology on a tableau.
/// `apriori_text`: contents of an a priori rankings file, or empty string for none.
#[wasm_bindgen]
pub fn run_factorial_typology(text: &str, apriori_text: &str) -> Result<FactorialTypologyResult, String> {
    let tableau = Tableau::parse(text)?;
    let apriori = if apriori_text.trim().is_empty() {
        Vec::new()
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        apriori::parse_apriori(apriori_text, &abbrevs)?
    };
    Ok(tableau.run_factorial_typology(&apriori))
}

/// Format factorial typology results as text for download.
/// `apriori_text`: contents of an a priori rankings file, or empty string for none.
#[wasm_bindgen]
pub fn format_factorial_typology_output(
    text: &str,
    filename: &str,
    apriori_text: &str,
    opts: &FtOptions,
) -> Result<String, String> {
    let tableau = Tableau::parse(text)?;
    let apriori = if apriori_text.trim().is_empty() {
        Vec::new()
    } else {
        let abbrevs: Vec<String> = tableau.constraints.iter().map(|c| c.abbrev()).collect();
        apriori::parse_apriori(apriori_text, &abbrevs)?
    };
    let result = tableau.run_factorial_typology(&apriori);
    let mut out = result.format_output(&tableau, filename);
    if opts.include_full_listing {
        out.push_str(&result.format_full_listing(&tableau, &apriori));
    }
    Ok(out)
}
