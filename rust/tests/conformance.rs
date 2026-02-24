//! Conformance tests: compare Rust output against VB6 OTSoft golden files.
//!
//! Golden files are collected by running VB6 OTSoft on Windows — see
//! `conformance/CHECKLIST.md` for instructions.
//!
//! Tests skip gracefully when golden files are missing, so `cargo test` always
//! passes even before any golden files have been collected.

use ot_soft::{FredOptions, FtOptions, MaxEntOptions};
use regex::Regex;
use serde::Deserialize;
use std::fs;
use std::path::{Path, PathBuf};

// ── Manifest schema ─────────────────────────────────────────────────────────

#[derive(Deserialize)]
struct Manifest {
    cases: Vec<TestCase>,
}

#[derive(Deserialize)]
struct TestCase {
    id: String,
    input_file: String,
    input_display_name: String,
    apriori_file: Option<String>,
    algorithm: String,
    params: serde_json::Value,
    golden_file: String,
}

// ── Helpers ─────────────────────────────────────────────────────────────────

/// Repository root (two levels up from rust/tests/).
fn repo_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("rust/ should be inside repo root")
        .to_path_buf()
}

/// Normalize output text for comparison: strip timestamps, version lines,
/// trailing whitespace, and `\r`.
fn normalize(text: &str) -> String {
    let text = text.replace('\r', "");

    // Strip date/time lines:
    //   VB6:  "2-16-2026, 8:39 p.m."
    //   Rust: "2-24-2026, 1:53 pm"
    let date_re = Regex::new(r"(?m)^\d{1,2}-\d{1,2}-\d{4},\s+\d{1,2}:\d{2}\s+[ap]\.?m\.?\s*$")
        .unwrap();
    let text = date_re.replace_all(&text, "<DATE>");

    // Strip OTSoft version lines like "OTSoft 2.7, release date 2/1/2026"
    let version_re = Regex::new(r"(?m)^OTSoft\s+\d+\.\d+.*$").unwrap();
    let text = version_re.replace_all(&text, "<VERSION>");

    // Normalize the broken-bar separator: VB6 uses ¦ (U+00A6) or its
    // Latin-1 lossy replacement (U+FFFD), Rust uses the same ¦. Normalize
    // any remaining replacement chars to ¦.
    let text = text.replace('\u{FFFD}', "\u{00A6}");

    // Strip trailing whitespace from each line
    text.lines()
        .map(|line| line.trim_end())
        .collect::<Vec<_>>()
        .join("\n")
}

// ── Dispatch: run algorithm and format output ───────────────────────────────

fn run_case(case: &TestCase, root: &Path) -> Result<String, String> {
    let input_path = root.join(&case.input_file);
    let input_text =
        fs::read_to_string(&input_path).map_err(|e| format!("read input: {e}"))?;

    let apriori_text = match &case.apriori_file {
        Some(path) => fs::read_to_string(root.join(path))
            .map_err(|e| format!("read apriori: {e}"))?,
        None => String::new(),
    };

    let filename = &case.input_display_name;

    match case.algorithm.as_str() {
        "rcd" => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_rcd_output(&input_text, filename, &apriori_text, &fred_opts)
        }
        "bcd" => {
            let specific = case.params["specific"].as_bool().unwrap_or(false);
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_bcd_output(&input_text, filename, specific, &fred_opts)
        }
        "lfcd" => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_lfcd_output(&input_text, filename, &apriori_text, &fred_opts)
        }
        "maxent" => {
            let opts = build_maxent_options(&case.params);
            ot_soft::format_maxent_output(&input_text, filename, &opts)
        }
        "factorial_typology" => {
            let opts = build_ft_options(&case.params);
            ot_soft::format_factorial_typology_output(
                &input_text,
                filename,
                &apriori_text,
                &opts,
            )
        }
        other => Err(format!("unknown algorithm: {other}")),
    }
}

fn build_fred_options(params: &serde_json::Value) -> FredOptions {
    let mut opts = FredOptions::new();
    if let Some(v) = params.get("include_fred").and_then(|v| v.as_bool()) {
        opts.include_fred = v;
    }
    if let Some(v) = params.get("use_mib").and_then(|v| v.as_bool()) {
        opts.use_mib = v;
    }
    if let Some(v) = params.get("show_details").and_then(|v| v.as_bool()) {
        opts.show_details = v;
    }
    if let Some(v) = params.get("include_mini_tableaux").and_then(|v| v.as_bool()) {
        opts.include_mini_tableaux = v;
    }
    opts
}

fn build_maxent_options(params: &serde_json::Value) -> MaxEntOptions {
    let mut opts = MaxEntOptions::new();
    if let Some(v) = params.get("iterations").and_then(|v| v.as_u64()) {
        opts.iterations = v as usize;
    }
    if let Some(v) = params.get("weight_min").and_then(|v| v.as_f64()) {
        opts.weight_min = v;
    }
    if let Some(v) = params.get("weight_max").and_then(|v| v.as_f64()) {
        opts.weight_max = v;
    }
    if let Some(v) = params.get("use_prior").and_then(|v| v.as_bool()) {
        opts.use_prior = v;
    }
    if let Some(v) = params.get("sigma_squared").and_then(|v| v.as_f64()) {
        opts.sigma_squared = v;
    }
    opts
}

fn build_ft_options(params: &serde_json::Value) -> FtOptions {
    let mut opts = FtOptions::new();
    if let Some(v) = params.get("include_full_listing").and_then(|v| v.as_bool()) {
        opts.include_full_listing = v;
    }
    opts
}

// ── Test ────────────────────────────────────────────────────────────────────

#[test]
fn conformance_tests() {
    let root = repo_root();
    let manifest_path = root.join("conformance/manifest.json");

    let manifest_text = match fs::read_to_string(&manifest_path) {
        Ok(t) => t,
        Err(_) => {
            eprintln!("conformance: manifest.json not found, skipping all tests");
            return;
        }
    };

    let manifest: Manifest =
        serde_json::from_str(&manifest_text).expect("failed to parse manifest.json");

    let mut ran = 0;
    let mut skipped = 0;
    let mut failures: Vec<String> = Vec::new();

    for case in &manifest.cases {
        let golden_path = root.join(&case.golden_file);
        let golden_bytes = match fs::read(&golden_path) {
            Ok(b) => b,
            Err(_) => {
                eprintln!("conformance: [SKIP] {} — golden file missing", case.id);
                skipped += 1;
                continue;
            }
        };
        // VB6 outputs may be Latin-1 encoded; read lossily
        let golden_text = String::from_utf8_lossy(&golden_bytes);

        let rust_output = match run_case(case, &root) {
            Ok(s) => s,
            Err(e) => {
                failures.push(format!("{}: Rust error: {e}", case.id));
                ran += 1;
                continue;
            }
        };

        let expected = normalize(&golden_text);
        let actual = normalize(&rust_output);

        if expected != actual {
            // Build a useful diff summary
            let expected_lines: Vec<&str> = expected.lines().collect();
            let actual_lines: Vec<&str> = actual.lines().collect();
            let mut diff = format!(
                "{}: output mismatch ({} expected lines, {} actual lines)\n",
                case.id,
                expected_lines.len(),
                actual_lines.len()
            );
            // Show first differing line
            for (i, (e, a)) in expected_lines.iter().zip(actual_lines.iter()).enumerate() {
                if e != a {
                    diff.push_str(&format!("  first diff at line {}:\n", i + 1));
                    diff.push_str(&format!("  expected: {:?}\n", e));
                    diff.push_str(&format!("  actual:   {:?}\n", a));
                    break;
                }
            }
            if expected_lines.len() != actual_lines.len() {
                diff.push_str(&format!(
                    "  line count: expected {}, actual {}\n",
                    expected_lines.len(),
                    actual_lines.len()
                ));
            }
            failures.push(diff);
        }

        ran += 1;
    }

    eprintln!(
        "conformance: {ran} ran, {skipped} skipped, {} failed",
        failures.len()
    );

    if !failures.is_empty() {
        panic!(
            "Conformance test failures:\n\n{}",
            failures.join("\n")
        );
    }
}
