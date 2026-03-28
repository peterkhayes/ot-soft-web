//! Conformance tests: compare Rust output against VB6 OTSoft golden files.
//!
//! Golden files are collected by running VB6 OTSoft on Windows — see
//! `conformance/CLAUDE.md` for instructions.
//!
//! Tests skip gracefully when golden files are missing, so `cargo test` always
//! passes even before any golden files have been collected.
//!
//! ## HTML conformance
//!
//! HTML test cases (format="html") compare the *semantic content* of tableaux
//! rather than byte-level HTML. Both VB6 and Rust produce structurally different
//! HTML (VB6 uses `<p class="test cl8">` inside `<TD>`, Rust uses `class="cl8"`
//! on `<td>` directly), so the comparison extracts cell grids — each cell as
//! (text content, CSS shading class) — and compares those.

use ot_soft::{FredOptions, FtOptions, GlaOptions, MaxEntOptions, NhgOptions};
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
    apriori_file: Option<String>,
    algorithm: String,
    #[serde(default)]
    format: OutputFormat,
    params: serde_json::Value,
    golden_file: String,
    /// Temporarily skip this case (with a reason string for documentation).
    #[serde(default)]
    skip: Option<String>,
    /// Section headers to strip before comparison (known VB6 divergences).
    /// Each entry is a substring matched against section headers like "4. Status of...".
    /// The entire section (up to the next numbered section or EOF) is removed from both outputs.
    #[serde(default)]
    ignore_sections: Vec<String>,
}

#[derive(Debug, Deserialize, Default, PartialEq)]
#[serde(rename_all = "lowercase")]
enum OutputFormat {
    #[default]
    Text,
    Html,
}

// ── Helpers ─────────────────────────────────────────────────────────────────

/// Repository root (two levels up from rust/tests/).
fn repo_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("rust/ should be inside repo root")
        .to_path_buf()
}

/// Normalize output text for comparison.
///
/// These normalizations absorb known differences between VB6 and Rust output
/// that don't reflect logical errors. Some (like date/version stripping) are
/// permanent; others (like whitespace collapsing) can be tightened over time.
fn normalize(text: &str) -> String {
    let text = text.replace('\r', "");

    // Strip date/time lines:
    //   VB6:  "2-16-2026, 8:39 p.m."
    //   Rust: "2-24-2026, 1:53 pm"
    let date_re = Regex::new(r"(?m)^\d{1,2}-\d{1,2}-\d{4},\s+\d{1,2}:\d{2}\s+[ap]\.?m\.?\s*$")
        .unwrap();
    let text = date_re.replace_all(&text, "<DATE>");

    // Strip OTSoft version lines like "OTSoft 2.7, release date 2/1/2026"
    // or "OTSoft version 2.7, release date 2/1/2026" (FT output uses "version" prefix).
    let version_re = Regex::new(r"(?m)^OTSoft\s+(?:version\s+)?\d+\.\d+.*$").unwrap();
    let text = version_re.replace_all(&text, "<VERSION>");

    // Normalize the broken-bar separator: VB6 uses ¦ (U+00A6) or its
    // Latin-1 lossy replacement (U+FFFD), Rust uses the same ¦. Normalize
    // any remaining replacement chars to ¦.
    let text = text.replace('\u{FFFD}', "\u{00A6}");

    // Treat tabs as spaces — VB6 and Rust may differ in tab vs space usage
    let text = text.replace('\t', " ");

    // Collapse runs of spaces into a single space (per line, after trimming).
    // This absorbs column-alignment differences between VB6 and Rust.
    let space_re = Regex::new(r" {2,}").unwrap();

    // Round floating-point numbers to 4 decimal places to absorb minor
    // precision differences between VB6 and Rust FP arithmetic.
    let float_re = Regex::new(r"\d+\.\d{5,}").unwrap();

    let text: String = text
        .lines()
        .map(|line| {
            let line = line.trim_end();
            let line = space_re.replace_all(line, " ");
            // Round long floats to 4 decimal places
            float_re
                .replace_all(&line, |caps: &regex::Captures| {
                    let s = &caps[0];
                    match s.parse::<f64>() {
                        Ok(v) => format!("{v:.4}"),
                        Err(_) => s.to_string(),
                    }
                })
                .into_owned()
        })
        .collect::<Vec<_>>()
        .join("\n");

    // Strip VB6-only "For a tabbed listing of the t-order found here, see the file"
    // lines (plus the file-path line that follows). Rust does not emit these.
    let tabbed_re = Regex::new(r"(?m)^For a tabbed listing[^\n]*\n[^\n]*\n").unwrap();
    let text = tabbed_re.replace_all(&text, "");

    // Normalize learning time lines: VB6 and Rust will have different timings.
    let learning_time_re = Regex::new(r"(?m)^Learning time:\s+\S+\s+minutes$").unwrap();
    let text = learning_time_re.replace_all(&text, "Learning time: <TIME> minutes");

    // Collapse runs of 3+ blank lines into exactly 2 blank lines.
    // VB6 and Rust may differ in spacing between sections.
    let blank_re = Regex::new(r"\n{4,}").unwrap();
    blank_re.replace_all(&text, "\n\n\n").into_owned()
}

/// Strip entire numbered sections whose headers contain any of the given substrings.
///
/// Sections are delimited by lines matching `\d+\. ` (e.g. "4. Status of...").
/// Everything from a matching header up to (but not including) the next section
/// header (or EOF) is removed.
fn strip_sections(text: &str, ignore: &[String]) -> String {
    if ignore.is_empty() {
        return text.to_string();
    }

    let section_header = Regex::new(r"^\d+\.\s").unwrap();
    let mut result = Vec::new();
    let mut skipping = false;

    for line in text.lines() {
        if section_header.is_match(line) {
            skipping = ignore.iter().any(|pat| line.contains(pat.as_str()));
        }
        if !skipping {
            result.push(line);
        }
    }

    result.join("\n")
}

// ── HTML cell-grid extraction ───────────────────────────────────────────────

/// A single cell in an extracted HTML tableau grid.
#[derive(Debug, PartialEq)]
struct HtmlCell {
    /// Text content with HTML tags and entities decoded, trimmed.
    text: String,
    /// CSS shading class (cl4, cl8, cl9, cl10), if any.
    class: Option<String>,
}

/// A tableau is a grid of rows × cells.
type TableGrid = Vec<Vec<HtmlCell>>;

/// Extract table grids from HTML.
///
/// Splits on `</table>` boundaries, so it works with both proper
/// `<table>...</table>` wrappers (Rust HTML) and VB6's bare `<tr>/<td>`
/// sequences that lack opening `<table>` tags but do have closing ones.
/// Classes are detected on both `<td>`/`<th>` attributes (Rust) and
/// nested `<p>` elements (VB6).
fn extract_html_tables(html: &str) -> Vec<TableGrid> {
    let html = html.replace('\r', "");
    let table_end_re = Regex::new(r"(?i)</table>").unwrap();
    let row_re = Regex::new(r"(?is)<tr[^>]*>(.*?)</tr>").unwrap();
    // Split cells on <td or <th openings rather than requiring matched
    // open/close tags — VB6 HTML often omits </td> closing tags.
    let cell_start_re = Regex::new(r"(?i)<(?:td|th)\b").unwrap();
    let class_re = Regex::new(r#"class="[^"]*\b(cl(?:4|8|9|10))\b[^"]*""#).unwrap();
    let tag_re = Regex::new(r"<[^>]+>").unwrap();

    let mut tables = Vec::new();

    for segment in table_end_re.split(&html) {
        let mut rows = Vec::new();

        for row_cap in row_re.captures_iter(segment) {
            let row_html = &row_cap[1];
            let mut cells = Vec::new();

            // Find positions of all <td / <th openings
            let starts: Vec<usize> = cell_start_re.find_iter(row_html)
                .map(|m| m.start())
                .collect();

            for (i, &start) in starts.iter().enumerate() {
                let end = starts.get(i + 1).copied().unwrap_or(row_html.len());
                let cell_html = &row_html[start..end];

                // Look for cl4/cl8/cl9/cl10 anywhere in the cell region
                let class = class_re
                    .captures(cell_html)
                    .map(|c| c[1].to_string());

                // Strip the opening tag, closing tag, and any nested tags
                let text = tag_re.replace_all(cell_html, "");
                let text = decode_html_entities(&text);
                let text = text.split_whitespace().collect::<Vec<_>>().join(" ");

                cells.push(HtmlCell { text, class });
            }

            if !cells.is_empty() {
                rows.push(cells);
            }
        }

        if !rows.is_empty() {
            tables.push(rows);
        }
    }

    tables
}

/// Decode the most common HTML entities to plain text.
fn decode_html_entities(s: &str) -> String {
    s.replace("&nbsp;", " ")
     .replace("&nbsp", " ")  // VB6 sometimes omits the trailing semicolon
     .replace("&amp;", "&")
     .replace("&lt;", "<")
     .replace("&gt;", ">")
     .replace("&quot;", "\"")
     .replace("&#x261E;", "☞")
     .replace("&#9758;", "☞")
}

/// Normalize a cell's CSS class for comparison.
///
/// VB6 inconsistently applies classes — it omits cl10/cl9 on cells that
/// have default styling, but Rust marks all cells explicitly. The
/// semantically meaningful distinction is the stratum border:
///   - cl4, cl8 → has stratum border (right)
///   - cl9, cl10, None → no border
///
/// We normalize to just "border" or None for comparison.
fn normalize_class(class: &Option<String>) -> Option<&'static str> {
    match class.as_deref() {
        Some("cl4") | Some("cl8") => Some("border"),
        _ => None, // cl9, cl10, None are all equivalent
    }
}

/// Compare two sets of HTML tableaux, returning a description of the first
/// difference found, or `None` if they match.
fn compare_html_tables(
    expected: &[TableGrid],
    actual: &[TableGrid],
    case_id: &str,
) -> Option<String> {
    if expected.len() != actual.len() {
        return Some(format!(
            "{case_id}: table count mismatch — expected {}, actual {}",
            expected.len(),
            actual.len(),
        ));
    }

    for (t_idx, (exp_table, act_table)) in expected.iter().zip(actual.iter()).enumerate() {
        if exp_table.len() != act_table.len() {
            return Some(format!(
                "{case_id}: table {t_idx} row count — expected {}, actual {}",
                exp_table.len(),
                act_table.len(),
            ));
        }

        for (r_idx, (exp_row, act_row)) in exp_table.iter().zip(act_table.iter()).enumerate() {
            if exp_row.len() != act_row.len() {
                return Some(format!(
                    "{case_id}: table {t_idx} row {r_idx} cell count — expected {}, actual {}",
                    exp_row.len(),
                    act_row.len(),
                ));
            }

            for (c_idx, (exp_cell, act_cell)) in exp_row.iter().zip(act_row.iter()).enumerate() {
                if normalize_class(&exp_cell.class) != normalize_class(&act_cell.class) {
                    return Some(format!(
                        "{case_id}: table {t_idx} row {r_idx} cell {c_idx} class mismatch\n\
                         \x20 expected: {:?} (text: {:?})\n\
                         \x20 actual:   {:?} (text: {:?})",
                        exp_cell.class, exp_cell.text,
                        act_cell.class, act_cell.text,
                    ));
                }
                // Normalize trailing colons/spaces: VB6 writes "/a/: " while
                // Rust writes "/a/" in form header cells.
                let exp_text = exp_cell.text.trim_end_matches(':').trim();
                let act_text = act_cell.text.trim_end_matches(':').trim();
                if exp_text != act_text {
                    return Some(format!(
                        "{case_id}: table {t_idx} row {r_idx} cell {c_idx} text mismatch\n\
                         \x20 expected: {:?}\n\
                         \x20 actual:   {:?}",
                        exp_cell.text, act_cell.text,
                    ));
                }
            }
        }
    }

    None
}

/// Compare normalized text outputs and return a diff description if they differ.
fn text_diff(expected: &str, actual: &str, case_id: &str) -> Option<String> {
    if expected == actual {
        return None;
    }

    let expected_lines: Vec<&str> = expected.lines().collect();
    let actual_lines: Vec<&str> = actual.lines().collect();
    let mut diff = format!(
        "{case_id}: output mismatch ({} expected lines, {} actual lines)\n",
        expected_lines.len(),
        actual_lines.len()
    );
    // Show first differing line
    for (i, (e, a)) in expected_lines.iter().zip(actual_lines.iter()).enumerate() {
        if e != a {
            diff.push_str(&format!("  first diff at line {}:\n", i + 1));
            diff.push_str(&format!("  expected: {e:?}\n"));
            diff.push_str(&format!("  actual:   {a:?}\n"));
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
    Some(diff)
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

    let filename = Path::new(&case.input_file)
        .file_name()
        .expect("input_file should have a filename")
        .to_str()
        .expect("input_file should be valid UTF-8");

    match (case.algorithm.as_str(), &case.format) {
        ("rcd", OutputFormat::Text) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_rcd_output(&input_text, filename, &apriori_text, &fred_opts)
        }
        ("rcd", OutputFormat::Html) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_rcd_html_output(&input_text, filename, &apriori_text, &fred_opts, ot_soft::AxisMode::NeverSwitch)
        }
        ("bcd", OutputFormat::Text) => {
            let specific = case.params["specific"].as_bool().unwrap_or(false);
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_bcd_output(&input_text, filename, specific, &fred_opts)
        }
        ("bcd", OutputFormat::Html) => {
            let specific = case.params["specific"].as_bool().unwrap_or(false);
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_bcd_html_output(&input_text, filename, specific, &fred_opts, ot_soft::AxisMode::NeverSwitch)
        }
        ("lfcd", OutputFormat::Text) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_lfcd_output(&input_text, filename, &apriori_text, &fred_opts)
        }
        ("lfcd", OutputFormat::Html) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_lfcd_html_output(&input_text, filename, &apriori_text, &fred_opts, ot_soft::AxisMode::NeverSwitch)
        }
        ("maxent", OutputFormat::Text) => {
            let opts = build_maxent_options(&case.params);
            ot_soft::format_maxent_output(&input_text, filename, &opts)
        }
        ("factorial_typology", OutputFormat::Text) => {
            let opts = build_ft_options(&case.params);
            ot_soft::format_factorial_typology_output(
                &input_text,
                filename,
                &apriori_text,
                &opts,
            )
        }
        (alg, fmt) => Err(format!("unsupported algorithm/format: {alg}/{fmt:?}")),
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
    if let Some(v) = params.get("diagnostics").and_then(|v| v.as_bool()) {
        opts.diagnostics = v;
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
        if let Some(reason) = &case.skip {
            eprintln!("conformance: [SKIP] {} — {reason}", case.id);
            skipped += 1;
            continue;
        }

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

        let diff_msg = if case.format == OutputFormat::Html {
            // HTML conformance: compare extracted cell grids semantically.
            let expected_tables = extract_html_tables(&golden_text);
            let actual_tables = extract_html_tables(&rust_output);
            compare_html_tables(&expected_tables, &actual_tables, &case.id)
        } else {
            // Text conformance: normalized byte comparison
            let expected = strip_sections(&normalize(&golden_text), &case.ignore_sections);
            let actual = strip_sections(&normalize(&rust_output), &case.ignore_sections);
            text_diff(&expected, &actual, &case.id)
        };

        if let Some(diff) = diff_msg {
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

// ── GLA structural conformance ───────────────────────────────────────────────
//
// GLA (Stochastic OT and online MaxEnt) is non-deterministic, so exact byte
// comparison against a golden file isn't possible.  Instead we verify that
// format_gla_output produces the correct structure: section headings present,
// mode name in header, and that SOT-only vs MaxEnt-only sections are correctly
// included or excluded.

/// Assert expected structure for a Stochastic OT formatted output.
fn assert_gla_sot_structure(output: &str, label: &str) {
    assert!(
        output.contains("GLA-Stochastic OT"),
        "{label}: missing 'GLA-Stochastic OT' in header"
    );
    assert!(output.contains("OTSoft 2.7"), "{label}: missing version line");
    assert!(
        output.contains("1. Ranking Values Found"),
        "{label}: missing section 1 (Ranking Values Found)"
    );
    assert!(
        output.contains("2. Matchup to Input Frequencies"),
        "{label}: missing section 2 (Matchup to Input Frequencies)"
    );
    assert!(
        output.contains("3. Testing the Grammar: Details"),
        "{label}: missing section 3 (Testing the Grammar: Details)"
    );
    assert!(
        output.contains("4. Ranking Values (sorted)"),
        "{label}: missing section 4 (Ranking Values sorted)"
    );
    assert!(
        output.contains("5. Ranking Value to Ranking Probability Conversion"),
        "{label}: missing section 5 (Pairwise Ranking Probabilities)"
    );
    assert!(
        output.contains("Log likelihood of data:"),
        "{label}: missing log likelihood line"
    );
    // MaxEnt-only sections must NOT appear
    assert!(
        !output.contains("1. Weights Found"),
        "{label}: SOT output should not contain 'Weights Found'"
    );
}

/// Assert expected structure for a GLA-MaxEnt formatted output.
fn assert_gla_maxent_structure(output: &str, label: &str) {
    assert!(
        output.contains("GLA-MaxEnt"),
        "{label}: missing 'GLA-MaxEnt' in header"
    );
    assert!(output.contains("OTSoft 2.7"), "{label}: missing version line");
    assert!(
        output.contains("1. Weights Found"),
        "{label}: missing section 1 (Weights Found)"
    );
    assert!(
        output.contains("2. Matchup to Input Frequencies"),
        "{label}: missing section 2 (Matchup to Input Frequencies)"
    );
    assert!(
        output.contains("3. Weights (sorted)"),
        "{label}: missing section 3 (Weights sorted)"
    );
    assert!(
        output.contains("Log likelihood of data:"),
        "{label}: missing log likelihood line"
    );
    // SOT-only sections must NOT appear in MaxEnt output
    assert!(
        !output.contains("Testing the Grammar: Details"),
        "{label}: MaxEnt output should not contain 'Testing the Grammar: Details'"
    );
    assert!(
        !output.contains("Ranking Value to Ranking Probability Conversion"),
        "{label}: MaxEnt output should not contain pairwise probability section"
    );
    assert!(
        !output.contains("1. Ranking Values Found"),
        "{label}: MaxEnt output should not contain 'Ranking Values Found'"
    );
}

#[test]
fn gla_structural_conformance() {
    let root = repo_root();

    // Both example files are tested; skip gracefully if an input file is missing.
    let cases: &[(&str, &str)] = &[
        (
            "examples/TinyIllustrativeFile.txt",
            "TinyIllustrativeFile.txt",
        ),
        (
            "examples/IlokanoHiatusResolution.txt",
            "IlokanoHiatusResolution.txt",
        ),
    ];

    for &(rel_path, filename) in cases {
        let input_text = match fs::read_to_string(root.join(rel_path)) {
            Ok(t) => t,
            Err(_) => {
                eprintln!("gla_structural: [SKIP] {filename} — input file missing");
                continue;
            }
        };

        // Use a short training run (500 cycles) so the test is fast.

        // Stochastic OT
        let mut sot_opts = GlaOptions::new();
        sot_opts.maxent_mode = false;
        sot_opts.cycles = 500;
        sot_opts.test_trials = 100;
        let sot_output = ot_soft::format_gla_output(&input_text, filename, &sot_opts)
            .unwrap_or_else(|e| panic!("GLA-SOT failed for {filename}: {e}"));
        assert_gla_sot_structure(&sot_output, &format!("{filename}/sot"));

        // MaxEnt
        let mut maxent_opts = GlaOptions::new();
        maxent_opts.maxent_mode = true;
        maxent_opts.cycles = 500;
        maxent_opts.test_trials = 100;
        let maxent_output = ot_soft::format_gla_output(&input_text, filename, &maxent_opts)
            .unwrap_or_else(|e| panic!("GLA-MaxEnt failed for {filename}: {e}"));
        assert_gla_maxent_structure(&maxent_output, &format!("{filename}/maxent"));
    }
}

// ── GLA convergence assertions ───────────────────────────────────────────────
//
// For TinyIllustrativeFile the RCD solution is known: *NoOns and *Coda rank
// strictly above Max and Dep.  After a full default training run, GLA should
// reliably produce ranking values (SOT) or weights (MaxEnt) that respect this
// ordering.  Constraint indices match the column order in the input file:
//   0 = *NoOns, 1 = *Coda, 2 = Max, 3 = Dep

#[test]
fn gla_convergence_tiny_illustrative() {
    let root = repo_root();
    let input_text = match fs::read_to_string(root.join("examples/TinyIllustrativeFile.txt")) {
        Ok(t) => t,
        Err(_) => {
            eprintln!("gla_convergence: [SKIP] TinyIllustrativeFile.txt not found");
            return;
        }
    };

    // Use the default cycle count (1M) and plasticity schedule.  On this tiny
    // 4-constraint / 3-form file a complete run takes well under a second.

    // ── Stochastic OT ──────────────────────────────────────────────────────
    let mut sot_opts = GlaOptions::new();
    sot_opts.maxent_mode = false;
    sot_opts.test_trials = 1000; // enough trials for reliable frequency estimates
    let sot = ot_soft::run_gla(&input_text, &sot_opts)
        .expect("GLA-SOT failed on TinyIllustrativeFile");

    // Constraint indices from TinyIllustrativeFile column order:
    //   0 = *NoOns,  1 = *Coda,  2 = Max,  3 = Dep
    let nosons = sot.get_ranking_value(0);
    let coda = sot.get_ranking_value(1);
    let max = sot.get_ranking_value(2);
    let dep = sot.get_ranking_value(3);

    // RCD stratum 1 (*NoOns, *Coda) must rank strictly above stratum 2 (Max, Dep).
    assert!(nosons > max, "SOT: *NoOns ({nosons:.4}) should rank above Max ({max:.4})");
    assert!(nosons > dep, "SOT: *NoOns ({nosons:.4}) should rank above Dep ({dep:.4})");
    assert!(coda > max,  "SOT: *Coda ({coda:.4}) should rank above Max ({max:.4})");
    assert!(coda > dep,  "SOT: *Coda ({coda:.4}) should rank above Dep ({dep:.4})");

    // Predicted winners: SOT test_probs are sampled — with 1000 trials and a
    // converged grammar the winner's frequency should be overwhelming.
    // Form 0 /a/: ?a=cand 0 wins over a=cand 1
    // Form 1 /tat/: ta=cand 0 wins over tat=cand 1
    // Form 2 /at/: ?a=cand 0 wins over ?at=cand 1, a=cand 2, at=cand 3
    for (form_idx, form_name) in [(0, "/a/"), (1, "/tat/"), (2, "/at/")] {
        let winner_prob = sot.get_test_prob(form_idx, 0);
        assert!(
            winner_prob > 0.5,
            "SOT: {form_name} winner (cand 0) should have predicted freq > 0.5, got {winner_prob:.4}"
        );
    }

    // ── Online MaxEnt ───────────────────────────────────────────────────────
    // MaxEnt weight ordering differs from SOT ranking: the grammar is correct iff
    //   w[Dep]   < w[*NoOns]  — so that ?a beats a  in /a/
    //   w[Max]   < w[*Coda]   — so that ta beats tat in /tat/
    // (The /at/ condition follows from both of the above.)
    //
    // Note: we compare weights directly rather than checking predicted probabilities,
    // because linearly-separable categorical data can drive weights to extreme values
    // that cause exp(-harmony) to underflow to 0 even for the correct winner.
    let mut me_opts = GlaOptions::new();
    me_opts.maxent_mode = true;
    me_opts.test_trials = 0;
    let me = ot_soft::run_gla(&input_text, &me_opts)
        .expect("GLA-MaxEnt failed on TinyIllustrativeFile");

    let nosons = me.get_ranking_value(0); // *NoOns  (index 0 = col order in file)
    let coda   = me.get_ranking_value(1); // *Coda
    let max_w  = me.get_ranking_value(2); // Max
    let dep    = me.get_ranking_value(3); // Dep

    // ?a wins /a/   iff H(?a) = w_Dep   < H(a)  = w_*NoOns
    assert!(
        dep < nosons,
        "MaxEnt: Dep weight ({dep:.4}) should be less than *NoOns ({nosons:.4})\n\
         (all weights: *NoOns={nosons:.4}, *Coda={coda:.4}, Max={max_w:.4}, Dep={dep:.4})"
    );
    // ta  wins /tat/ iff H(ta) = w_Max   < H(tat) = w_*Coda
    assert!(
        max_w < coda,
        "MaxEnt: Max weight ({max_w:.4}) should be less than *Coda ({coda:.4})\n\
         (all weights: *NoOns={nosons:.4}, *Coda={coda:.4}, Max={max_w:.4}, Dep={dep:.4})"
    );
}

// ── GLA structural conformance for AnttilaFinnishGenitivePlurals ─────────────
//
// This file has frequency data and 25 inputs, making it a more realistic test
// than TinyIllustrativeFile.  We verify structural correctness and that the
// grammar converges to produce reasonable predictions.

#[test]
fn gla_structural_anttila() {
    let root = repo_root();
    let input_text = match fs::read_to_string(root.join("examples/AnttilaFinnishGenitivePlurals.txt")) {
        Ok(t) => t,
        Err(_) => {
            eprintln!("gla_structural_anttila: [SKIP] AnttilaFinnishGenitivePlurals.txt not found");
            return;
        }
    };

    let filename = "AnttilaFinnishGenitivePlurals.txt";

    // Stochastic OT — short run for speed
    let mut sot_opts = GlaOptions::new();
    sot_opts.maxent_mode = false;
    sot_opts.cycles = 500;
    sot_opts.test_trials = 100;
    let sot_output = ot_soft::format_gla_output(&input_text, filename, &sot_opts)
        .unwrap_or_else(|e| panic!("GLA-SOT failed for Anttila: {e}"));
    assert_gla_sot_structure(&sot_output, "Anttila/sot");

    // GLA-MaxEnt — short run for speed
    let mut maxent_opts = GlaOptions::new();
    maxent_opts.maxent_mode = true;
    maxent_opts.cycles = 500;
    maxent_opts.test_trials = 100;
    let maxent_output = ot_soft::format_gla_output(&input_text, filename, &maxent_opts)
        .unwrap_or_else(|e| panic!("GLA-MaxEnt failed for Anttila: {e}"));
    assert_gla_maxent_structure(&maxent_output, "Anttila/maxent");
}

// ── NHG structural conformance ──────────────────────────────────────────────
//
// NHG is stochastic, so we verify structural correctness and basic convergence
// properties rather than exact output matching.

fn assert_nhg_structure(output: &str, label: &str) {
    assert!(
        output.contains("Noisy Harmonic Grammar"),
        "{label}: missing 'Noisy Harmonic Grammar' in header"
    );
    assert!(output.contains("OTSoft 2.7"), "{label}: missing version line");
    assert!(
        output.contains("1. Weights Found"),
        "{label}: missing section 1 (Weights Found)"
    );
    assert!(
        output.contains("2. Matchup to Input Frequencies"),
        "{label}: missing section 2 (Matchup to Input Frequencies)"
    );
    assert!(
        output.contains("Log likelihood of data:"),
        "{label}: missing log likelihood line"
    );
}

#[test]
fn nhg_structural_conformance() {
    let root = repo_root();

    let cases: &[(&str, &str)] = &[
        (
            "examples/TinyIllustrativeFile.txt",
            "TinyIllustrativeFile.txt",
        ),
        (
            "examples/IlokanoHiatusResolution.txt",
            "IlokanoHiatusResolution.txt",
        ),
        (
            "examples/AnttilaFinnishGenitivePlurals.txt",
            "AnttilaFinnishGenitivePlurals.txt",
        ),
    ];

    for &(rel_path, filename) in cases {
        let input_text = match fs::read_to_string(root.join(rel_path)) {
            Ok(t) => t,
            Err(_) => {
                eprintln!("nhg_structural: [SKIP] {filename} — input file missing");
                continue;
            }
        };

        let mut opts = NhgOptions::new();
        opts.cycles = 500;
        opts.test_trials = 100;
        let output = ot_soft::format_nhg_output(&input_text, filename, &opts)
            .unwrap_or_else(|e| panic!("NHG failed for {filename}: {e}"));
        assert_nhg_structure(&output, filename);
    }
}

// ── Unit tests for HTML extraction ──────────────────────────────────────────

#[cfg(test)]
mod html_extraction_tests {
    use super::*;

    #[test]
    fn extract_rust_style_html() {
        let html = r#"
        <table>
          <tr>
            <th>/a/</th>
            <th class="cl8">*NoOns</th>
            <th class="cl10">*Coda</th>
          </tr>
          <tr>
            <td>&#x261E;&nbsp;?a</td>
            <td class="cl8">&nbsp;</td>
            <td class="cl9">*</td>
          </tr>
          <tr>
            <td>&nbsp;&nbsp;&nbsp;a</td>
            <td class="cl8">*!</td>
            <td class="cl9">&nbsp;</td>
          </tr>
        </table>
        "#;

        let tables = extract_html_tables(html);
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.len(), 3); // header + winner + loser

        // Header row
        assert_eq!(table[0][0].text, "/a/");
        assert_eq!(table[0][1].text, "*NoOns");
        assert_eq!(table[0][1].class, Some("cl8".into()));
        assert_eq!(table[0][2].class, Some("cl10".into()));

        // Winner row
        assert_eq!(table[1][0].text, "☞ ?a");
        assert_eq!(table[1][1].class, Some("cl8".into()));
        assert_eq!(table[1][2].class, Some("cl9".into()));

        // Loser row
        assert_eq!(table[2][1].text, "*!");
        assert_eq!(table[2][1].class, Some("cl8".into()));
    }

    #[test]
    fn extract_vb6_style_html() {
        // VB6 puts classes on <p> elements nested inside <TD>
        let html = r#"
        <TABLE>
          <TR>
            <TD>/a/: </TD>
            <TD ALIGN=Center>
            <p class="test cl8">
            *NoOns</TD>
            <TD ALIGN=Center>
            *Coda</TD>
          </TR>
          <TR>
            <TD>☞ ?a</TD>
            <TD ALIGN=Center>
            <p class="test cl8">
            &nbsp;</TD>
            <TD ALIGN=Center>
            <p class="test cl9">
            *</TD>
          </TR>
        </TABLE>
        "#;

        let tables = extract_html_tables(html);
        assert_eq!(tables.len(), 1);
        let table = &tables[0];

        // Header: *NoOns cell has cl8
        assert_eq!(table[0][1].text, "*NoOns");
        assert_eq!(table[0][1].class, Some("cl8".into()));
        // *Coda has no class
        assert_eq!(table[0][2].text, "*Coda");
        assert_eq!(table[0][2].class, None);

        // Winner: cl8 on first violation cell, cl9 on second
        assert_eq!(table[1][1].class, Some("cl8".into()));
        assert_eq!(table[1][2].class, Some("cl9".into()));
    }
}
