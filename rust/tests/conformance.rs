//! Conformance tests: compare Rust output against VB6 OTSoft golden files.
//!
//! Golden files are collected by running VB6 OTSoft on Windows — see
//! `conformance/CHECKLIST.md` for instructions.
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
    #[serde(default)]
    format: OutputFormat,
    params: serde_json::Value,
    golden_file: String,
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

/// Extract all `<table>` elements from HTML as cell grids.
///
/// Works on both VB6 HTML (classes on `<p>` inside `<td>`) and Rust HTML
/// (classes on `<td>`/`<th>` directly). The extraction is intentionally
/// forgiving of malformed HTML.
fn extract_html_tables(html: &str) -> Vec<TableGrid> {
    let html = html.replace('\r', "");
    let table_re = Regex::new(r"(?is)<table[^>]*>(.*?)</table>").unwrap();
    let row_re = Regex::new(r"(?is)<tr[^>]*>(.*?)</tr>").unwrap();
    let cell_re = Regex::new(r"(?is)<(td|th)[^>]*>(.*?)</(?:td|th)>").unwrap();
    let class_re = Regex::new(r#"class="[^"]*\b(cl(?:4|8|9|10))\b[^"]*""#).unwrap();
    let tag_re = Regex::new(r"<[^>]+>").unwrap();

    let mut tables = Vec::new();

    for table_cap in table_re.captures_iter(&html) {
        let table_html = &table_cap[1];
        let mut rows = Vec::new();

        for row_cap in row_re.captures_iter(table_html) {
            let row_html = &row_cap[1];
            let mut cells = Vec::new();

            for cell_cap in cell_re.captures_iter(row_html) {
                let cell_tag_and_content = &cell_cap[0]; // full <td ...>...</td>
                let cell_content = &cell_cap[2];

                // Look for cl4/cl8/cl9/cl10 anywhere in the cell (tag attrs or nested <p>)
                let class = class_re
                    .captures(cell_tag_and_content)
                    .map(|c| c[1].to_string());

                // Strip HTML tags, decode common entities, normalize whitespace
                let text = tag_re.replace_all(cell_content, "");
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
     .replace("&amp;", "&")
     .replace("&lt;", "<")
     .replace("&gt;", ">")
     .replace("&quot;", "\"")
     .replace("&#x261E;", "☞")
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
                if exp_cell.class != act_cell.class {
                    return Some(format!(
                        "{case_id}: table {t_idx} row {r_idx} cell {c_idx} class mismatch\n\
                         \x20 expected: {:?} (text: {:?})\n\
                         \x20 actual:   {:?} (text: {:?})",
                        exp_cell.class, exp_cell.text,
                        act_cell.class, act_cell.text,
                    ));
                }
                if exp_cell.text != act_cell.text {
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

    match (case.algorithm.as_str(), &case.format) {
        ("rcd", OutputFormat::Text) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_rcd_output(&input_text, filename, &apriori_text, &fred_opts)
        }
        ("rcd", OutputFormat::Html) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_rcd_html_output(&input_text, filename, &apriori_text, &fred_opts)
        }
        ("bcd", OutputFormat::Text) => {
            let specific = case.params["specific"].as_bool().unwrap_or(false);
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_bcd_output(&input_text, filename, specific, &fred_opts)
        }
        ("bcd", OutputFormat::Html) => {
            let specific = case.params["specific"].as_bool().unwrap_or(false);
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_bcd_html_output(&input_text, filename, specific, &fred_opts)
        }
        ("lfcd", OutputFormat::Text) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_lfcd_output(&input_text, filename, &apriori_text, &fred_opts)
        }
        ("lfcd", OutputFormat::Html) => {
            let fred_opts = build_fred_options(&case.params);
            ot_soft::format_lfcd_html_output(&input_text, filename, &apriori_text, &fred_opts)
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

        if case.format == OutputFormat::Html {
            // HTML conformance: compare extracted cell grids semantically
            let expected_tables = extract_html_tables(&golden_text);
            let actual_tables = extract_html_tables(&rust_output);

            if let Some(diff) = compare_html_tables(&expected_tables, &actual_tables, &case.id) {
                failures.push(diff);
            }
        } else {
            // Text conformance: normalized byte comparison
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
