//! A priori constraint rankings
//!
//! Supports parsing of tab-delimited a priori rankings files and enforcement
//! of those rankings within RCD and LFCD.
//!
//! File format (same as VB6 APrioriRankings.bas):
//!
//! ```text
//! [tab]  Abbr1  [tab]  Abbr2  [tab]  ...
//! Abbr1  [tab]  [1?]   [tab]  [1?]   [tab]  ...
//! Abbr2  [tab]  [1?]   [tab]  [1?]   [tab]  ...
//! ...
//! ```
//!
//! A non-empty cell at row i / column j means constraint i a priori dominates
//! constraint j (i >> j).

/// Parse an a priori rankings file.
///
/// Returns `table[i][j] = true` when constraint `i` a priori dominates constraint `j`.
/// Returns an empty `Vec` if `text` is blank (treated as "no a priori rankings").
///
/// Validates:
/// - Column/row labels match `abbrevs` exactly, in order
/// - No constraint dominates itself
/// - No mutual domination (A >> B and B >> A simultaneously)
/// - No circular chains (detected by running a mini-RCD on the table)
///
/// Reproduces VB6 APrioriRankings.bas:ReadAPrioriRankingsAsTable
pub fn parse_apriori(text: &str, abbrevs: &[String]) -> Result<Vec<Vec<bool>>, String> {
    if text.trim().is_empty() {
        return Ok(vec![]);
    }

    let nc = abbrevs.len();
    if nc == 0 {
        return Ok(vec![]);
    }

    // Collect non-empty lines, stripping Windows \r
    let lines: Vec<&str> = text
        .lines()
        .map(|l| l.trim_end_matches('\r'))
        .filter(|l| !l.trim().is_empty())
        .collect();

    if lines.is_empty() {
        return Ok(vec![]);
    }

    if lines.len() < nc + 1 {
        return Err(format!(
            "A priori rankings file has {} line(s) but {} are required ({} constraint rows + 1 header)",
            lines.len(),
            nc + 1,
            nc,
        ));
    }

    // ── Header row ───────────────────────────────────────────────────────────
    // Format: leading tab, then nc constraint abbreviations
    let header: Vec<&str> = lines[0].split('\t').collect();
    for (i, expected) in abbrevs.iter().enumerate() {
        let col = i + 1; // skip leading-tab column
        let actual = if col < header.len() { header[col].trim() } else { "" };
        if actual != expected.as_str() {
            return Err(format!(
                "A priori rankings file: column {} header is '{}', expected constraint abbreviation '{}'",
                col, actual, expected
            ));
        }
    }

    // ── Data rows ─────────────────────────────────────────────────────────────
    let mut table = vec![vec![false; nc]; nc];

    for (i, expected_abbrev) in abbrevs.iter().enumerate() {
        let fields: Vec<&str> = lines[i + 1].split('\t').collect();

        let row_label = fields.first().map(|s| s.trim()).unwrap_or("");
        if row_label != expected_abbrev.as_str() {
            return Err(format!(
                "A priori rankings file: row {} label is '{}', expected constraint abbreviation '{}'",
                i + 2,
                row_label,
                expected_abbrev
            ));
        }

        for (j, cell_val) in table[i].iter_mut().enumerate() {
            let col = j + 1;
            let cell = if col < fields.len() { fields[col].trim() } else { "" };
            if !cell.is_empty() {
                *cell_val = true;
            }
        }
    }

    // ── Validation ────────────────────────────────────────────────────────────

    // No self-domination
    for i in 0..nc {
        if table[i][i] {
            return Err(format!(
                "A priori rankings: constraint '{}' dominates itself, which is impossible",
                abbrevs[i]
            ));
        }
    }

    // No mutual domination
    for i in 0..nc {
        for j in (i + 1)..nc {
            if table[i][j] && table[j][i] {
                return Err(format!(
                    "A priori rankings: '{}' and '{}' mutually dominate each other, which is impossible",
                    abbrevs[i], abbrevs[j]
                ));
            }
        }
    }

    // No circular chains (mini-RCD over the a priori table)
    if has_circular_chain(&table, nc) {
        return Err(
            "A priori rankings contain a circular dominance chain (A >> B >> ... >> A)".to_string(),
        );
    }

    Ok(table)
}

/// Detect circular chains in an a priori table using a simplified RCD.
///
/// Reproduces VB6 APrioriRankings.bas:FindAPrioriContradiction
fn has_circular_chain(table: &[Vec<bool>], nc: usize) -> bool {
    let mut ranked = vec![false; nc];
    let mut remaining = nc;

    loop {
        if remaining == 0 {
            return false; // All ranked — no contradiction
        }

        // Mark constraints demotable if any unranked constraint dominates them
        let mut demotable = vec![false; nc];
        for i in 0..nc {
            if !ranked[i] {
                for j in 0..nc {
                    if !ranked[j] && table[i][j] {
                        demotable[j] = true;
                    }
                }
            }
        }

        // Install non-demotable unranked constraints
        let mut installed_any = false;
        for i in 0..nc {
            if !ranked[i] && !demotable[i] {
                ranked[i] = true;
                remaining -= 1;
                installed_any = true;
            }
        }

        if !installed_any {
            return true; // Stuck — circular chain detected
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn abbrevs(names: &[&str]) -> Vec<String> {
        names.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn test_empty_text_returns_empty() {
        let result = parse_apriori("", &abbrevs(&["A", "B"])).unwrap();
        assert!(result.is_empty());
    }

    #[test]
    fn test_valid_table() {
        // A dominates B: table[0][1] = true
        let text = "\tA\tB\nA\t\t1\nB\t\t\n";
        let result = parse_apriori(text, &abbrevs(&["A", "B"])).unwrap();
        assert_eq!(result.len(), 2);
        assert!(!result[0][0]); // A not >> A
        assert!(result[0][1]);  // A >> B
        assert!(!result[1][0]); // B not >> A
        assert!(!result[1][1]); // B not >> B
    }

    #[test]
    fn test_self_domination_rejected() {
        let text = "\tA\tB\nA\t1\t\nB\t\t\n";
        let result = parse_apriori(text, &abbrevs(&["A", "B"]));
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("itself"));
    }

    #[test]
    fn test_mutual_domination_rejected() {
        let text = "\tA\tB\nA\t\t1\nB\t1\t\n";
        let result = parse_apriori(text, &abbrevs(&["A", "B"]));
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("mutually"));
    }

    #[test]
    fn test_circular_chain_rejected() {
        // A >> B, B >> C, C >> A
        let text = "\tA\tB\tC\nA\t\t1\t\nB\t\t\t1\nC\t1\t\t\n";
        let result = parse_apriori(text, &abbrevs(&["A", "B", "C"]));
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("circular"));
    }

    #[test]
    fn test_wrong_column_header_rejected() {
        let text = "\tA\tX\nA\t\t\nB\t\t\n";
        let result = parse_apriori(text, &abbrevs(&["A", "B"]));
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("'X'"));
    }

    #[test]
    fn test_wrong_row_label_rejected() {
        let text = "\tA\tB\nA\t\t\nX\t\t\n";
        let result = parse_apriori(text, &abbrevs(&["A", "B"]));
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("'X'"));
    }
}
