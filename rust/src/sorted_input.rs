//! Sorted input file generation
//!
//! Produces a copy of the input tableau file with constraints reordered by
//! ranking stratum (highest-ranked first) and candidates sorted by harmony
//! (winner first, then rivals in decreasing harmony order).
//!
//! Matches VB6 `PrintTableaux.bas:SaveSortedInputFile`.

use crate::rcd::RCDResult;
use crate::tableau::Tableau;

/// Format a sorted copy of the input file based on an RCD/BCD/LFCD result.
///
/// Constraints are reordered by stratum (ascending stratum = higher rank;
/// within the same stratum, original order is preserved for stability).
/// Candidates (excluding the winner) are sorted by harmony: compared
/// left-to-right through the sorted constraints, fewer violations = more
/// harmonic. The winner always stays first.
///
/// Output is tab-delimited text in the same format as the original input file.
pub fn format_sorted_input(tableau: &Tableau, result: &RCDResult) -> String {
    let sorted_constraint_indices = build_sorted_constraint_indices(tableau, result);
    let mut out = String::new();

    // Header row 1: constraint full names
    out.push_str("\t\t");
    for (i, &ci) in sorted_constraint_indices.iter().enumerate() {
        if i > 0 {
            out.push('\t');
        }
        out.push_str(&tableau.constraints[ci].full_name());
    }
    out.push('\n');

    // Header row 2: constraint abbreviations
    out.push_str("\t\t");
    for (i, &ci) in sorted_constraint_indices.iter().enumerate() {
        if i > 0 {
            out.push('\t');
        }
        out.push_str(&tableau.constraints[ci].abbrev());
    }
    out.push('\n');

    // Data rows: one row per candidate, grouped by input form
    for form in &tableau.forms {
        let sorted_candidates = sort_candidates_by_harmony(form, &sorted_constraint_indices);

        for (cand_idx, &orig_idx) in sorted_candidates.iter().enumerate() {
            let cand = &form.candidates[orig_idx];

            // Column 1: input form (only on first candidate)
            if cand_idx == 0 {
                out.push_str(&form.input);
            }
            out.push('\t');

            // Column 2: candidate form
            out.push_str(&cand.form);
            out.push('\t');

            // Column 3: frequency (blank if 0)
            if cand.frequency > 0 {
                out.push_str(&cand.frequency.to_string());
            }

            // Columns 4+: violations in sorted constraint order (blank if 0)
            for &ci in &sorted_constraint_indices {
                out.push('\t');
                let v = cand.violations[ci];
                if v > 0 {
                    out.push_str(&v.to_string());
                }
            }
            out.push('\n');
        }
    }

    out
}

/// Build constraint indices sorted by stratum (ascending = higher rank).
/// Within the same stratum, preserve original order (stable sort).
fn build_sorted_constraint_indices(tableau: &Tableau, result: &RCDResult) -> Vec<usize> {
    let mut indices: Vec<usize> = (0..tableau.constraints.len()).collect();
    indices.sort_by_key(|&i| result.get_stratum(i).unwrap_or(usize::MAX));
    indices
}

/// Sort candidates within an input form by harmony.
///
/// Returns indices into `form.candidates`. The winner (index 0, the candidate
/// with frequency > 0) stays first; remaining candidates are sorted by
/// lexicographic comparison of violations through the sorted constraints
/// (fewer violations = more harmonic = earlier).
fn sort_candidates_by_harmony(
    form: &crate::tableau::InputForm,
    sorted_constraint_indices: &[usize],
) -> Vec<usize> {
    // Winner is always index 0 in the parsed tableau
    let mut rival_indices: Vec<usize> = (1..form.candidates.len()).collect();

    rival_indices.sort_by(|&a, &b| {
        for &ci in sorted_constraint_indices {
            let va = form.candidates[a].violations[ci];
            let vb = form.candidates[b].violations[ci];
            match va.cmp(&vb) {
                std::cmp::Ordering::Less => return std::cmp::Ordering::Less,
                std::cmp::Ordering::Greater => return std::cmp::Ordering::Greater,
                std::cmp::Ordering::Equal => continue,
            }
        }
        std::cmp::Ordering::Equal
    });

    let mut result = Vec::with_capacity(form.candidates.len());
    result.push(0); // winner first
    result.extend(rival_indices);
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    fn load_tiny() -> (Tableau, RCDResult) {
        let text = std::fs::read_to_string("../examples/TinyIllustrativeFile.txt")
            .expect("Failed to load tiny example");
        let tableau = Tableau::parse(&text).unwrap();
        let result = tableau.run_rcd();
        (tableau, result)
    }

    #[test]
    fn test_sorted_input_constraints_in_stratum_order() {
        let (tableau, result) = load_tiny();
        let output = format_sorted_input(&tableau, &result);
        let lines: Vec<&str> = output.lines().collect();

        // First line: full names in stratum order
        // RCD on tiny example: stratum 1 = {*NoOns, *Coda}, stratum 2 = {Max, Dep}
        let header_names: Vec<&str> = lines[0].split('\t').skip(2).collect();
        assert_eq!(header_names, vec!["*No Onset", "*Coda", "Max(t)", "Dep(?)"]);

        // Second line: abbreviations in same order
        let header_abbrevs: Vec<&str> = lines[1].split('\t').skip(2).collect();
        assert_eq!(header_abbrevs, vec!["*NoOns", "*Coda", "Max", "Dep"]);
    }

    #[test]
    fn test_sorted_input_candidates_sorted_by_harmony() {
        let (tableau, result) = load_tiny();
        let output = format_sorted_input(&tableau, &result);
        let lines: Vec<&str> = output.lines().collect();

        // Skip 2 header lines. Third input form "at" has 4 candidates.
        // After sorting by harmony through sorted constraints (*NoOns, *Coda, Max, Dep):
        //   Winner ?a: 0,0,1,1  (stays first)
        //   Rivals sorted: ?at(0,1,0,1), a(1,0,1,0), at(1,1,0,0)
        // Find the "at" form lines (starts after a's 2 lines + tat's 2 lines = line index 6)
        let at_lines: Vec<&str> = lines[6..10].to_vec();
        let candidates: Vec<&str> = at_lines
            .iter()
            .map(|l| l.split('\t').nth(1).unwrap())
            .collect();
        assert_eq!(candidates, vec!["?a", "?at", "a", "at"]);
    }

    #[test]
    fn test_sorted_input_winner_always_first() {
        let (tableau, result) = load_tiny();
        let output = format_sorted_input(&tableau, &result);
        let lines: Vec<&str> = output.lines().collect();

        // First data line for each form should have the input in column 0
        // and the winner in column 1
        assert!(lines[2].starts_with("a\t?a\t"));
        assert!(lines[4].starts_with("tat\tta\t"));
        assert!(lines[6].starts_with("at\t?a\t"));
    }

    #[test]
    fn test_sorted_input_blank_for_zero() {
        let (tableau, result) = load_tiny();
        let output = format_sorted_input(&tableau, &result);
        let lines: Vec<&str> = output.lines().collect();

        // First candidate of "a": ?a with freq=1, violations 0,0,0,1
        // In sorted order (*NoOns, *Coda, Max, Dep): 0,0,0,1
        // Should be: "a\t?a\t1\t\t\t\t1"
        assert_eq!(lines[2], "a\t?a\t1\t\t\t\t1");

        // Second candidate of "a": a with freq=0, violations 1,0,0,0
        // Should be: "\ta\t\t1"  (freq blank, then 1 for *NoOns, rest blank)
        assert_eq!(lines[3], "\ta\t\t1\t\t\t");
    }
}
