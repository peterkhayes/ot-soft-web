//! Hasse diagram DOT generation
//!
//! Ports `Fred.bas:PrepareHasseDiagram()` from VB6.
//! Produces a GraphViz DOT string encoding the partial order over constraints.

const NO_ARGUMENT: u8 = 0;
const DISJUNCTIVE: u8 = 1;
const CERTAIN: u8 = 2;

/// Generate a GraphViz DOT string for a FRed Hasse diagram.
///
/// `valhalla`: the final ERC basis from FRed (each ERC is a string of 'W', 'L', 'e').
/// `abbrevs`: constraint abbreviations, in the same column order as the ERCs.
///
/// Ports `Fred.bas:PrepareHasseDiagram()`.
pub(crate) fn fred_hasse_dot(valhalla: &[&str], abbrevs: &[&str]) -> String {
    let n = abbrevs.len();

    // ranking_array[i][j]: relationship from constraint i to constraint j
    let mut ranking_array = vec![vec![NO_ARGUMENT; n]; n];

    for erc in valhalla {
        let chars: Vec<char> = erc.chars().collect();
        if chars.len() < n {
            continue;
        }
        let w_count = chars.iter().filter(|&&c| c == 'W').count();

        if w_count == 1 {
            // Certain ranking: the single W dominates all L positions
            let dominator = chars.iter().position(|&c| c == 'W').unwrap();
            for (j, &ch) in chars.iter().enumerate() {
                if ch == 'L' {
                    ranking_array[dominator][j] = CERTAIN;
                }
            }
        } else if w_count > 1 {
            // Disjunctive ranking: each W dominates each L (if not already Certain)
            for (i, &ch) in chars.iter().enumerate() {
                if ch == 'W' {
                    for (j, &ch2) in chars.iter().enumerate() {
                        if ch2 == 'L' && ranking_array[i][j] == NO_ARGUMENT {
                            ranking_array[i][j] = DISJUNCTIVE;
                        }
                    }
                }
            }
        }
    }

    let mut dot = String::from("digraph G {\n");

    // Nodes: 1-indexed labels matching VB6 output
    for (i, abbrev) in abbrevs.iter().enumerate() {
        dot.push_str(&format!(
            "   {} [label=\"{}\",fontsize = 14]\n",
            i + 1,
            abbrev
        ));
    }

    // Edges
    for (i, row) in ranking_array.iter().enumerate() {
        for (j, &rel) in row.iter().enumerate() {
            match rel {
                CERTAIN => {
                    dot.push_str(&format!(
                        "   {} -> {} [fontsize = 11]\n",
                        i + 1,
                        j + 1
                    ));
                }
                DISJUNCTIVE => {
                    dot.push_str(&format!(
                        "   {} -> {} [fontsize = 11,style=dotted,label=\"or\"]\n",
                        i + 1,
                        j + 1
                    ));
                }
                _ => {}
            }
        }
    }

    dot.push('}');
    dot
}

// ─────────────────────────────────────────────────────────────────────────────
// Tests
// ─────────────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tableau::Tableau;

    fn load_tiny() -> String {
        std::fs::read_to_string("../examples/tiny/input.txt")
            .expect("Failed to load examples/tiny/input.txt")
    }

    #[test]
    fn test_fred_hasse_dot_tiny() {
        let text = load_tiny();
        let tableau = Tableau::parse(&text).expect("parse failed");
        let result = tableau.run_fred(false); // Skeletal Basis

        let abbrevs: Vec<&str> = result.constraint_abbrevs.iter().map(|s| s.as_str()).collect();
        let valhalla: Vec<&str> = result.valhalla.iter().map(|s| s.as_str()).collect();
        let dot = fred_hasse_dot(&valhalla, &abbrevs);

        // Should have 4 nodes for the 4 constraints
        assert!(dot.contains("digraph G {"));
        assert!(dot.contains("[label=\"*NoOns\",fontsize = 14]"));
        assert!(dot.contains("[label=\"*Coda\",fontsize = 14]"));
        assert!(dot.contains("[label=\"Max\",fontsize = 14]"));
        assert!(dot.contains("[label=\"Dep\",fontsize = 14]"));

        // WeeL → *NoOns >> Dep (certain): constraint 1 → constraint 4
        // eWLe → *Coda >> Max (certain): constraint 2 → constraint 3
        assert!(dot.contains("[fontsize = 11]"));
        // No disjunctive edges for this example
        assert!(!dot.contains("style=dotted"));
    }

    #[test]
    fn test_fred_hasse_dot_empty_valhalla() {
        let dot = fred_hasse_dot(&[], &["A", "B"]);
        assert_eq!(dot, "digraph G {\n   1 [label=\"A\",fontsize = 14]\n   2 [label=\"B\",fontsize = 14]\n}");
    }

    #[test]
    fn test_fred_hasse_dot_disjunctive() {
        // ERC "WWeLL" → two W's, so disjunctive edges
        let abbrevs = ["C1", "C2", "C3", "C4", "C5"];
        let dot = fred_hasse_dot(&["WWeLL"], &abbrevs);
        assert!(dot.contains("style=dotted"));
        assert!(dot.contains("\"or\""));
        // No certain edges
        let lines: Vec<&str> = dot.lines().collect();
        let certain: Vec<_> = lines.iter()
            .filter(|l| l.contains("->") && !l.contains("dotted"))
            .collect();
        assert!(certain.is_empty());
    }

    #[test]
    fn test_fred_hasse_dot_certain_not_overridden_by_disjunctive() {
        // First ERC: certain C1 >> C3; Second ERC: disjunctive involving same pair
        // Certain should not be overridden
        let abbrevs = ["C1", "C2", "C3"];
        let dot = fred_hasse_dot(&["WeL", "WWL"], &abbrevs);
        // C1 >> C3 is Certain from "WeL"
        // "WWL": both C1 and C2 are W, C3 is L → C1→C3 stays Certain, C2→C3 becomes Disjunctive
        assert!(dot.contains("style=dotted")); // C2→C3 is disjunctive
        // Certain edge line for 1→3 should not have "dotted"
        let lines: Vec<&str> = dot.lines().collect();
        let edge_1_3: Vec<_> = lines.iter()
            .filter(|l| l.contains("1 -> 3"))
            .collect();
        assert_eq!(edge_1_3.len(), 1);
        assert!(!edge_1_3[0].contains("dotted"));
    }
}
