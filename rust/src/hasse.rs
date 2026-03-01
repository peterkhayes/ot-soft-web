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

/// Generate a GraphViz DOT string for a GLA Hasse diagram.
///
/// `ranking_values`: the final ranking values from GLA (one per constraint, in
///   the same order as `abbrevs`).
/// `abbrevs`: constraint abbreviations, in the same order as `ranking_values`.
///
/// Ports `boersma.frm:PrintPairwiseRankingProbabilities()`.
pub(crate) fn gla_hasse_dot(ranking_values: &[f64], abbrevs: &[&str]) -> String {
    let n = abbrevs.len();

    // Sort constraints by descending ranking value (highest first), matching VB6 SlotFiller.
    let mut sorted_order: Vec<usize> = (0..n).collect();
    sorted_order.sort_by(|&a, &b| {
        ranking_values[b]
            .partial_cmp(&ranking_values[a])
            .unwrap_or(std::cmp::Ordering::Equal)
    });

    let mut dot = String::from("digraph G {\n");

    // Nodes: 1-indexed by original constraint order (matching VB6 node numbering)
    for (i, abbrev) in abbrevs.iter().enumerate() {
        dot.push_str(&format!(
            "   {} [label=\"{}\",fontsize = 14]\n",
            i + 1,
            abbrev
        ));
    }

    // Edges: iterate over all sorted pairs (sorted_i < sorted_j)
    for sorted_i in 0..n {
        for sorted_j in (sorted_i + 1)..n {
            let orig_i = sorted_order[sorted_i];
            let orig_j = sorted_order[sorted_j];
            let diff = ranking_values[orig_i] - ranking_values[orig_j];

            if let Some(prob_str) = lookup_gla_probability(diff) {
                let prob_val: f64 = prob_str.parse().unwrap_or(0.0);
                dot.push_str(&format!("   {} -> {} [fontsize=11", orig_i + 1, orig_j + 1));
                if prob_val < 0.95 {
                    dot.push_str(",style=dotted");
                }
                dot.push_str(&format!(",label={},fontsize=11]\n", prob_str));
            } else {
                // Probability exceeds all 443 table entries (P > 0.999):
                // only emit an edge for constraints adjacent in sorted order.
                if sorted_j == sorted_i + 1 {
                    dot.push_str(&format!(
                        "   {} -> {} [fontsize=11,label= 1 ]\n",
                        orig_i + 1,
                        orig_j + 1
                    ));
                }
            }
        }
    }

    dot.push('}');
    dot
}

/// Look up the probability string for a ranking value difference.
///
/// Ports the lookup in `boersma.frm:PrintPairwiseRankingProbabilities()`.
/// Returns `None` if `diff` exceeds all 443 table thresholds (P > 0.999).
fn lookup_gla_probability(diff: f64) -> Option<&'static str> {
    for &(threshold, prob) in PROBABILITY_TABLE {
        if diff < threshold {
            return Some(prob);
        }
    }
    None
}

/// 443-entry threshold/probability table.
///
/// Each entry is `(threshold, probability_string)`.  Find the first entry where
/// `diff < threshold`; the corresponding string is the probability label.
///
/// Ported verbatim from `boersma.frm:LookUpProbabilities()` (entries 1–443).
// The value 3.14 in this table is a data value, not an approximation of π.
#[allow(clippy::approx_constant)]
static PROBABILITY_TABLE: &[(f64, &str)] = &[
    (0.0, "0.5"),
    (0.01, "0.501"),
    (0.02, "0.503"),
    (0.03, "0.504"),
    (0.04, "0.506"),
    (0.05, "0.507"),
    (0.06, "0.508"),
    (0.07, "0.51"),
    (0.08, "0.511"),
    (0.09, "0.513"),
    (0.1, "0.514"),
    (0.11, "0.516"),
    (0.12, "0.517"),
    (0.13, "0.518"),
    (0.14, "0.52"),
    (0.15, "0.521"),
    (0.16, "0.523"),
    (0.17, "0.524"),
    (0.18, "0.525"),
    (0.19, "0.527"),
    (0.2, "0.528"),
    (0.21, "0.53"),
    (0.22, "0.531"),
    (0.23, "0.532"),
    (0.24, "0.534"),
    (0.25, "0.535"),
    (0.26, "0.537"),
    (0.27, "0.538"),
    (0.28, "0.539"),
    (0.29, "0.541"),
    (0.3, "0.542"),
    (0.31, "0.544"),
    (0.32, "0.545"),
    (0.33, "0.546"),
    (0.34, "0.548"),
    (0.35, "0.549"),
    (0.36, "0.551"),
    (0.37, "0.552"),
    (0.38, "0.553"),
    (0.39, "0.555"),
    (0.4, "0.556"),
    (0.41, "0.558"),
    (0.42, "0.559"),
    (0.43, "0.56"),
    (0.44, "0.562"),
    (0.45, "0.563"),
    (0.46, "0.565"),
    (0.47, "0.566"),
    (0.48, "0.567"),
    (0.49, "0.569"),
    (0.5, "0.57"),
    (0.51, "0.572"),
    (0.52, "0.573"),
    (0.53, "0.574"),
    (0.54, "0.576"),
    (0.55, "0.577"),
    (0.56, "0.578"),
    (0.57, "0.58"),
    (0.58, "0.581"),
    (0.59, "0.583"),
    (0.6, "0.584"),
    (0.61, "0.585"),
    (0.62, "0.587"),
    (0.63, "0.588"),
    (0.64, "0.59"),
    (0.65, "0.591"),
    (0.66, "0.592"),
    (0.67, "0.594"),
    (0.68, "0.595"),
    (0.69, "0.596"),
    (0.7, "0.598"),
    (0.71, "0.599"),
    (0.72, "0.6"),
    (0.73, "0.602"),
    (0.74, "0.603"),
    (0.75, "0.605"),
    (0.76, "0.606"),
    (0.77, "0.607"),
    (0.78, "0.609"),
    (0.79, "0.61"),
    (0.8, "0.611"),
    (0.81, "0.613"),
    (0.82, "0.614"),
    (0.83, "0.615"),
    (0.84, "0.617"),
    (0.85, "0.618"),
    (0.86, "0.619"),
    (0.87, "0.621"),
    (0.88, "0.622"),
    (0.89, "0.623"),
    (0.9, "0.625"),
    (0.91, "0.626"),
    (0.92, "0.628"),
    (0.93, "0.629"),
    (0.94, "0.63"),
    (0.95, "0.632"),
    (0.96, "0.633"),
    (0.97, "0.634"),
    (0.98, "0.636"),
    (0.99, "0.637"),
    (1.0, "0.638"),
    (1.01, "0.639"),
    (1.02, "0.641"),
    (1.03, "0.642"),
    (1.04, "0.643"),
    (1.05, "0.645"),
    (1.06, "0.646"),
    (1.07, "0.647"),
    (1.08, "0.649"),
    (1.09, "0.65"),
    (1.1, "0.651"),
    (1.11, "0.653"),
    (1.12, "0.654"),
    (1.13, "0.655"),
    (1.14, "0.657"),
    (1.15, "0.658"),
    (1.16, "0.659"),
    (1.17, "0.66"),
    (1.18, "0.662"),
    (1.19, "0.663"),
    (1.2, "0.664"),
    (1.21, "0.666"),
    (1.22, "0.667"),
    (1.23, "0.668"),
    (1.24, "0.669"),
    (1.25, "0.671"),
    (1.26, "0.672"),
    (1.27, "0.673"),
    (1.28, "0.675"),
    (1.29, "0.676"),
    (1.3, "0.677"),
    (1.31, "0.678"),
    (1.32, "0.68"),
    (1.33, "0.681"),
    (1.34, "0.682"),
    (1.35, "0.683"),
    (1.36, "0.685"),
    (1.37, "0.686"),
    (1.38, "0.687"),
    (1.39, "0.688"),
    (1.4, "0.69"),
    (1.41, "0.691"),
    (1.42, "0.692"),
    (1.43, "0.693"),
    (1.44, "0.695"),
    (1.45, "0.696"),
    (1.46, "0.697"),
    (1.47, "0.698"),
    (1.48, "0.7"),
    (1.49, "0.701"),
    (1.5, "0.702"),
    (1.51, "0.703"),
    (1.52, "0.705"),
    (1.53, "0.706"),
    (1.54, "0.707"),
    (1.55, "0.708"),
    (1.56, "0.709"),
    (1.57, "0.711"),
    (1.58, "0.712"),
    (1.59, "0.713"),
    (1.6, "0.714"),
    (1.61, "0.715"),
    (1.62, "0.717"),
    (1.63, "0.718"),
    (1.64, "0.719"),
    (1.65, "0.72"),
    (1.66, "0.721"),
    (1.67, "0.723"),
    (1.68, "0.724"),
    (1.69, "0.725"),
    (1.7, "0.726"),
    (1.71, "0.727"),
    (1.72, "0.728"),
    (1.73, "0.73"),
    (1.74, "0.731"),
    (1.75, "0.732"),
    (1.76, "0.733"),
    (1.77, "0.734"),
    (1.78, "0.735"),
    (1.79, "0.737"),
    (1.8, "0.738"),
    (1.81, "0.739"),
    (1.82, "0.74"),
    (1.83, "0.741"),
    (1.84, "0.742"),
    (1.85, "0.743"),
    (1.86, "0.745"),
    (1.87, "0.746"),
    (1.88, "0.747"),
    (1.89, "0.748"),
    (1.9, "0.749"),
    (1.91, "0.75"),
    (1.92, "0.751"),
    (1.93, "0.752"),
    (1.94, "0.754"),
    (1.95, "0.755"),
    (1.96, "0.756"),
    (1.97, "0.757"),
    (1.98, "0.758"),
    (1.99, "0.759"),
    (2.0, "0.76"),
    (2.01, "0.761"),
    (2.02, "0.762"),
    (2.03, "0.764"),
    (2.04, "0.765"),
    (2.05, "0.766"),
    (2.06, "0.767"),
    (2.07, "0.768"),
    (2.08, "0.769"),
    (2.09, "0.77"),
    (2.1, "0.771"),
    (2.11, "0.772"),
    (2.12, "0.773"),
    (2.13, "0.774"),
    (2.14, "0.775"),
    (2.15, "0.776"),
    (2.16, "0.777"),
    (2.17, "0.779"),
    (2.18, "0.78"),
    (2.19, "0.781"),
    (2.2, "0.782"),
    (2.21, "0.783"),
    (2.22, "0.784"),
    (2.23, "0.785"),
    (2.24, "0.786"),
    (2.25, "0.787"),
    (2.26, "0.788"),
    (2.27, "0.789"),
    (2.28, "0.79"),
    (2.29, "0.791"),
    (2.3, "0.792"),
    (2.31, "0.793"),
    (2.32, "0.794"),
    (2.33, "0.795"),
    (2.34, "0.796"),
    (2.35, "0.797"),
    (2.36, "0.798"),
    (2.37, "0.799"),
    (2.38, "0.8"),
    (2.39, "0.801"),
    (2.4, "0.802"),
    (2.41, "0.803"),
    (2.42, "0.804"),
    (2.43, "0.805"),
    (2.44, "0.806"),
    (2.45, "0.807"),
    (2.46, "0.808"),
    (2.47, "0.809"),
    (2.48, "0.81"),
    (2.49, "0.811"),
    (2.5, "0.812"),
    (2.51, "0.813"),
    (2.52, "0.814"),
    (2.54, "0.815"),
    (2.55, "0.816"),
    (2.56, "0.817"),
    (2.57, "0.818"),
    (2.58, "0.819"),
    (2.59, "0.82"),
    (2.6, "0.821"),
    (2.61, "0.822"),
    (2.62, "0.823"),
    (2.63, "0.824"),
    (2.64, "0.825"),
    (2.65, "0.826"),
    (2.66, "0.827"),
    (2.68, "0.828"),
    (2.69, "0.829"),
    (2.7, "0.83"),
    (2.71, "0.831"),
    (2.72, "0.832"),
    (2.73, "0.833"),
    (2.74, "0.834"),
    (2.75, "0.835"),
    (2.77, "0.836"),
    (2.78, "0.837"),
    (2.79, "0.838"),
    (2.8, "0.839"),
    (2.81, "0.84"),
    (2.82, "0.841"),
    (2.84, "0.842"),
    (2.85, "0.843"),
    (2.86, "0.844"),
    (2.87, "0.845"),
    (2.88, "0.846"),
    (2.89, "0.847"),
    (2.91, "0.848"),
    (2.92, "0.849"),
    (2.93, "0.85"),
    (2.94, "0.851"),
    (2.95, "0.852"),
    (2.97, "0.853"),
    (2.98, "0.854"),
    (2.99, "0.855"),
    (3.0, "0.856"),
    (3.02, "0.857"),
    (3.03, "0.858"),
    (3.04, "0.859"),
    (3.05, "0.86"),
    (3.07, "0.861"),
    (3.08, "0.862"),
    (3.09, "0.863"),
    (3.11, "0.864"),
    (3.12, "0.865"),
    (3.13, "0.866"),
    (3.14, "0.867"),
    (3.16, "0.868"),
    (3.17, "0.869"),
    (3.18, "0.87"),
    (3.2, "0.871"),
    (3.21, "0.872"),
    (3.22, "0.873"),
    (3.24, "0.874"),
    (3.25, "0.875"),
    (3.27, "0.876"),
    (3.28, "0.877"),
    (3.29, "0.878"),
    (3.31, "0.879"),
    (3.32, "0.88"),
    (3.34, "0.881"),
    (3.35, "0.882"),
    (3.36, "0.883"),
    (3.38, "0.884"),
    (3.39, "0.885"),
    (3.41, "0.886"),
    (3.42, "0.887"),
    (3.44, "0.888"),
    (3.45, "0.889"),
    (3.47, "0.89"),
    (3.48, "0.891"),
    (3.5, "0.892"),
    (3.51, "0.893"),
    (3.53, "0.894"),
    (3.54, "0.895"),
    (3.56, "0.896"),
    (3.57, "0.897"),
    (3.59, "0.898"),
    (3.61, "0.899"),
    (3.62, "0.9"),
    (3.64, "0.901"),
    (3.65, "0.902"),
    (3.67, "0.903"),
    (3.69, "0.904"),
    (3.7, "0.905"),
    (3.72, "0.906"),
    (3.74, "0.907"),
    (3.75, "0.908"),
    (3.77, "0.909"),
    (3.79, "0.91"),
    (3.81, "0.911"),
    (3.82, "0.912"),
    (3.84, "0.913"),
    (3.86, "0.914"),
    (3.88, "0.915"),
    (3.9, "0.916"),
    (3.91, "0.917"),
    (3.93, "0.918"),
    (3.95, "0.919"),
    (3.97, "0.92"),
    (3.99, "0.921"),
    (4.01, "0.922"),
    (4.03, "0.923"),
    (4.05, "0.924"),
    (4.07, "0.925"),
    (4.09, "0.926"),
    (4.11, "0.927"),
    (4.13, "0.928"),
    (4.15, "0.929"),
    (4.17, "0.93"),
    (4.19, "0.931"),
    (4.21, "0.932"),
    (4.23, "0.933"),
    (4.25, "0.934"),
    (4.28, "0.935"),
    (4.3, "0.936"),
    (4.32, "0.937"),
    (4.34, "0.938"),
    (4.37, "0.939"),
    (4.39, "0.94"),
    (4.41, "0.941"),
    (4.44, "0.942"),
    (4.46, "0.943"),
    (4.49, "0.944"),
    (4.51, "0.945"),
    (4.54, "0.946"),
    (4.56, "0.947"),
    (4.59, "0.948"),
    (4.62, "0.949"),
    (4.64, "0.95"),
    (4.67, "0.951"),
    (4.7, "0.952"),
    (4.73, "0.953"),
    (4.76, "0.954"),
    (4.79, "0.955"),
    (4.82, "0.956"),
    (4.85, "0.957"),
    (4.88, "0.958"),
    (4.91, "0.959"),
    (4.94, "0.96"),
    (4.97, "0.961"),
    (5.01, "0.962"),
    (5.04, "0.963"),
    (5.08, "0.964"),
    (5.11, "0.965"),
    (5.15, "0.966"),
    (5.19, "0.967"),
    (5.22, "0.968"),
    (5.26, "0.969"),
    (5.3, "0.97"),
    (5.35, "0.971"),
    (5.39, "0.972"),
    (5.43, "0.973"),
    (5.48, "0.974"),
    (5.52, "0.975"),
    (5.57, "0.976"),
    (5.62, "0.977"),
    (5.68, "0.978"),
    (5.73, "0.979"),
    (5.78, "0.98"),
    (5.84, "0.981"),
    (5.9, "0.982"),
    (5.97, "0.983"),
    (6.04, "0.984"),
    (6.11, "0.985"),
    (6.18, "0.986"),
    (6.26, "0.987"),
    (6.34, "0.988"),
    (6.44, "0.989"),
    (6.53, "0.99"),
    (6.64, "0.991"),
    (6.76, "0.992"),
    (6.88, "0.993"),
    (7.03, "0.994"),
    (7.2, "0.995"),
    (7.39, "0.996"),
    (7.63, "0.997"),
    (7.94, "0.998"),
    (8.4, "0.9985"),
    (8.43, "0.9986"),
    (8.49, "0.9987"),
    (8.56, "0.9988"),
    (8.63, "0.9989"),
    (8.7, "0.999"),
];

// ─────────────────────────────────────────────────────────────────────────────
// Tests
// ─────────────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tableau::Tableau;

    fn load_tiny() -> String {
        std::fs::read_to_string("../examples/TinyIllustrativeFile/input.txt")
            .expect("Failed to load examples/TinyIllustrativeFile/input.txt")
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

    #[test]
    fn test_gla_hasse_dot_nodes() {
        // With 4 constraints, there should be 4 nodes labeled with abbreviations.
        let abbrevs = ["A", "B", "C", "D"];
        let values = [10.0_f64, 8.0, 6.0, 4.0];
        let dot = gla_hasse_dot(&values, &abbrevs);

        assert!(dot.starts_with("digraph G {"));
        assert!(dot.contains("[label=\"A\",fontsize = 14]"));
        assert!(dot.contains("[label=\"B\",fontsize = 14]"));
        assert!(dot.contains("[label=\"C\",fontsize = 14]"));
        assert!(dot.contains("[label=\"D\",fontsize = 14]"));
    }

    #[test]
    fn test_gla_hasse_dot_large_diff_adjacent_only() {
        // A diff of 20.0 exceeds all thresholds (P > 0.999).
        // Only adjacent pairs in sorted order get an edge.
        let abbrevs = ["A", "B", "C"];
        // Sorted order: A > B > C. Adjacent in sorted order: (A,B) and (B,C).
        // (A,C) is not adjacent (sorted_j = 2, sorted_i = 0, diff != 1).
        let values = [30.0_f64, 10.0, 0.0]; // diffs: A-B=20, A-C=30, B-C=10 — all > 8.7
        let dot = gla_hasse_dot(&values, &abbrevs);

        // A->B and B->C should have "1" labels
        assert!(dot.contains("1 -> 2"), "expected A->B edge");
        assert!(dot.contains("2 -> 3"), "expected B->C edge");
        // A->C (1->3) should NOT appear (not adjacent in sorted order)
        assert!(!dot.contains("1 -> 3"), "unexpected A->C edge");
        assert!(dot.contains("label= 1 "));
    }

    #[test]
    fn test_gla_hasse_dot_small_diff_dotted() {
        // A diff of 0.05 is less than threshold[6]=0.05... wait, let's use diff=0.005
        // which is < threshold[2]=0.01 → prob "0.501" < 0.95 → dotted.
        let abbrevs = ["A", "B"];
        let values = [100.005_f64, 100.0]; // diff = 0.005
        let dot = gla_hasse_dot(&values, &abbrevs);
        assert!(dot.contains("style=dotted"));
        assert!(dot.contains("label=0.501"));
    }

    #[test]
    fn test_gla_hasse_dot_high_prob_solid() {
        // diff of 4.64 is exactly at threshold[389]=4.64 → 4.64 < 4.64 is false,
        // so we use threshold[390]=4.67, prob "0.951" >= 0.95 → solid.
        // Use diff = 4.65 (>= 4.64, < 4.67) → prob "0.951", no dotted.
        let abbrevs = ["A", "B"];
        let values = [104.65_f64, 100.0]; // diff = 4.65
        let dot = gla_hasse_dot(&values, &abbrevs);
        assert!(!dot.contains("style=dotted"));
        assert!(dot.contains("label=0.951"));
    }

    #[test]
    fn test_gla_hasse_dot_sort_order() {
        // Even if constraints are given in non-sorted order, the edges should
        // go from the higher-ranked to the lower-ranked constraint.
        // Constraint 0 (abbrev "Low") has ranking value 1.0.
        // Constraint 1 (abbrev "High") has ranking value 10.0.
        // Sorted order: 1 (High) then 0 (Low). Edge should be 2 -> 1 (orig indices).
        let abbrevs = ["Low", "High"];
        let values = [1.0_f64, 10.0]; // High (idx 1) > Low (idx 0)
        let dot = gla_hasse_dot(&values, &abbrevs);
        // Node 1 = Low, Node 2 = High. Edge should be 2 -> 1.
        assert!(dot.contains("2 -> 1"), "expected High->Low edge");
        assert!(!dot.contains("1 -> 2"), "unexpected Low->High edge");
    }

    #[test]
    fn test_probability_table_length() {
        assert_eq!(PROBABILITY_TABLE.len(), 443);
    }
}
