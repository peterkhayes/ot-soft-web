//! Biased Constraint Demotion (BCD) algorithm
//!
//! BCD extends RCD with a bias toward ranking Faithfulness constraints low.
//! It uses several heuristics: faithfulness delay, activeness filtering,
//! optional specificity, and minimal faithfulness subset selection.
//!
//! Source: BCD.bas (programmed by Bruce Tesar)
//! Reference: Prince & Tesar (2004)

use crate::rcd::RCDResult;
use crate::tableau::Tableau;

/// Precompute violation subset relationships between constraints.
/// `result[i][j]` is true if constraint i's violations are a subset of constraint j's
/// (i.e., for all forms/rivals, violations(i) <= violations(j)), AND constraint i
/// has at least one nonzero violation somewhere.
///
/// Reproduces VB6 BCD.bas:LocateViolationSubsets and LFCD.bas:LocateViolationSubsets
pub(crate) fn locate_violation_subsets(tableau: &Tableau) -> Vec<Vec<bool>> {
    let nc = tableau.constraints.len();
    let mut subset = vec![vec![false; nc]; nc];

    for (outer, row) in subset.iter_mut().enumerate() {
        for (inner, cell) in row.iter_mut().enumerate() {
            if outer == inner {
                continue;
            }

            let mut is_subset = true;
            let mut has_any_violation = false;

            'form_loop: for form in &tableau.forms {
                let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                    Some(w) => w,
                    None => continue,
                };

                // Compare winner violations
                if winner.violations[outer] > winner.violations[inner] {
                    is_subset = false;
                    break;
                }
                if winner.violations[outer] > 0 {
                    has_any_violation = true;
                }

                // Compare rival violations
                for rival in form.candidates.iter().filter(|c| c.frequency == 0) {
                    if rival.violations[outer] > rival.violations[inner] {
                        is_subset = false;
                        break 'form_loop;
                    }
                    if rival.violations[outer] > 0 {
                        has_any_violation = true;
                    }
                }
            }

            // Subset(outer, inner) = true only if is_subset AND outer has at least one violation
            *cell = is_subset && has_any_violation;
        }
    }

    subset
}

/// Immutable recursion context for BCD faith-subset evaluation.
struct BcdContext<'a> {
    tableau: &'a Tableau,
    is_faithfulness: &'a [bool],
    current_stratum: usize,
    strata: &'a [usize],
    still_informative: &'a [Vec<bool>],
}

/// Simulate ranking a faithfulness subset and count how many markedness constraints
/// are subsequently freed up (become non-demotable).
///
/// Reproduces VB6 BCD.bas:CheckMarkednessRelease
fn check_markedness_release(
    ctx: &BcdContext,
    faith_subset: &[usize],
) -> usize {
    let tableau = ctx.tableau;
    let is_faithfulness = ctx.is_faithfulness;
    let current_stratum = ctx.current_stratum;
    let strata = ctx.strata;
    let still_informative = ctx.still_informative;
    let nc = tableau.constraints.len();

    // Make local copies of state
    let mut local_strata = strata.to_vec();
    let mut local_informative: Vec<Vec<bool>> = still_informative.to_vec();
    let mut local_current_stratum = current_stratum;

    // Put the faith subset in the current stratum
    for &c_idx in faith_subset {
        local_strata[c_idx] = local_current_stratum;
    }

    let mut total_markedness_freed = 0;

    loop {
        // Remove pairs decided by constraints in current stratum
        for (c_idx, &c_stratum) in local_strata.iter().enumerate() {
            if c_stratum != local_current_stratum {
                continue;
            }
            for (form_idx, form) in tableau.forms.iter().enumerate() {
                let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                    Some(w) => w,
                    None => continue,
                };
                for (rival_idx, rival) in form.candidates.iter().enumerate() {
                    if rival.frequency > 0 {
                        continue;
                    }
                    if rival.violations[c_idx] > winner.violations[c_idx] {
                        local_informative[form_idx][rival_idx] = false;
                    }
                }
            }
        }

        local_current_stratum += 1;

        // Mark demotable constraints
        let mut demotable = vec![false; nc];

        for (form_idx, form) in tableau.forms.iter().enumerate() {
            let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                Some(w) => w,
                None => continue,
            };
            for (rival_idx, rival) in form.candidates.iter().enumerate() {
                if rival.frequency > 0 {
                    continue;
                }
                if !local_informative[form_idx][rival_idx] {
                    continue;
                }
                for (c_idx, demot) in demotable.iter_mut().enumerate() {
                    if local_strata[c_idx] == 0
                        && winner.violations[c_idx] > rival.violations[c_idx]
                    {
                        *demot = true;
                    }
                }
            }
        }

        // Rank non-demotable markedness constraints
        let mut markedness_count = 0;
        for (c_idx, ls) in local_strata.iter_mut().enumerate() {
            if *ls == 0 && !demotable[c_idx] && !is_faithfulness[c_idx] {
                markedness_count += 1;
                *ls = local_current_stratum;
            }
        }

        total_markedness_freed += markedness_count;

        if markedness_count == 0 {
            break;
        }
    }

    total_markedness_freed
}

/// Mutable state tracking the best faithfulness subset found during enumeration.
struct BcdSearchState {
    current_subset: Vec<usize>,
    best_subset: Vec<usize>,
    best_mark_count: usize,
    tie_flag: bool,
}

/// Recursively enumerate all subsets of `rankable_faith` of size `target_size`,
/// evaluate each, and track the best (most markedness freed).
///
/// Reproduces VB6 BCD.bas:EvaluateFaithSubsets
fn evaluate_faith_subsets(
    rankable_faith: &[usize],
    target_size: usize,
    start_idx: usize,
    state: &mut BcdSearchState,
    ctx: &BcdContext,
) {
    if state.current_subset.len() == target_size {
        // Evaluate this complete subset
        let mark_count = check_markedness_release(ctx, &state.current_subset);

        // Check for tie (only non-zero ties count)
        if mark_count == state.best_mark_count && state.best_mark_count > 0 {
            state.tie_flag = true;
        }

        if mark_count > state.best_mark_count {
            state.tie_flag = false;
            state.best_mark_count = mark_count;
            state.best_subset = state.current_subset.clone();
        }

        return;
    }

    // Recurse, building subsets
    let remaining_needed = target_size - state.current_subset.len();
    let max_start = rankable_faith.len().saturating_sub(remaining_needed - 1);

    for idx in start_idx..max_start {
        state.current_subset.push(rankable_faith[idx]);
        evaluate_faith_subsets(
            rankable_faith,
            target_size,
            idx + 1,
            state,
            ctx,
        );
        state.current_subset.pop();
    }
}

/// Identify demotable and active constraints for the current stratum.
///
/// A constraint is demotable if it prefers a loser; active if it prefers a winner.
fn mark_demotable_and_active(
    tableau: &Tableau,
    strata: &[usize],
    still_informative: &[Vec<bool>],
) -> (Vec<bool>, Vec<bool>) {
    let nc = strata.len();
    let mut demotable = vec![false; nc];
    let mut active = vec![false; nc];

    for (form_idx, form) in tableau.forms.iter().enumerate() {
        let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
            Some(w) => w,
            None => continue,
        };
        for (rival_idx, rival) in form.candidates.iter().enumerate() {
            if rival.frequency > 0 || !still_informative[form_idx][rival_idx] {
                continue;
            }
            for c_idx in 0..nc {
                if strata[c_idx] != 0 {
                    continue;
                }
                if winner.violations[c_idx] > rival.violations[c_idx] {
                    demotable[c_idx] = true;
                } else if winner.violations[c_idx] < rival.violations[c_idx] {
                    active[c_idx] = true;
                }
            }
        }
    }

    (demotable, active)
}

/// Apply the faithfulness delay and subset search to assign constraints to the current stratum.
/// Returns the tie_flag for this iteration.
#[allow(clippy::too_many_arguments)]
fn rank_bcd_stratum(
    tableau: &Tableau,
    strata: &mut [usize],
    current_stratum: usize,
    demotable: &[bool],
    active: &[bool],
    is_faithfulness: &[bool],
    violation_subsets: &[Vec<bool>],
    specific_bcd: bool,
    still_informative: &[Vec<bool>],
) -> bool {
    let nc = strata.len();
    let mut tie_flag = false;

    // --- Faithfulness Delay ---
    crate::ot_log!("Faithfulness Delay:");
    let mut rankable_marked = false;
    let mut faith_is_active = false;

    for c_idx in 0..nc {
        if strata[c_idx] == 0 && !demotable[c_idx] {
            if is_faithfulness[c_idx] && active[c_idx] {
                faith_is_active = true;
            } else if !is_faithfulness[c_idx] {
                rankable_marked = true;
                crate::ot_log!("  Markedness constraint {} prefers no losers, joins stratum #{}.",
                    tableau.constraints[c_idx].abbrev(), current_stratum);
                strata[c_idx] = current_stratum;
            }
        }
    }

    if rankable_marked {
        crate::ot_log!("  Faithfulness constraints are excluded from stratum.");
        return tie_flag;
    }

    crate::ot_log!("  No rankable markedness constraints are available for this stratum.");

    // No markedness could be ranked — must deal with faithfulness

    // --- Avoid The Inactive ---
    crate::ot_log!("Avoid The Inactive:");
    if !faith_is_active {
        crate::ot_log!("  No active Faithfulness constraints are available for this stratum.");
        // No active faithfulness — rank all remaining non-demotable (terminal case)
        for c_idx in 0..nc {
            if strata[c_idx] == 0 && !demotable[c_idx] {
                crate::ot_log!("  Faithfulness constraint {} joins stratum #{} by default.",
                    tableau.constraints[c_idx].abbrev(), current_stratum);
                strata[c_idx] = current_stratum;
            }
        }
        return tie_flag;
    }

    crate::ot_log!("  Active, as-yet-unranked Faithfulness constraints available for this stratum:");
    for c_idx in 0..nc {
        if strata[c_idx] == 0 && is_faithfulness[c_idx] && !demotable[c_idx] && active[c_idx] {
            crate::ot_log!("    {}", tableau.constraints[c_idx].abbrev());
        }
    }

    // --- Favor Specificity (if specific BCD) ---
    let mut subsetted = vec![false; nc];
    if specific_bcd {
        crate::ot_log!("Favor Specificity:");
        let mut any_subsetted = false;
        for c_idx in 0..nc {
            if strata[c_idx] != 0 || !is_faithfulness[c_idx] {
                continue;
            }
            for inner in 0..nc {
                if inner == c_idx || strata[inner] != 0 || !is_faithfulness[inner] {
                    continue;
                }
                if violation_subsets[inner][c_idx] {
                    subsetted[c_idx] = true;
                    any_subsetted = true;
                    crate::ot_log!("  {} cannot be installed in this stratum, because {} is more specific.",
                        tableau.constraints[c_idx].abbrev(), tableau.constraints[inner].abbrev());
                    break;
                }
            }
        }
        if !any_subsetted {
            crate::ot_log!("  (no cases found)");
        }
    }

    // Build rankable faithfulness list
    let rankable_faith: Vec<usize> = (0..nc)
        .filter(|&c_idx| {
            strata[c_idx] == 0
                && is_faithfulness[c_idx]
                && !demotable[c_idx]
                && active[c_idx]
                && !subsetted[c_idx]
        })
        .collect();

    if rankable_faith.is_empty() {
        for c_idx in 0..nc {
            if strata[c_idx] == 0 && !demotable[c_idx] {
                crate::ot_log!("  Faithfulness constraint {} joins stratum #{} by default.",
                    tableau.constraints[c_idx].abbrev(), current_stratum);
                strata[c_idx] = current_stratum;
            }
        }
        return tie_flag;
    }

    // --- Smallest Effective Faithfulness Sets ---
    crate::ot_log!("Smallest Effective Faithfulness Sets");
    crate::ot_log!("  Evaluating subsets of increasing size:");
    let ctx = BcdContext {
        tableau,
        is_faithfulness,
        current_stratum,
        strata,
        still_informative,
    };
    let mut search = BcdSearchState {
        current_subset: Vec::new(),
        best_subset: Vec::new(),
        best_mark_count: 0,
        tie_flag: false,
    };
    for subset_size in 1..=rankable_faith.len() {
        crate::ot_log!("  Subset size = {}", subset_size);
        search.current_subset.clear();
        evaluate_faith_subsets(&rankable_faith, subset_size, 0, &mut search, &ctx);
        if search.best_mark_count > 0 {
            break;
        }
    }
    tie_flag = search.tie_flag;

    if search.best_mark_count == 0 {
        crate::ot_log!("  No subset released any markedness constraints.");
        for &c_idx in &rankable_faith {
            strata[c_idx] = current_stratum;
        }
    } else {
        for &c_idx in &search.best_subset {
            crate::ot_log!("  Faithfulness constraint {} joins stratum #{} as member of best subset.",
                tableau.constraints[c_idx].abbrev(), current_stratum);
            strata[c_idx] = current_stratum;
        }
        if tie_flag {
            crate::ot_log!("  Note: This is an arbitrary choice, arising from the tie noted above.");
        }
    }

    tie_flag
}

/// Mark pairs as no longer informative once a newly-ranked constraint prefers the winner.
fn update_informativeness(
    tableau: &Tableau,
    strata: &[usize],
    current_stratum: usize,
    still_informative: &mut [Vec<bool>],
) {
    for (c_idx, &c_stratum) in strata.iter().enumerate() {
        if c_stratum != current_stratum {
            continue;
        }
        for (form_idx, form) in tableau.forms.iter().enumerate() {
            let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                Some(w) => w,
                None => continue,
            };
            for (rival_idx, rival) in form.candidates.iter().enumerate() {
                if rival.frequency > 0 {
                    continue;
                }
                if rival.violations[c_idx] > winner.violations[c_idx] {
                    still_informative[form_idx][rival_idx] = false;
                }
            }
        }
    }
}

impl Tableau {
    /// Run Biased Constraint Demotion to find a ranking.
    ///
    /// `specific_bcd`: if true, favors more specific faithfulness constraints
    /// (the mnuSpecificBCD option in VB6).
    pub fn run_bcd(&self, specific_bcd: bool) -> RCDResult {
        let nc = self.constraints.len();
        let is_faithfulness: Vec<bool> = self.constraints.iter()
            .map(|c| c.is_faithfulness())
            .collect();
        let violation_subsets = if specific_bcd {
            locate_violation_subsets(self)
        } else {
            vec![]
        };

        let mut strata = vec![0usize; nc];
        let mut current_stratum = 0usize;
        let mut still_informative: Vec<Vec<bool>> = self.forms.iter().map(|form| {
            form.candidates.iter().map(|_| true).collect()
        }).collect();
        let mut upper_tie_flag = false;
        let mut tie_flag = false;

        let count_pairs = |si: &Vec<Vec<bool>>| -> usize {
            self.forms.iter().enumerate()
                .map(|(fi, form)| {
                    form.candidates.iter().enumerate()
                        .filter(|&(ri, c)| c.frequency == 0 && si[fi][ri])
                        .count()
                })
                .sum()
        };

        crate::ot_log!("****** Application of Biased Constraint Demotion ******");
        if specific_bcd {
            crate::ot_log!("Version of BCD used: modified, gives priority to more specific Faithfulness constraints");
        }
        crate::ot_log!("Starting BCD with {} pairs", count_pairs(&still_informative));

        loop {
            current_stratum += 1;
            if tie_flag {
                upper_tie_flag = true;
            }

            crate::ot_log!("");
            crate::ot_log!("****** Now doing Stratum #{} ******", current_stratum);

            let (demotable, active) = mark_demotable_and_active(self, &strata, &still_informative);

            tie_flag = rank_bcd_stratum(
                self, &mut strata, current_stratum, &demotable, &active,
                &is_faithfulness, &violation_subsets, specific_bcd, &still_informative,
            );

            self.log_results_so_far(&strata, current_stratum);

            // Check termination
            let unranked_remain = strata.contains(&0);
            let a_constraint_was_ranked = strata.contains(&current_stratum);

            crate::ot_log!("After stratum {}: {} pairs remaining",
                current_stratum, count_pairs(&still_informative));

            if !unranked_remain {
                crate::ot_log!("");
                crate::ot_log!("Ranking is complete and yields successful grammar.");
                crate::ot_log!("BCD SUCCEEDED with {} strata", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum, true);
                result.set_tie_warning(upper_tie_flag || tie_flag);
                result.compute_extra_analyses(self, &[], true);
                return result;
            }

            if !a_constraint_was_ranked {
                crate::ot_log!("");
                crate::ot_log!("Ranking has failed. This constraint set cannot derive only winners.");
                crate::ot_log!("BCD FAILED: no constraints added to stratum {}", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum - 1, false);
                result.set_tie_warning(upper_tie_flag || tie_flag);
                return result;
            }

            update_informativeness(self, &strata, current_stratum, &mut still_informative);

            if current_stratum > nc {
                crate::ot_log!("");
                crate::ot_log!("Ranking has failed. This constraint set cannot derive only winners.");
                crate::ot_log!("BCD FAILED: safety limit reached at stratum {}", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum, false);
                result.set_tie_warning(upper_tie_flag || tie_flag);
                return result;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::tableau::Tableau;

    fn load_tiny_example() -> String {
        std::fs::read_to_string("../examples/TinyIllustrativeFile.txt")
            .expect("Failed to load examples/TinyIllustrativeFile.txt")
    }

    #[test]
    fn test_faithfulness_detection() {
        let tiny = load_tiny_example();
        let tableau = Tableau::parse(&tiny).unwrap();

        // *No Onset -> markedness (starts with *N)
        assert!(!tableau.constraints[0].is_faithfulness(), "*No Onset should be markedness");
        // *Coda -> markedness (starts with *C)
        assert!(!tableau.constraints[1].is_faithfulness(), "*Coda should be markedness");
        // Max(t) -> faithfulness (starts with Max)
        assert!(tableau.constraints[2].is_faithfulness(), "Max(t) should be faithfulness");
        // Dep(?) -> faithfulness (starts with Dep)
        assert!(tableau.constraints[3].is_faithfulness(), "Dep(?) should be faithfulness");
    }

    #[test]
    fn test_faithfulness_detection_patterns() {
        // Test all documented patterns
        use crate::tableau::Constraint;

        let cases = vec![
            ("Ident-IO", true),
            ("Faithfulness", true),
            ("Id(voice)", true),
            ("Max-IO", true),
            ("Dep-IO", true),
            ("Map-IO", true),
            ("*Map-IO", true),
            ("F:Max", true),
            ("*NoCoda", false),
            ("Align-Left", false),
            ("Onset", false),
            ("f:max", false), // lowercase f: is NOT faithfulness (case-sensitive)
        ];

        for (name, expected) in cases {
            let c = Constraint::new(name.to_string(), name.to_string());
            assert_eq!(c.is_faithfulness(), expected, "Failed for constraint '{}'", name);
        }
    }

    #[test]
    fn test_bcd_tiny_example() {
        let tiny = load_tiny_example();
        let tableau = Tableau::parse(&tiny).unwrap();

        let result = tableau.run_bcd(false);

        // BCD should succeed
        assert!(result.success(), "BCD should find a valid ranking");

        // For the tiny example, markedness constraints should be ranked first
        // (faithfulness delay ensures *NoOns and *Coda go to stratum 1)
        assert_eq!(result.get_stratum(0).unwrap(), 1, "*NoOns should be in stratum 1");
        assert_eq!(result.get_stratum(1).unwrap(), 1, "*Coda should be in stratum 1");

        // Faithfulness constraints should be in stratum 2
        assert_eq!(result.get_stratum(2).unwrap(), 2, "Max should be in stratum 2");
        assert_eq!(result.get_stratum(3).unwrap(), 2, "Dep should be in stratum 2");

        // No ties in this simple example
        assert!(!result.tie_warning(), "No tie warning expected for tiny example");
    }

    #[test]
    fn test_bcd_specific_tiny_example() {
        let tiny = load_tiny_example();
        let tableau = Tableau::parse(&tiny).unwrap();

        let result = tableau.run_bcd(true);

        // Specific BCD should also succeed on tiny example
        assert!(result.success(), "Specific BCD should find a valid ranking");

        // Same result as regular BCD for this example
        assert_eq!(result.get_stratum(0).unwrap(), 1, "*NoOns should be in stratum 1");
        assert_eq!(result.get_stratum(1).unwrap(), 1, "*Coda should be in stratum 1");
        assert_eq!(result.get_stratum(2).unwrap(), 2, "Max should be in stratum 2");
        assert_eq!(result.get_stratum(3).unwrap(), 2, "Dep should be in stratum 2");
    }

    #[test]
    fn test_bcd_log_capture() {
        crate::clear_log();
        let tiny = load_tiny_example();
        let tableau = Tableau::parse(&tiny).unwrap();
        let _result = tableau.run_bcd(false);
        let log = crate::get_log();

        // Header
        assert!(log.contains("Application of Biased Constraint Demotion"), "Should have BCD header");
        // Stratum processing
        assert!(log.contains("Now doing Stratum #1"), "Should log stratum 1");
        assert!(log.contains("Now doing Stratum #2"), "Should log stratum 2");
        // Faithfulness delay
        assert!(log.contains("Faithfulness Delay:"), "Should have faithfulness delay section");
        assert!(log.contains("Markedness constraint"), "Should log markedness ranking");
        assert!(log.contains("Faithfulness constraints are excluded"), "Should note faith exclusion");
        // Results summary
        assert!(log.contains("Results so far:"), "Should have results summary");
        // Success
        assert!(log.contains("SUCCEEDED"), "Should log success");
    }

    #[test]
    fn test_ilokano_bcd_specific_subsets() {
        // Specific BCD should detect Max-stem ⊆ Max and delay Max to a later stratum.
        // VB6's "Favor Specificity" menu item is hidden (Visible=0), so there is no
        // valid VB6 golden file to compare against — this test validates Rust's behavior
        // directly against the algorithm specification.
        let text = std::fs::read_to_string("../examples/IlokanoHiatusResolution.txt")
            .expect("Failed to load IlokanoHiatusResolution.txt");
        let tableau = Tableau::parse(&text).unwrap();

        let max_idx = tableau.constraints.iter().position(|c| c.abbrev() == "Max").unwrap();
        let max_stem_idx = tableau.constraints.iter().position(|c| c.abbrev() == "Max-stem").unwrap();

        // Max-stem's violations should be a subset of Max's violations
        let subsets = super::locate_violation_subsets(&tableau);
        assert!(subsets[max_stem_idx][max_idx], "Max-stem should be a subset of Max");
        assert!(!subsets[max_idx][max_stem_idx], "Max should NOT be a subset of Max-stem");

        let result_plain = tableau.run_bcd(false);
        let result_specific = tableau.run_bcd(true);

        // In plain BCD, Max and Max-stem land in the same stratum
        assert_eq!(result_plain.get_stratum(max_idx), result_plain.get_stratum(max_stem_idx));

        // In specific BCD, Max-stem ranks earlier (more specific) and Max is delayed
        let max_stratum = result_specific.get_stratum(max_idx).unwrap();
        let max_stem_stratum = result_specific.get_stratum(max_stem_idx).unwrap();
        assert!(max_stratum > max_stem_stratum,
            "Specific BCD should rank Max (stratum {}) after Max-stem (stratum {})",
            max_stratum, max_stem_stratum);
    }

}
