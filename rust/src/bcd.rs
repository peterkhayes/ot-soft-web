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

/// Simulate ranking a faithfulness subset and count how many markedness constraints
/// are subsequently freed up (become non-demotable).
///
/// Reproduces VB6 BCD.bas:CheckMarkednessRelease
fn check_markedness_release(
    tableau: &Tableau,
    faith_subset: &[usize],
    is_faithfulness: &[bool],
    current_stratum: usize,
    strata: &[usize],
    still_informative: &[Vec<bool>],
) -> usize {
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

/// Recursively enumerate all subsets of `rankable_faith` of size `target_size`,
/// evaluate each, and track the best (most markedness freed).
///
/// Reproduces VB6 BCD.bas:EvaluateFaithSubsets
#[allow(clippy::too_many_arguments)]
fn evaluate_faith_subsets(
    rankable_faith: &[usize],
    target_size: usize,
    current_subset: &mut Vec<usize>,
    start_idx: usize,
    best_subset: &mut Vec<usize>,
    best_mark_count: &mut usize,
    tie_flag: &mut bool,
    tableau: &Tableau,
    is_faithfulness: &[bool],
    current_stratum: usize,
    strata: &[usize],
    still_informative: &[Vec<bool>],
) {
    if current_subset.len() == target_size {
        // Evaluate this complete subset
        let mark_count = check_markedness_release(
            tableau,
            current_subset,
            is_faithfulness,
            current_stratum,
            strata,
            still_informative,
        );

        // Check for tie (only non-zero ties count)
        if mark_count == *best_mark_count && *best_mark_count > 0 {
            *tie_flag = true;
        }

        if mark_count > *best_mark_count {
            *tie_flag = false;
            *best_mark_count = mark_count;
            *best_subset = current_subset.clone();
        }

        return;
    }

    // Recurse, building subsets
    let remaining_needed = target_size - current_subset.len();
    let max_start = rankable_faith.len().saturating_sub(remaining_needed - 1);

    for idx in start_idx..max_start {
        current_subset.push(rankable_faith[idx]);
        evaluate_faith_subsets(
            rankable_faith,
            target_size,
            current_subset,
            idx + 1,
            best_subset,
            best_mark_count,
            tie_flag,
            tableau,
            is_faithfulness,
            current_stratum,
            strata,
            still_informative,
        );
        current_subset.pop();
    }
}

impl Tableau {
    /// Run Biased Constraint Demotion to find a ranking.
    ///
    /// `specific_bcd`: if true, favors more specific faithfulness constraints
    /// (the mnuSpecificBCD option in VB6).
    pub fn run_bcd(&self, specific_bcd: bool) -> RCDResult {
        let nc = self.constraints.len();

        // Detect faithfulness constraints
        let is_faithfulness: Vec<bool> = self.constraints.iter()
            .map(|c| c.is_faithfulness())
            .collect();

        // Precompute violation subsets if using specific BCD
        let violation_subsets = if specific_bcd {
            locate_violation_subsets(self)
        } else {
            vec![]
        };

        // Initialize state
        let mut strata = vec![0usize; nc];
        let mut current_stratum = 0usize;

        // Build still_informative[form_idx][rival_idx]
        let mut still_informative: Vec<Vec<bool>> = self.forms.iter().map(|form| {
            form.candidates.iter().map(|_| true).collect()
        }).collect();

        // Tie tracking
        let mut upper_tie_flag = false;
        let mut tie_flag = false;

        // Helper: count still-informative loser pairs
        let count_pairs = |si: &Vec<Vec<bool>>| -> usize {
            self.forms.iter().enumerate()
                .map(|(fi, form)| {
                    form.candidates.iter().enumerate()
                        .filter(|&(ri, c)| c.frequency == 0 && si[fi][ri])
                        .count()
                })
                .sum()
        };

        crate::ot_log!("Starting BCD with {} pairs", count_pairs(&still_informative));

        loop {
            current_stratum += 1;

            // Propagate tie flag
            if tie_flag {
                upper_tie_flag = true;
            }

            // Initialize demotable, active, subsetted
            let mut demotable = vec![false; nc];
            let mut active = vec![false; nc];
            let mut subsetted = vec![false; nc];

            // Mark demotable and active constraints
            for (form_idx, form) in self.forms.iter().enumerate() {
                let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                    Some(w) => w,
                    None => continue,
                };
                for (rival_idx, rival) in form.candidates.iter().enumerate() {
                    if rival.frequency > 0 {
                        continue;
                    }
                    if !still_informative[form_idx][rival_idx] {
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

            // FAITHFULNESS DELAY: Try to rank markedness constraints first
            let mut rankable_marked = false;
            let mut faith_is_active = false;

            for c_idx in 0..nc {
                if strata[c_idx] == 0 && !demotable[c_idx] {
                    if is_faithfulness[c_idx] && active[c_idx] {
                        faith_is_active = true;
                    } else if !is_faithfulness[c_idx] {
                        // Non-demotable markedness: install in current stratum
                        rankable_marked = true;
                        strata[c_idx] = current_stratum;
                    }
                }
            }

            if !rankable_marked {
                // No markedness could be ranked — must deal with faithfulness
                if faith_is_active {
                    // FAVOR SPECIFICITY (if specific BCD)
                    if specific_bcd {
                        for c_idx in 0..nc {
                            if strata[c_idx] != 0 || !is_faithfulness[c_idx] {
                                continue;
                            }
                            for inner in 0..nc {
                                if inner == c_idx || strata[inner] != 0 || !is_faithfulness[inner] {
                                    continue;
                                }
                                // If inner's violations are a subset of c_idx's,
                                // then c_idx is subsetted (blocked by more specific inner)
                                if violation_subsets[inner][c_idx] {
                                    subsetted[c_idx] = true;
                                    break;
                                }
                            }
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
                        // Shouldn't happen since faith_is_active is true, but handle gracefully
                        // Rank all remaining non-demotable
                        for c_idx in 0..nc {
                            if strata[c_idx] == 0 && !demotable[c_idx] {
                                strata[c_idx] = current_stratum;
                            }
                        }
                    } else {
                        // FIND MINIMAL FAITHFULNESS SET
                        let mut best_subset: Vec<usize> = Vec::new();
                        let mut best_mark_count = 0usize;
                        tie_flag = false;

                        for subset_size in 1..=rankable_faith.len() {
                            let mut current_subset = Vec::new();
                            evaluate_faith_subsets(
                                &rankable_faith,
                                subset_size,
                                &mut current_subset,
                                0,
                                &mut best_subset,
                                &mut best_mark_count,
                                &mut tie_flag,
                                self,
                                &is_faithfulness,
                                current_stratum,
                                &strata,
                                &still_informative,
                            );
                            if best_mark_count > 0 {
                                break;
                            }
                        }

                        if best_mark_count == 0 {
                            // No subset released any markedness; rank all rankable faith
                            for &c_idx in &rankable_faith {
                                strata[c_idx] = current_stratum;
                            }
                        } else {
                            // Rank the best subset
                            for &c_idx in &best_subset {
                                strata[c_idx] = current_stratum;
                            }
                        }
                    }
                } else {
                    // No active faithfulness — rank all remaining non-demotable (terminal case)
                    for c_idx in 0..nc {
                        if strata[c_idx] == 0 && !demotable[c_idx] {
                            strata[c_idx] = current_stratum;
                        }
                    }
                }
            }

            // CHECK TERMINATION
            let mut unranked_remain = false;
            let mut a_constraint_was_ranked = false;

            for &c_stratum in &strata {
                if c_stratum == 0 {
                    unranked_remain = true;
                } else if c_stratum == current_stratum {
                    a_constraint_was_ranked = true;
                }
            }

            crate::ot_log!("After stratum {}: {} pairs remaining",
                current_stratum, count_pairs(&still_informative));

            if !unranked_remain {
                crate::ot_log!("BCD SUCCEEDED with {} strata", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum, true);
                result.set_tie_warning(upper_tie_flag || tie_flag);
                result.compute_extra_analyses(self);
                return result;
            }

            if !a_constraint_was_ranked {
                crate::ot_log!("BCD FAILED: no constraints added to stratum {}", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum - 1, false);
                result.set_tie_warning(upper_tie_flag || tie_flag);
                return result;
            }

            // UPDATE INFORMATIVENESS
            for (c_idx, &c_stratum) in strata.iter().enumerate() {
                if c_stratum != current_stratum {
                    continue;
                }
                for (form_idx, form) in self.forms.iter().enumerate() {
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

            // Safety check
            if current_stratum > nc {
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
        std::fs::read_to_string("../examples/TinyIllustrativeFile/input.txt")
            .expect("Failed to load examples/TinyIllustrativeFile/input.txt")
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

}
