//! Low Faithfulness Constraint Demotion (LFCD) algorithm
//!
//! LFCD extends RCD with a stronger bias toward ranking Faithfulness constraints low.
//! Each stratum-building step applies four heuristics in order:
//!
//! 1. **Favor Markedness**: If any non-demotable Markedness constraint exists, install it
//!    and block all Faithfulness.
//! 2. **Favor Activeness**: Mark Faithfulness constraints as active if they prefer the winner
//!    for at least one non-superset informative pair.  Demote inactive ones when active ones exist.
//! 3. **Favor Specificity**: Demote a Faithfulness constraint if a more specific one exists
//!    (where "more specific" means its violations are a subset).
//! 4. **Favor Autonomy**: Among remaining Faithfulness, install those with the fewest
//!    "helpers" (other constraints that also prefer the winner for the same pair).
//!
//! Source: LowFaithfulnessConstraintDemotion.bas (programmed by Bruce Hayes)

use crate::bcd::locate_violation_subsets;
use crate::rcd::RCDResult;
use crate::tableau::Tableau;

/// Check whether a rival candidate has a superset of the winner's violations.
///
/// Returns true if for every constraint c, winner_viols[c] <= rival_viols[c].
/// Such a rival will lose regardless of constraint ranking, so it carries no
/// useful ranking information.
///
/// Reproduces VB6 LFCD.bas:Superset
fn is_superset_rival(winner_viols: &[usize], rival_viols: &[usize]) -> bool {
    winner_viols.iter()
        .zip(rival_viols.iter())
        .all(|(&w, &r)| w <= r)
}

impl Tableau {
    /// Run Low Faithfulness Constraint Demotion to find a ranking.
    pub fn run_lfcd(&self) -> RCDResult {
        self.run_lfcd_with_apriori(&[])
    }

    /// Run LFCD enforcing a priori constraint rankings.
    ///
    /// `apriori[i][j] = true` means constraint i must rank above constraint j.
    ///
    /// Reproduces VB6 LowFaithfulnessConstraintDemotion.bas:Main
    pub fn run_lfcd_with_apriori(&self, apriori: &[Vec<bool>]) -> RCDResult {
        let nc = self.constraints.len();

        // Detect faithfulness constraints
        let is_faithfulness: Vec<bool> = self.constraints.iter()
            .map(|c| c.is_faithfulness())
            .collect();

        // Precompute violation subsets for specificity and autonomy checks
        let violation_subsets = locate_violation_subsets(self);

        // Initialize state
        let mut strata = vec![0usize; nc];
        let mut current_stratum = 0usize;

        // still_informative[form_idx][rival_idx]: whether this pair is still relevant
        let mut still_informative: Vec<Vec<bool>> = self.forms.iter().map(|form| {
            form.candidates.iter().map(|_| true).collect()
        }).collect();

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

        crate::ot_log!("Starting LFCD with {} pairs", count_pairs(&still_informative));

        // Safety: at most nc+1 iterations
        loop {
            current_stratum += 1;
            if current_stratum > nc + 1 {
                crate::ot_log!("LFCD FAILED: safety limit reached at stratum {}", current_stratum);
                return RCDResult::new(strata, current_stratum - 1, false);
            }

            // Per-stratum arrays (reset each iteration)
            let mut demotable = vec![false; nc];
            let mut active = vec![false; nc];
            // num_helpers initialized to usize::MAX as a sentinel for "not yet assessed"
            let mut num_helpers = vec![usize::MAX; nc];

            // ===== AVOID PREFERENCE FOR LOSERS =====
            //
            // A constraint is demotable if it prefers a loser over a winner in any
            // still-informative pair.
            for (form_idx, form) in self.forms.iter().enumerate() {
                let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                    Some(w) => w,
                    None => continue,
                };
                for (rival_idx, rival) in form.candidates.iter().enumerate() {
                    if rival.frequency > 0 || !still_informative[form_idx][rival_idx] {
                        continue;
                    }
                    for c_idx in 0..nc {
                        if strata[c_idx] == 0
                            && winner.violations[c_idx] > rival.violations[c_idx]
                        {
                            demotable[c_idx] = true;
                        }
                    }
                }
            }

            // ===== ENFORCE A PRIORI RANKINGS =====
            //
            // Any constraint that is a priori dominated by an unranked constraint
            // cannot join the current stratum.
            if !apriori.is_empty() {
                for outer in 0..nc {
                    if strata[outer] == 0 {
                        for inner in 0..nc {
                            if apriori[outer][inner] {
                                demotable[inner] = true;
                            }
                        }
                    }
                }
            }

            // ===== FAVOR MARKEDNESS =====
            //
            // If any non-demotable Markedness constraint exists, install it and
            // shut out all Faithfulness constraints for this stratum.
            let there_is_rankable_markedness = (0..nc).any(|c_idx| {
                strata[c_idx] == 0 && !demotable[c_idx] && !is_faithfulness[c_idx]
            });

            if there_is_rankable_markedness {
                // Shut out all Faithfulness
                for c_idx in 0..nc {
                    if is_faithfulness[c_idx] {
                        demotable[c_idx] = true;
                    }
                }
                // Fall through to ReportStrata
            } else {
                // ===== FAVOR ACTIVENESS =====
                //
                // A Faithfulness constraint is "active" if it prefers the winner for at
                // least one still-informative, non-superset pair.  Superset rivals are
                // excluded because they carry no ranking information.
                let mut at_least_one_faith_active = false;

                for c_idx in 0..nc {
                    if strata[c_idx] != 0 || !is_faithfulness[c_idx] || demotable[c_idx] {
                        continue;
                    }
                    'find_active_pair: for (form_idx, form) in self.forms.iter().enumerate() {
                        let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                            Some(w) => w,
                            None => continue,
                        };
                        for (rival_idx, rival) in form.candidates.iter().enumerate() {
                            if rival.frequency > 0 || !still_informative[form_idx][rival_idx] {
                                continue;
                            }
                            if is_superset_rival(&winner.violations, &rival.violations) {
                                continue;
                            }
                            if rival.violations[c_idx] > winner.violations[c_idx] {
                                active[c_idx] = true;
                                at_least_one_faith_active = true;
                                break 'find_active_pair; // Only need one pair for proof
                            }
                        }
                    }
                }

                if at_least_one_faith_active {
                    // Demote inactive Faithfulness (since active ones are available)
                    for c_idx in 0..nc {
                        if strata[c_idx] == 0
                            && is_faithfulness[c_idx]
                            && !demotable[c_idx]
                            && !active[c_idx]
                        {
                            demotable[c_idx] = true;
                        }
                    }

                    // ===== FAVOR SPECIFICITY =====
                    //
                    // Block a Faithfulness constraint C if any other non-demotable
                    // Faithfulness constraint InnerC has subset[InnerC][C] = true,
                    // meaning InnerC's violations are a subset of C's (InnerC is more specific).
                    for c_idx in 0..nc {
                        if strata[c_idx] != 0
                            || demotable[c_idx]
                            || !active[c_idx]
                            || !is_faithfulness[c_idx]
                        {
                            continue;
                        }
                        for inner in 0..nc {
                            if inner == c_idx
                                || strata[inner] != 0
                                || !is_faithfulness[inner]
                                || demotable[inner]
                            {
                                continue;
                            }
                            if violation_subsets[inner][c_idx] {
                                demotable[c_idx] = true;
                                break;
                            }
                        }
                    }

                    // ===== FAVOR AUTONOMY =====
                    //
                    // Among remaining non-demotable Faithfulness constraints, prefer those
                    // that rule out a loser with the fewest "helpers" (other constraints
                    // also preferring the winner for the same pair).
                    //
                    // A helper is excluded if it is a Faithfulness constraint whose
                    // violations are a superset of the target constraint's (i.e., it is a
                    // broader version of the same constraint).
                    for (form_idx, form) in self.forms.iter().enumerate() {
                        let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
                            Some(w) => w,
                            None => continue,
                        };
                        for (rival_idx, rival) in form.candidates.iter().enumerate() {
                            if rival.frequency > 0 || !still_informative[form_idx][rival_idx] {
                                continue;
                            }
                            if is_superset_rival(&winner.violations, &rival.violations) {
                                continue;
                            }

                            for c_idx in 0..nc {
                                if strata[c_idx] != 0
                                    || demotable[c_idx]
                                    || !is_faithfulness[c_idx]
                                {
                                    continue;
                                }
                                // Only process constraints that prefer winner for this pair
                                if rival.violations[c_idx] <= winner.violations[c_idx] {
                                    continue;
                                }

                                // Count helpers: other constraints that also prefer winner,
                                // excluding "superset faithfulness" relatives of c_idx
                                let mut local_helpers = 0usize;
                                for inner in 0..nc {
                                    if inner == c_idx {
                                        continue;
                                    }
                                    if rival.violations[inner] > winner.violations[inner] {
                                        // Superset faithfulness constraints don't count
                                        let is_superset_faith =
                                            violation_subsets[c_idx][inner]
                                                && is_faithfulness[inner];
                                        if !is_superset_faith {
                                            local_helpers += 1;
                                        }
                                    }
                                }

                                // Record the minimum helpers seen for this constraint
                                if local_helpers < num_helpers[c_idx] {
                                    num_helpers[c_idx] = local_helpers;
                                }
                            }
                        }
                    }

                    // Find the overall lowest helper count among non-demotable Faithfulness.
                    // Unassessed constraints (num_helpers == usize::MAX) are excluded.
                    let mut lowest_helpers = usize::MAX;
                    for c_idx in 0..nc {
                        if is_faithfulness[c_idx] && !demotable[c_idx] && strata[c_idx] == 0
                            && num_helpers[c_idx] < lowest_helpers
                        {
                            lowest_helpers = num_helpers[c_idx];
                        }
                    }

                    // Demote constraints with more helpers than the minimum
                    for c_idx in 0..nc {
                        if strata[c_idx] == 0
                            && !demotable[c_idx]
                            && num_helpers[c_idx] > lowest_helpers
                        {
                            demotable[c_idx] = true;
                        }
                    }
                }
                // else: no active Faithfulness — fall through to ReportStrata.
                // Non-loser-favoring inactive Faithfulness remain non-demotable
                // and will be assigned at ReportStrata (last-stratum case).
            }

            // ===== REPORT STRATA (termination check) =====
            //
            // Case I:   Some non-demotable, some demotable → assign non-demotable, continue
            // Case II:  All demotable → failure
            // Case III: None demotable → success (all go into current stratum)
            let mut some_are_non_demotable = false;
            let mut some_are_demotable = false;

            for c_idx in 0..nc {
                if strata[c_idx] == 0 {
                    if !demotable[c_idx] {
                        some_are_non_demotable = true;
                        strata[c_idx] = current_stratum;
                    } else {
                        some_are_demotable = true;
                    }
                }
            }

            crate::ot_log!("After stratum {}: {} pairs remaining",
                current_stratum, count_pairs(&still_informative));

            if !some_are_demotable {
                // Case III: success
                crate::ot_log!("LFCD SUCCEEDED with {} strata", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum, true);
                result.compute_extra_analyses(self);
                return result;
            } else if !some_are_non_demotable {
                // Case II: failure — assign remaining to current stratum for diagnostics
                crate::ot_log!("LFCD FAILED: all remaining constraints are demotable at stratum {}", current_stratum);
                for c_idx in 0..nc {
                    if strata[c_idx] == 0 {
                        strata[c_idx] = current_stratum;
                    }
                }
                return RCDResult::new(strata, current_stratum, false);
            }

            // Case I: some non-demotable constraints were assigned; continue loop

            // ===== UPDATE INFORMATIVENESS =====
            //
            // A pair is no longer informative once a ranked constraint prefers the winner.
            for c_idx in 0..nc {
                if strata[c_idx] != current_stratum {
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
        }
    }
}

#[cfg(test)]
mod tests {
    use crate::tableau::Tableau;

    fn load_tiny_example() -> String {
        std::fs::read_to_string("../examples/tiny/input.txt")
            .expect("Failed to load examples/tiny/input.txt")
    }

    #[test]
    fn test_lfcd_tiny_example() {
        let tiny = load_tiny_example();
        let tableau = Tableau::parse(&tiny).unwrap();

        let result = tableau.run_lfcd();

        // LFCD should succeed on the tiny example
        assert!(result.success(), "LFCD should find a valid ranking");

        // *NoOns and *Coda are Markedness → stratum 1 (FAVOR MARKEDNESS)
        // Max and Dep are Faithfulness → stratum 2
        assert_eq!(result.get_stratum(0).unwrap(), 1, "*NoOns should be in stratum 1");
        assert_eq!(result.get_stratum(1).unwrap(), 1, "*Coda should be in stratum 1");
        assert_eq!(result.get_stratum(2).unwrap(), 2, "Max should be in stratum 2");
        assert_eq!(result.get_stratum(3).unwrap(), 2, "Dep should be in stratum 2");
    }

    #[test]
    fn test_lfcd_output_format() {
        let tiny = load_tiny_example();
        let tableau = Tableau::parse(&tiny).unwrap();
        let result = tableau.run_lfcd();

        let output = result.format_output_with_algorithm(
            &tableau,
            "test.txt",
            "Low Faithfulness Constraint Demotion",
        );
        assert!(output.contains(
            "Results of Applying Low Faithfulness Constraint Demotion to test.txt"
        ));
        assert!(!output.contains("Caution"));
    }
}
