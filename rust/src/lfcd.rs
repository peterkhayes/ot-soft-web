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
fn is_superset_rival(winner_viols: &[i32], rival_viols: &[i32]) -> bool {
    winner_viols.iter()
        .zip(rival_viols.iter())
        .all(|(&w, &r)| w <= r)
}

/// Mark constraints as demotable if they prefer a loser in any still-informative pair.
fn mark_loser_preferring(
    tableau: &Tableau,
    strata: &[usize],
    still_informative: &[Vec<bool>],
    demotable: &mut [bool],
) {
    for (form_idx, form) in tableau.forms.iter().enumerate() {
        let winner = match form.candidates.iter().find(|c| c.frequency > 0) {
            Some(w) => w,
            None => continue,
        };
        for (rival_idx, rival) in form.candidates.iter().enumerate() {
            if rival.frequency > 0 || !still_informative[form_idx][rival_idx] {
                continue;
            }
            for (c_idx, &stratum) in strata.iter().enumerate() {
                if stratum == 0 && winner.violations[c_idx] > rival.violations[c_idx] {
                    demotable[c_idx] = true;
                }
            }
        }
    }
}

/// Demote constraints that are a priori dominated by an unranked constraint.
fn enforce_apriori_demotions(apriori: &[Vec<bool>], strata: &[usize], demotable: &mut [bool]) {
    if apriori.is_empty() {
        return;
    }
    let nc = strata.len();
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

/// Apply the four LFCD faithfulness heuristics: markedness, activeness, specificity, autonomy.
fn apply_faithfulness_heuristics(
    tableau: &Tableau,
    strata: &[usize],
    still_informative: &[Vec<bool>],
    is_faithfulness: &[bool],
    violation_subsets: &[Vec<bool>],
    demotable: &mut [bool],
) {
    let nc = strata.len();

    // Favor Markedness: if any non-demotable Markedness exists, shut out all Faithfulness
    let there_is_rankable_markedness = (0..nc).any(|c_idx| {
        strata[c_idx] == 0 && !demotable[c_idx] && !is_faithfulness[c_idx]
    });

    if there_is_rankable_markedness {
        for c_idx in 0..nc {
            if is_faithfulness[c_idx] {
                demotable[c_idx] = true;
            }
        }
        return;
    }

    // Favor Activeness: mark Faithfulness constraints as active if they prefer
    // the winner for at least one still-informative, non-superset pair.
    let mut active = vec![false; nc];
    let mut at_least_one_faith_active = false;

    for c_idx in 0..nc {
        if strata[c_idx] != 0 || !is_faithfulness[c_idx] || demotable[c_idx] {
            continue;
        }
        'find_active_pair: for (form_idx, form) in tableau.forms.iter().enumerate() {
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
                    break 'find_active_pair;
                }
            }
        }
    }

    if !at_least_one_faith_active {
        return;
    }

    // Demote inactive Faithfulness (since active ones are available)
    for c_idx in 0..nc {
        if strata[c_idx] == 0 && is_faithfulness[c_idx] && !demotable[c_idx] && !active[c_idx] {
            demotable[c_idx] = true;
        }
    }

    // Favor Specificity: demote a constraint if a more specific one exists
    for c_idx in 0..nc {
        if strata[c_idx] != 0 || demotable[c_idx] || !active[c_idx] || !is_faithfulness[c_idx] {
            continue;
        }
        for inner in 0..nc {
            if inner == c_idx || strata[inner] != 0 || !is_faithfulness[inner] || demotable[inner] {
                continue;
            }
            if violation_subsets[inner][c_idx] {
                demotable[c_idx] = true;
                break;
            }
        }
    }

    // Favor Autonomy: prefer constraints with fewest helpers
    apply_favor_autonomy(
        tableau, strata, still_informative, is_faithfulness, violation_subsets, demotable,
    );
}

/// Among remaining non-demotable Faithfulness constraints, prefer those that rule out
/// a loser with the fewest "helpers" (other constraints also preferring the winner).
fn apply_favor_autonomy(
    tableau: &Tableau,
    strata: &[usize],
    still_informative: &[Vec<bool>],
    is_faithfulness: &[bool],
    violation_subsets: &[Vec<bool>],
    demotable: &mut [bool],
) {
    let nc = strata.len();
    let mut num_helpers = vec![usize::MAX; nc];

    for (form_idx, form) in tableau.forms.iter().enumerate() {
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
                if strata[c_idx] != 0 || demotable[c_idx] || !is_faithfulness[c_idx] {
                    continue;
                }
                if rival.violations[c_idx] <= winner.violations[c_idx] {
                    continue;
                }

                let mut local_helpers = 0usize;
                for inner in 0..nc {
                    if inner == c_idx {
                        continue;
                    }
                    if rival.violations[inner] > winner.violations[inner] {
                        let is_superset_faith =
                            violation_subsets[c_idx][inner] && is_faithfulness[inner];
                        if !is_superset_faith {
                            local_helpers += 1;
                        }
                    }
                }

                if local_helpers < num_helpers[c_idx] {
                    num_helpers[c_idx] = local_helpers;
                }
            }
        }
    }

    let lowest_helpers = (0..nc)
        .filter(|&c| is_faithfulness[c] && !demotable[c] && strata[c] == 0)
        .map(|c| num_helpers[c])
        .min()
        .unwrap_or(usize::MAX);

    for c_idx in 0..nc {
        if strata[c_idx] == 0 && !demotable[c_idx] && num_helpers[c_idx] > lowest_helpers {
            demotable[c_idx] = true;
        }
    }
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
        let is_faithfulness: Vec<bool> = self.constraints.iter()
            .map(|c| c.is_faithfulness())
            .collect();
        let violation_subsets = locate_violation_subsets(self);

        let mut strata = vec![0usize; nc];
        let mut current_stratum = 0usize;
        let mut still_informative: Vec<Vec<bool>> = self.forms.iter().map(|form| {
            form.candidates.iter().map(|_| true).collect()
        }).collect();

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

        loop {
            current_stratum += 1;
            if current_stratum > nc + 1 {
                crate::ot_log!("LFCD FAILED: safety limit reached at stratum {}", current_stratum);
                return RCDResult::new(strata, current_stratum - 1, false);
            }

            let mut demotable = vec![false; nc];

            mark_loser_preferring(self, &strata, &still_informative, &mut demotable);
            enforce_apriori_demotions(apriori, &strata, &mut demotable);
            apply_faithfulness_heuristics(
                self, &strata, &still_informative, &is_faithfulness,
                &violation_subsets, &mut demotable,
            );

            // Assign non-demotable constraints to current stratum and check termination
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
                crate::ot_log!("LFCD SUCCEEDED with {} strata", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum, true);
                result.compute_extra_analyses(self, apriori, false);
                return result;
            } else if !some_are_non_demotable {
                crate::ot_log!("LFCD FAILED: all remaining constraints are demotable at stratum {}", current_stratum);
                for s in strata.iter_mut() {
                    if *s == 0 {
                        *s = current_stratum;
                    }
                }
                return RCDResult::new(strata, current_stratum, false);
            }

            update_informativeness(self, &strata, current_stratum, &mut still_informative);
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

}
