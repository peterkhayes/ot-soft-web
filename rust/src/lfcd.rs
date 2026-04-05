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
    crate::ot_log!("Avoid Preference For Losers:");
    let mut found_one = false;
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
                    // Log only the first evidence for each constraint
                    if !demotable[c_idx] {
                        crate::ot_log!("  {} is excluded from stratum; prefers loser *[{}] for /{}/.  ",
                            tableau.constraints[c_idx].abbrev(),
                            rival.form,
                            form.input);
                        found_one = true;
                    }
                    demotable[c_idx] = true;
                }
            }
        }
    }
    if !found_one {
        crate::ot_log!("  Search found no unranked constraints that prefer losers.");
    }
}

/// Demote constraints that are a priori dominated by an unranked constraint.
fn enforce_apriori_demotions(
    tableau: &Tableau,
    apriori: &[Vec<bool>],
    strata: &[usize],
    demotable: &mut [bool],
) {
    crate::ot_log!("Enforce a priori rankings:");
    if apriori.is_empty() {
        crate::ot_log!("  Search found no constraints that must be demoted due to an a priori ranking.");
        return;
    }
    let nc = strata.len();
    let mut found_one = false;
    for outer in 0..nc {
        if strata[outer] == 0 {
            for inner in 0..nc {
                if apriori[outer][inner] && !demotable[inner] {
                    crate::ot_log!("  {} is excluded from stratum.", tableau.constraints[inner].abbrev());
                    crate::ot_log!("    It is dominated a priori by {}, which has yet to be ranked.",
                        tableau.constraints[outer].abbrev());
                    found_one = true;
                }
                if apriori[outer][inner] {
                    demotable[inner] = true;
                }
            }
        }
    }
    if !found_one {
        crate::ot_log!("  Search found no constraints that must be demoted due to an a priori ranking.");
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

    // --- Favor Markedness ---
    crate::ot_log!("Favor Markedness:");
    let there_is_rankable_markedness = (0..nc).any(|c_idx| {
        strata[c_idx] == 0 && !demotable[c_idx] && !is_faithfulness[c_idx]
    });

    if there_is_rankable_markedness {
        for c_idx in 0..nc {
            if strata[c_idx] == 0 && !demotable[c_idx] && !is_faithfulness[c_idx] {
                crate::ot_log!("  {} is a Markedness constraint that favors no losers, joins new stratum.",
                    tableau.constraints[c_idx].abbrev());
            }
        }
        for c_idx in 0..nc {
            if is_faithfulness[c_idx] {
                demotable[c_idx] = true;
            }
        }
        crate::ot_log!("  Faithfulness constraints are excluded from stratum.");
        return;
    }
    crate::ot_log!("  There are no rankable Markedness constraints.");

    // --- Favor Activeness ---
    crate::ot_log!("");
    crate::ot_log!("Favor Activeness:");
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
                    crate::ot_log!("  {} is shown to be active by ruling out *[{}] for /{}/.  ",
                        tableau.constraints[c_idx].abbrev(), rival.form, form.input);
                    break 'find_active_pair;
                }
            }
        }
    }

    if !at_least_one_faith_active {
        crate::ot_log!("  Only remaining rankable constraints are inactive Faithfulness constraints.");
        crate::ot_log!("  All of them join the current stratum:");
        for c_idx in 0..nc {
            if strata[c_idx] == 0 && is_faithfulness[c_idx] && !demotable[c_idx] {
                crate::ot_log!("    {}", tableau.constraints[c_idx].abbrev());
            }
        }
        return;
    }

    // Demote inactive Faithfulness (since active ones are available)
    let mut all_active = true;
    for c_idx in 0..nc {
        if strata[c_idx] == 0 && is_faithfulness[c_idx] && !demotable[c_idx] && !active[c_idx] {
            crate::ot_log!("  {} is excluded from stratum because it is inactive.",
                tableau.constraints[c_idx].abbrev());
            demotable[c_idx] = true;
            all_active = false;
        }
    }
    if all_active {
        crate::ot_log!("  All unranked Faithfulness constraints are active.");
    }

    // --- Favor Specificity ---
    crate::ot_log!("");
    crate::ot_log!("Favor Specificity:");
    let mut found_specificity = false;
    for c_idx in 0..nc {
        if strata[c_idx] != 0 || demotable[c_idx] || !active[c_idx] || !is_faithfulness[c_idx] {
            continue;
        }
        for inner in 0..nc {
            if inner == c_idx || strata[inner] != 0 || !is_faithfulness[inner] || demotable[inner] {
                continue;
            }
            if violation_subsets[inner][c_idx] {
                crate::ot_log!("  {} is excluded from stratum because {} is more specific.",
                    tableau.constraints[c_idx].abbrev(), tableau.constraints[inner].abbrev());
                demotable[c_idx] = true;
                found_specificity = true;
                break;
            }
        }
    }
    if !found_specificity {
        crate::ot_log!("  (no cases found)");
    }

    // --- Favor Autonomy ---
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
    crate::ot_log!("");
    crate::ot_log!("Favor Autonomy:");
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
                let mut helper_names = Vec::new();
                for inner in 0..nc {
                    if inner == c_idx {
                        continue;
                    }
                    if rival.violations[inner] > winner.violations[inner] {
                        let is_superset_faith =
                            violation_subsets[c_idx][inner] && is_faithfulness[inner];
                        if !is_superset_faith {
                            local_helpers += 1;
                            helper_names.push(tableau.constraints[inner].abbrev().to_string());
                        }
                    }
                }

                if local_helpers < num_helpers[c_idx] {
                    num_helpers[c_idx] = local_helpers;
                    let plural = if local_helpers == 1 { "" } else { "s" };
                    crate::ot_log!("  {} is assigned {} helper{}, based on /{}/ -/-> *[{}].",
                        tableau.constraints[c_idx].abbrev(), local_helpers, plural,
                        form.input, rival.form);
                    for h in &helper_names {
                        crate::ot_log!("    {}", h);
                    }
                }
            }
        }
    }

    let lowest_helpers = (0..nc)
        .filter(|&c| is_faithfulness[c] && !demotable[c] && strata[c] == 0)
        .map(|c| num_helpers[c])
        .min()
        .unwrap_or(usize::MAX);

    if lowest_helpers < usize::MAX {
        crate::ot_log!("");
        crate::ot_log!("  Lowest number of helpers:  {}", lowest_helpers);
    } else {
        crate::ot_log!("");
        crate::ot_log!("  (none found; no non-superset Faithfulness constraint favors a winner)");
    }

    for c_idx in 0..nc {
        if strata[c_idx] == 0 && !demotable[c_idx] {
            if num_helpers[c_idx] <= lowest_helpers {
                crate::ot_log!("  Constraint {} joins the current stratum, having {} helpers.",
                    tableau.constraints[c_idx].abbrev(), num_helpers[c_idx]);
            } else {
                let plural = if num_helpers[c_idx] == 1 { "helper" } else { "helpers" };
                crate::ot_log!("  Constraint {} is excluded from stratum because it has {} {}.",
                    tableau.constraints[c_idx].abbrev(), num_helpers[c_idx], plural);
                demotable[c_idx] = true;
            }
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

        crate::ot_log!("****** Application of Low Faithfulness Constraint Demotion ******");
        crate::ot_log!("Starting LFCD with {} pairs", count_pairs(&still_informative));

        loop {
            current_stratum += 1;
            if current_stratum > nc + 1 {
                crate::ot_log!("");
                crate::ot_log!("Ranking has failed. This constraint set is unable to derive only winners.");
                crate::ot_log!("LFCD FAILED: safety limit reached at stratum {}", current_stratum);
                return RCDResult::new(strata, current_stratum - 1, false);
            }

            crate::ot_log!("");
            crate::ot_log!("****** Now doing Stratum #{} ******", current_stratum);

            let mut demotable = vec![false; nc];

            mark_loser_preferring(self, &strata, &still_informative, &mut demotable);
            enforce_apriori_demotions(self, apriori, &strata, &mut demotable);
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

            self.log_results_so_far(&strata, current_stratum);
            crate::ot_log!("After stratum {}: {} pairs remaining",
                current_stratum, count_pairs(&still_informative));

            if !some_are_demotable {
                crate::ot_log!("");
                crate::ot_log!("Ranking is complete and yields successful grammar.");
                crate::ot_log!("LFCD SUCCEEDED with {} strata", current_stratum);
                let mut result = RCDResult::new(strata, current_stratum, true);
                result.compute_extra_analyses(self, apriori, false);
                return result;
            } else if !some_are_non_demotable {
                crate::ot_log!("");
                crate::ot_log!("Ranking has failed. This constraint set is unable to derive only winners.");
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
    fn test_lfcd_log_capture() {
        crate::clear_log();
        let tiny = load_tiny_example();
        let tableau = Tableau::parse(&tiny).unwrap();
        let _result = tableau.run_lfcd();
        let log = crate::get_log();

        // Header
        assert!(log.contains("Application of Low Faithfulness Constraint Demotion"), "Should have LFCD header");
        // Stratum processing
        assert!(log.contains("Now doing Stratum #1"), "Should log stratum 1");
        // Heuristic sections
        assert!(log.contains("Avoid Preference For Losers:"), "Should have loser-preferring section");
        assert!(log.contains("Favor Markedness:"), "Should have markedness section");
        assert!(log.contains("is a Markedness constraint"), "Should log markedness ranking");
        assert!(log.contains("Faithfulness constraints are excluded"), "Should note faith exclusion");
        // Results summary
        assert!(log.contains("Results so far:"), "Should have results summary");
        // Success
        assert!(log.contains("SUCCEEDED"), "Should log success");
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
