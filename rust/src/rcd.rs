//! Recursive Constraint Demotion (RCD) algorithm
//!
//! This module implements the RCD algorithm for finding stratified constraint
//! rankings in Optimality Theory. The algorithm iteratively ranks constraints
//! into strata by identifying which constraints never prefer losers.

use wasm_bindgen::prelude::*;
use crate::tableau::Tableau;
use crate::fred::FRedResult;

/// Classification of constraint necessity
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ConstraintNecessity {
    Necessary,
    UnnecessaryButShownForFaithfulness,
    CompletelyUnnecessary,
}

/// A mini-tableau showing a simplified winner-loser comparison
#[derive(Debug, Clone)]
pub struct MiniTableau {
    pub form_index: usize,
    pub winner_index: usize,
    pub loser_index: usize,
    pub included_constraints: Vec<usize>,
}

/// Result of running RCD algorithm
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct RCDResult {
    /// Stratum number for each constraint (1-indexed)
    constraint_strata: Vec<usize>,
    /// Total number of strata
    num_strata: usize,
    /// Whether a valid ranking was found
    success: bool,
    /// Necessity classification for each constraint (not exposed to WASM)
    #[wasm_bindgen(skip)]
    constraint_necessity: Vec<ConstraintNecessity>,
    /// FRed ranking-argumentation result (not exposed to WASM)
    #[wasm_bindgen(skip)]
    fred_result: Option<FRedResult>,
    /// Mini-tableaux for ranking arguments (not exposed to WASM)
    #[wasm_bindgen(skip)]
    mini_tableaux: Vec<MiniTableau>,
    /// Whether BCD encountered tied faithfulness subsets (arbitrary choice made)
    tie_warning: bool,
}

#[wasm_bindgen]
impl RCDResult {
    pub fn num_strata(&self) -> usize {
        self.num_strata
    }

    pub fn success(&self) -> bool {
        self.success
    }

    pub fn get_stratum(&self, constraint_index: usize) -> Option<usize> {
        self.constraint_strata.get(constraint_index).copied()
    }

    pub fn tie_warning(&self) -> bool {
        self.tie_warning
    }
}

impl RCDResult {
    /// Create a new RCDResult (used by BCD to construct results)
    pub(crate) fn new(constraint_strata: Vec<usize>, num_strata: usize, success: bool) -> Self {
        RCDResult {
            constraint_strata,
            num_strata,
            success,
            constraint_necessity: Vec::new(),
            fred_result: None,
            mini_tableaux: Vec::new(),
            tie_warning: false,
        }
    }

    /// Compute additional analyses (necessity, FRed ranking arguments, mini-tableaux)
    pub(crate) fn compute_extra_analyses(&mut self, tableau: &Tableau) {
        self.constraint_necessity = tableau.compute_constraint_necessity(self);
        self.fred_result = Some(tableau.run_fred(false)); // Skeletal Basis mode
        self.mini_tableaux = self.generate_mini_tableaux(tableau);
    }

    pub(crate) fn set_tie_warning(&mut self, value: bool) {
        self.tie_warning = value;
    }

    /// Generate mini-tableaux showing simplified ranking arguments
    fn generate_mini_tableaux(&self, tableau: &Tableau) -> Vec<MiniTableau> {
        let mut mini_tableaux = Vec::new();

        for (form_idx, form) in tableau.forms.iter().enumerate() {
            // Find the winner
            let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
                Some(idx) => idx,
                None => continue,
            };

            let winner = &form.candidates[winner_idx];

            // Check each loser
            for (loser_idx, loser) in form.candidates.iter().enumerate() {
                if loser.frequency > 0 {
                    continue; // Skip winners
                }

                // Count constraints preferring each candidate
                let mut winner_preferring = Vec::new();
                let mut loser_preferring = Vec::new();
                let mut included_constraints = Vec::new();

                for c_idx in 0..tableau.constraints.len() {
                    let w_viol = winner.violations[c_idx];
                    let l_viol = loser.violations[c_idx];

                    // Skip if neither violates OR both violate equally
                    if w_viol == l_viol {
                        continue;
                    }

                    included_constraints.push(c_idx);

                    if w_viol < l_viol {
                        winner_preferring.push(c_idx);
                    } else if l_viol < w_viol {
                        loser_preferring.push(c_idx);
                    }
                }

                // Include if exactly one winner-preferring and at least one loser-preferring
                if winner_preferring.len() == 1 && !loser_preferring.is_empty() {
                    mini_tableaux.push(MiniTableau {
                        form_index: form_idx,
                        winner_index: winner_idx,
                        loser_index: loser_idx,
                        included_constraints,
                    });
                }
            }
        }

        mini_tableaux
    }

    /// Generate formatted text output for the RCD analysis
    pub fn format_output(&self, tableau: &Tableau, filename: &str) -> String {
        self.format_output_with_algorithm(tableau, filename, "Recursive Constraint Demotion")
    }

    /// Generate formatted text output with a configurable algorithm name
    pub(crate) fn format_output_with_algorithm(&self, tableau: &Tableau, filename: &str, algorithm_name: &str) -> String {
        let mut output = String::new();

        // Header
        output.push_str(&format!("Results of Applying {} to {}\n", algorithm_name, filename));
        output.push_str("\n\n");

        // Date and version (current date/time and version)
        let now = chrono::Local::now();
        output.push_str(&format!("{}\n\n", now.format("%-m-%-d-%Y, %-I:%M %p").to_string().to_lowercase()));
        output.push_str("OTSoft 2.7, release date 2/1/2026\n");
        output.push_str("\n\n");

        if self.tie_warning {
            output.push_str("Caution: The BCD algorithm has selected arbitrarily among tied Faithfulness constraint subsets.\n");
            output.push_str("You may wish to try changing the order of the Faithfulness constraints in the input file,\n");
            output.push_str("to see whether this results in a different ranking.\n\n\n");
        }

        // Section 1: Result
        output.push_str("1. Result\n\n");

        if self.success {
            output.push_str("A ranking was found that generates the correct outputs.\n\n");
        } else {
            output.push_str("No ranking was found.\n\n");
        }

        // List strata
        for stratum in 1..=self.num_strata {
            output.push_str(&format!("   Stratum #{}\n", stratum));

            for (c_idx, &c_stratum) in self.constraint_strata.iter().enumerate() {
                if c_stratum == stratum {
                    if let Some(constraint) = tableau.get_constraint(c_idx) {
                        let full_name = constraint.full_name();
                        let abbrev = constraint.abbrev();
                        // Format: left-align full name in ~42 chars, then abbreviation
                        output.push_str(&format!("      {:<42}{}\n", full_name, abbrev));
                    }
                }
            }
        }
        output.push_str("\n");

        // Section 2: Tableaux
        output.push_str("2. Tableaux\n\n");

        for form in &tableau.forms {
            output.push_str("\n");
            output.push_str(&format!("/{}/: \n", form.input));

            // Build constraint header with stratum separators
            // Use broken bar character (U+00A6) for within-stratum, | for between-strata
            let sep_char = '\u{00A6}'; // Broken bar (¦)

            let mut header = String::new();

            // Calculate max candidate width for alignment
            let max_cand_width = form.candidates.iter()
                .map(|c| c.form.len())
                .max()
                .unwrap_or(0)
                .max(2); // At least 2 chars for marker + space

            // Add initial spacing
            for _ in 0..max_cand_width + 2 {
                header.push(' ');
            }

            for (c_idx, constraint) in tableau.constraints.iter().enumerate() {
                let c_stratum = self.constraint_strata[c_idx];

                // Add separator before this constraint
                if c_idx > 0 {
                    let prev_stratum = self.constraint_strata[c_idx - 1];
                    if c_stratum != prev_stratum {
                        header.push('|');
                    } else {
                        header.push(sep_char);
                    }
                }

                header.push_str(&constraint.abbrev());
            }
            output.push_str(&header);
            output.push_str("\n");

            // Find winner
            let winner_idx = form.candidates.iter()
                .position(|c| c.frequency > 0);

            // Output each candidate
            for (cand_idx, candidate) in form.candidates.iter().enumerate() {
                let is_winner = Some(cand_idx) == winner_idx;
                let marker = if is_winner { ">" } else { " " };

                // Find the first fatal violation (if any) for this loser
                let first_fatal_idx = if !is_winner && winner_idx.is_some() {
                    let winner = &form.candidates[winner_idx.unwrap()];
                    candidate.violations.iter().enumerate()
                        .position(|(idx, &viols)| viols > winner.violations[idx])
                } else {
                    None
                };

                // Candidate surface form (right-aligned to match expected output)
                output.push_str(&format!("{}{:<width$} ", marker, candidate.form, width = max_cand_width));

                // Violations with stratum separators
                for (c_idx, &viols) in candidate.violations.iter().enumerate() {
                    let c_stratum = self.constraint_strata[c_idx];
                    let constraint_abbrev = &tableau.constraints[c_idx].abbrev();
                    let col_width = constraint_abbrev.len();

                    // Add separator
                    if c_idx > 0 {
                        let prev_stratum = self.constraint_strata[c_idx - 1];
                        if c_stratum != prev_stratum {
                            output.push('|');
                        } else {
                            output.push(sep_char);
                        }
                    }

                    // Mark violation as fatal only if it's the FIRST fatal violation
                    let is_fatal = first_fatal_idx == Some(c_idx);

                    if viols == 0 {
                        // Empty cell with proper width
                        for _ in 0..col_width {
                            output.push(' ');
                        }
                    } else {
                        let viol_str = if is_fatal {
                            format!("{}!", viols)
                        } else {
                            format!("{}", viols)
                        };
                        // Place violation - left-aligned with trailing spaces, except narrow columns
                        if col_width <= 3 {
                            // Narrow columns: no leading space
                            output.push_str(&viol_str);
                            for _ in viol_str.len()..col_width {
                                output.push(' ');
                            }
                        } else {
                            // Wider columns: add leading space
                            output.push(' ');
                            output.push_str(&viol_str);
                            for _ in (1 + viol_str.len())..col_width {
                                output.push(' ');
                            }
                        }
                    }
                }
                output.push_str("\n");
            }
            output.push_str("\n");
        }

        // Section 3: Status of Proposed Constraints
        if !self.constraint_necessity.is_empty() {
            output.push_str("3. Status of Proposed Constraints:  Necessary or Unnecessary\n\n");

            for (c_idx, necessity) in self.constraint_necessity.iter().enumerate() {
                let constraint = &tableau.constraints[c_idx];
                let status = match necessity {
                    ConstraintNecessity::Necessary => "Necessary",
                    ConstraintNecessity::UnnecessaryButShownForFaithfulness =>
                        "Not necessary (but included to show Faithfulness violations\n              of a winning candidate)",
                    ConstraintNecessity::CompletelyUnnecessary => "Not necessary",
                };
                output.push_str(&format!("   {}  {}\n", constraint.abbrev(), status));
            }

            // Check if mass deletion is possible
            let mass_deletion_possible = self.check_mass_deletion(tableau);
            if mass_deletion_possible {
                output.push_str("\nA check has determined that the grammar will still work even if the \n");
                output.push_str("constraints marked above as unnecessary are removed en masse.\n\n\n");
            } else {
                output.push_str("\n\n");
            }
        }

        // Section 4: Ranking Arguments (FRed)
        if let Some(ref fred) = self.fred_result {
            output.push_str(&fred.format_section4());
        }

        // Section 5: Mini-Tableaux
        if !self.mini_tableaux.is_empty() {
            output.push_str("5. Mini-Tableaux\n\n");
            output.push_str("The following small tableaux may be useful in presenting ranking arguments. \n");
            output.push_str("They include all winner-rival comparisons in which there is just one \n");
            output.push_str("winner-preferring constraint and at least one loser-preferring constraint.  \n");
            output.push_str("Constraints not violated by either candidate are omitted.\n\n");

            for mini in &self.mini_tableaux {
                self.format_mini_tableau(tableau, mini, &mut output);
            }
        }

        output
    }

    /// Check if mass deletion of unnecessary constraints still allows RCD to succeed
    fn check_mass_deletion(&self, tableau: &Tableau) -> bool {
        use crate::tableau::{InputForm, Candidate};

        // Create a modified tableau with all unnecessary constraints removed
        let forms = tableau.forms.iter().map(|form| {
            let candidates = form.candidates.iter().map(|cand| {
                let mut violations = cand.violations.clone();
                for (c_idx, necessity) in self.constraint_necessity.iter().enumerate() {
                    if *necessity != ConstraintNecessity::Necessary {
                        violations[c_idx] = 0;
                    }
                }
                Candidate {
                    form: cand.form.clone(),
                    frequency: cand.frequency,
                    violations,
                }
            }).collect();

            InputForm {
                input: form.input.clone(),
                candidates,
            }
        }).collect();

        let modified_tableau = Tableau {
            constraints: tableau.constraints.clone(),
            forms,
        };

        // Run RCD on modified tableau (without computing extra analyses to avoid recursion)
        let test_result = modified_tableau.run_rcd_internal(false, &[]);
        test_result.success
    }

    /// Format a mini-tableau
    fn format_mini_tableau(&self, tableau: &Tableau, mini: &MiniTableau, output: &mut String) {
        let form = &tableau.forms[mini.form_index];
        let winner = &form.candidates[mini.winner_index];
        let loser = &form.candidates[mini.loser_index];

        output.push_str(&format!("\n/{}/: \n", form.input));

        // Build header with only included constraints
        let sep_char = '\u{00A6}'; // Broken bar (¦)
        let max_cand_width = winner.form.len().max(loser.form.len()).max(2);

        let mut header = String::new();
        for _ in 0..max_cand_width + 2 {
            header.push(' ');
        }

        for (i, &c_idx) in mini.included_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];

            // Add separator before this constraint
            if i > 0 {
                let prev_c_idx = mini.included_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    header.push('|');
                } else {
                    header.push(sep_char);
                }
            }

            header.push_str(&constraint.abbrev());
        }
        output.push_str(&header);
        output.push_str("\n");

        // Output winner
        output.push_str(&format!(">{:<width$} ", winner.form, width = max_cand_width));
        for (i, &c_idx) in mini.included_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];
            let col_width = constraint.abbrev().len();

            // Add separator
            if i > 0 {
                let prev_c_idx = mini.included_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    output.push('|');
                } else {
                    output.push(sep_char);
                }
            }

            let viols = winner.violations[c_idx];
            if viols == 0 {
                for _ in 0..col_width {
                    output.push(' ');
                }
            } else {
                let viol_str = format!("{}", viols);
                if col_width <= 3 {
                    output.push_str(&viol_str);
                    for _ in viol_str.len()..col_width {
                        output.push(' ');
                    }
                } else {
                    output.push(' ');
                    output.push_str(&viol_str);
                    for _ in (1 + viol_str.len())..col_width {
                        output.push(' ');
                    }
                }
            }
        }
        output.push_str("\n");

        // Output loser (without fatal violation markers in mini-tableaux)
        output.push_str(&format!(" {:<width$} ", loser.form, width = max_cand_width));

        for (i, &c_idx) in mini.included_constraints.iter().enumerate() {
            let constraint = &tableau.constraints[c_idx];
            let c_stratum = self.constraint_strata[c_idx];
            let col_width = constraint.abbrev().len();

            // Add separator
            if i > 0 {
                let prev_c_idx = mini.included_constraints[i - 1];
                let prev_stratum = self.constraint_strata[prev_c_idx];
                if c_stratum != prev_stratum {
                    output.push('|');
                } else {
                    output.push(sep_char);
                }
            }

            let viols = loser.violations[c_idx];

            if viols == 0 {
                for _ in 0..col_width {
                    output.push(' ');
                }
            } else {
                let viol_str = format!("{}", viols);
                if col_width <= 3 {
                    output.push_str(&viol_str);
                    for _ in viol_str.len()..col_width {
                        output.push(' ');
                    }
                } else {
                    output.push(' ');
                    output.push_str(&viol_str);
                    for _ in (1 + viol_str.len())..col_width {
                        output.push(' ');
                    }
                }
            }
        }
        output.push_str("\n\n");
    }
}

impl Tableau {
    /// Run Recursive Constraint Demotion to find a ranking
    pub fn run_rcd(&self) -> RCDResult {
        self.run_rcd_internal(true, &[])
    }

    /// Run RCD enforcing a priori constraint rankings.
    ///
    /// `apriori[i][j] = true` means constraint i must rank above constraint j.
    pub fn run_rcd_with_apriori(&self, apriori: &[Vec<bool>]) -> RCDResult {
        self.run_rcd_internal(true, apriori)
    }

    /// Internal RCD implementation
    fn run_rcd_internal(&self, compute_extra_analyses: bool, apriori: &[Vec<bool>]) -> RCDResult {
        let num_constraints = self.constraints.len();
        let mut constraint_strata = vec![0; num_constraints];
        let mut current_stratum = 0;

        // Track which winner-loser pairs are still informative
        let mut informative_pairs: Vec<(usize, usize, usize)> = Vec::new();

        // Build list of all winner-loser pairs
        // (form_index, winner_index, loser_index)
        for (form_idx, form) in self.forms.iter().enumerate() {
            // Winner is the candidate with non-zero frequency
            if let Some(winner_idx) = form.candidates.iter().position(|c| c.frequency > 0) {
                for (loser_idx, candidate) in form.candidates.iter().enumerate() {
                    if loser_idx != winner_idx && candidate.frequency == 0 {
                        informative_pairs.push((form_idx, winner_idx, loser_idx));
                    }
                }
            }
        }

        // Debug: Starting RCD
        #[cfg(target_arch = "wasm32")]
        {
            use wasm_bindgen::prelude::*;
            #[wasm_bindgen]
            extern "C" {
                #[wasm_bindgen(js_namespace = console)]
                fn log(s: &str);
            }
            log(&format!("Starting RCD with {} pairs", informative_pairs.len()));
        }

        // RCD main loop
        loop {
            current_stratum += 1;

            // Find constraints that are "demotable" (prefer a loser in any informative pair)
            let mut demotable = vec![false; num_constraints];

            for &(form_idx, winner_idx, loser_idx) in &informative_pairs {
                let winner = &self.forms[form_idx].candidates[winner_idx];
                let loser = &self.forms[form_idx].candidates[loser_idx];

                for c_idx in 0..num_constraints {
                    // Skip constraints already ranked
                    if constraint_strata[c_idx] != 0 {
                        continue;
                    }

                    let winner_viols = winner.violations[c_idx];
                    let loser_viols = loser.violations[c_idx];

                    // Constraint prefers loser if loser has fewer violations
                    if loser_viols < winner_viols {
                        demotable[c_idx] = true;
                    }
                }
            }

            // ENFORCE A PRIORI RANKINGS
            // Any constraint that is a priori dominated by an unranked constraint
            // cannot join the current stratum.
            if !apriori.is_empty() {
                for outer in 0..num_constraints {
                    if constraint_strata[outer] == 0 {
                        for inner in 0..num_constraints {
                            if apriori[outer][inner] {
                                demotable[inner] = true;
                            }
                        }
                    }
                }
            }

            // All non-demotable, unranked constraints go into current stratum
            let mut added_any = false;
            for c_idx in 0..num_constraints {
                if constraint_strata[c_idx] == 0 && !demotable[c_idx] {
                    constraint_strata[c_idx] = current_stratum;
                    added_any = true;
                }
            }

            // Debug logging (after removing pairs)
            #[cfg(target_arch = "wasm32")]
            {
                use wasm_bindgen::prelude::*;
                #[wasm_bindgen]
                extern "C" {
                    #[wasm_bindgen(js_namespace = console)]
                    fn log(s: &str);
                }
                log(&format!("After stratum {}: {} pairs remaining",
                    current_stratum, informative_pairs.len()));
            }

            // Check if all constraints are ranked
            let all_ranked = constraint_strata.iter().all(|&s| s != 0);

            // If all constraints ranked, we're done (success even if pairs remain - those are ties)
            if all_ranked {
                #[cfg(target_arch = "wasm32")]
                {
                    use wasm_bindgen::prelude::*;
                    #[wasm_bindgen]
                    extern "C" {
                        #[wasm_bindgen(js_namespace = console)]
                        fn log(s: &str);
                    }
                    log(&format!("RCD SUCCEEDED: all constraints ranked in {} strata ({} pairs unresolved - ties)",
                        current_stratum, informative_pairs.len()));
                }

                // Create initial result
                let mut result = RCDResult {
                    constraint_strata,
                    num_strata: current_stratum,
                    success: true,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };

                // Compute additional analyses only if requested
                if compute_extra_analyses {
                    result.compute_extra_analyses(self);
                }

                return result;
            }

            // If no constraints added but some still unranked, algorithm failed
            if !added_any {
                #[cfg(target_arch = "wasm32")]
                {
                    use wasm_bindgen::prelude::*;
                    #[wasm_bindgen]
                    extern "C" {
                        #[wasm_bindgen(js_namespace = console)]
                        fn log(s: &str);
                    }
                    log(&format!("RCD FAILED: no constraints added to stratum {}", current_stratum));
                }
                return RCDResult {
                    constraint_strata,
                    num_strata: current_stratum - 1,
                    success: false,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };
            }

            // Remove pairs that are now decided by current stratum
            informative_pairs.retain(|&(form_idx, winner_idx, loser_idx)| {
                let winner = &self.forms[form_idx].candidates[winner_idx];
                let loser = &self.forms[form_idx].candidates[loser_idx];

                // Check if any constraint in current stratum decides this pair
                for c_idx in 0..num_constraints {
                    if constraint_strata[c_idx] == current_stratum {
                        let winner_viols = winner.violations[c_idx];
                        let loser_viols = loser.violations[c_idx];

                        // If this constraint prefers the winner, pair is decided
                        if winner_viols < loser_viols {
                            return false; // Remove this pair
                        }
                    }
                }
                true // Keep this pair
            });

            // If all pairs decided, we're done
            if informative_pairs.is_empty() {
                // Check if all constraints are ranked
                let all_ranked = constraint_strata.iter().all(|&s| s != 0);

                // Unranked constraints go in final stratum
                if !all_ranked {
                    for c_idx in 0..num_constraints {
                        if constraint_strata[c_idx] == 0 {
                            constraint_strata[c_idx] = current_stratum + 1;
                        }
                    }
                    current_stratum += 1;
                }

                #[cfg(target_arch = "wasm32")]
                {
                    use wasm_bindgen::prelude::*;
                    #[wasm_bindgen]
                    extern "C" {
                        #[wasm_bindgen(js_namespace = console)]
                        fn log(s: &str);
                    }
                    log(&format!("RCD SUCCEEDED with {} strata", current_stratum));
                }

                // Create initial result
                let mut result = RCDResult {
                    constraint_strata,
                    num_strata: current_stratum,
                    success: true,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };

                // Compute additional analyses only if requested
                if compute_extra_analyses {
                    result.compute_extra_analyses(self);
                }

                return result;
            }

            // Safety check: avoid infinite loop
            if current_stratum > num_constraints {
                return RCDResult {
                    constraint_strata,
                    num_strata: current_stratum,
                    success: false,
                    constraint_necessity: Vec::new(),
                    fred_result: None,
                    mini_tableaux: Vec::new(),
                    tie_warning: false,
                };
            }
        }
    }

    /// Compute constraint necessity for each constraint
    pub(crate) fn compute_constraint_necessity(&self, rcd_result: &RCDResult) -> Vec<ConstraintNecessity> {
        let mut necessity = vec![ConstraintNecessity::Necessary; self.constraints.len()];

        // Only analyze if RCD succeeded
        if !rcd_result.success {
            return necessity;
        }

        for c_idx in 0..self.constraints.len() {
            // Test if constraint is necessary
            if !self.is_constraint_necessary(c_idx) {
                // Constraint is unnecessary - check if violated by any winner
                if self.is_violated_by_winner(c_idx) {
                    necessity[c_idx] = ConstraintNecessity::UnnecessaryButShownForFaithfulness;
                } else {
                    necessity[c_idx] = ConstraintNecessity::CompletelyUnnecessary;
                }
            }
        }

        necessity
    }

    /// Test if a constraint is necessary by running RCD without it
    fn is_constraint_necessary(&self, constraint_idx: usize) -> bool {
        // Create modified tableau with constraint violations zeroed
        let modified_tableau = self.clone_with_constraint_removed(constraint_idx);

        // Run RCD on modified tableau (without computing extra analyses to avoid recursion)
        let test_result = modified_tableau.run_rcd_internal(false, &[]);

        // Constraint is necessary if RCD fails without it
        !test_result.success
    }

    /// Check if any winner violates this constraint
    fn is_violated_by_winner(&self, constraint_idx: usize) -> bool {
        for form in &self.forms {
            // Find the winner (candidate with non-zero frequency)
            if let Some(winner) = form.candidates.iter().find(|c| c.frequency > 0) {
                if winner.violations[constraint_idx] > 0 {
                    return true;
                }
            }
        }
        false
    }

    /// Clone the tableau with a constraint's violations set to zero
    fn clone_with_constraint_removed(&self, constraint_idx: usize) -> Tableau {
        use crate::tableau::{InputForm, Candidate};

        let forms = self.forms.iter().map(|form| {
            let candidates = form.candidates.iter().map(|cand| {
                let mut violations = cand.violations.clone();
                violations[constraint_idx] = 0;
                Candidate {
                    form: cand.form.clone(),
                    frequency: cand.frequency,
                    violations,
                }
            }).collect();

            InputForm {
                input: form.input.clone(),
                candidates,
            }
        }).collect();

        Tableau {
            constraints: self.constraints.clone(),
            forms,
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
    fn test_rcd_tiny_example() {
        let tiny_example = load_tiny_example();
        let tableau = Tableau::parse(&tiny_example).expect("Failed to parse tiny example");

        let result = tableau.run_rcd();

        // Should succeed
        assert!(result.success, "RCD should find a valid ranking");

        // Should have 2 strata according to expected output
        assert_eq!(result.num_strata, 2, "Should have 2 strata");

        // Check constraint rankings
        // Expected from full_output.txt:
        // Stratum #1: *NoOns, *Coda
        // Stratum #2: Max, Dep

        let constraint_0 = tableau.get_constraint(0).unwrap(); // *NoOns
        let constraint_1 = tableau.get_constraint(1).unwrap(); // *Coda
        let constraint_2 = tableau.get_constraint(2).unwrap(); // Max
        let constraint_3 = tableau.get_constraint(3).unwrap(); // Dep

        println!("Constraint 0 ({}): stratum {}", constraint_0.abbrev(), result.get_stratum(0).unwrap());
        println!("Constraint 1 ({}): stratum {}", constraint_1.abbrev(), result.get_stratum(1).unwrap());
        println!("Constraint 2 ({}): stratum {}", constraint_2.abbrev(), result.get_stratum(2).unwrap());
        println!("Constraint 3 ({}): stratum {}", constraint_3.abbrev(), result.get_stratum(3).unwrap());

        // *NoOns and *Coda should be in stratum 1
        assert_eq!(result.get_stratum(0).unwrap(), 1, "*NoOns should be in stratum 1");
        assert_eq!(result.get_stratum(1).unwrap(), 1, "*Coda should be in stratum 1");

        // Max and Dep should be in stratum 2
        assert_eq!(result.get_stratum(2).unwrap(), 2, "Max should be in stratum 2");
        assert_eq!(result.get_stratum(3).unwrap(), 2, "Dep should be in stratum 2");
    }

    fn extract_sections(text: &str, sections: &[usize]) -> String {
        let mut result = String::new();
        let lines: Vec<&str> = text.lines().collect();
        let mut current_section = 0;
        let mut in_section = false;

        for line in lines {
            // Check if this is a section header
            if let Some(section_num) = line.strip_prefix(char::is_numeric) {
                if let Some(num_str) = section_num.chars().next() {
                    if let Some(num) = num_str.to_digit(10) {
                        current_section = num as usize;
                        in_section = sections.contains(&current_section);
                    }
                }
            }

            // Check for next section starting
            if line.starts_with(char::is_numeric) && line.contains(". ") {
                if let Some(dot_pos) = line.find(". ") {
                    if let Ok(num) = line[..dot_pos].parse::<usize>() {
                        current_section = num;
                        in_section = sections.contains(&current_section);
                    }
                }
            }

            if in_section {
                result.push_str(line);
                result.push('\n');
            }

            // Stop after last requested section
            if !sections.is_empty() && current_section > *sections.iter().max().unwrap() {
                break;
            }
        }

        result
    }

    #[test]
    fn test_rcd_output_format() {
        let tiny_example = load_tiny_example();
        let tableau = Tableau::parse(&tiny_example).expect("Failed to parse tiny example");
        let result = tableau.run_rcd();

        // Generate formatted output
        let generated = result.format_output(&tableau, "TinyIllustrativeFile.txt");

        // Load expected output (file is ISO-8859 encoded)
        let expected_bytes = std::fs::read("../examples/tiny/full_output.txt")
            .expect("Failed to load examples/tiny/full_output.txt");
        let expected = String::from_utf8_lossy(&expected_bytes).to_string();

        // Extract header (first line)
        let gen_lines: Vec<&str> = generated.lines().collect();
        let exp_lines: Vec<&str> = expected.lines().collect();

        assert_eq!(gen_lines[0], exp_lines[0], "Header should match");

        // Extract and compare Section 1 (Result)
        let gen_section1 = extract_sections(&generated, &[1]);
        let exp_section1 = extract_sections(&expected, &[1]);

        assert_eq!(
            gen_section1.trim(),
            exp_section1.trim(),
            "Section 1 (Result) should match"
        );

        // For Section 2, verify it contains the key elements rather than exact formatting
        let gen_section2 = extract_sections(&generated, &[2]);

        // Verify section 2 header
        assert!(gen_section2.contains("2. Tableaux"), "Should have Section 2 header");

        // Verify all input forms are present
        assert!(gen_section2.contains("/a/:"), "Should contain /a/ form");
        assert!(gen_section2.contains("/tat/:"), "Should contain /tat/ form");
        assert!(gen_section2.contains("/at/:"), "Should contain /at/ form");

        // Verify constraint names appear
        assert!(gen_section2.contains("*NoOns"), "Should contain *NoOns constraint");
        assert!(gen_section2.contains("*Coda"), "Should contain *Coda constraint");
        assert!(gen_section2.contains("Max"), "Should contain Max constraint");
        assert!(gen_section2.contains("Dep"), "Should contain Dep constraint");

        // Verify winners are marked
        assert!(gen_section2.contains(">?a"), "Should mark >?a as winner");
        assert!(gen_section2.contains(">ta"), "Should mark >ta as winner");

        // Verify fatal violations are marked
        assert!(gen_section2.contains("1!"), "Should contain fatal violation markers");
    }

    #[test]
    fn test_rcd_sections_3_4_5() {
        let tiny_example = load_tiny_example();
        let tableau = Tableau::parse(&tiny_example).expect("Failed to parse tiny example");
        let result = tableau.run_rcd();

        // Generate formatted output
        let generated = result.format_output(&tableau, "TinyIllustrativeFile.txt");

        // Verify Section 3 is present
        assert!(generated.contains("3. Status of Proposed Constraints"), "Should have Section 3 header");
        assert!(generated.contains("Necessary") || generated.contains("Not necessary"),
            "Should classify constraints as necessary or not");

        // Verify Section 4 is present with correct FRed-based ranking statements.
        assert!(generated.contains("4. Ranking Arguments"), "Should have Section 4 header");
        assert!(generated.contains("Skeletal Basis"), "Should mention Skeletal Basis");
        assert!(generated.contains("*Coda >> Max"), "Should contain '*Coda >> Max' from FRed");
        assert!(generated.contains("*NoOns >> Dep"), "Should contain '*NoOns >> Dep' from FRed");

        // Verify Section 5 is present
        assert!(generated.contains("5. Mini-Tableaux"), "Should have Section 5 header");
    }
}
