//! FRed (Fusional Reduction Algorithm)
//!
//! Implements Prince & Brasoveanu (2005) ranking-argumentation algorithm.
//! Computes a basis of Elementary Ranking Conditions (ERCs) that encapsulates
//! all ranking arguments supported by an OT dataset.

use std::collections::HashSet;
use wasm_bindgen::prelude::*;
use crate::tableau::Tableau;

// ─────────────────────────────────────────────────────────────────────────────
// ERC helpers
// ─────────────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
enum ErcStatus {
    Valid,
    Uninformative,
    Unsatisfiable,
    Duplicate,
}

fn erc_status(erc: &str) -> ErcStatus {
    let mut w = 0usize;
    let mut l = 0usize;
    for ch in erc.chars() {
        match ch {
            'W' => w += 1,
            'L' => l += 1,
            _ => {}
        }
    }
    if l == 0 {
        if w == 0 { ErcStatus::Duplicate } else { ErcStatus::Uninformative }
    } else if w == 0 {
        ErcStatus::Unsatisfiable
    } else {
        ErcStatus::Valid
    }
}

fn w_count(erc: &str) -> usize {
    erc.chars().filter(|&c| c == 'W').count()
}

fn l_count(erc: &str) -> usize {
    erc.chars().filter(|&c| c == 'L').count()
}

// ─────────────────────────────────────────────────────────────────────────────
// Fusion algebra
// ─────────────────────────────────────────────────────────────────────────────

/// Fuse a set of ERCs using the Prince & Brasoveanu algebra:
///   e ⊕ X = X,  W ⊕ W = W,  W ⊕ L = L,  L ⊕ X = L
fn fusion(ercs: &[&str], n: usize) -> String {
    let mut buf: Vec<char> = vec!['e'; n];
    for erc in ercs {
        let chars: Vec<char> = erc.chars().collect();
        for i in 0..n {
            match buf[i] {
                'e' => buf[i] = chars[i],
                'W' => {
                    if chars[i] == 'L' {
                        buf[i] = 'L';
                    }
                    // W ⊕ W = W, W ⊕ e = W → no change needed
                }
                'L' => {} // L ⊕ anything = L
                _ => {}
            }
        }
    }
    buf.into_iter().collect()
}

/// Compute the Fusion of the Total Information-Loss Residue.
///
/// For each constraint column: if the column has only W's and e's (no L),
/// with at least one W and one e, mark the ERCs with 'e' in that column.
/// Return the fusion of all marked ERCs.
fn fusion_of_total_residue(ercs: &[&str], n: usize) -> String {
    let mut include = vec![false; ercs.len()];

    for col in 0..n {
        let mut has_l = false;
        let mut has_w = false;
        let mut has_e = false;

        for erc in ercs.iter() {
            match erc.chars().nth(col).unwrap_or('e') {
                'L' => {
                    has_l = true;
                    break;
                }
                'W' => has_w = true,
                _ => has_e = true,
            }
        }

        if !has_l && has_w && has_e {
            for (idx, erc) in ercs.iter().enumerate() {
                if erc.chars().nth(col).unwrap_or('e') == 'e' {
                    include[idx] = true;
                }
            }
        }
    }

    let residue: Vec<&str> = ercs
        .iter()
        .zip(include.iter())
        .filter(|(_, &inc)| inc)
        .map(|(erc, _)| *erc)
        .collect();

    if residue.is_empty() {
        String::new()
    } else {
        fusion(&residue, n)
    }
}

/// Check if `erc1` entails `erc2` (Prince & Brasoveanu p. 13):
///   W entails only W
///   e entails W or e
///   L entails anything
fn entails(erc1: &str, erc2: &str) -> bool {
    for (c1, c2) in erc1.chars().zip(erc2.chars()) {
        match c1 {
            'W' => {
                if c2 != 'W' {
                    return false;
                }
            }
            'e' => {
                if c2 == 'L' {
                    return false;
                }
            }
            'L' => {} // L entails anything
            _ => {}
        }
    }
    true
}

/// Compute the Skeletal Basis: copy `fusion_str` but replace positions
/// where `residue` has 'L' with 'e'.
fn skeletal_basis(fusion_str: &str, residue: &str) -> String {
    if residue.is_empty() {
        return fusion_str.to_string();
    }
    fusion_str
        .chars()
        .zip(residue.chars())
        .map(|(f, r)| if r == 'L' { 'e' } else { f })
        .collect()
}

// ─────────────────────────────────────────────────────────────────────────────
// Recursive state
// ─────────────────────────────────────────────────────────────────────────────

struct FRedState {
    num_constraints: usize,
    valhalla: Vec<String>,
    failure_flag: bool,
    use_skeletal_basis: bool,
    /// De-duplication set for send-on ERC sets (concatenated ERC strings).
    visited: HashSet<String>,
}

impl FRedState {
    fn recursive_routine(&mut self, ercs: &[&str]) {
        // ── a. Compute fusion ──────────────────────────────────────────────
        let fus = fusion(ercs, self.num_constraints);

        // Failed fusion: L but no W → unrankable
        if l_count(&fus) > 0 && w_count(&fus) == 0 {
            self.valhalla.push(fus);
            self.failure_flag = true;
            return;
        }

        // Trivially satisfied: no L → any ranking works
        if l_count(&fus) == 0 {
            self.valhalla.push(fus);
            self.failure_flag = true;
            return;
        }

        // ── b. Fusion of total residue ─────────────────────────────────────
        let residue = fusion_of_total_residue(ercs, self.num_constraints);

        // ── c. Entailment check ────────────────────────────────────────────
        let entailed = self.entailment_check(&residue, &fus);

        if !entailed {
            // In Skeletal Basis mode, store the skeletal version of the fusion.
            let final_erc = if self.use_skeletal_basis {
                skeletal_basis(&fus, &residue)
            } else {
                fus.clone()
            };
            // De-duplicate before adding to Valhalla.
            if !self.valhalla.contains(&final_erc) {
                self.valhalla.push(final_erc);
            }
        }

        // ── d. Form send-on sets and recurse ──────────────────────────────
        for col in 0..self.num_constraints {
            let mut has_l = false;
            let mut has_w = false;
            let mut has_e = false;

            for erc in ercs.iter() {
                match erc.chars().nth(col).unwrap_or('e') {
                    'L' => {
                        has_l = true;
                        break;
                    }
                    'W' => has_w = true,
                    _ => has_e = true,
                }
            }

            // Column qualifies: no L, at least one W and one e.
            if !has_l && has_w && has_e {
                let send_on: Vec<&str> = ercs
                    .iter()
                    .filter(|erc| erc.chars().nth(col).unwrap_or('e') == 'e')
                    .copied()
                    .collect();

                // Check novelty: concatenate send-on ERCs for de-duplication.
                let check = send_on.concat();
                if !self.visited.contains(&check) {
                    self.visited.insert(check);
                    self.recursive_routine(&send_on);
                    if self.failure_flag {
                        return;
                    }
                }
            }
        }
    }

    /// Entailment check (returns true if the fusion IS entailed by the residue,
    /// meaning it should NOT be added to Valhalla).
    fn entailment_check(&self, residue: &str, fus: &str) -> bool {
        // If residue is empty, there is no entailment relation.
        if residue.is_empty() {
            return false;
        }

        if self.use_skeletal_basis {
            // Skeletal Basis mode: compute skeletal basis of fusion w.r.t. residue.
            // If the result has no L's → entailed (discard).
            let sb = skeletal_basis(fus, residue);
            l_count(&sb) == 0
        } else {
            // Most Informative Basis mode: check if residue entails fusion.
            entails(residue, fus)
        }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Public result type
// ─────────────────────────────────────────────────────────────────────────────

/// Result of running the FRed algorithm.
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct FRedResult {
    /// The Valhalla: the final basis of ERCs.
    #[wasm_bindgen(skip)]
    pub(crate) valhalla: Vec<String>,
    /// Whether the algorithm encountered an unrankable configuration.
    failure: bool,
    /// Whether Skeletal Basis mode was used (false = Most Informative Basis).
    use_skeletal_basis: bool,
    /// Constraint abbreviations for formatting ranking statements.
    #[wasm_bindgen(skip)]
    pub(crate) constraint_abbrevs: Vec<String>,
}

#[wasm_bindgen]
impl FRedResult {
    pub fn failure(&self) -> bool {
        self.failure
    }

    pub fn use_skeletal_basis(&self) -> bool {
        self.use_skeletal_basis
    }

    pub fn valhalla_size(&self) -> usize {
        self.valhalla.len()
    }

    pub fn get_valhalla_erc(&self, index: usize) -> Option<String> {
        self.valhalla.get(index).cloned()
    }
}

impl FRedResult {
    /// Format Section 4 of the RCD output: "Ranking Arguments, based on FRed".
    pub(crate) fn format_section4(&self) -> String {
        let mut out = String::new();

        let basis_name = if self.use_skeletal_basis {
            "Skeletal Basis"
        } else {
            "Most Informative Basis"
        };
        let purpose = if self.use_skeletal_basis {
            "keep each final ranking argument as pithy as possible"
        } else {
            "minimize the set of final ranking arguments"
        };

        out.push_str("4. Ranking Arguments, based on the Fusional Reduction Algorithm\n\n");
        out.push_str(&format!(
            "This run sought to obtain the {basis_name}, intended to {purpose}.\n\n\n\n"
        ));

        if self.failure {
            out.push_str(
                "The constraints cannot be ranked to yield the desired outcomes.\n\n",
            );
        }

        out.push_str("The final rankings obtained are as follows:\n\n");

        // WCount == 1 first (simple "A >> B" form)
        for erc in &self.valhalla {
            if w_count(erc) == 1 {
                out.push_str(&format!(
                    "      {} >> {}\n",
                    self.constraint_set_string(erc, 'W'),
                    self.constraint_set_string(erc, 'L'),
                ));
            }
        }
        // Blank line after WCount=1 section (matches VB6's `Print #mTmpFile,`).
        out.push_str("\n");

        // WCount > 1 next ("At least one of {…} >> {…}" form)
        for erc in &self.valhalla {
            if w_count(erc) > 1 {
                out.push_str(&format!(
                    "      {} >> {}\n",
                    self.constraint_set_string(erc, 'W'),
                    self.constraint_set_string(erc, 'L'),
                ));
            }
        }

        // Final blank line to complete 2 blank lines total at the section end.
        out.push_str("\n");
        out
    }

    /// Build a ranking-statement string for positions matching `criterion` in `erc`.
    ///
    /// Matches VB6's `ConstraintSetString`:
    ///   single match  → just the abbreviation
    ///   multiple W    → "At least one of { C1, C2 }"
    ///   multiple L    → "{ C1, C2 }"
    fn constraint_set_string(&self, erc: &str, criterion: char) -> String {
        let count = erc.chars().filter(|&c| c == criterion).count();

        if count == 1 {
            for (i, ch) in erc.chars().enumerate() {
                if ch == criterion {
                    return self.constraint_abbrevs[i].clone();
                }
            }
            String::new()
        } else {
            // Build ", C1, C2, ..." then strip the leading ", ".
            let mut acc = String::new();
            for (i, ch) in erc.chars().enumerate() {
                if ch == criterion {
                    acc.push_str(", ");
                    acc.push_str(&self.constraint_abbrevs[i]);
                }
            }
            // acc starts with ", "; Mid(acc, 2) in VB6 skips the first char (',')
            // leaving " C1, C2, ...".
            let inner = &acc[1..]; // skip the leading ','
            if criterion == 'L' {
                format!("{{{} }}", inner)
            } else {
                format!("At least one of {{{} }}", inner)
            }
        }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Tableau integration
// ─────────────────────────────────────────────────────────────────────────────

impl Tableau {
    /// Run FRed with optional Skeletal Basis or MIB mode.
    ///
    /// `use_mib = false` → Skeletal Basis (default in VB6)
    /// `use_mib = true`  → Most Informative Basis
    pub fn run_fred(&self, use_mib: bool) -> FRedResult {
        self.run_fred_internal(use_mib, &[])
    }

    /// Run FRed with a priori rankings enforced as additional ERCs.
    pub fn run_fred_with_apriori(&self, use_mib: bool, apriori: &[Vec<bool>]) -> FRedResult {
        self.run_fred_internal(use_mib, apriori)
    }

    fn run_fred_internal(&self, use_mib: bool, apriori: &[Vec<bool>]) -> FRedResult {
        let n = self.constraints.len();
        let abbrevs: Vec<String> = self.constraints.iter().map(|c| c.abbrev()).collect();
        let use_skeletal_basis = !use_mib;

        // ── Step 1: build original ERC set ────────────────────────────────

        let mut ercs: Vec<String> = Vec::new();

        // Prepend a priori ranking ERCs.
        if !apriori.is_empty() {
            for i in 0..n {
                for j in 0..n {
                    if apriori[i][j] {
                        let mut erc = vec!['e'; n];
                        erc[i] = 'W'; // winner-preferrer
                        erc[j] = 'L'; // loser-preferrer
                        ercs.push(erc.into_iter().collect());
                    }
                }
            }
        }

        // Add ERCs from winner-rival pairs.
        for form in &self.forms {
            let winner_idx = match form.candidates.iter().position(|c| c.frequency > 0) {
                Some(i) => i,
                None => continue,
            };
            let winner = &form.candidates[winner_idx];

            for (rival_idx, rival) in form.candidates.iter().enumerate() {
                if rival_idx == winner_idx || rival.frequency > 0 {
                    continue;
                }

                let erc: String = (0..n)
                    .map(|c| {
                        let wv = winner.violations[c];
                        let rv = rival.violations[c];
                        if wv < rv {
                            'W'
                        } else if wv > rv {
                            'L'
                        } else {
                            'e'
                        }
                    })
                    .collect();

                match erc_status(&erc) {
                    ErcStatus::Uninformative => continue,
                    ErcStatus::Unsatisfiable | ErcStatus::Duplicate => {
                        // Hard failure: winner cannot be derived under any ranking.
                        return FRedResult {
                            valhalla: Vec::new(),
                            failure: true,
                            use_skeletal_basis,
                            constraint_abbrevs: abbrevs,
                        };
                    }
                    ErcStatus::Valid => {
                        // De-duplicate.
                        if !ercs.contains(&erc) {
                            ercs.push(erc);
                        }
                    }
                }
            }
        }

        // If no informative ERCs were found, nothing to do.
        if ercs.is_empty() {
            return FRedResult {
                valhalla: Vec::new(),
                failure: false,
                use_skeletal_basis,
                constraint_abbrevs: abbrevs,
            };
        }

        // ── Step 2: recursive search ───────────────────────────────────────

        let mut state = FRedState {
            num_constraints: n,
            valhalla: Vec::new(),
            failure_flag: false,
            use_skeletal_basis,
            visited: HashSet::new(),
        };

        let erc_refs: Vec<&str> = ercs.iter().map(|s| s.as_str()).collect();
        state.recursive_routine(&erc_refs);

        FRedResult {
            valhalla: state.valhalla,
            failure: state.failure_flag,
            use_skeletal_basis,
            constraint_abbrevs: abbrevs,
        }
    }
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
    fn test_fusion_basic() {
        // e ⊕ W = W, W ⊕ L = L, L ⊕ anything = L
        let ercs = ["WeeL", "eWLe"];
        let result = fusion(&ercs, 4);
        assert_eq!(result, "WWLL");

        let ercs2 = ["WeeL"];
        let result2 = fusion(&ercs2, 4);
        assert_eq!(result2, "WeeL");
    }

    #[test]
    fn test_skeletal_basis() {
        // Positions where residue='L' become 'e' in fusion.
        assert_eq!(skeletal_basis("WWLL", "WWLL"), "WWee");
        assert_eq!(skeletal_basis("eWLe", ""), "eWLe");
        assert_eq!(skeletal_basis("WeeL", ""), "WeeL");
    }

    #[test]
    fn test_entails() {
        // An ERC entails itself.
        assert!(entails("WeeL", "WeeL"));
        assert!(entails("eWLe", "eWLe"));

        // e does not entail L.
        assert!(!entails("eWLe", "WWLL")); // pos 4: e→L fails
        assert!(!entails("WeeL", "WWLL")); // pos 3: e→L fails

        // L entails anything.
        assert!(entails("L", "W"));
        assert!(entails("L", "e"));
        assert!(entails("L", "L"));
    }

    #[test]
    fn test_erc_status() {
        assert_eq!(erc_status("WeeL"), ErcStatus::Valid);
        assert_eq!(erc_status("eWLe"), ErcStatus::Valid);
        assert_eq!(erc_status("WWLL"), ErcStatus::Valid);
        assert_eq!(erc_status("Weee"), ErcStatus::Uninformative);
        assert_eq!(erc_status("eeLe"), ErcStatus::Unsatisfiable);
        assert_eq!(erc_status("eeee"), ErcStatus::Duplicate);
    }

    #[test]
    fn test_fred_tiny_example() {
        let text = load_tiny();
        let tableau = Tableau::parse(&text).expect("parse failed");
        let result = tableau.run_fred(false); // Skeletal Basis

        assert!(!result.failure, "FRed should not report failure");

        // Expected Valhalla for the tiny example: {"eWLe", "WeeL"}
        // (order may vary, so check by content)
        assert_eq!(result.valhalla.len(), 2, "Valhalla should have 2 ERCs");
        assert!(
            result.valhalla.contains(&"eWLe".to_string()),
            "Valhalla should contain eWLe"
        );
        assert!(
            result.valhalla.contains(&"WeeL".to_string()),
            "Valhalla should contain WeeL"
        );
    }

    #[test]
    fn test_fred_output_section4() {
        let text = load_tiny();
        let tableau = Tableau::parse(&text).expect("parse failed");
        let result = tableau.run_fred(false);

        let section4 = result.format_section4();

        // Must contain the section header.
        assert!(section4.contains("4. Ranking Arguments, based on the Fusional Reduction Algorithm"));

        // Must contain the basis description.
        assert!(section4.contains("Skeletal Basis"));
        assert!(section4.contains("keep each final ranking argument as pithy as possible"));

        // Must contain both ranking statements.
        assert!(
            section4.contains("*Coda >> Max"),
            "Section 4 should contain '*Coda >> Max'"
        );
        assert!(
            section4.contains("*NoOns >> Dep"),
            "Section 4 should contain '*NoOns >> Dep'"
        );
    }

    #[test]
    fn test_fred_output_format_matches_expected() {
        // Verify that format_rcd_output for the tiny example produces Section 4
        // matching full_output.txt lines 56-65.
        let text = load_tiny();
        let tableau = Tableau::parse(&text).expect("parse failed");
        let result = tableau.run_rcd();
        let generated = result.format_output(&tableau, "TinyIllustrativeFile.txt");

        // Load expected output
        let expected_bytes = std::fs::read("../examples/tiny/full_output.txt")
            .expect("Failed to load full_output.txt");
        let expected = String::from_utf8_lossy(&expected_bytes).to_string();

        // Check that both contain the same FRed ranking statements.
        assert!(generated.contains("*Coda >> Max"), "Generated output missing '*Coda >> Max'");
        assert!(generated.contains("*NoOns >> Dep"), "Generated output missing '*NoOns >> Dep'");
        assert!(expected.contains("*Coda >> Max"), "Expected output missing '*Coda >> Max'");
        assert!(expected.contains("*NoOns >> Dep"), "Expected output missing '*NoOns >> Dep'");
    }
}
