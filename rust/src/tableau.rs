//! OT Tableau data structures and parsing
//!
//! This module contains the core data structures for representing Optimality Theory
//! tableaux, including constraints, candidates, input forms, and the tableau itself.

use wasm_bindgen::prelude::*;

/// Represents a constraint in an OT tableau
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct Constraint {
    full_name: String,
    abbrev: String,
}

#[wasm_bindgen]
impl Constraint {
    #[wasm_bindgen(getter)]
    pub fn full_name(&self) -> String {
        self.full_name.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn abbrev(&self) -> String {
        self.abbrev.clone()
    }
}

/// Represents a candidate output form
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct Candidate {
    pub(crate) form: String,
    pub(crate) frequency: usize,
    pub(crate) violations: Vec<usize>,
}

#[wasm_bindgen]
impl Candidate {
    #[wasm_bindgen(getter)]
    pub fn form(&self) -> String {
        self.form.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn frequency(&self) -> usize {
        self.frequency
    }

    pub fn get_violation(&self, constraint_index: usize) -> Option<usize> {
        self.violations.get(constraint_index).copied()
    }
}

/// Represents an input form with its candidates
#[wasm_bindgen]
#[derive(Debug, Clone)]
pub struct InputForm {
    pub(crate) input: String,
    pub(crate) candidates: Vec<Candidate>,
}

#[wasm_bindgen]
impl InputForm {
    #[wasm_bindgen(getter)]
    pub fn input(&self) -> String {
        self.input.clone()
    }

    pub fn candidate_count(&self) -> usize {
        self.candidates.len()
    }

    pub fn get_candidate(&self, index: usize) -> Option<Candidate> {
        self.candidates.get(index).cloned()
    }
}

/// Represents an OT tableau
#[wasm_bindgen]
#[derive(Debug)]
pub struct Tableau {
    pub(crate) constraints: Vec<Constraint>,
    pub(crate) forms: Vec<InputForm>,
}

#[wasm_bindgen]
impl Tableau {
    /// Parse a tableau from tab-delimited text
    /// Expected format:
    /// - Row 1: Constraint full names (first 3 columns blank)
    /// - Row 2: Constraint abbreviations (first 3 columns blank)
    /// - Row 3+: Input, Output, Frequency, Violations...
    pub fn parse(text: &str) -> Result<Tableau, String> {
        let lines: Vec<&str> = text.lines().collect();

        if lines.len() < 3 {
            return Err("Tableau must have at least 3 lines (full names, abbrevs, and one data row)".to_string());
        }

        // Parse constraint names (line 1)
        // Skip first 3 columns (Input, Output, Frequency headers - which are blank)
        let full_names: Vec<&str> = lines[0]
            .split('\t')
            .skip(3)
            .filter(|s| !s.trim().is_empty())
            .collect();

        // Parse constraint abbreviations (line 2)
        let abbrevs: Vec<&str> = lines[1]
            .split('\t')
            .skip(3)
            .filter(|s| !s.trim().is_empty())
            .collect();

        if full_names.len() != abbrevs.len() {
            return Err(format!(
                "Mismatch between full names ({}) and abbreviations ({})",
                full_names.len(),
                abbrevs.len()
            ));
        }

        if full_names.is_empty() {
            return Err("No constraints found in tableau".to_string());
        }

        let constraints: Vec<Constraint> = full_names
            .iter()
            .zip(abbrevs.iter())
            .map(|(full, abbr)| Constraint {
                full_name: full.trim().to_string(),
                abbrev: abbr.trim().to_string(),
            })
            .collect();

        // Parse data rows (lines 3+)
        let mut forms: Vec<InputForm> = Vec::new();
        let mut current_input: Option<String> = None;
        let mut current_candidates: Vec<Candidate> = Vec::new();

        for line in lines.iter().skip(2) {
            // Skip completely empty lines
            if line.trim().is_empty() {
                continue;
            }

            let parts: Vec<&str> = line.split('\t').collect();

            // Need at least 3 parts (input, output, frequency)
            if parts.len() < 3 {
                continue;
            }

            let input_cell = parts[0].trim();
            let output_cell = parts[1].trim();
            let frequency_cell = parts[2].trim();

            // Skip rows where output is empty (invalid row)
            if output_cell.is_empty() {
                continue;
            }

            // Parse frequency
            let frequency = frequency_cell.parse::<usize>().unwrap_or(0);

            // If there's an input form, start a new input group
            if !input_cell.is_empty() {
                // Save previous input if exists
                if let Some(input) = current_input.take() {
                    if !current_candidates.is_empty() {
                        forms.push(InputForm {
                            input,
                            candidates: current_candidates.clone(),
                        });
                        current_candidates.clear();
                    }
                }
                current_input = Some(input_cell.to_string());
            }

            // Ensure we have a current input (could be from previous row)
            if current_input.is_none() {
                return Err(format!("Found output '{}' without an input form", output_cell));
            }

            // Parse violations (columns 4 onwards)
            let mut violations: Vec<usize> = Vec::new();
            for i in 0..constraints.len() {
                let col_index = 3 + i;
                if col_index < parts.len() {
                    let viol_str = parts[col_index].trim();
                    violations.push(viol_str.parse::<usize>().unwrap_or(0));
                } else {
                    violations.push(0);
                }
            }

            current_candidates.push(Candidate {
                form: output_cell.to_string(),
                frequency,
                violations,
            });
        }

        // Don't forget the last input form
        if let Some(input) = current_input {
            if !current_candidates.is_empty() {
                forms.push(InputForm {
                    input,
                    candidates: current_candidates,
                });
            }
        }

        if forms.is_empty() {
            return Err("No valid input forms found in tableau".to_string());
        }

        Ok(Tableau { constraints, forms })
    }

    pub fn constraint_count(&self) -> usize {
        self.constraints.len()
    }

    pub fn form_count(&self) -> usize {
        self.forms.len()
    }

    pub fn get_constraint(&self, index: usize) -> Option<Constraint> {
        self.constraints.get(index).cloned()
    }

    pub fn get_form(&self, index: usize) -> Option<InputForm> {
        self.forms.get(index).cloned()
    }

    /// Get a summary string of the tableau
    pub fn summary(&self) -> String {
        format!(
            "Tableau with {} constraints and {} input forms",
            self.constraints.len(),
            self.forms.len()
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn load_tiny_example() -> String {
        std::fs::read_to_string("../examples/tiny/input.txt")
            .expect("Failed to load examples/tiny/input.txt")
    }

    #[test]
    fn test_parse_tiny_example() {
        let tiny_example = load_tiny_example();
        let tableau = Tableau::parse(&tiny_example).expect("Failed to parse tiny example");

        // Test constraint count
        assert_eq!(tableau.constraint_count(), 4, "Should have 4 constraints");

        // Test constraint names
        let constraint_0 = tableau.get_constraint(0).unwrap();
        assert_eq!(constraint_0.full_name(), "*No Onset");
        assert_eq!(constraint_0.abbrev(), "*NoOns");

        let constraint_3 = tableau.get_constraint(3).unwrap();
        assert_eq!(constraint_3.full_name(), "Dep(?)");
        assert_eq!(constraint_3.abbrev(), "Dep");

        // Test form count
        assert_eq!(tableau.form_count(), 3, "Should have 3 input forms");

        // Test first input form "a"
        let form_0 = tableau.get_form(0).unwrap();
        assert_eq!(form_0.input(), "a");
        assert_eq!(form_0.candidate_count(), 2);

        // Test first candidate of "a" -> "?a" with frequency 1
        let cand_0 = form_0.get_candidate(0).unwrap();
        assert_eq!(cand_0.form(), "?a");
        assert_eq!(cand_0.frequency(), 1);
        assert_eq!(cand_0.get_violation(0), Some(0)); // *NoOns
        assert_eq!(cand_0.get_violation(1), Some(0)); // *Coda
        assert_eq!(cand_0.get_violation(2), Some(0)); // Max
        assert_eq!(cand_0.get_violation(3), Some(1)); // Dep

        // Test second candidate of "a" -> "a" with frequency 0
        let cand_1 = form_0.get_candidate(1).unwrap();
        assert_eq!(cand_1.form(), "a");
        assert_eq!(cand_1.frequency(), 0);
        assert_eq!(cand_1.get_violation(0), Some(1)); // *NoOns
        assert_eq!(cand_1.get_violation(1), Some(0)); // *Coda
        assert_eq!(cand_1.get_violation(2), Some(0)); // Max
        assert_eq!(cand_1.get_violation(3), Some(0)); // Dep
    }

    #[test]
    fn test_parse_second_input_form() {
        let tiny_example = load_tiny_example();
        let tableau = Tableau::parse(&tiny_example).expect("Failed to parse tiny example");

        // Test second input form "tat"
        let form_1 = tableau.get_form(1).unwrap();
        assert_eq!(form_1.input(), "tat");
        assert_eq!(form_1.candidate_count(), 2);

        // Test "tat" -> "ta"
        let cand_0 = form_1.get_candidate(0).unwrap();
        assert_eq!(cand_0.form(), "ta");
        assert_eq!(cand_0.frequency(), 1);
        assert_eq!(cand_0.get_violation(0), Some(0)); // *NoOns
        assert_eq!(cand_0.get_violation(1), Some(0)); // *Coda
        assert_eq!(cand_0.get_violation(2), Some(1)); // Max (deleted t)
        assert_eq!(cand_0.get_violation(3), Some(0)); // Dep

        // Test "tat" -> "tat"
        let cand_1 = form_1.get_candidate(1).unwrap();
        assert_eq!(cand_1.form(), "tat");
        assert_eq!(cand_1.frequency(), 0);
        assert_eq!(cand_1.get_violation(0), Some(0)); // *NoOns
        assert_eq!(cand_1.get_violation(1), Some(1)); // *Coda
        assert_eq!(cand_1.get_violation(2), Some(0)); // Max
        assert_eq!(cand_1.get_violation(3), Some(0)); // Dep
    }

    #[test]
    fn test_parse_third_input_form() {
        let tiny_example = load_tiny_example();
        let tableau = Tableau::parse(&tiny_example).expect("Failed to parse tiny example");

        // Test third input form "at"
        let form_2 = tableau.get_form(2).unwrap();
        assert_eq!(form_2.input(), "at");
        assert_eq!(form_2.candidate_count(), 4);

        // Test all four candidates
        let candidates = vec![
            ("?a", 1, vec![0, 0, 1, 1]),
            ("?at", 0, vec![0, 1, 0, 1]),
            ("a", 0, vec![1, 0, 1, 0]),
            ("at", 0, vec![1, 1, 0, 0]),
        ];

        for (i, (form, freq, viols)) in candidates.iter().enumerate() {
            let cand = form_2.get_candidate(i).unwrap();
            assert_eq!(cand.form(), *form);
            assert_eq!(cand.frequency(), *freq);
            for (j, expected_viol) in viols.iter().enumerate() {
                assert_eq!(
                    cand.get_violation(j),
                    Some(*expected_viol),
                    "Candidate {} violation {} mismatch",
                    i,
                    j
                );
            }
        }
    }

    #[test]
    fn test_parse_empty_input() {
        let result = Tableau::parse("");
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("at least 3 lines"));
    }

    #[test]
    fn test_parse_no_constraints() {
        let input = "\t\t\t\n\t\t\t\na\tb\t1";
        let result = Tableau::parse(input);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("No constraints found"));
    }

    #[test]
    fn test_parse_mismatched_headers() {
        let input = "\t\t\tCon1\tCon2\n\t\t\tC1\na\tb\t1\t0\t0";
        let result = Tableau::parse(input);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("Mismatch"));
    }

    #[test]
    fn test_parse_output_without_input() {
        let input = "\t\t\tCon1\n\t\t\tC1\n\toutput\t1\t0";
        let result = Tableau::parse(input);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("without an input"));
    }
}
