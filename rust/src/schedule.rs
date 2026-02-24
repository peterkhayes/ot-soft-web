//! Multi-stage learning schedule for GLA and NHG.
//!
//! All online algorithms (GLA, NHG) use a multi-stage plasticity schedule.
//! By default, 4 stages with geometrically interpolated plasticity from
//! initial (high) to final (low). A custom schedule can be parsed from
//! tab/space-delimited text with columns:
//!
//!   `Trials  PlastMark  PlastFaith  NoiseMark  NoiseFaith`
//!
//! Reproduces VB6 `DetermineLearningSchedule` (boersma.frm / NoisyHarmonicGrammar.frm).

/// Parameters for one stage of the learning schedule.
#[derive(Debug, Clone)]
pub struct LearningStage {
    /// Number of training trials in this stage.
    pub trials: usize,
    /// Plasticity applied to Markedness constraints.
    pub plast_mark: f64,
    /// Plasticity applied to Faithfulness constraints.
    pub plast_faith: f64,
    /// Noise sigma for Markedness constraints (GLA StochasticOT evaluation).
    pub noise_mark: f64,
    /// Noise sigma for Faithfulness constraints (GLA StochasticOT evaluation).
    pub noise_faith: f64,
}

/// Multi-stage plasticity schedule.
#[derive(Debug, Clone)]
pub struct LearningSchedule {
    pub stages: Vec<LearningStage>,
}

impl LearningSchedule {
    /// Build the default 4-stage schedule.
    ///
    /// Plasticity is geometrically interpolated between `initial_plasticity`
    /// and `final_plasticity`. Noise defaults to σ=2 for all stages.
    /// Trials are split evenly (integer division by 4).
    ///
    /// Reproduces VB6 `DetermineLearningSchedule` (non-custom path).
    pub fn default_4stage(cycles: usize, initial_plasticity: f64, final_plasticity: f64) -> Self {
        let p1 = initial_plasticity;
        let p4 = final_plasticity;
        let p2 = (p1 * p1 * p4).powf(1.0 / 3.0);
        let p3 = (p1 * p4 * p4).powf(1.0 / 3.0);
        let trials_per_stage = cycles / 4;

        let stages = [p1, p2, p3, p4]
            .iter()
            .map(|&p| LearningStage {
                trials: trials_per_stage,
                plast_mark: p,
                plast_faith: p,
                noise_mark: 2.0,
                noise_faith: 2.0,
            })
            .collect();

        LearningSchedule { stages }
    }

    /// Parse a custom learning schedule from whitespace-delimited text.
    ///
    /// The first line is a header and is ignored. Subsequent lines each define
    /// one stage with five columns: `Trials PlastMark PlastFaith NoiseMark NoiseFaith`.
    ///
    /// Reproduces VB6 `DetermineLearningSchedule` (custom-file path).
    pub fn parse(text: &str) -> Result<Self, String> {
        let mut lines = text.lines();
        // Skip the header row
        lines.next();

        let mut stages = Vec::new();
        for (i, line) in lines.enumerate() {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }
            let cols: Vec<&str> = line.split_whitespace().collect();
            if cols.len() < 5 {
                return Err(format!(
                    "Learning schedule line {}: expected 5 columns (Trials PlastMark PlastFaith NoiseMark NoiseFaith), got {}",
                    i + 2,
                    cols.len()
                ));
            }
            let row = i + 2; // 1-based, accounting for header
            let parse_usize = |s: &str, name: &str| {
                s.parse::<usize>().map_err(|_| {
                    format!("Learning schedule line {row}: invalid {name} '{s}'")
                })
            };
            let parse_f64 = |s: &str, name: &str| {
                s.parse::<f64>().map_err(|_| {
                    format!("Learning schedule line {row}: invalid {name} '{s}'")
                })
            };
            stages.push(LearningStage {
                trials: parse_usize(cols[0], "Trials")?,
                plast_mark: parse_f64(cols[1], "PlastMark")?,
                plast_faith: parse_f64(cols[2], "PlastFaith")?,
                noise_mark: parse_f64(cols[3], "NoiseMark")?,
                noise_faith: parse_f64(cols[4], "NoiseFaith")?,
            });
        }

        if stages.is_empty() {
            return Err("Custom learning schedule: no stages found (file must have a header row followed by data rows)".to_string());
        }

        Ok(LearningSchedule { stages })
    }

    /// Total number of training trials across all stages.
    pub fn total_cycles(&self) -> usize {
        self.stages.iter().map(|s| s.trials).sum()
    }

    /// True if this schedule was built from a custom text (non-default).
    /// Heuristic: stages > 4 or any stage has plast_mark != plast_faith.
    pub fn is_custom(&self) -> bool {
        self.stages.len() != 4
            || self.stages.iter().any(|s| (s.plast_mark - s.plast_faith).abs() > 1e-12)
    }

    /// Format schedule as a human-readable description for output files.
    pub fn format_description(&self) -> String {
        let mut out = String::new();
        if !self.is_custom() {
            // Compact form for the default 4-stage schedule
            let p1 = self.stages.first().map(|s| s.plast_mark).unwrap_or(0.0);
            let p4 = self.stages.last().map(|s| s.plast_mark).unwrap_or(0.0);
            let total = self.total_cycles();
            out.push_str(&format!(
                "   Cycles: {total}\n   Initial plasticity: {p1:.3}\n   Final plasticity: {p4:.3}\n"
            ));
        } else {
            // Full stage-by-stage listing for custom schedules
            out.push_str(&format!(
                "   Custom learning schedule ({} stages):\n",
                self.stages.len()
            ));
            out.push_str("   Stage  Trials    PlastMark  PlastFaith  NoiseMark  NoiseFaith\n");
            for (i, s) in self.stages.iter().enumerate() {
                out.push_str(&format!(
                    "   {:>5}  {:>7}  {:>9.3}  {:>9.3}  {:>9.3}  {:>9.3}\n",
                    i + 1,
                    s.trials,
                    s.plast_mark,
                    s.plast_faith,
                    s.noise_mark,
                    s.noise_faith,
                ));
            }
        }
        out
    }
}

/// Default template text for a 4-stage custom schedule (used in the web UI).
pub const CUSTOM_SCHEDULE_TEMPLATE: &str =
    "Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n\
     15000\t2\t2\t2\t2\n\
     15000\t0.2\t0.2\t2\t2\n\
     15000\t0.02\t0.02\t2\t2\n\
     15000\t0.002\t0.002\t2\t2";

// ─── Tests ───────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_4stage_plasticity() {
        let s = LearningSchedule::default_4stage(1_000_000, 2.0, 0.001);
        assert_eq!(s.stages.len(), 4);
        assert_eq!(s.stages[0].trials, 250_000);
        assert_eq!(s.stages[3].trials, 250_000);
        assert_eq!(s.total_cycles(), 1_000_000);

        // First and last stages should equal initial and final plasticity
        assert!((s.stages[0].plast_mark - 2.0).abs() < 1e-9);
        assert!((s.stages[3].plast_mark - 0.001).abs() < 1e-9);

        // Geometric interpolation: ratio p2/p1 == p3/p2 == p4/p3
        let r1 = s.stages[1].plast_mark / s.stages[0].plast_mark;
        let r2 = s.stages[2].plast_mark / s.stages[1].plast_mark;
        let r3 = s.stages[3].plast_mark / s.stages[2].plast_mark;
        assert!((r1 - r2).abs() < 1e-9, "ratios should be equal: {} vs {}", r1, r2);
        assert!((r2 - r3).abs() < 1e-9, "ratios should be equal: {} vs {}", r2, r3);

        // All noise values should be 2.0 by default
        for stage in &s.stages {
            assert_eq!(stage.noise_mark, 2.0);
            assert_eq!(stage.noise_faith, 2.0);
            // M/F plasticity equal by default
            assert_eq!(stage.plast_mark, stage.plast_faith);
        }

        assert!(!s.is_custom());
    }

    #[test]
    fn test_parse_custom_schedule() {
        let text = "Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n\
                    15000\t2\t1\t2\t2\n\
                    15000\t0.2\t0.1\t2\t2\n\
                    15000\t0.02\t0.01\t2\t2\n";
        let s = LearningSchedule::parse(text).unwrap();
        assert_eq!(s.stages.len(), 3);
        assert_eq!(s.stages[0].trials, 15000);
        assert!((s.stages[0].plast_mark - 2.0).abs() < 1e-9);
        assert!((s.stages[0].plast_faith - 1.0).abs() < 1e-9);
        assert!((s.stages[1].plast_mark - 0.2).abs() < 1e-9);
        assert!((s.stages[1].plast_faith - 0.1).abs() < 1e-9);
        assert_eq!(s.total_cycles(), 45000);
        assert!(s.is_custom());
    }

    #[test]
    fn test_parse_empty_fails() {
        let result = LearningSchedule::parse("Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n");
        assert!(result.is_err());
    }

    #[test]
    fn test_parse_bad_column_count() {
        let result = LearningSchedule::parse("header\n15000\t2\t1\n");
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("5 columns"));
    }

    #[test]
    fn test_parse_template() {
        // The built-in template should parse successfully
        let s = LearningSchedule::parse(CUSTOM_SCHEDULE_TEMPLATE).unwrap();
        assert_eq!(s.stages.len(), 4);
        assert_eq!(s.total_cycles(), 60_000);
    }
}
