# `bcd_specific` — Wrong title

**Symptom:** Our title says "Biased Constraint Demotion (Specific)", VB6 says "Biased Constraint Demotion".

**VB6 source:** `Main.frm:5655-5656` always sets `gAlgorithmName = "Biased Constraint Demotion"` regardless of specificity mode. The specificity note is a separate paragraph after the header (`Main.frm:5678-5681`): "Version of BCD: specific Faithfulness constraints get priority."

**Fix:** In `lib.rs`, `format_bcd_output` and `format_bcd_html_output`: always use `"Biased Constraint Demotion"` as the algorithm name. Optionally add the separate specificity paragraph to match VB6.
