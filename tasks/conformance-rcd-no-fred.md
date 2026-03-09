# `rcd_no_fred` — Extra trailing blank line

**Symptom:** 53 actual lines vs 52 expected — one extra trailing blank line.

**VB6 source:** `Main.frm:6349-6351` uses `PrintPara` for the mass deletion message, which ends with two `Print` statements (`Module1.bas:243-244`) producing exactly `\n\n` after the text. Our code (`rcd.rs:467`) ends with `\n\n\n` — one extra newline that's invisible when more sections follow, but creates an extra blank line when mass deletion is the last section.

**Fix:** In `rcd.rs`, change mass deletion success trailing from `\n\n\n` to `\n\n`, and add a leading `\n` to section 4 (FRed) and section 5 (mini-tableaux) headers to preserve inter-section spacing. Alternatively, normalize trailing newlines at the end of `format_output_with_algorithm`.
