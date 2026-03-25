---
status: done
type: bug
priority: medium
depends_on: []
---

# NHG download: add raw count columns

## Resolution

Updated `NhgResult::format_output()` in `rust/src/nhg.rs` to match the VB6 format:

- Changed column headers from `Input%` / `Gen%` to `Input Fr.` / `Gen Fr.` / `Input #` / `Gen. #`
- Changed data from percentages (0-100) to fractions (0-1 with 3 decimal places)
- Added raw input count (`cand.frequency`) and generated count (`gen_prop * test_trials`)
- When a candidate has zero input frequency, `Input #` is blank (matching VB6 behavior)

The web UI display (NhgPanel.tsx) is unaffected — it renders its own HTML table independently.

## Reference

- BPH comparison file: `ComparisonsForPeter/4. NoisyHarmonicGrammar/BPH_AnttilaFinnishGenitivePluralsDraftOutput.txt`
- VB6 columns: "Input Fr.", "Gen Fr.", "Input #", "Gen. #"

## Acceptance Criteria

- [x] Download output includes raw count columns matching VB6 format
