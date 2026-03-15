---
status: open
type: bug
priority: medium
depends_on: []
---

# NHG download: add raw count columns

## Description

BPH's VB6 NHG output includes "Input #" and "Gen. #" columns showing raw counts alongside the frequency/percentage columns. Our download file only shows "Input%" and "Gen%".

The web UI may already show counts, but the download output file needs to include them to match VB6.

**Note:** This may have already been fixed since the comparison was made (2026-03-07). Verify current behavior before implementing.

## Reference

- BPH comparison file: `ComparisonsForPeter/4. NoisyHarmonicGrammar/BPH_AnttilaFinnishGenitivePluralsDraftOutput.txt`
- VB6 columns: "Input Fr.", "Gen Fr.", "Input #", "Gen. #"

## Acceptance Criteria

- [ ] Download output includes raw count columns matching VB6 format
