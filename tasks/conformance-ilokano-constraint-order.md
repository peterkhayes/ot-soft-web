---
status: open
type: bug
priority: high
depends_on: []
---

# Conformance: Ilokano constraint ordering within strata differs from VB6

## Affected cases
- `ilokano_rcd_defaults`
- `ilokano_rcd_no_fred`
- `ilokano_rcd_mib`
- `ilokano_bcd_defaults`
- `ilokano_bcd_specific`
- `ilokano_lfcd_defaults`
- `ilokano_lfcd_mib`

## Description

All RCD/BCD/LFCD conformance cases for the Ilokano Hiatus Resolution dataset fail because the ordering of constraints within strata differs between VB6 and Rust.

Example diff (line 30):
```
expected: " Dep(h)¦Dep(+cons)¦Max¦Max-stem¦*VV¦Id(lo)¦*Nonhigh glide|Dep|Id(syl)¦Id(hi)"
actual:   " Dep|Dep(h)¦Dep(+cons)¦Max¦Max-stem¦*VV|Id(syl)¦Id(hi)|Id(lo)¦*Nonhigh glide"
```

This affects all algorithms (RCD, BCD, LFCD) and both Skeletal and MIB basis options, so the root cause is likely in how constraints are sorted or iterated when building strata output, not in the ranking algorithm itself.

The TinyIllustrativeFile dataset (4 constraints) passes, suggesting the issue only manifests with larger, more complex inputs (Ilokano has 10 constraints).

## Investigation

1. Compare the constraint iteration order in `format_rcd_output` vs VB6's `Main.frm`.
2. Check whether VB6 preserves original input file constraint order within strata, while Rust sorts or reorders them.
3. The separator characters (`¦` vs `|`) indicate stratum boundaries — verify the stratum assignments are also correct, not just the ordering within a stratum.

## Acceptance Criteria
- [ ] All 7 Ilokano conformance cases pass without a skip
- [ ] `make conformance-test` reports 0 failures for Ilokano cases
