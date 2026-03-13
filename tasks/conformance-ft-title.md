---
status: done
type: bug
priority: medium
depends_on: []
---

# Conformance: Factorial Typology output header includes filename; VB6 does not

## Affected cases
- `TinyIllustrativeFile_ft_defaults`
- `TinyIllustrativeFile_ft_full_listing`
- `TinyIllustrativeFile_ft_apriori`
- `ilokano_ft_defaults`
- `ilokano_ft_full_listing`

## Description

Rust outputs:
```
Results of Factorial Typology Search for TinyIllustrativeFile.txt
```

VB6 outputs:
```
Results of Factorial Typology Search
```

The Rust implementation appends ` for {filename}` to the FT header, which VB6 does not. Additionally, VB6 and Rust differ in overall line counts (e.g., 73 expected vs 69 actual for `ft_defaults`), suggesting there may be additional structural differences beyond the title line.

## Fix

Investigate `factorial_typology.rs` / `format_factorial_typology_output` to find where the header is generated. Compare section-by-section against the golden file to identify all structural differences, then align Rust output with VB6.

## Acceptance Criteria
- [ ] All five FT conformance cases pass without a skip
- [ ] `make conformance-test` reports 0 failures for FT cases
