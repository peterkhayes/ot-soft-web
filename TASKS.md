# OTSoft: Open Tasks

## Bugs

- [Stochastic OT: investigate accuracy discrepancy with VB6](tasks/stochastic-ot-accuracy.md)
- ~~Fix RCD HTML tableaux download test~~

## UX

- ~~RCD: default "Show details" to off~~

## Features

- ~~Diagnostics if ranking fails~~
- [HTML shading customization](tasks/html-shading-options.md)
- [Excel file parsing (.xlsx)](tasks/excel-parsing.md)
- [Praat export](tasks/praat-export.md)
- [R export](tasks/r-export.md)
- [HowIRanked log file](tasks/how-i-ranked.md)

## Code Quality

### Rust

- ~~Extract version string into a constant~~
- ~~Pass FredOptions struct to apply_fred_options~~
- ~~Extract broken-bar character into a named constant~~
- ~~Remove unused _apriori parameter from apply_fred_options~~
- ~~Extract apriori parsing helper in lib.rs~~
- ~~Break up long formatting functions~~

### Web

- ~~Extract shared DownloadButton component~~
- ~~Use useLocalStorage hook consistently in App.tsx~~
- ~~Extract responsive breakpoint constants~~ (already done)
- ~~Break up GlaPanel.tsx~~
- ~~Standardize result state types with a generic~~
- ~~Add aria-labels to interactive SVG elements~~
- ~~Break up NhgPanel.tsx~~
- ~~Break up FactorialTypologyPanel.tsx~~
- ~~Break up RcdPanel.tsx~~
- ~~Break up MaxEntPanel.tsx~~

## Testing

- [Collect remaining VB6 golden files](tasks/golden-files.md)
- [Edge case test examples](tasks/edge-case-examples.md)
- [Conformance: Ilokano constraint ordering within strata](tasks/conformance-ilokano-constraint-order.md)
- [Collect golden files for BasicRCD_FT and Anttila examples](tasks/collect-basicrcd-anttila-golden.md)
