# Conformance Tests

Conformance tests compare Rust algorithm output against golden files collected from VB6 OTSoft on Windows. Use `/conformance` to run, collect, add, or manage test cases.

## Key Files

| File | Purpose |
|------|---------|
| `manifest.json` | Test case definitions (source of truth) |
| `golden/` | Golden output files collected from VB6 OTSoft |
| `automation/` | Python scripts + HTTP server for automated VB6 collection (see `automation/README.md`) |
| `rust/tests/conformance.rs` | Rust test runner that compares against golden files |

## manifest.json Schema

Each test case has:

- **id**: Unique identifier (e.g. `TinyIllustrativeFile_rcd_defaults`)
- **description**: Human-readable description
- **skip** (optional): Reason string — skips the case with a logged message
- **ignore_sections** (optional): List of section header substrings to strip before comparison. Sections matching `N. <header containing substring>` through the next numbered section are removed from both expected and actual output. Use for known VB6 bugs where we intentionally diverge.
- **input_file**: Path to input file relative to repo root
- **apriori_file**: Path to a priori rankings file, or `null`
- **algorithm**: One of `rcd`, `bcd`, `lfcd`, `maxent`, `factorial_typology`
- **format**: `"text"` (default) or `"html"` — determines comparison strategy
- **params**: Algorithm-specific parameters (varies by algorithm)
- **golden_file**: Path to expected output, relative to repo root

### Algorithm params

- **rcd/bcd/lfcd**: `include_fred`, `use_mib`, `show_details`, `include_mini_tableaux` (booleans); bcd also has `specific`
- **maxent**: `iterations`, `weight_min`, `weight_max`, `use_prior`, `sigma_squared`
- **factorial_typology**: `include_full_listing`

## Comparison Strategies

- **Text**: Normalizes dates/versions, collapses whitespace, rounds floats to 4 decimal places, collapses 3+ blank lines to 2.
- **HTML**: Extracts semantic cell grids by splitting on `</table>` boundaries and `<td`/`<th` openings (handles VB6's malformed HTML). Filters to shaded tables only, normalizes CSS classes to border/no-border, and compares mini-tableaux in sorted order with canonicalized column positions.

Tests skip gracefully when golden files are missing.
