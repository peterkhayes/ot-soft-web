# Rust Library

This directory contains the core computational logic for OT-Soft, compiled to WebAssembly.

## Principles

### Data and Logic Only

The Rust code should contain:
- ✅ Data structures (structs, enums)
- ✅ Parsing logic
- ✅ Computational algorithms (OT ranking, typology, etc.)
- ✅ Data validation

The Rust code should NOT contain:
- ❌ HTML generation
- ❌ CSS styling
- ❌ DOM manipulation
- ❌ Any presentation logic

**All presentation logic belongs in the Web codebase.**

## Module Structure

### `src/lib.rs` (40 lines)
Main entry point and public API:
- `init()` - WASM initialization
- `parse_tableau(text)` - Parse tableau from text
- `run_rcd(text)` - Run RCD algorithm on text
- Re-exports public types from modules

### `src/tableau.rs` (350+ lines)
Data structures and parsing:
- `Constraint` - Constraint name and abbreviation
- `Candidate` - Output form with frequency and violation profile
- `InputForm` - Groups candidates by underlying input
- `Tableau` - Complete OT tableau
- `Tableau::parse()` - Tab-delimited file parser
  - Row 1: Constraint full names
  - Row 2: Constraint abbreviations
  - Row 3+: Input, Output, Frequency, Violations
- Tests for all parsing functionality (7 tests)

### `src/rcd.rs` (270+ lines)
Recursive Constraint Demotion algorithm:
- `RCDResult` - Constraint strata and success flag
- `Tableau::run_rcd()` - Main RCD algorithm
  - Identifies winner-loser pairs
  - Iteratively ranks constraints into strata
  - Non-demotable constraints rank high
  - Handles ties (unresolved pairs)
- Console logging for WASM debugging
- Tests for RCD algorithm (1 test)

## Design Principles

- **Separation of concerns**: Parsing, algorithm, and API are separate modules
- **Tests live with code**: Each module has its own test section
- **WASM-first**: Debug logging only for WASM builds (not test builds)
- **Public fields**: Internal fields are `pub(crate)` for module access

## Testing

Run tests with:
```bash
cargo test
```

### Test Coverage

Current tests for parsing:
- ✅ `test_parse_tiny_example` - Validates parsing of the tiny example file
- ✅ `test_parse_second_input_form` - Tests second input form parsing
- ✅ `test_parse_third_input_form` - Tests third input form with 4 candidates
- ✅ `test_parse_empty_input` - Error handling for empty input
- ✅ `test_parse_no_constraints` - Error handling for missing constraints
- ✅ `test_parse_mismatched_headers` - Error handling for header mismatch
- ✅ `test_parse_output_without_input` - Error handling for orphaned outputs

All tests verify:
- Constraint names and abbreviations
- Input form grouping
- Candidate forms and frequencies
- Violation counts (including empty = 0)
- Error messages for invalid input
