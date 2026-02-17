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

## Current Implementation

### Data Structures

- `Constraint` - Stores constraint name and abbreviation
- `Candidate` - Output form with violation profile
- `InputForm` - Groups candidates by underlying input
- `Tableau` - Complete OT tableau with constraints and forms

### Parsing

- `Tableau::parse()` - Parses tab-delimited OT tableau files
  - Handles variable whitespace
  - Filters empty columns
  - Validates structure

### Exports to JavaScript

- `parse_tableau(text)` - Returns a Tableau object for JavaScript to format and display

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
