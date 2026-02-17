# Rust Library

Core computational logic for OT-Soft, compiled to WebAssembly.

## Principles

**Separation of concerns**: Rust handles data and logic, JavaScript handles presentation.

The Rust code should contain:
- ✅ Data structures, parsing, algorithms, validation

The Rust code should NOT contain:
- ❌ HTML, CSS, DOM manipulation, presentation logic

## Architecture

```
src/
  lib.rs          Public API - WASM exports
  tableau.rs      Data structures and parsing
  rcd.rs          Recursive Constraint Demotion algorithm
```

### Module Responsibilities

- **lib.rs**: Entry points for JavaScript (`parse_tableau`, `run_rcd`)
- **tableau.rs**: Parse tab-delimited files into `Tableau` structs
- **rcd.rs**: Find stratified constraint rankings

See inline documentation in each module for details.

## Testing

```bash
cargo test
```

Tests live in each module's `#[cfg(test)]` section.
