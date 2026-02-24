# Rust Library

Core data structures, parsing, and algorithms compiled to WebAssembly. No presentation logic belongs here.

## Modules

```
src/
  lib.rs        WASM exports — public API surface for JavaScript
  tableau.rs    Data structures and tab-delimited file parsing
  apriori.rs    A priori rankings (partial orderings fed into RCD/LFCD)
  rcd.rs        Recursive Constraint Demotion
  bcd.rs        Biased Constraint Demotion (faithfulness delay + specificity)
  lfcd.rs       Low Faithfulness Constraint Demotion
  fred.rs       Ranking argumentation (FRED / MIB)
  maxent.rs     Maximum Entropy learning
  nhg.rs        Noisy Harmonic Grammar
  gla.rs        Gradual Learning Algorithm (Stochastic OT + online MaxEnt)
  factorial_typology.rs  Factorial Typology (FastRCD, cross-classification, T-order)
  hasse.rs      Hasse diagram DOT generation (FRed Hasse: fred_hasse_dot)
```

## Testing

Tests live in each module's `#[cfg(test)]` section. Every new algorithm or public function must have tests that validate correctness against expected outputs — use the example files in `examples/` where possible. Use `make test` to run the full test suite.

## Linting

Run `make lint` to check Clippy linting. All warnings are treated as errors (`-D warnings`). New code must pass lint before committing.
