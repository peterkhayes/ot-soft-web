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
```

Tests live in each module's `#[cfg(test)]` section. Use `make test` to run them.
