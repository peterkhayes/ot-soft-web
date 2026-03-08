# OTSoft (modern version)

This repository ports an older program that does linguistic analysis to a modern web tech stack. See `TASKS.md` for current porting status.

## Principles

### Fidelity to the original

Maintain **logical** fidelity with original code as closely as possible. If the original code has bugs, reproduce them in the Rust code to the degree possible.

To help ensure this, the files in `examples` can be used as a test suite.

However, maintain **visual** fidelity only loosely. Keep the available options, buttons, and menus the same, but use normal HTML/CSS with simple, readable styling.

### Modularity

All data and computational logic should live in the Rust codebase. The Web codebase should only include presentational logic.

## Workflow Commands

- `/port-feature` ports a VB6 feature to the modern stack.
- `/conformance` manages conformance tests against VB6 golden files.

## Commands

Always use `make` targets rather than invoking `cargo`, `npm`, or other tools directly. See the Makefile for available targets (`make test`, `make lint`, `make build`, `make check`, `make serve`, etc.).
