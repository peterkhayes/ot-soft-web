# OTSoft (modern version)

This repository ports an older program that does linguistic analysis to a modern web tech stack. See `TASKS.md` for current porting status.

## Principles

### Fidelity to the original

Maintain **logical** fidelity with original code as closely as possible. If the original code has bugs, reproduce them in the Rust code to the degree possible.

To help ensure this, the files in `examples` can be used as a test suite.

However, maintain **visual** fidelity only loosely. Keep the available options, buttons, and menus the same, but use normal HTML/CSS with simple, readable styling.

### Modularity

All data and computational logic should live in the Rust codebase. The Web codebase should only include presentational logic.

## Workflow for Adding Features

Follow these steps when porting a new feature from the VB6 source:

1. **Read the source documentation** — Start with the relevant sections in `source/ALGORITHMS.md`, `source/INPUT_FORMAT.md`, `source/OUTPUTS.md`, and `source/USER_INTERFACE.md` to understand the feature's parameters, algorithm, outputs, and UI.

2. **Read the VB6 source code** — Find and read the original implementation in `source/vb6/`. Use the documentation as a guide to locate the right files and functions. Pay attention to edge cases and any bugs that should be reproduced.

3. **Implement in Rust**
   a. Write the implementation following patterns established in existing Rust modules.
   b. Write tests that validate against expected outputs (add example files to `examples/` if needed).
   c. Run `make lint` to check for style issues, then `make test` to confirm correctness.
   d. If a source file grows too large, split it into separate modules.

4. **Implement Web UI**
   a. Add any necessary UI elements (buttons, inputs, displays) to the web interface.
   b. All business logic must remain in Rust — the web layer only handles presentation and calls into Wasm.
   c. Write a test in `web/tests/flows/` for the new feature. For deterministic algorithms, use inline snapshots on download content. For stochastic algorithms, use structural assertions only (headings, labels, download filename). See `web/CLAUDE.md` for patterns.
   d. Run `make web-test` to confirm the tests pass.

5. **Update documentation** — Update all relevant docs: mark completed items in `TASKS.md`, add any newly discovered tasks, and update any `CLAUDE.md` files whose descriptions have become inaccurate (e.g. new modules, changed file structure, new patterns).

6. **Finish** — After every completed task, always do all of these steps in order:
   a. `make lint` — check Rust style (no warnings allowed)
   b. `make test` — run Rust tests
   c. `make build` — rebuild WASM
   d. `make web-test` — run web tests
   e. `git commit` - write a descriptive message
   f. `git push` — push to remote

## Commands

Always use `make` targets rather than invoking `cargo`, `npm`, or other tools directly. See the Makefile for available targets (`make test`, `make lint`, `make build`, `make check`, `make serve`, etc.).
