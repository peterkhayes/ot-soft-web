# ot-soft (modern version)

This repository ports an older program that does linguistic analysis to a modern web tech stack. There are three sub-folders, each with their own `claude.md` file.

## Contents

### `source`

The original source code, plus screenshots of the original user interface.

### Rust

Rust library that implements the same features as the original source code. This will be compiled to Wasm and used in the Web interface.

### Web interface

Lightweight HTML/CSS/Javascript layer that wraps the Wasm code with a GUI.

### Examples

Contains folders with expected input/output files.

## Principles

### Work in progress

This project is in-progress, and not all features of the original source work yet.

### Fidelity to the original

Maintain **logical** fidelity with original code as closely as possible. If the original code has bugs, reproduce them in the Rust code to the degree possible.

To help ensure this, the files in `examples` can be used as a test suite.

However, maintain **visual** fidelity only loosely. Keep the available options, buttons, and menus the same, but use normal HTML/CSS with simple, readable styling.

### Modularity

All data and computational logic should live in the Rust codebase. The Web codebase should only include presentational logic.
