# OT-Soft

A modern web-based linguistic analysis tool, ported from VB6 to Rust/WebAssembly.

## Project Structure

- `source/` - Original VB6 source code and UI screenshots
- `rust/` - Rust implementation compiled to WebAssembly
- `web/` - HTML/CSS/JavaScript web interface

## Setup

### Prerequisites

- [Rust](https://rustup.rs/) (latest stable)
- [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/)

Install wasm-pack if you haven't already:
```bash
curl https://rustwasm.github.io/wasm-pack/installer/init.sh -sSf | sh
```

### Building

1. Build the Rust code to WebAssembly:
```bash
cd rust
wasm-pack build --target web --out-dir ../web/pkg
```

2. Serve the web interface:
```bash
cd ../web
python3 -m http.server 8000
```

3. Open http://localhost:8000 in your browser

## Current Status

The basic shell is working with the following features:

### Implemented
- ✅ Tab-delimited OT tableau parser
- ✅ Data structures for constraints, candidates, and input forms
- ✅ Web interface for loading and displaying tableaux
- ✅ "Tiny" example file parsing

### Testing

The web interface is currently running at http://localhost:8000

Try it out by:
1. Clicking "Load Tiny Example" to parse the included tiny illustrative file
2. Or upload your own tab-delimited .txt tableau file

The tiny example file format:
```
Row 1: Full constraint names (tab-delimited, first two columns empty)
Row 2: Constraint abbreviations (tab-delimited, first two columns empty)
Row 3+: Input form, candidate form, violation counts...
```

## Development

### Running Tests

Test the Rust parsing logic:
```bash
cd rust
cargo test
```

All tests should pass before committing changes.

### Rebuilding

After making changes to the Rust code, rebuild with:
```bash
cd rust && PATH="$HOME/.cargo/bin:$PATH" wasm-pack build --target web --out-dir ../web/pkg
```

The web interface will automatically pick up the changes on refresh.

## Principles

- **Logical fidelity**: Maintain the same computational logic as the original VB6 code
- **Visual freedom**: Modern, simple HTML/CSS design (no need to match the original UI exactly)
- **Clear separation**: All logic in Rust, only presentation in JavaScript
