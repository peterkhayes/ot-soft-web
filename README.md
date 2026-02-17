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
- Python 3 (for development server)

Install wasm-pack if you haven't already:
```bash
curl https://rustwasm.github.io/wasm-pack/installer/init.sh -sSf | sh
```

### Quick Start

Build and run the development server:
```bash
make dev
```

Then open http://localhost:8000 in your browser

### Available Commands

```bash
make help     # Show all available commands
make build    # Build Rust to WebAssembly
make test     # Run Rust tests
make serve    # Start development server
make dev      # Build and serve (one command)
make clean    # Clean build artifacts
make check    # Check Rust code without building
make fmt      # Format Rust code
```

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
make test
```

All tests should pass before committing changes.

### Development Workflow

1. Make changes to Rust code in `rust/src/`
2. Rebuild: `make build`
3. Refresh browser to see changes

Or use one command:
```bash
make dev
```

This builds and starts the server in one step.

### Code Quality

```bash
make check   # Quick syntax check without building
make fmt     # Auto-format Rust code
```

## Principles

- **Logical fidelity**: Maintain the same computational logic as the original VB6 code
- **Visual freedom**: Modern, simple HTML/CSS design (no need to match the original UI exactly)
- **Clear separation**: All logic in Rust, only presentation in JavaScript
