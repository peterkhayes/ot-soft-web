# OTSoft (Web Version)

A modern web-based linguistic analysis tool, ported from VB6 to Rust/WebAssembly.

## Project Structure

- `source/` - Original VB6 source code, UI screenshots, and reference documentation
- `rust/` - Rust implementation compiled to WebAssembly
- `web/` - HTML/CSS/JavaScript web interface
- `examples/` - Test input/output files

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

Then open http://localhost:8000 in your browser.

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

## Development

### Running Tests

```bash
make test
```

All tests should pass before committing changes.

### Development Workflow

1. Make changes to Rust code in `rust/src/`
2. Rebuild: `make build`
3. Refresh browser to see changes

Or use `make dev` to build and serve in one step.
