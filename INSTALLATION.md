# Installation

## Prerequisites

- [Rust](https://rustup.rs/) (latest stable)
- [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/)
- [Node.js](https://nodejs.org/) (for the Vite dev server)

## Running the App

Build and run the development server:

```bash
make dev
```

It will serve the application from http://localhost:5173.

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

### Project Structure

- `source/` - original VB6 source code, UI screenshots, and reference documentation written by Claude
- `rust/` - rust code
- `web/` - web app code
- `examples/` - test input/output files

See `CLAUDE.md` for more information.
