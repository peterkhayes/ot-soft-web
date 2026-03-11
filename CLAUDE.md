# OTSoft (modern version)

This repository ports an older program that does linguistic analysis to a modern web tech stack. See `TASKS.md` for open tasks.

## Principles

### Fidelity to the original

Maintain **logical** fidelity with original code as closely as possible. If the original code has bugs, reproduce them in the Rust code to the degree possible.

To help ensure this, the files in `examples` can be used as a test suite.

However, maintain **visual** fidelity only loosely. Keep the available options, buttons, and menus the same, but use normal HTML/CSS with simple, readable styling.

### Modularity

All data and computational logic should live in the Rust codebase. The Web codebase should only include presentational logic.

## Commands

Always use `make` targets rather than invoking `cargo`, `npm`, or other tools directly.

| Command | Purpose |
|---------|---------|
| `make build` | Compile Rust to WASM |
| `make test` | Run Rust tests (quiet by default) |
| `make lint` | Clippy lint (warnings = errors) |
| `make check` | Rust type-check only |
| `make web-test` | Run web tests (quiet by default) |
| `make web-check` | TypeScript type-check |
| `make web-lint` | ESLint the web frontend |
| `make web-fmt` | Format web code (Prettier) |
| `make precommit` | All checks (lint, test, build, web) |
| `make serve` | Start Vite dev server |
| `make dev` | Build WASM + start dev server |
| `make conformance-test` | Run conformance tests vs VB6 golden files |

Flags can be passed as `VAR=1`:

| Flag | Applies to | Effect |
|------|-----------|--------|
| `VERBOSE` | `test`, `web-test` | Show full output |
| `FIX` | `lint`, `web-lint` | Auto-fix issues |
| `UPDATE` | `web-test` | Update inline snapshots |
| `CHECK` | `web-fmt` | Check only, no writes |
