# OTSoft (Web Version)

[OTSoft](https://brucehayes.org/otsoft/) (short for Optimality Theory Software) is originally a Windows program meant to facilitate analysis in [Optimality Theory](https://en.wikipedia.org/wiki/Optimality_theory) and related frameworks by using algorithms to do tasks that are too large or complex to be done reliably by hand.

OTSoft is written in [Visual Basic 6](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation) and is distributed as [a Windows installation file](https://brucehayes.org/otsoft/OTSoft2.5.zip). This leads to difficulties in portability, especially for non-Windows computers.

This project seeks to address those by creating a web version (Rust/WASM/TypeScript) that can be opened on any device with a modern web browser.

## Philosophy

This project is developed using an "interesting" development workflow that is newly possible in 2026. In short:

- The original source code is copied verbatim into this repo.
- [Claude Code](https://code.claude.com/docs/en/overview) analyses that code and, with some human supervision, writes the new Rust and Web code to match.
- A corpus of test files, outputted from the original code, is used as a verification mechanism.

This solves the following problems:

- **Bruce Hayes** (the primary OTSoft author) isn't familiar with web programming technologies.
- **Peter Hayes** (me, his son) is familiar with modern software development practices, but isn't familiar with linguistics.
- **Claude Code** can write decent code very quickly and cheaply.

### Technologies

The web is a universally-available software platform, more portable than Python, R, or other common scientific languages. Web applications can be loaded on any device without installation.

Given that I wanted a web app, I chose the following technologies:

- **[Rust](https://rust-lang.org/)** is a fast and safe language for the algorithmic code.
- **[WebAssembly (Wasm)](https://webassembly.org/)** and **[wasm-pack](https://github.com/drager/wasm-pack)** allow Rust code to be compiled into a form that runs in modern web browsers.
- **[TypeScript](https://www.typescriptlang.org/)** is a type-safe language that compiles to JavaScript, which is the only language supported by web browsers for client-side code.
- **[React](https://react.dev/)** is a user-interface framework for JavaScript/TypeScript.
- **[Vite](https://vite.dev/)** is a build tool for web applications.

A pure-TypeScript implementation is also possible, but would likely be slower than the Rust/Wasm version.

### Status

At present, the port is not complete. It has most of the core OTSoft functionality, but is missing many features. I hope to continue development as my Claude Code token limits allow.

Once complete, development can move to a workflow where new changes made by Bruce Hayes (or collaborators) are pulled into this repo; Claude can then analyze the diff and implement the changes.

### Risks

This software will produce divergent results compared to the original code if Claude Code makes mistakes in its code, and if those mistakes are not caught via testing.

As a result, this project is highly reliant on its test suite, which itself relies on a well-chosen corpus of test cases (input + parameters + output). That corpus is not well-developed at this time. **Until that point, this project should not be used for real work.**

## Project Structure

- `source/` - original VB6 source code, UI screenshots, and reference documentation written by Claude
- `rust/` - rust code
- `web/` - web app code
- `examples/` - test input/output files

## Setup

### Prerequisites

- [Rust](https://rustup.rs/) (latest stable)
- [wasm-pack](https://rustwasm.github.io/wasm-pack/installer/)
- [Node.js](https://nodejs.org/) (for the Vite dev server)

### Quick Start

Build and run the development server:

```bash
make dev
```

Then open http://localhost:5173 in your browser.

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

See `CLAUDE.md`.
