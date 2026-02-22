# OTSoft (Web Version)

[OTSoft](https://brucehayes.org/otsoft/) (short for Optimality Theory Software) is a Windows program meant to facilitate analysis in [Optimality Theory](https://en.wikipedia.org/wiki/Optimality_theory) and related frameworks by using algorithms to do tasks that are too large or complex to be done reliably by hand.

OTSoft is written in [Visual Basic 6](https://learn.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/visual-basic-6.0-documentation) and is distributed as [a Windows installation file](https://brucehayes.org/otsoft/OTSoft2.5.zip). This leads to difficulties in portability, especially for non-Windows computers.

This project seeks to address those by creating a web version (Rust/WASM/TypeScript) that can be opened on any device with a modern web browser.

## Usage

Information on what OTSoft does and how to use it can be found on [the main webpage](https://brucehayes.org/otsoft/).

To use this web version, simply open https://peterkhayes.github.io/ot-soft-web in your web browser. No installation is required.

## About

This project uses an "interesting" development workflow newly possible in 2026. In short:

- The original source code was copied verbatim into this repository.
- [Claude Code](https://code.claude.com/docs/en/overview) produced a detailed analysis of the code.
- With some human supervision, Claude Code ports the code to new technologies.
- A corpus of test data files (generated using the VB6 version) is used as a verification mechanism.

This produces a division of labor as follows:

- **Bruce Hayes** (the primary OTSoft author) understands the source code and the expected behavior.
- **Peter Hayes** (me, his son) understands web technologies and software development workflows.
- **Claude Code** can write decent code very quickly and cheaply, if given clear specifications.

### Status

At present, the port is not complete. It has most of the core OTSoft functionality, but is missing many features. It also has not been properly validated for correctness.

It remains to be seen if the development approach will succeed, but it has been an interesting journey so far, and I hope to continue as my Claude Code token limits allow.

Once complete, development can move to a workflow where new changes made by Bruce Hayes (or collaborators) are pulled into this repo; Claude can then analyze the diff and implement the changes.

### Risks

This software will produce divergent results compared to the original code if Claude Code makes mistakes in its port, and if those mistakes are not caught via testing.

As a result, this project is highly reliant on its test suite, which itself relies on a well-chosen corpus of test cases (input + parameters + output). But that corpus is not well-developed at this time. **Until that point, this project should not be used for real work.**

### Technologies

The web is a universally-available software platform, more portable than Python, R, or other common scientific languages. Web applications can be loaded on any device without installation.

Given that I wanted a web app, I chose the following technologies:

- **[Rust](https://rust-lang.org/)** is a fast and safe language for the algorithmic code.
- **[WebAssembly (Wasm)](https://webassembly.org/)** and **[wasm-pack](https://github.com/drager/wasm-pack)** allow Rust code to be compiled into a form that runs in modern web browsers.
- **[TypeScript](https://www.typescriptlang.org/)** is a type-safe language that compiles to JavaScript, which is the only language supported by web browsers for client-side code.
- **[React](https://react.dev/)** is a user-interface framework for JavaScript/TypeScript.
- **[Vite](https://vite.dev/)** is a build tool for web applications.

A pure-TypeScript implementation is also possible, but would likely be slower than the Rust/Wasm version.

## Development

The following steps are only required if you want to modify the source code. As mentioned above, if you simply want to _use_ the software, you can open the web application at https://peterkhayes.github.io/ot-soft-web.

1. Install [Claude Code](https://code.claude.com/docs/en/overview#get-started). Other AI coding tools such as [Cursor](https://cursor.com/), [Codex](https://openai.com/codex/), or [Antigravity](https://antigravity.google/) would likely work as well.
2. Tell Claude Code to follow the setup instructions in [INSTALLATION.md](INSTALLATION.md), or follow them yourself.
3. Open http://localhost:5173 in your browser to use the application.
