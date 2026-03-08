# Port a VB6 Feature

This skill guides you through porting a feature from the original VB6 OTSoft to the modern Rust + Web stack.

Ask the user which feature to port. If unclear, read `TASKS.md` and suggest the next unimplemented feature.

---

## Step 1 — Read the source documentation

Read the relevant sections in these files to understand the feature's parameters, algorithm, outputs, and UI:

- `source/ALGORITHMS.md`
- `source/INPUT_FORMAT.md`
- `source/OUTPUTS.md`
- `source/USER_INTERFACE.md`

Summarize what you've learned and confirm with the user before proceeding.

---

## Step 2 — Read the VB6 source code

Find and read the original implementation in `source/vb6/`. Use the documentation from Step 1 as a guide to locate the right files and functions. Pay attention to edge cases and any bugs that should be reproduced (per the project's fidelity principle).

---

## Step 3 — Implement in Rust

a. Write the implementation following patterns established in existing Rust modules (see `rust/CLAUDE.md` for module layout).
b. Write tests that validate against expected outputs (add example files to `examples/` if needed).
c. Run `make lint` to check for style issues, then `make test` to confirm correctness.
d. If a source file grows too large, split it into separate modules.

---

## Step 4 — Implement Web UI

a. Add any necessary UI elements (buttons, inputs, displays) to the web interface.
b. All business logic must remain in Rust — the web layer only handles presentation and calls into WASM (see `web/CLAUDE.md` for patterns and conventions).
c. Write a test in `web/tests/flows/` for the new feature. For deterministic algorithms, use inline snapshots on download content. For stochastic algorithms, use structural assertions only (headings, labels, download filename).
d. Run `make web-lint`, `make web-check`, and `make web-test` to confirm the web code is clean and tests pass.

---

## Step 5 — Update documentation

Update all relevant docs:

- Mark completed items in `TASKS.md` and add any newly discovered tasks.
- Update any `CLAUDE.md` files whose descriptions have become inaccurate (e.g. new modules, changed file structure, new patterns).

---

## Step 6 — Finish

After every completed task, always do all of these steps in order:

1. `make precommit` — run all checks (Rust lint + tests + build, web lint + format check + type-check + tests)
2. `git commit` — write a descriptive message
3. `git push` — push to remote
