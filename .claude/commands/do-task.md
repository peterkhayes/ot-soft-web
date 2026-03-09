# Do a Task

General-purpose workflow for completing a task from TASKS.md. Use `/port-feature` instead if the task involves porting a VB6 feature.

Ask the user which task to work on. If unclear, read `TASKS.md` and suggest options.

---

## Step 1 — Understand the task

Read relevant source files, documentation, and code to understand what needs to be done. Key docs:

- `TASKS.md` and any linked files in `tasks/`
- `CLAUDE.md` files (`rust/CLAUDE.md`, `web/CLAUDE.md`, `source/CLAUDE.md`)
- Relevant source code in `rust/src/` and/or `web/src/`

Summarize your understanding and confirm the approach with the user before proceeding.

---

## Step 2 — Implement

Do the work. Follow the conventions in the relevant `CLAUDE.md` files for whichever part of the codebase you're modifying.

---

## Step 3 — Validate

Run the appropriate checks for the code you changed:

- **Rust changes**: `make lint` then `make test`
- **Web changes**: `make web-lint`, `make web-check`, `make web-test`
- **Both**: `make precommit` (runs everything)

Fix any failures before proceeding.

---

## Step 4 — Update documentation

Update all relevant docs:

- Mark completed items in `TASKS.md` and add any newly discovered tasks.
- Update any `CLAUDE.md` files whose descriptions have become inaccurate (e.g. new modules, changed file structure, new patterns).

---

## Step 5 — Finish

After every completed task, always do all of these steps in order:

1. If code changed since Step 3, run `make precommit` again. Otherwise skip — checks already passed.
2. `git commit` — write a descriptive message
3. `git push` — push to remote
