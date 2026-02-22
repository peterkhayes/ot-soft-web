# Update VB6 Source Code

This skill imports a new version of the OTSoft VB6 source code, replaces the checked-in copy, analyzes what changed, and prepares a task list for porting the changes.

Follow these steps in order. At each step requiring user input, pause and wait for a response before continuing.

---

## Step 1 — Check model

This workflow involves carefully reading VB6 source code, writing technical documentation, and generating precise task descriptions. It benefits from using the most capable model available to you.

Check which model you are currently running on. If you are not running the most capable model in your current Claude family (typically an Opus-tier model), suggest the user switch before continuing:

> "This workflow involves careful code analysis and documentation. For best results, I recommend using the most capable model you have access to. You can switch by running `/model` in the CLI."

Do not block on this — just suggest it and continue if the user acknowledges.

---

## Step 2 — Verify clean git state

Run `git status` and `git branch --show-current`.

- If the working directory is **not clean** (uncommitted changes or untracked files), stop and inform the user. Do not proceed.
- If the working directory is clean but the current branch is **not `main`**, switch to `main` with `git checkout main` and inform the user you did so.

---

## Step 3 — Ask for source location

Ask the user:

> "Where should I find the new OTSoft source code? Please provide either a local file path or a URL."

Wait for a response before continuing.

---

## Step 4 — Prepare temp directory

Create the directory `.tmp/vb6-update/` at the project root. `.tmp/` is already listed in `.gitignore`.

---

## Step 5 — Download or copy source into temp directory

Based on the user's answer in Step 3:

- If it is a **local path to a directory**: copy the contents into `.tmp/vb6-update/`.
- If it is a **local path to a `.zip` file**: copy the zip to `.tmp/` and unzip it into `.tmp/vb6-update/`.
- If it is a **URL to a `.zip` file**: download the zip to `.tmp/` using `curl` or `wget`, then unzip it into `.tmp/vb6-update/`.

After extracting, inspect the contents of `.tmp/vb6-update/` to locate the root of the source code — the directory that contains `Module1.bas`. It may be in a subdirectory rather than at the extract root. Note this path for use in the next steps.

---

## Step 6 — Check versions

Find the version number and release date in both the **current** and **new** source by reading `Module1.bas` from each. Look for these lines:

```
Public Const gMyVersionNumber As String = "..."
Public Const gMyReleaseDate As String = "..."
```

- Current source: `source/vb6/Module1.bas`.
- New source: the `Module1.bas` found in Step 5.

Report both versions to the user. If the new version number is **lower** or new release date is **earlier** than the current version, ask the user to confirm they want to proceed before continuing.

---

## Step 7 — Replace source/vb6/

Delete the entire contents of `source/vb6/` and replace them with **all files** from the source root identified in Step 5. The result should be an exact copy — every file present in the new source, nothing added or omitted.

Important:

- Replace **only** `source/vb6/`.
- Do not modify `source/*.md` or any other files in the `source/` folder.

After copying, delete `.tmp/vb6-update/` and any downloaded zip file in `.tmp/`. Do not commit yet.

---

## Step 8 — Analyze the diff

Run `git diff --stat` to see what changed at a high level. It should only contain changes from `source/vb6/`. If other files have changed, determine what went wrong.

Then run `git diff` to see what changed in detail. Read the changed files carefully.

Summarize the changes for the user in plain language — what algorithms were modified, what new options or outputs were added, what was removed or renamed. Ask:

> "Does this summary look right? Any corrections or additions before I update the documentation?"

Wait for a response before continuing.

---

## Step 9 — Update source documentation

Based on the diff and the user's feedback from Step 8, update the relevant files in `source/`:

- `source/ALGORITHMS.md` — if any algorithms were added, changed, or removed
- `source/INPUT_FORMAT.md` — if the input format changed
- `source/OUTPUTS.md` — if outputs changed
- `source/USER_INTERFACE.md` — if UI options or behavior changed

Try to carefully follow the source code diffs and the user's instructions. Continue to ask for clarification as needed.

---

## Step 10 — Update TASKS.md

Read `TASKS.md`. Based on the diff and the updated documentation, add new tasks for each change that needs to be ported to Rust/Wasm and the web UI. Modify existing tasks that may have been affected by the changes. Keep the existing format and legend.

---

## Step 11 — Commit and push

Stage and commit all changes: the updated `source/vb6/` files, updated `source/*.md` docs, and updated `TASKS.md`.

Use a commit message like: `Update VB6 source to v{new_version} ({release_date})`

Push to `main`.

---

## Step 12 — Suggest next steps

Tell the user:

> "Source code updates are complete! Should I start working on [first unimplemented task from TASKS.md]?"
