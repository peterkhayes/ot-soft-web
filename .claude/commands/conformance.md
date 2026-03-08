# Conformance Tests

Manage conformance tests that compare Rust output against VB6 OTSoft golden files. Read `conformance/CLAUDE.md` for full documentation on the conformance test system.

Ask the user what they'd like to do:

---

## Option A — Run conformance tests

Run `make conformance-test` and report results. If any tests fail, analyze the diff output and suggest fixes.

---

## Option B — Collect golden files from Windows

Trigger golden file collection from VB6 OTSoft on the remote Windows machine.

1. Ask the user if they want to collect **all** cases or filter by regex (e.g. `"rcd_defaults$"`).
2. Run `./conformance/automation/remote_run.sh` with the appropriate arguments.
3. After collection completes, run `git pull` to fetch the new golden files.
4. Run `make conformance-test` to verify the new golden files match Rust output.
5. Report any new failures and suggest next steps.

If the server is not reachable, suggest `./conformance/automation/remote_run.sh --status` to check, or `--reload` to restart after a git pull on the Windows side.

---

## Option C — Add a new test case

Walk the user through adding a new conformance test case:

1. Ask which **algorithm** (`rcd`, `bcd`, `lfcd`, `maxent`, `factorial_typology`) and **input file**.
2. Ask about algorithm-specific parameters (refer to `conformance/CLAUDE.md` for the params schema).
3. Ask whether this is a **text** or **html** format test.
4. Choose a descriptive **case ID** following the convention: `{InputFileName}_{algorithm}_{variant}`.
5. Determine the **golden file path**: `conformance/golden/{InputFileName}/{algorithm}_{variant}.{txt|htm}`.
6. Read `conformance/manifest.json`, add the new case entry, and write it back.
7. Check if `rust/tests/conformance.rs` has a match arm for this algorithm/format combination. If not, add one.
8. Remind the user to collect the golden file (Option B) and then run tests (Option A).

---

## Option D — Review test status

1. Read `conformance/manifest.json` and summarize:
   - Total cases, how many have golden files, how many are skipped (and why).
   - Group by algorithm and input file.
2. Run `make conformance-test` and report pass/fail/skip counts.
3. Highlight any cases with `"skip"` set and suggest whether the skip reason is still valid.

---

## Option E — Update or fix a test case

Ask which test case to modify. Common operations:

- **Skip/unskip** a case: set or remove the `"skip"` field in `manifest.json`.
- **Update params**: modify algorithm parameters for a case.
- **Re-collect a golden file**: trigger targeted collection with `--filter`.
- **Investigate a failure**: run the Rust algorithm for the failing case, compare against the golden file, and analyze the diff.
