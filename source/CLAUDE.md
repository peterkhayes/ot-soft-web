# Source Files

This directory contains the original VB6 source code and UI screenshots from the legacy OT-Soft application.

## Contents

- `vb6/` - Original Visual Basic 6 source code
- `screenshots/` - Screenshots of the original application UI

## Guidelines

**DO NOT modify any code files in the source/vb6 folder.**

All code changes should be made in the `rust/` or `web/` directories instead.

## Authoritative Sources for VB6 Defaults

When verifying that modern code matches VB6 defaults, use these sources **in priority order**:

1. **The `.frm` source files** — form-level initialization code is the ground truth for what the UI actually showed users.
2. **`source/vb6/Module1.bas`** — hardcoded VB6 constants (e.g. `DefaultUpperFaith`, `DefaultLowerFaith`).
3. **`source/ALGORITHMS.md` and `source/USER_INTERFACE.md`** — useful as a cross-reference, but may lag behind the source code.

**Do NOT use `source/vb6/OTSoftRememberUserChoices.txt` as a source of defaults.** This file is a saved user state (it even references a specific user's file paths), not the application's built-in defaults. Its values for things like plasticity and test trials may differ significantly from the real defaults.
