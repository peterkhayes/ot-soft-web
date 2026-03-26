---
status: done
---

# Break up MaxEntPanel.tsx

At 393 lines, `MaxEntPanel.tsx` mixes configuration options, runner logic, and result display. Split into focused sub-components following the same pattern as the `gla/` subdirectory.
