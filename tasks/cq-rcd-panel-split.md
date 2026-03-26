---
status: done
---

# Break up RcdPanel.tsx

At 413 lines, `RcdPanel.tsx` handles algorithm selection, FRed options, a priori rankings, runner logic, and result display. Split into focused sub-components following the same pattern as the `gla/` subdirectory.
