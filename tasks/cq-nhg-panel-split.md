---
status: done
---

# Break up NhgPanel.tsx

At 522 lines, `NhgPanel.tsx` is the second-largest panel after GlaPanel (which was already split). It handles framework options, multiple parameter groups, runner logic, and result display. Split into focused sub-components following the same pattern as the `gla/` subdirectory.
