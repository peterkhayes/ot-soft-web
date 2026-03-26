---
status: done
---

# Break up GlaPanel.tsx

At 778 lines, `GlaPanel.tsx` handles too many concerns:
- Framework selection (maxent/stochastic modes)
- Multiple parameter categories (cycles, plasticity, noise, learning schedule)
- Two separate runner systems (single run and multiple runs)
- Four download handlers
- Complex conditional UI rendering

Split into smaller sub-components for parameter groups and result display.
