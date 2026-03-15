---
status: open
type: ux
priority: low
depends_on: []
---

# RCD: consider defaulting "Show details" to off

## Description

The "Show details of argumentation" checkbox currently defaults to ON (matching VB6). BPH's feedback suggests most users don't need it — the output file is much longer with details enabled.

Consider changing the default to OFF to produce shorter, more focused output by default while keeping the option available.

**Note:** This may have already been addressed since the comparison was made (2026-03-07). Verify current default before changing.

## Reference

- BPH comment: "Output file is much longer with checking of menu item 'Show details of argumentation'; probably most users..."
- BPH comparison output (without details): 160 lines
- PKH comparison output (with details): ~2000+ lines

## Acceptance Criteria

- [ ] Decide whether to change default (requires BPH input — sentence was incomplete)
- [ ] If changing, update default in `FredOptions::new()` and `useLocalStorage` initial value
