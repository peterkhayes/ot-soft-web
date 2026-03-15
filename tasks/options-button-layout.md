---
status: done
type: ux
priority: low
depends_on: []
---

# Options button layout should always be vertical

## Description

Option buttons/checkboxes are not consistently laid out across panels. Some panels combine options into a single horizontal line, which breaks on mobile. Options should always be stacked vertically for consistency and mobile-friendliness.

## Fix

Audit all panels (RcdPanel, NhgPanel, GlaPanel, MaxEntPanel, FrameworkPanel, FactorialTypologyPanel) and ensure option groups always use vertical stacking (`flex-direction: column`), never horizontal/inline layout. However, _independent_ option groups can live on the same row on desktop.

### Bad

Label
() Option 1   () Option 2    () Option 3

### Good (single group)

Label
() Option 1
() Option 2
() Option 3

### Good (multiple groups, desktop only)

Label A               Label B
() Option A1          () Option B1
() Option A2          () Option B2
() Option A3          () Option B3

## Files

- `web/style.css`
- `web/src/components/*.tsx` — check each panel's option layout
