---
status: open
type: ux
priority: low
depends_on: []
---

# Options button layout should always be vertical

## Description

Option buttons/checkboxes are not consistently laid out across panels. Some panels combine options into a single horizontal line, which breaks on mobile. Options should always be stacked vertically for consistency and mobile-friendliness.

## Fix

Audit all panels (RcdPanel, NhgPanel, GlaPanel, MaxEntPanel, FrameworkPanel, FactorialTypologyPanel) and ensure option groups always use vertical stacking (`flex-direction: column`), never horizontal/inline layout.

## Files

- `web/style.css`
- `web/src/components/*.tsx` — check each panel's option layout
