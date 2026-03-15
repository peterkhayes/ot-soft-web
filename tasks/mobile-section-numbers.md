---
status: done
type: ux
priority: low
depends_on: []
---

# Mobile: section numbers waste vertical space

## Description

On mobile layout (max-width 768px), the `.panel-header` switches to `flex-direction: column`, which stacks the section number (e.g. "01") below the section name. This wastes a full line of vertical space per panel.

On desktop the number sits in the top-right corner via `justify-content: space-between`. The mobile layout should keep the number in the top-right corner too, rather than stacking it below.

## Fix

Remove or override the `flex-direction: column` for `.panel-header` at the mobile breakpoint, keeping `flex-direction: row` with `justify-content: space-between` so the number stays top-right at all screen sizes.

## Files

- `web/style.css` — `.panel-header` and its mobile media query
