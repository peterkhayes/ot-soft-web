---
status: done
type: testing
priority: high
depends_on: []
---

# Conformance: HTML golden files use bare `<tr>/<td>` without `<table>` wrappers

## Affected cases
- `TinyIllustrativeFile_rcd_defaults_html`
- `TinyIllustrativeFile_bcd_defaults_html`
- `TinyIllustrativeFile_lfcd_defaults_html`
- `ilokano_rcd_defaults_html`

## Fix Applied

### Rust HTML output changes (`rcd.rs`, `fred.rs`)
- **Strata listing**: changed from `<ul>/<li>` to a 3-column `<table>` (Stratum / Constraint Name / Abbreviation), matching VB6.
- **Ranking arguments**: changed from `<pre>` dump to a 2-column `<table>` (ranking | &nbsp;), matching VB6. Added `ranking_strings()` helper on `FRedResult`.
- **Necessity table**: added header row (Constraint / Status) and switched from abbreviations to full constraint names, matching VB6.
- **Mini-tableau constraint sorting**: HTML mini-tableaux now sort constraints via `vb6_sort_constraint_slice`, matching the text formatter and VB6's `PrintTableaux.SortTheConstraints`.

### Conformance test changes (`conformance.rs`)
1. **Split on `</table>` boundaries** instead of requiring matched `<table>...</table>` wrappers — VB6 omits opening `<table>` tags.
2. **Split cells on `<td`/`<th` openings** instead of requiring matched open/close tags — VB6 omits `</td>` closing tags.
3. **Added `&#9758;` and `&nbsp` entity decoding** (VB6 uses decimal entity and omits trailing semicolons).
4. **Normalize CSS classes** to just "border" vs "no border" (cl4/cl8 → border; cl9/cl10/None → no border) since VB6 inconsistently applies background classes.

Also removed `skip` entries from all 4 HTML test cases in `conformance/manifest.json`.
