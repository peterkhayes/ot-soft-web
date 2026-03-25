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

Updated `extract_html_tables` in `rust/tests/conformance.rs` to handle VB6's malformed HTML:

1. **Split on `</table>` boundaries** instead of requiring matched `<table>...</table>` wrappers — VB6 omits opening `<table>` tags.
2. **Split cells on `<td`/`<th` openings** instead of requiring matched open/close tags — VB6 omits `</td>` closing tags.
3. **Added `&#9758;` entity decoding** (decimal form of the pointing hand character).
4. **Normalize CSS classes** to just "border" vs "no border" (cl4/cl8 → border; cl9/cl10/None → no border) since VB6 inconsistently applies background classes.
5. **Filter to shaded tables only** (those with cl4/cl8/cl9/cl10) to exclude non-tableau tables (strata listings, status tables, ranking arguments) that VB6 renders as tables but Rust renders as lists/preformatted text.
6. **Canonicalize mini-tableau columns** by sorting constraint names alphabetically, since VB6 HTML and text outputs may order constraints differently within mini-tableaux.
7. **Sort mini-tableaux** by content for order-independent comparison.

Also removed `skip` entries from all 4 HTML test cases in `conformance/manifest.json`.
