---
status: open
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

## Description

The `extract_html_tables` function in `rust/tests/conformance.rs` searches for `<table>...</table>` wrappers. VB6 OTSoft golden HTML files contain `<tr>/<td>` elements directly in the `<body>` without enclosing `<table>` tags, so extraction yields 0 tables while Rust output (which uses proper `<table>` elements) yields 8–13 tables.

## Reproduction

Run `make conformance-test` — all four `_html` cases report:
```
table count mismatch — expected 0, actual 8
```

## Fix Options

1. **Update `extract_html_tables`** to also collect bare `<tr>` sequences (not inside a `<table>`) into synthetic table grids.
2. **Re-collect golden files** once Rust output format is confirmed correct, comparing Rust HTML against itself (not VB6 HTML). This would require reconsidering the purpose of HTML conformance tests.

## Acceptance Criteria
- [ ] All four HTML conformance cases pass without a skip
- [ ] `make conformance-test` reports 0 failures
