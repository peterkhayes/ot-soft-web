---
status: done
type: ux
priority: medium
depends_on: []
---

# Render Pairwise Ranking Probabilities as HTML table

## Description

The "Pairwise Ranking Probabilities" section in GLA results currently renders the data as a pre-formatted text string (`<pre>`) from the Rust `gla_pairwise_probabilities()` function. Other results sections render as proper HTML tables. This section should too.

## Fix

Either:
1. Return structured data from Rust (e.g. a 2D array + headers) instead of a pre-formatted string, and render it as an HTML `<table>` in the React component, or
2. Parse the pre-formatted string in TypeScript and convert it to a table.

Option 1 is preferred for consistency with the modular architecture (data logic in Rust, presentation in web).

## Files

- `rust/src/gla.rs` — `gla_pairwise_probabilities()` or related formatting
- `web/src/components/GlaPanel.tsx` — rendering (~lines 725-730)
