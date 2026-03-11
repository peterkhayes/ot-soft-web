---
status: open
type: ux
priority: medium
depends_on: []
---

# GLA panel layout polish

## Issues

1. **Controls take too much vertical space on desktop.** Each option group (Stochastic OT options, Learning schedule, Multiple runs) is stacked vertically, pushing the CTA buttons below the fold. On desktop, group related controls into rows (e.g. side-by-side columns for smaller option groups).

2. **Multiple runs count should be a text input.** The current dropdown with 10/100/1000 mirrors the VB6 menu structure, but there's no technical reason for these specific values. Replace the `<select>` with a numeric `<input>` to allow any count.

3. **Pairwise Ranking Probabilities section duplicates its header.** The web UI renders a `<h3>` "Pairwise Ranking Probabilities" and then the `pairwiseTable` text also contains the "5. Ranking Value to Ranking Probability Conversion" header from `format_pairwise_probabilities()`. Either strip the header from the formatted text or remove the `<h3>`.
