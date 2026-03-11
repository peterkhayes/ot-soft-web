---
status: open
type: ux
priority: medium
depends_on: []
---

# GLA "Run x times & Download" needs progress indicator

## Description

The main GLA "Run" button has a progress bar (via `useChunkedRunner`), but the "Run x times & Download" button only shows "Running…" text with no progress feedback. For large run counts this can take a long time with no indication of progress.

## Fix

Add a progress indicator for multiple runs. The multiple-runs path currently runs synchronously in a `setTimeout`. It should either:

1. Reuse `useChunkedRunner` to run each iteration with progress (e.g. "Run 3/100"), or
2. At minimum show a progress bar counting completed runs out of the total.

## Files

- `web/src/components/GlaPanel.tsx` — multiple runs handler (~line 218) and UI (~line 629)
- `web/src/hooks/useChunkedRunner.ts` — potentially extend for multi-run progress
