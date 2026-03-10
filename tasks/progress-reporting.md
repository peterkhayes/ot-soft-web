---
status: open
type: ux
priority: medium
depends_on: []
---

# Progress reporting for long-running jobs

## Description
Long-running algorithms (e.g. GLA with many iterations) give no feedback during execution. Requires WASM progress callbacks.

## Acceptance Criteria
- [ ] Progress is reported during long-running algorithm execution
- [ ] WASM callback mechanism for progress updates is implemented
