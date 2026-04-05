---
status: done
type: feature
priority: low
depends_on: []
---

# HowIRanked log file

## Description
VB6 can produce a "HowIRanked" log file explaining the ranking decisions. Port this feature.

## Acceptance Criteria
- [x] Log file output available that explains ranking decisions

## Resolution
Implemented a stateful logger infrastructure with thread-local buffer that captures all `ot_log!` calls.
Detailed step-by-step decision logging added for all three ranking algorithms (RCD, BCD, LFCD),
matching VB6's narrative structure. Web UI provides a "Download Log" button after each algorithm run.
