---
status: done
type: ux
priority: medium
depends_on: []
---

# Add "Done" indicator when an algorithm finishes

## Description
Users have no visual feedback when an algorithm completes. The Run button shows a loading animation during execution, and results appear when done, but there's no explicit "done" signal — easy to miss if scrolled away or not watching.

## Options Considered

1. **Visual flash/highlight on the status message** — Brief CSS animation (pulse/glow) on the success/failure status line when it first appears. Low-friction, unobtrusive. Works well for fast algorithms.
2. **Toast notification** — Small temporary banner (top/bottom of screen) that auto-dismisses. Works even if the user has scrolled away from the panel.
3. **Browser notification** — OS-level notification via Notifications API. Good for long-running algorithms when user switches tabs. Requires permission; heavy for sub-second operations.
4. **Audio beep** — Short sound on completion. Would need a user preference toggle.
5. **Combination** — e.g., visual highlight always + optional sound toggle.

**Recommendation**: Option 1 (visual flash) as baseline — simplest, always works, no preferences needed. A brief pulse animation on the status message (1-2 scale/glow pulses) when it first renders.

## Acceptance Criteria
- [ ] A visible "Done" or completion indicator appears after any algorithm finishes running
