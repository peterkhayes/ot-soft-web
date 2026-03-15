---
status: done
type: bug
priority: medium
depends_on: []
---

# Conformance: MaxEnt output differs from VB6 (title typo + structural differences)

## Affected cases
- `TinyIllustrativeFile_maxent_defaults`

## Resolution

Rewrote `MaxEntResult::format_output` to match VB6 section structure:
1. Fixed title: "Result" (singular) matching VB6's `PrintAHeader`
2. Added all 4 VB6 sections: Constraints/weights, Inputs/candidates/frequencies, Weights Found, Tableaux
3. Added learning time footer with normalization in conformance tests
4. Fixed section spacing to match VB6 blank-line patterns

Also fixed manifest: golden file was collected with 0 iterations (VB6 default), not 5. Updated `iterations: 0` to match.
