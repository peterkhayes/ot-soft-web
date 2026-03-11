---
status: open
type: ux
priority: medium
depends_on: []
---

# Reduce test log output to save tokens

## Description

Test commands (`make test`, `make web-test`, etc.) produce verbose output even on success, which wastes context window tokens when used with AI coding assistants. Test output should be minimal by default, showing full logs only on failure or when explicitly requested.

## Options

1. **Quiet by default, verbose on failure:** Suppress stdout on passing tests, show full output only for failures.
2. **Verbose flag:** Add `make test-verbose` / `make web-test-verbose` targets that pass verbose flags, while the default targets stay quiet.
3. **Both:** Quiet by default with full output on failure, plus an explicit verbose flag for debugging.

## Files

- `Makefile` — test targets
