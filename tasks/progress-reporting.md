---
status: open
type: ux
priority: medium
depends_on: []
---

# Progress reporting for long-running jobs

## Description
Long-running algorithms (GLA default 1M cycles) block the main thread with no feedback. The solution is chunked execution: break computation into pieces, yield to the browser between chunks via `requestAnimationFrame`, and update a progress bar.

## Plan

### Rust: `GlaRunner` struct (gla.rs, lib.rs)

An internal `ChunkedRunner` trait defines the contract (not exported via wasm_bindgen):

```rust
trait ChunkedRunner {
    fn run_chunk(&mut self, max_work: usize) -> bool; // true when done
    fn progress(&self) -> Vec<f64>; // [completed, total]
}
```

`GlaRunner` is a `#[wasm_bindgen]` struct that holds the algorithm state currently local to `run_gla_with_schedule()`. It exposes:

- `new(text, opts)` — parse tableau, build schedule, initialize state
- `run_chunk(max_trials)` — advance up to N trials, return true when done (learning + testing)
- `progress()` — return `[completed, total]`
- `into_result()` — extract `GlaResult` after completion

Existing `run_gla` / `run_gla_with_schedule` stay unchanged for tests.

### TypeScript: `useChunkedRunner` hook (hooks/useChunkedRunner.ts)

A `ChunkedRunner` TS interface mirrors the Rust trait. A `RunnerState<T>` discriminated union tracks status:

```typescript
type RunnerState<T> =
  | { status: 'idle' }
  | { status: 'running'; completed: number; total: number }
  | { status: 'done'; result: T }
  | { status: 'error'; error: string }
```

The hook uses `useReducer` for clean state transitions. It takes a runner factory, a result extractor, and a chunk size. Returns `{ state, run }`. Uses `requestAnimationFrame` between chunks.

### GlaPanel update

Replace the current `setTimeout` + synchronous `run_gla()` with `useChunkedRunner`. Show a progress bar when `status === 'running'`.

### Scope

GLA only for now. NHG and Factorial Typology can reuse the same hook and follow the same runner pattern later.

## Acceptance Criteria
- [ ] `GlaRunner` struct with chunked execution in Rust
- [ ] `ChunkedRunner` trait in Rust (internal, for consistency)
- [ ] `useChunkedRunner` hook with `RunnerState` union and `useReducer`
- [ ] GlaPanel uses chunked runner with progress bar
- [ ] Existing `run_gla` still works for tests
