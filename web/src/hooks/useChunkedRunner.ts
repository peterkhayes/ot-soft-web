import { useCallback, useReducer, useRef } from 'react'

const COMPLETION_SOUND_THRESHOLD_MS = 3_000

/** Play a short two-tone chime via the Web Audio API. */
function playCompletionSound(): void {
  try {
    const ctx = new AudioContext()
    const notes = [660, 880] // E5 → A5 ascending
    notes.forEach((freq, i) => {
      const osc = ctx.createOscillator()
      const gain = ctx.createGain()
      osc.type = 'sine'
      osc.frequency.value = freq
      gain.gain.setValueAtTime(0.15, ctx.currentTime)
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + i * 0.15 + 0.2)
      osc.connect(gain)
      gain.connect(ctx.destination)
      osc.start(ctx.currentTime + i * 0.15)
      osc.stop(ctx.currentTime + i * 0.15 + 0.2)
    })
  } catch {
    // AudioContext may not be available in all environments
  }
}

/** Mirrors the Rust ChunkedRunner trait. */
export interface ChunkedRunner {
  run_chunk(max_work: number): boolean
  progress(): Float64Array
  free(): void
}

export type RunnerState<T> =
  | { status: 'idle' }
  | { status: 'running'; completed: number; total: number }
  | { status: 'done'; result: T }
  | { status: 'error'; error: string }

type Action<T> =
  | { type: 'start' }
  | { type: 'progress'; completed: number; total: number }
  | { type: 'done'; result: T }
  | { type: 'error'; error: string }
  | { type: 'reset' }

function reducer<T>(state: RunnerState<T>, action: Action<T>): RunnerState<T> {
  switch (action.type) {
    case 'start':
      return { status: 'running', completed: 0, total: 1 }
    case 'progress':
      if (state.status !== 'running') return state
      return { status: 'running', completed: action.completed, total: action.total }
    case 'done':
      return { status: 'done', result: action.result }
    case 'error':
      return { status: 'error', error: action.error }
    case 'reset':
      return { status: 'idle' }
  }
}

/**
 * Generic hook for running a chunked WASM computation with progress reporting.
 *
 * Breaks computation into chunks via `requestAnimationFrame`, yielding to
 * the browser between chunks so the UI stays responsive.
 *
 * @param createRunner - Factory that creates and returns the WASM runner.
 * @param extractResult - Extracts the final result from the completed runner.
 * @param chunkSize - Number of work units per chunk (default 50,000).
 */
export function useChunkedRunner<R extends ChunkedRunner, T>(
  createRunner: () => R,
  extractResult: (runner: R) => T,
  chunkSize = 50_000,
): { state: RunnerState<T>; run: () => void; reset: () => void } {
  const [state, dispatch] = useReducer(reducer<T>, { status: 'idle' })
  const runnerRef = useRef<R | null>(null)
  const rafRef = useRef<number>(0)

  const run = useCallback(() => {
    // Cancel any in-progress run
    if (rafRef.current) cancelAnimationFrame(rafRef.current)
    if (runnerRef.current) {
      runnerRef.current.free()
      runnerRef.current = null
    }

    let runner: R
    try {
      runner = createRunner()
    } catch (e) {
      dispatch({ type: 'error', error: String(e) })
      return
    }
    runnerRef.current = runner
    dispatch({ type: 'start' })
    const startTime = performance.now()
    let lastProgressTime = 0

    function step() {
      try {
        const done = runner.run_chunk(chunkSize)
        if (!done) {
          const now = performance.now()
          if (now - lastProgressTime > 100) {
            const p = runner.progress()
            dispatch({ type: 'progress', completed: p[0], total: p[1] })
            lastProgressTime = now
          }
          rafRef.current = requestAnimationFrame(step)
        } else {
          const result = extractResult(runner)
          runner.free()
          runnerRef.current = null
          if (performance.now() - startTime > COMPLETION_SOUND_THRESHOLD_MS) {
            playCompletionSound()
          }
          dispatch({ type: 'done', result })
        }
      } catch (e) {
        runner.free()
        runnerRef.current = null
        dispatch({ type: 'error', error: String(e) })
      }
    }

    rafRef.current = requestAnimationFrame(step)
  }, [createRunner, extractResult, chunkSize])

  const reset = useCallback(() => {
    if (rafRef.current) cancelAnimationFrame(rafRef.current)
    if (runnerRef.current) {
      runnerRef.current.free()
      runnerRef.current = null
    }
    dispatch({ type: 'reset' })
  }, [])

  return { state, run, reset }
}
