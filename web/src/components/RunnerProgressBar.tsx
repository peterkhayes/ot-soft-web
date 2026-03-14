import type { RunnerState } from '../hooks/useChunkedRunner.ts'

interface RunnerProgressBarProps {
  state: RunnerState<unknown>
}

/**
 * Progress bar shown while a chunked runner is active.
 * Renders nothing when the runner is idle, done, or has zero total work.
 */
function RunnerProgressBar({ state }: RunnerProgressBarProps) {
  if (state.status !== 'running' || state.total <= 0) return null
  const pct = Math.round((state.completed / state.total) * 100)
  return (
    <div className="progress-bar-container">
      <div className="progress-bar-fill" style={{ width: `${pct}%` }} />
      <span className="progress-bar-label">{pct}%</span>
    </div>
  )
}

export default RunnerProgressBar
