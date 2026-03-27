import { DEFAULT_SCHEDULE_TEMPLATE } from '../../constants.ts'
import TextFileEditor from '../TextFileEditor.tsx'
import type { GlaParams, SetGlaParams } from './types.ts'

interface GlaOptionsGridProps {
  params: GlaParams
  setParams: SetGlaParams
  scheduleError: string | null
  setScheduleError: (err: string | null) => void
}

function GlaOptionsGrid({
  params,
  setParams,
  scheduleError,
  setScheduleError,
}: GlaOptionsGridProps) {
  const {
    exactProportions,
    useCustomSchedule,
    customSchedule,
    multipleRunsCount,
    maxentMode,
    generateHistory,
    generateFullHistory,
    generateCandidateProbHistory,
  } = params

  return (
    <div className="options-three-col">
      <div className="nhg-options">
        <div className="nhg-options-label">Learning schedule</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={exactProportions}
            onChange={(e) =>
              setParams(
                e.target.checked
                  ? { exactProportions: true, useCustomSchedule: false }
                  : { exactProportions: false },
              )
            }
          />
          Present data in exact proportions
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={useCustomSchedule}
            onChange={(e) =>
              setParams(
                e.target.checked
                  ? { useCustomSchedule: true, exactProportions: false }
                  : { useCustomSchedule: false },
              )
            }
          />
          Use custom learning schedule
        </label>
        {useCustomSchedule && (
          <TextFileEditor
            value={customSchedule}
            onChange={(text) => {
              setParams({ customSchedule: text })
              setScheduleError(null)
            }}
            defaultValue={DEFAULT_SCHEDULE_TEMPLATE}
            hint="Columns: Trials, PlastMark, PlastFaith, NoiseMark, NoiseFaith (tab or space separated)"
            error={scheduleError}
            testId="gla-schedule-file-input"
          />
        )}
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Multiple runs</div>
        <input
          type="number"
          className="param-input"
          aria-label="Number of runs"
          value={multipleRunsCount}
          min={1}
          onChange={(e) =>
            setParams({ multipleRunsCount: Math.max(1, parseInt(e.target.value) || 1) })
          }
        />
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Output options</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={generateHistory}
            onChange={(e) => setParams({ generateHistory: e.target.checked })}
          />
          Generate history of {maxentMode ? 'weights' : 'ranking values'}
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={generateFullHistory}
            onChange={(e) => setParams({ generateFullHistory: e.target.checked })}
          />
          Generate full history (with input/output annotations)
        </label>
        {maxentMode && (
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={generateCandidateProbHistory}
              onChange={(e) => setParams({ generateCandidateProbHistory: e.target.checked })}
            />
            Generate history of candidate probabilities
          </label>
        )}
      </div>
    </div>
  )
}

export default GlaOptionsGrid
