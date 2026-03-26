import { DEFAULT_SCHEDULE_TEMPLATE } from '../../constants.ts'
import TextFileEditor from '../TextFileEditor.tsx'
import type { NhgParams, SetNhgParams } from './types.ts'

interface NhgScheduleOptionsProps {
  params: NhgParams
  setParams: SetNhgParams
  scheduleError: string | null
  setScheduleError: (err: string | null) => void
}

function NhgScheduleOptions({
  params,
  setParams,
  scheduleError,
  setScheduleError,
}: NhgScheduleOptionsProps) {
  const {
    exactProportions,
    useCustomSchedule,
    customSchedule,
    generateHistory,
    generateFullHistory,
  } = params

  return (
    <div className="options-two-col">
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
            testId="nhg-schedule-file-input"
          />
        )}
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Output options</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={generateHistory}
            onChange={(e) => setParams({ generateHistory: e.target.checked })}
          />
          Generate history of weights
        </label>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={generateFullHistory}
            onChange={(e) => setParams({ generateFullHistory: e.target.checked })}
          />
          Generate full history (with input/output annotations)
        </label>
      </div>
    </div>
  )
}

export default NhgScheduleOptions
