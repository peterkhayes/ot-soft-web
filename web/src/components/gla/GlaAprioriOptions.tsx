import TextFileEditor from '../TextFileEditor.tsx'
import type { GlaParams, SetGlaParams } from './types.ts'

interface GlaAprioriOptionsProps {
  params: GlaParams
  setParams: SetGlaParams
}

function GlaAprioriOptions({ params, setParams }: GlaAprioriOptionsProps) {
  const { showApriori, aprioriText, aprioriGap } = params

  return (
    <div className="nhg-options">
      <div className="nhg-options-label">A priori rankings</div>
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={showApriori}
          onChange={(e) => {
            setParams({ showApriori: e.target.checked })
            if (!e.target.checked) setParams({ aprioriText: '' })
          }}
        />
        Use a priori rankings
      </label>
      {showApriori && (
        <>
          <TextFileEditor
            value={aprioriText}
            onChange={(text) => setParams({ aprioriText: text })}
            hint="Tab-delimited constraint × constraint matrix (abbreviations must match current tableau)."
            placeholder="Load from file or paste content here…"
            testId="gla-apriori-file-input"
          />
          {aprioriText.trim() && (
            <label className="param-label" style={{ marginTop: '0.5rem' }}>
              Constraints ranked a priori must differ by
              <input
                type="number"
                className="param-input"
                value={aprioriGap}
                min={0.001}
                step={1}
                onChange={(e) =>
                  setParams({ aprioriGap: Math.max(0.001, parseFloat(e.target.value) || 20) })
                }
              />
            </label>
          )}
        </>
      )}
    </div>
  )
}

export default GlaAprioriOptions
