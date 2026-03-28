import TextFileEditor from '../TextFileEditor.tsx'
import type { Algorithm, RcdParams, SetRcdParams } from './types.ts'

interface RcdOptionsProps {
  params: RcdParams
  setParams: SetRcdParams
  aprioriText: string
  setAprioriText: (text: string) => void
  showApriori: boolean
  setShowApriori: (show: boolean) => void
}

function RcdOptions({
  params,
  setParams,
  aprioriText,
  setAprioriText,
  showApriori,
  setShowApriori,
}: RcdOptionsProps) {
  const { algorithm, includeFred, useMib, showDetails, includeMiniTableaux, diagnostics } = params
  const supportsApriori = algorithm === 'rcd' || algorithm === 'lfcd'

  return (
    <>
      <div className="nhg-options">
        <div className="nhg-options-label">Algorithm</div>
        <select
          className="algorithm-select"
          value={algorithm}
          onChange={(e) => setParams({ algorithm: e.target.value as Algorithm })}
          aria-label="Algorithm"
        >
          <option value="rcd">RCD</option>
          <option value="bcd">BCD</option>
          <option value="bcd-specific">BCD (Specific)</option>
          <option value="lfcd">LFCD</option>
        </select>
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Ranking argumentation</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeFred}
            onChange={(e) => setParams({ includeFred: e.target.checked })}
          />
          Include ranking arguments
        </label>
        {includeFred && (
          <>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={useMib}
                onChange={(e) => setParams({ useMib: e.target.checked })}
              />
              Use Most Informative Basis
            </label>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={showDetails}
                onChange={(e) => setParams({ showDetails: e.target.checked })}
              />
              Show details of argumentation
            </label>
            <label className="nhg-checkbox nhg-checkbox-indent">
              <input
                type="checkbox"
                checked={includeMiniTableaux}
                onChange={(e) => setParams({ includeMiniTableaux: e.target.checked })}
              />
              Include illustrative mini-tableaux
            </label>
          </>
        )}
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Diagnostics</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={diagnostics}
            onChange={(e) => setParams({ diagnostics: e.target.checked })}
          />
          Diagnostics if ranking fails
        </label>
      </div>

      {supportsApriori && (
        <div className="nhg-options">
          <div className="nhg-options-label">A priori rankings</div>
          <label className="nhg-checkbox">
            <input
              type="checkbox"
              checked={showApriori}
              onChange={(e) => {
                setShowApriori(e.target.checked)
                if (!e.target.checked) setAprioriText('')
              }}
            />
            Use a priori rankings
          </label>
          {showApriori && (
            <TextFileEditor
              value={aprioriText}
              onChange={setAprioriText}
              hint="Tab-delimited constraint x constraint matrix (abbreviations must match current tableau)."
              placeholder="Load from file or paste content here..."
              testId="rcd-apriori-file-input"
            />
          )}
        </div>
      )}
    </>
  )
}

export default RcdOptions
