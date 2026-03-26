import TextFileEditor from '../TextFileEditor.tsx'
import type { FtParams, SetFtParams } from './types.ts'

interface FtOptionsProps {
  params: FtParams
  setParams: SetFtParams
  aprioriText: string
  setAprioriText: (text: string) => void
  showApriori: boolean
  setShowApriori: (show: boolean) => void
}

function FtOptions({
  params,
  setParams,
  aprioriText,
  setAprioriText,
  showApriori,
  setShowApriori,
}: FtOptionsProps) {
  const { includeFullListing, includeFtsum, includeCompactSum } = params

  return (
    <>
      <div className="nhg-options">
        <div className="nhg-options-label">Output options</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeFullListing}
            onChange={(e) => setParams({ includeFullListing: e.target.checked })}
          />
          Include rankings in results
        </label>

        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeFtsum}
            onChange={(e) => setParams({ includeFtsum: e.target.checked })}
          />
          Generate FTSum file
        </label>

        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={includeCompactSum}
            onChange={(e) => setParams({ includeCompactSum: e.target.checked })}
          />
          Generate CompactSum file
        </label>
      </div>

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
            testId="ft-apriori-file-input"
          />
        )}
      </div>
    </>
  )
}

export default FtOptions
