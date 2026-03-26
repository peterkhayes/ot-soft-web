import type { NhgParams, SetNhgParams } from './types.ts'

interface NhgNoiseOptionsProps {
  params: NhgParams
  setParams: SetNhgParams
}

function NhgNoiseOptions({ params, setParams }: NhgNoiseOptionsProps) {
  const {
    noiseByCell,
    postMultNoise,
    noiseForZeroCells,
    lateNoise,
    exponentialNhg,
    demiGaussians,
    negativeWeightsOk,
    resolveTiesBySkipping,
  } = params

  return (
    <div className="nhg-options">
      <div className="nhg-options-label">Noise variant options</div>
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={noiseByCell}
          onChange={(e) => setParams({ noiseByCell: e.target.checked })}
        />
        Apply noise by tableau cell, not by constraint
      </label>
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={postMultNoise}
          onChange={(e) => {
            setParams(
              e.target.checked
                ? { postMultNoise: true }
                : { postMultNoise: false, noiseForZeroCells: false },
            )
          }}
        />
        Apply noise after multiplication of weights by violations
      </label>
      {postMultNoise && (
        <label className="nhg-checkbox nhg-checkbox-indent">
          <input
            type="checkbox"
            checked={noiseForZeroCells}
            onChange={(e) => setParams({ noiseForZeroCells: e.target.checked })}
          />
          Include noise even in cells with no violation
        </label>
      )}
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={lateNoise}
          onChange={(e) => setParams({ lateNoise: e.target.checked })}
        />
        Add noise to candidates, after harmony calculation
      </label>
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={exponentialNhg}
          onChange={(e) => setParams({ exponentialNhg: e.target.checked })}
        />
        Employ Exponential NHG
      </label>
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={demiGaussians}
          onChange={(e) => setParams({ demiGaussians: e.target.checked })}
        />
        Use positive demi-Gaussians
      </label>
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={negativeWeightsOk}
          onChange={(e) => setParams({ negativeWeightsOk: e.target.checked })}
        />
        Allow constraint weights to go negative
      </label>
      <label className="nhg-checkbox">
        <input
          type="checkbox"
          checked={resolveTiesBySkipping}
          onChange={(e) => setParams({ resolveTiesBySkipping: e.target.checked })}
        />
        Resolve ties by skipping trial
      </label>
    </div>
  )
}

export default NhgNoiseOptions
