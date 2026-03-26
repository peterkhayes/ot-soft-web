import type { GlaParams, SetGlaParams } from './types.ts'

interface GlaFrameworkOptionsProps {
  params: GlaParams
  setParams: SetGlaParams
}

function GlaFrameworkOptions({ params, setParams }: GlaFrameworkOptionsProps) {
  const { maxentMode, magriUpdateRule, negativeWeightsOk, gaussianPrior, sigma } = params

  return (
    <div className="nhg-options">
      <div className="nhg-options-label">Framework</div>
      <label className="nhg-checkbox">
        <input
          type="radio"
          name="gla-mode"
          checked={!maxentMode}
          onChange={() => setParams({ maxentMode: false })}
        />
        Stochastic OT (ranking values)
      </label>
      {!maxentMode && (
        <label className="nhg-checkbox nhg-checkbox-indent">
          <input
            type="checkbox"
            checked={magriUpdateRule}
            onChange={(e) => setParams({ magriUpdateRule: e.target.checked })}
          />
          Use the Magri update rule
        </label>
      )}
      <label className="nhg-checkbox">
        <input
          type="radio"
          name="gla-mode"
          checked={maxentMode}
          onChange={() => setParams({ maxentMode: true })}
        />
        Online MaxEnt (weights)
      </label>
      {maxentMode && (
        <>
          <label className="nhg-checkbox nhg-checkbox-indent">
            <input
              type="checkbox"
              checked={negativeWeightsOk}
              onChange={(e) => setParams({ negativeWeightsOk: e.target.checked })}
            />
            Allow constraint weights to go negative
          </label>
          <label className="nhg-checkbox nhg-checkbox-indent">
            <input
              type="checkbox"
              checked={gaussianPrior}
              onChange={(e) => setParams({ gaussianPrior: e.target.checked })}
            />
            Gaussian prior (L2 regularization)
          </label>
          {gaussianPrior && (
            <label className="param-label" style={{ marginLeft: '3rem' }}>
              σ
              <input
                type="number"
                className="param-input"
                value={sigma}
                min={0.0001}
                step={0.1}
                onChange={(e) =>
                  setParams({ sigma: Math.max(0.0001, parseFloat(e.target.value) || 1) })
                }
              />
            </label>
          )}
        </>
      )}
    </div>
  )
}

export default GlaFrameworkOptions
