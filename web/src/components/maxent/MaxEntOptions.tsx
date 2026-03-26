import type { MaxEntParams, SetMaxEntParams } from './types.ts'

interface MaxEntOptionsProps {
  params: MaxEntParams
  setParams: SetMaxEntParams
}

function MaxEntOptions({ params, setParams }: MaxEntOptionsProps) {
  const { usePrior, sigmaSquared, sortByWeight, generateHistory, generateOutputProbHistory } =
    params

  return (
    <>
      <div className="nhg-options">
        <div className="nhg-options-label">Prior</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={usePrior}
            onChange={(e) => setParams({ usePrior: e.target.checked })}
          />
          Gaussian prior (L2 regularization)
        </label>
        {usePrior && (
          <label className="param-label" style={{ marginLeft: '1.5rem' }}>
            &sigma;&sup2;
            <input
              type="number"
              className="param-input"
              value={sigmaSquared}
              min={0.0001}
              step={0.1}
              onChange={(e) =>
                setParams({ sigmaSquared: Math.max(0.0001, parseFloat(e.target.value) || 1) })
              }
            />
          </label>
        )}
      </div>

      <div className="nhg-options">
        <div className="nhg-options-label">Output options</div>
        <label className="nhg-checkbox">
          <input
            type="checkbox"
            checked={sortByWeight}
            onChange={(e) => setParams({ sortByWeight: e.target.checked })}
          />
          Sort constraints by weight
        </label>
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
            checked={generateOutputProbHistory}
            onChange={(e) => setParams({ generateOutputProbHistory: e.target.checked })}
          />
          Generate history of output probabilities
        </label>
      </div>
    </>
  )
}

export default MaxEntOptions
