import type { MaxEntParams, SetMaxEntParams } from './types.ts'

interface MaxEntParameterInputsProps {
  params: MaxEntParams
  setParams: SetMaxEntParams
}

function MaxEntParameterInputs({ params, setParams }: MaxEntParameterInputsProps) {
  const { iterations, weightMin, weightMax } = params

  return (
    <div className="maxent-params">
      <label className="param-label">
        Iterations
        <input
          type="number"
          className="param-input"
          value={iterations}
          min={1}
          max={100000}
          onChange={(e) => setParams({ iterations: Math.max(1, parseInt(e.target.value) || 1) })}
        />
      </label>
      <label className="param-label">
        Weight min
        <input
          type="number"
          className="param-input"
          value={weightMin}
          step={0.1}
          onChange={(e) => setParams({ weightMin: parseFloat(e.target.value) || 0 })}
        />
      </label>
      <label className="param-label">
        Weight max
        <input
          type="number"
          className="param-input"
          value={weightMax}
          min={0.1}
          step={1}
          onChange={(e) =>
            setParams({ weightMax: Math.max(0.1, parseFloat(e.target.value) || 50) })
          }
        />
      </label>
    </div>
  )
}

export default MaxEntParameterInputs
