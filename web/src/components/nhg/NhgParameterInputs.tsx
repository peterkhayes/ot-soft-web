import type { NhgParams, SetNhgParams } from './types.ts'

interface NhgParameterInputsProps {
  params: NhgParams
  setParams: SetNhgParams
}

function NhgParameterInputs({ params, setParams }: NhgParameterInputsProps) {
  const { cycles, initialPlasticity, finalPlasticity, testTrials } = params

  return (
    <div className="maxent-params">
      <label className="param-label">
        Cycles
        <input
          type="number"
          className="param-input"
          value={cycles}
          min={1}
          max={10000000}
          onChange={(e) => setParams({ cycles: Math.max(1, parseInt(e.target.value) || 1) })}
        />
      </label>
      <label className="param-label">
        Initial plasticity
        <input
          type="number"
          className="param-input"
          value={initialPlasticity}
          min={0.0001}
          step={0.1}
          onChange={(e) =>
            setParams({ initialPlasticity: Math.max(0.0001, parseFloat(e.target.value) || 2) })
          }
        />
      </label>
      <label className="param-label">
        Final plasticity
        <input
          type="number"
          className="param-input"
          value={finalPlasticity}
          min={0.000001}
          step={0.001}
          onChange={(e) =>
            setParams({
              finalPlasticity: Math.max(0.000001, parseFloat(e.target.value) || 0.002),
            })
          }
        />
      </label>
      <label className="param-label">
        Test trials
        <input
          type="number"
          className="param-input"
          value={testTrials}
          min={1}
          max={100000}
          onChange={(e) => setParams({ testTrials: Math.max(1, parseInt(e.target.value) || 2000) })}
        />
      </label>
    </div>
  )
}

export default NhgParameterInputs
