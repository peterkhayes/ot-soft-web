export type Framework = 'classical-ot' | 'maxent' | 'stochastic-ot' | 'nhg'

interface FrameworkPanelProps {
  framework: Framework
  onFrameworkChange: (fw: Framework) => void
}

const FRAMEWORKS: { value: Framework; label: string; description: string }[] = [
  {
    value: 'classical-ot',
    label: 'Classical OT',
    description: 'Categorical ranking — RCD, BCD, LFCD',
  },
  {
    value: 'maxent',
    label: 'Maximum Entropy',
    description: 'Probabilistic — weighted constraint violations',
  },
  {
    value: 'stochastic-ot',
    label: 'Stochastic OT',
    description: 'Probabilistic — Gradual Learning Algorithm',
  },
  {
    value: 'nhg',
    label: 'Noisy Harmonic Grammar',
    description: 'Probabilistic — noisy weighted constraints',
  },
]

function FrameworkPanel({ framework, onFrameworkChange }: FrameworkPanelProps) {
  return (
    <section className="analysis-panel">
      <div className="panel-header">
        <h2>Framework</h2>
        <span className="panel-number">03</span>
      </div>

      <div className="framework-options">
        {FRAMEWORKS.map(({ value, label, description }) => (
          <label
            key={value}
            className={`framework-option ${framework === value ? 'selected' : ''}`}
          >
            <input
              type="radio"
              name="framework"
              value={value}
              checked={framework === value}
              onChange={() => onFrameworkChange(value)}
            />
            <span className="framework-label">{label}</span>
            <span className="framework-description">{description}</span>
          </label>
        ))}
      </div>
    </section>
  )
}

export default FrameworkPanel
