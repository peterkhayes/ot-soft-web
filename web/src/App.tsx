import '../style.css'

import { useEffect, useMemo, useRef, useState } from 'react'

import type { Tableau } from '../pkg/ot_soft.js'
import { AxisMode, format_sorted_input_file, parse_tableau } from '../pkg/ot_soft.js'
import FactorialTypologyPanel from './components/FactorialTypologyPanel.tsx'
import type { Framework } from './components/FrameworkPanel.tsx'
import FrameworkPanel from './components/FrameworkPanel.tsx'
import GlaPanel from './components/GlaPanel.tsx'
import InputPanel from './components/InputPanel.tsx'
import MaxEntPanel from './components/MaxEntPanel.tsx'
import NhgPanel from './components/NhgPanel.tsx'
import RcdPanel from './components/RcdPanel.tsx'
import TableauPanel from './components/TableauPanel.tsx'

const NOT_IMPLEMENTED: Record<string, string> = {}

const COLLAPSED_MAX_HEIGHT = 900
const COLLAPSED_MAX_HEIGHT_MOBILE = 1350
const MOBILE_BREAKPOINT = 768

function ExpandableSection({ children }: { children: React.ReactNode }) {
  const [expanded, setExpanded] = useState(false)
  const [overflows, setOverflows] = useState(false)
  const bodyRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const el = bodyRef.current
    if (!el) return
    const ro = new ResizeObserver(() => {
      const maxHeight =
        window.innerWidth <= MOBILE_BREAKPOINT ? COLLAPSED_MAX_HEIGHT_MOBILE : COLLAPSED_MAX_HEIGHT
      setOverflows(el.scrollHeight > maxHeight + 100)
    })
    ro.observe(el)
    return () => ro.disconnect()
  }, [])

  return (
    <div className="expandable-section">
      <div
        ref={bodyRef}
        className={`expandable-body${!expanded ? ' expandable-body--collapsed' : ''}${overflows ? ' expandable-body--overflows' : ''}`}
      >
        {children}
      </div>
      {overflows && (
        <div className="expand-bar">
          <button className="expand-toggle" onClick={() => setExpanded((e) => !e)}>
            {expanded ? (
              <>
                <svg
                  width="14"
                  height="14"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.5"
                >
                  <polyline points="18 15 12 9 6 15" />
                </svg>
                Collapse
              </>
            ) : (
              <>
                <svg
                  width="14"
                  height="14"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.5"
                >
                  <polyline points="6 9 12 15 18 9" />
                </svg>
                Expand
              </>
            )}
          </button>
        </div>
      )}
    </div>
  )
}

function App() {
  const [wasmReady, setWasmReady] = useState(false)
  const [wasmError, setWasmError] = useState<string | null>(null)
  const [currentTableau, setCurrentTableau] = useState<Tableau | null>(null)
  const [currentTableauText, setCurrentTableauText] = useState<string | null>(null)
  const [currentInputFilename, setCurrentInputFilename] = useState<string | null>(null)
  const [parseError, setParseError] = useState<string | null>(null)
  const VALID_FRAMEWORKS = new Set<Framework>(['classical-ot', 'maxent', 'stochastic-ot', 'nhg'])
  const [framework, setFrameworkRaw] = useState<Framework>(() => {
    try {
      const stored = localStorage.getItem('otsoft:framework')
      if (stored && VALID_FRAMEWORKS.has(stored as Framework)) return stored as Framework
    } catch {}
    return 'maxent'
  })
  function setFramework(fw: Framework) {
    try {
      localStorage.setItem('otsoft:framework', fw)
    } catch {}
    setFrameworkRaw(fw)
  }
  const VALID_AXIS_MODES = new Set([
    AxisMode.SwitchAll,
    AxisMode.SwitchWhereNeeded,
    AxisMode.NeverSwitch,
  ])
  const [axisMode, setAxisModeRaw] = useState<AxisMode>(() => {
    try {
      const stored = localStorage.getItem('otsoft:params:tableau-axis')
      if (stored !== null) {
        const num = Number(stored)
        if (VALID_AXIS_MODES.has(num as AxisMode)) return num as AxisMode
      }
    } catch {}
    return AxisMode.SwitchAll
  })
  function setAxisMode(mode: AxisMode) {
    try {
      localStorage.setItem('otsoft:params:tableau-axis', String(mode))
    } catch {}
    setAxisModeRaw(mode)
  }

  const [sortByHarmony, setSortByHarmonyRaw] = useState(() => {
    try {
      const stored = localStorage.getItem('otsoft:params:sort-by-harmony')
      if (stored !== null) return stored === 'true'
    } catch {}
    return true // VB6 default: checked
  })
  function setSortByHarmony(v: boolean) {
    try {
      localStorage.setItem('otsoft:params:sort-by-harmony', String(v))
    } catch {}
    setSortByHarmonyRaw(v)
  }

  // Compute sorted tableau when sort-by-harmony is enabled.
  // Uses RCD to determine constraint ranking order, matching VB6 default behavior.
  const sortedTableau = useMemo(() => {
    if (!sortByHarmony || !currentTableauText) return null
    try {
      const sortedText = format_sorted_input_file(currentTableauText, '', 'rcd')
      return parse_tableau(sortedText)
    } catch {
      return null
    }
  }, [sortByHarmony, currentTableauText])

  // loadCountRef is used as a React key on algorithm panels to force a full remount
  // (resetting their internal state) whenever a new tableau is loaded.
  const loadCountRef = useRef(0)

  useEffect(() => {
    async function initWasm() {
      try {
        const wasm = await import('../pkg/ot_soft.js')
        await wasm.default()
        setWasmReady(true)
        console.log('OTSoft WebAssembly module loaded successfully')
      } catch (err) {
        console.error('Failed to load WASM module:', err)
        setWasmError('Error loading WebAssembly module. Check console for details.')
      }
    }
    initWasm()
  }, [])

  function handleTableauLoaded(tableau: Tableau, text: string, filename: string) {
    loadCountRef.current += 1
    setCurrentTableau(tableau)
    setCurrentTableauText(text)
    setCurrentInputFilename(filename)
    setParseError(null)
  }

  function handleParseError(error: string) {
    setParseError(error)
    setCurrentTableau(null)
    setCurrentTableauText(null)
  }

  function handleReset() {
    setCurrentTableau(null)
    setCurrentTableauText(null)
    setCurrentInputFilename(null)
    setParseError(null)
  }

  if (wasmError) {
    return (
      <>
        <div className="grain-overlay"></div>
        <div className="container">
          <Header />
          <main>
            <div
              className="status-message"
              style={{ background: '#ffe8e8', borderLeftColor: '#e74c3c' }}
            >
              <span>{wasmError}</span>
            </div>
          </main>
          <Footer />
        </div>
      </>
    )
  }

  if (!wasmReady) {
    return (
      <>
        <div className="grain-overlay"></div>
        <div className="container">
          <Header />
          <main>
            <div className="status-message">
              <div className="loading-indicator"></div>
              <span>Initializing analysis environment...</span>
            </div>
          </main>
          <Footer />
        </div>
      </>
    )
  }

  const notImplementedName = NOT_IMPLEMENTED[framework]

  return (
    <>
      <div className="grain-overlay"></div>
      <div className="container">
        <Header />
        <main>
          <div className="content">
            <ExpandableSection>
              <InputPanel
                onTableauLoaded={handleTableauLoaded}
                onParseError={handleParseError}
                onReset={handleReset}
                loadedFilename={currentInputFilename}
              />
            </ExpandableSection>

            {(currentTableau || parseError) && (
              <ExpandableSection>
                <section className="output-panel">
                  <div className="panel-header">
                    <h2>Tableau Analysis</h2>
                    <span className="panel-number">02</span>
                  </div>
                  {!parseError && (
                    <div className="options-two-col">
                      <div className="axis-mode-options">
                        <span className="axis-mode-label">Axis layout</span>
                        {(
                          [
                            [AxisMode.SwitchAll, 'Switch all'],
                            [AxisMode.SwitchWhereNeeded, 'Switch where needed'],
                            [AxisMode.NeverSwitch, 'Never switch'],
                          ] as const
                        ).map(([value, label]) => (
                          <label key={value} className="axis-mode-option">
                            <input
                              type="radio"
                              name="axis-mode"
                              value={value}
                              checked={axisMode === value}
                              onChange={() => setAxisMode(value)}
                            />
                            {label}
                          </label>
                        ))}
                      </div>
                      <div className="axis-mode-options">
                        <span className="axis-mode-label">Display</span>
                        <label className="axis-mode-option">
                          <input
                            type="checkbox"
                            checked={sortByHarmony}
                            onChange={(e) => setSortByHarmony(e.target.checked)}
                          />
                          Sort candidates by harmony
                        </label>
                      </div>
                    </div>
                  )}
                  {parseError ? (
                    <div className="tableau-container">
                      {'Error parsing tableau:\n\n' + parseError}
                    </div>
                  ) : (
                    <TableauPanel tableau={sortedTableau ?? currentTableau!} axisMode={axisMode} />
                  )}
                </section>
              </ExpandableSection>
            )}

            {currentTableau && !parseError && (
              <ExpandableSection>
                <FrameworkPanel framework={framework} onFrameworkChange={setFramework} />
              </ExpandableSection>
            )}

            {currentTableau && !parseError && framework === 'classical-ot' && (
              <ExpandableSection>
                <RcdPanel
                  key={loadCountRef.current}
                  tableau={currentTableau}
                  tableauText={currentTableauText!}
                  inputFilename={currentInputFilename}
                  axisMode={axisMode}
                />
              </ExpandableSection>
            )}

            {currentTableau && !parseError && framework === 'classical-ot' && (
              <ExpandableSection>
                <FactorialTypologyPanel
                  key={loadCountRef.current}
                  tableau={currentTableau}
                  tableauText={currentTableauText!}
                  inputFilename={currentInputFilename}
                />
              </ExpandableSection>
            )}

            {currentTableau && !parseError && framework === 'maxent' && (
              <ExpandableSection>
                <MaxEntPanel
                  key={loadCountRef.current}
                  tableau={currentTableau}
                  tableauText={currentTableauText!}
                  inputFilename={currentInputFilename}
                />
              </ExpandableSection>
            )}

            {currentTableau && !parseError && framework === 'nhg' && (
              <ExpandableSection>
                <NhgPanel
                  key={loadCountRef.current}
                  tableau={currentTableau}
                  tableauText={currentTableauText!}
                  inputFilename={currentInputFilename}
                />
              </ExpandableSection>
            )}

            {currentTableau && !parseError && framework === 'stochastic-ot' && (
              <ExpandableSection>
                <GlaPanel
                  key={loadCountRef.current}
                  tableau={currentTableau}
                  tableauText={currentTableauText!}
                  inputFilename={currentInputFilename}
                />
              </ExpandableSection>
            )}

            {currentTableau && !parseError && notImplementedName && (
              <ExpandableSection>
                <section className="analysis-panel">
                  <div className="panel-header">
                    <h2>Analysis</h2>
                    <span className="panel-number">04</span>
                  </div>
                  <div className="rcd-status failure">
                    {notImplementedName} is not yet implemented.
                  </div>
                </section>
              </ExpandableSection>
            )}
          </div>
        </main>
        <Footer />
      </div>
    </>
  )
}

function Header() {
  return (
    <header className="masthead">
      <div className="header-ornament"></div>
      <h1 className="site-title">OTSoft Web</h1>
      <p className="site-subtitle">Optimality Theory Software</p>
      <div className="header-divider"></div>
    </header>
  )
}

function Footer() {
  return (
    <footer className="site-footer">
      <div className="footer-divider"></div>
      <p>
        OTSoft &middot; Version 2.7 &middot;{' '}
        <a
          href="https://github.com/peterkhayes/ot-soft-web/blob/main/README.md"
          target="_blank"
          rel="noopener noreferrer"
        >
          About
        </a>
      </p>
    </footer>
  )
}

export default App
