import { useState, useEffect, useRef } from 'react'
import type { Tableau } from '../pkg/ot_soft.js'
import InputPanel from './components/InputPanel.tsx'
import TableauPanel from './components/TableauPanel.tsx'
import FrameworkPanel from './components/FrameworkPanel.tsx'
import type { Framework } from './components/FrameworkPanel.tsx'
import RcdPanel from './components/RcdPanel.tsx'

const NOT_IMPLEMENTED: Record<string, string> = {
  'maxent': 'Maximum Entropy',
  'stochastic-ot': 'Stochastic OT',
  'nhg': 'Noisy Harmonic Grammar',
}

function App() {
  const [wasmReady, setWasmReady] = useState(false)
  const [wasmError, setWasmError] = useState<string | null>(null)
  const [currentTableau, setCurrentTableau] = useState<Tableau | null>(null)
  const [currentTableauText, setCurrentTableauText] = useState<string | null>(null)
  const [currentInputFilename, setCurrentInputFilename] = useState<string | null>(null)
  const [parseError, setParseError] = useState<string | null>(null)
  const [framework, setFramework] = useState<Framework>('classical-ot')
  const loadCountRef = useRef(0)

  useEffect(() => {
    async function initWasm() {
      try {
        const wasm = await import('../pkg/ot_soft.js')
        await wasm.default()
        setWasmReady(true)
        console.log('OT-Soft WebAssembly module loaded successfully')
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

  if (wasmError) {
    return (
      <>
        <div className="grain-overlay"></div>
        <div className="container">
          <Header />
          <main>
            <div className="status-message" style={{ background: '#ffe8e8', borderLeftColor: '#e74c3c' }}>
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
            <InputPanel
              onTableauLoaded={handleTableauLoaded}
              onParseError={handleParseError}
            />

            {(currentTableau || parseError) && (
              <section className="output-panel">
                <div className="panel-header">
                  <h2>Tableau Analysis</h2>
                  <span className="panel-number">02</span>
                </div>
                {parseError ? (
                  <div className="tableau-container">
                    {'Error parsing tableau:\n\n' + parseError}
                  </div>
                ) : (
                  <TableauPanel tableau={currentTableau!} />
                )}
              </section>
            )}

            {currentTableau && !parseError && (
              <FrameworkPanel
                framework={framework}
                onFrameworkChange={setFramework}
              />
            )}

            {currentTableau && !parseError && framework === 'classical-ot' && (
              <RcdPanel
                key={loadCountRef.current}
                tableau={currentTableau}
                tableauText={currentTableauText!}
                inputFilename={currentInputFilename}
              />
            )}

            {currentTableau && !parseError && notImplementedName && (
              <section className="analysis-panel">
                <div className="panel-header">
                  <h2>Analysis</h2>
                  <span className="panel-number">04</span>
                </div>
                <div className="rcd-status failure">
                  {notImplementedName} is not yet implemented.
                </div>
              </section>
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
      <h1 className="site-title">OT-Soft</h1>
      <p className="site-subtitle">Optimality Theory Analysis</p>
      <div className="header-divider"></div>
    </header>
  )
}

function Footer() {
  return (
    <footer className="site-footer">
      <div className="footer-divider"></div>
      <p>OT-Soft &middot; Version 2.7</p>
    </footer>
  )
}

export default App
