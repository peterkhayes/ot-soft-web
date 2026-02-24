import type { Viz } from '@viz-js/viz'
import { useEffect, useRef, useState } from 'react'

import { useBlobDownload } from '../contexts/blobDownloadContext.ts'

interface HasseDiagramProps {
  dotString: string
  downloadName?: string
}

let vizInstance: Viz | null = null

async function getViz(): Promise<Viz> {
  if (vizInstance) return vizInstance
  const { instance } = await import('@viz-js/viz')
  vizInstance = await instance()
  return vizInstance
}

function HasseDiagram({ dotString, downloadName = 'HasseDiagram' }: HasseDiagramProps) {
  const [svgEl, setSvgEl] = useState<SVGSVGElement | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const containerRef = useRef<HTMLDivElement>(null)
  const blobDownload = useBlobDownload()

  useEffect(() => {
    let cancelled = false
    setLoading(true)
    setError(null)
    setSvgEl(null)

    getViz()
      .then((viz) => {
        if (cancelled) return
        try {
          const el = viz.renderSVGElement(dotString)
          setSvgEl(el)
        } catch (err) {
          setError(String(err))
        } finally {
          setLoading(false)
        }
      })
      .catch((err) => {
        if (!cancelled) {
          setError(String(err))
          setLoading(false)
        }
      })

    return () => {
      cancelled = true
    }
  }, [dotString])

  useEffect(() => {
    if (!containerRef.current) return
    containerRef.current.innerHTML = ''
    if (svgEl) containerRef.current.appendChild(svgEl)
  }, [svgEl])

  function handleDownloadSvg() {
    const svg = containerRef.current?.querySelector('svg')
    if (!svg) return
    const svgStr = new XMLSerializer().serializeToString(svg)
    const blob = new Blob([svgStr], { type: 'image/svg+xml' })
    blobDownload(blob, `${downloadName}.svg`)
  }

  function handleDownloadPng() {
    const svg = containerRef.current?.querySelector('svg')
    if (!svg) return
    const svgStr = new XMLSerializer().serializeToString(svg)
    const vb = svg.viewBox.baseVal
    const width = (vb.width || svg.width.baseVal.value || 400) * 2
    const height = (vb.height || svg.height.baseVal.value || 300) * 2
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')!
    const img = new Image()
    const svgBlob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' })
    const svgUrl = URL.createObjectURL(svgBlob)
    img.onload = () => {
      ctx.fillStyle = 'white'
      ctx.fillRect(0, 0, width, height)
      ctx.drawImage(img, 0, 0, width, height)
      URL.revokeObjectURL(svgUrl)
      canvas.toBlob((blob) => {
        if (blob) blobDownload(blob, `${downloadName}.png`)
      }, 'image/png')
    }
    img.src = svgUrl
  }

  return (
    <div className="hasse-diagram">
      <div className="hasse-header">
        <h3 className="hasse-title">Hasse Diagram</h3>
        <div className="hasse-actions">
          <button className="hasse-export-button" onClick={handleDownloadSvg} disabled={!svgEl}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            SVG
          </button>
          <button className="hasse-export-button" onClick={handleDownloadPng} disabled={!svgEl}>
            <svg
              className="button-icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            PNG
          </button>
        </div>
      </div>
      {loading && <div className="hasse-loading">Rendering diagram…</div>}
      {error && <div className="hasse-error">Error rendering diagram: {error}</div>}
      <div ref={containerRef} className="hasse-svg-container" />
    </div>
  )
}

export default HasseDiagram
