import type { Viz } from '@viz-js/viz'
import { useEffect, useRef, useState } from 'react'

import { useBlobDownload } from '../contexts/blobDownloadContext.ts'
import DownloadMenu from './DownloadMenu.tsx'

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
    <div className="hasse-diagram" data-testid="hasse-diagram">
      <div className="hasse-header">
        <h3 className="hasse-title">Hasse Diagram</h3>
        {svgEl && (
          <DownloadMenu
            items={[
              { label: 'Download SVG', onClick: handleDownloadSvg },
              { label: 'Download PNG', onClick: handleDownloadPng },
            ]}
          />
        )}
      </div>
      {loading && <div className="hasse-loading">Rendering diagram…</div>}
      {error && <div className="hasse-error">Error rendering diagram: {error}</div>}
      <div ref={containerRef} className="hasse-svg-container" />
    </div>
  )
}

export default HasseDiagram
