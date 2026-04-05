import { useCallback, useEffect, useRef, useState } from 'react'

export interface DownloadMenuItem {
  label: string
  onClick: () => void
}

interface DownloadMenuProps {
  items: DownloadMenuItem[]
}

/**
 * Accessible menu button (WAI-ARIA menu button pattern) that groups
 * multiple download actions behind a single "Download" trigger.
 *
 * - Arrow keys navigate menu items
 * - Escape / click-outside closes the menu
 * - Each item has role="menuitem"
 */
function DownloadMenu({ items }: DownloadMenuProps) {
  const [open, setOpen] = useState(false)
  const buttonRef = useRef<HTMLButtonElement>(null)
  const menuRef = useRef<HTMLDivElement>(null)
  const itemRefs = useRef<(HTMLButtonElement | null)[]>([])

  // Close on click outside
  useEffect(() => {
    if (!open) return
    function handleClick(e: MouseEvent) {
      if (
        menuRef.current &&
        !menuRef.current.contains(e.target as Node) &&
        buttonRef.current &&
        !buttonRef.current.contains(e.target as Node)
      ) {
        setOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClick)
    return () => document.removeEventListener('mousedown', handleClick)
  }, [open])

  // Focus first item when menu opens
  useEffect(() => {
    if (open) {
      itemRefs.current[0]?.focus()
    }
  }, [open])

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (!open) return

      const focusedIdx = itemRefs.current.findIndex((el) => el === document.activeElement)

      switch (e.key) {
        case 'ArrowDown': {
          e.preventDefault()
          const next = focusedIdx < items.length - 1 ? focusedIdx + 1 : 0
          itemRefs.current[next]?.focus()
          break
        }
        case 'ArrowUp': {
          e.preventDefault()
          const prev = focusedIdx > 0 ? focusedIdx - 1 : items.length - 1
          itemRefs.current[prev]?.focus()
          break
        }
        case 'Escape':
          e.preventDefault()
          setOpen(false)
          buttonRef.current?.focus()
          break
        case 'Tab':
          setOpen(false)
          break
        case 'Home': {
          e.preventDefault()
          itemRefs.current[0]?.focus()
          break
        }
        case 'End': {
          e.preventDefault()
          itemRefs.current[items.length - 1]?.focus()
          break
        }
      }
    },
    [open, items.length],
  )

  function handleItemClick(item: DownloadMenuItem) {
    item.onClick()
    setOpen(false)
    buttonRef.current?.focus()
  }

  return (
    <div className="download-menu" onKeyDown={handleKeyDown}>
      <button
        ref={buttonRef}
        className="download-button"
        aria-haspopup="true"
        aria-expanded={open}
        onClick={() => setOpen((prev) => !prev)}
      >
        <svg
          className="button-icon"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          aria-hidden="true"
        >
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
          <polyline points="7 10 12 15 17 10" />
          <line x1="12" y1="15" x2="12" y2="3" />
        </svg>
        Download
        <svg
          className="button-icon download-menu-chevron"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          aria-hidden="true"
        >
          <polyline points="6 9 12 15 18 9" />
        </svg>
      </button>
      {open && (
        <div ref={menuRef} role="menu" className="download-menu-popup">
          {items.map((item, i) => (
            <button
              key={item.label}
              ref={(el) => {
                itemRefs.current[i] = el
              }}
              role="menuitem"
              tabIndex={-1}
              className="download-menu-item"
              onClick={() => handleItemClick(item)}
            >
              {item.label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

export default DownloadMenu
