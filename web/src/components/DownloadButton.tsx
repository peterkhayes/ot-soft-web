interface DownloadButtonProps {
  onClick: () => void
  children: React.ReactNode
}

/** Secondary button with a download icon. */
function DownloadButton({ onClick, children }: DownloadButtonProps) {
  return (
    <button className="download-button" onClick={onClick}>
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
      {children}
    </button>
  )
}

export default DownloadButton
