interface RunButtonProps {
  isLoading: boolean
  onClick: () => void
  label: string
  disabled?: boolean
}

/** Primary run button with a play icon at rest and an hourglass while loading. */
function RunButton({ isLoading, onClick, label, disabled }: RunButtonProps) {
  return (
    <button
      className={`primary-button${isLoading ? ' primary-button--loading' : ''}`}
      onClick={onClick}
      disabled={disabled ?? isLoading}
    >
      {isLoading ? (
        <svg
          className="button-icon"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          <path d="M5 22h14" />
          <path d="M5 2h14" />
          <path d="M17 22v-4.172a2 2 0 0 0-.586-1.414L12 12l-4.414 4.414A2 2 0 0 0 7 17.828V22" />
          <path d="M7 2v4.172a2 2 0 0 0 .586 1.414L12 12l4.414-4.414A2 2 0 0 0 17 6.172V2" />
        </svg>
      ) : (
        <svg
          className="button-icon"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
        >
          <polygon points="5 3 19 12 5 21 5 3" />
        </svg>
      )}
      {label}
    </button>
  )
}

export default RunButton
