import { useRef, useState } from 'react'

interface TextFileEditorProps {
  value: string
  onChange: (value: string) => void
  /** If provided, a "Reset to default" button appears whenever value differs from this. */
  defaultValue?: string
  /** Small descriptive text shown above the toolbar. */
  hint?: string
  rows?: number
  /** Textarea placeholder shown when value is empty. */
  placeholder?: string
  /** External error message (e.g. from algorithm validation) shown below the textarea. */
  error?: string | null
  /** data-testid placed on the hidden file input, for test targeting. */
  testId?: string
}

function TextFileEditor({
  value,
  onChange,
  defaultValue,
  hint,
  rows = 6,
  placeholder,
  error,
  testId,
}: TextFileEditorProps) {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [filename, setFilename] = useState<string | null>(null)

  function handleFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    file
      .text()
      .then((text) => {
        onChange(text)
        setFilename(file.name)
      })
      .catch((err) => console.error('Error reading file:', err))
    // Reset so the same file can be re-uploaded
    e.target.value = ''
  }

  // Only show the filename when the content is still what came from the file
  // (i.e., hasn't been overwritten by reset-to-default or manual edit)
  const isAtDefault = defaultValue !== undefined && value === defaultValue
  const buttonLabel = filename && !isAtDefault ? `Loaded: ${filename}` : 'Load from file\u2026'
  const hasChanged = defaultValue !== undefined && value !== defaultValue

  return (
    <div className="text-file-editor">
      {hint && <div className="text-file-editor-hint">{hint}</div>}
      <div className="text-file-editor-toolbar">
        <input
          ref={fileInputRef}
          type="file"
          accept=".txt"
          className="file-input-hidden"
          data-testid={testId}
          onChange={handleFileChange}
        />
        <button
          type="button"
          className="text-file-editor-btn"
          onClick={() => fileInputRef.current?.click()}
        >
          {buttonLabel}
        </button>
        {hasChanged && (
          <button
            type="button"
            className="text-file-editor-btn text-file-editor-btn--muted"
            onClick={() => {
              onChange(defaultValue!)
              setFilename(null)
            }}
          >
            Reset to default
          </button>
        )}
      </div>
      <textarea
        className="schedule-textarea"
        value={value}
        onChange={(e) => {
          onChange(e.target.value)
          setFilename(null)
        }}
        rows={rows}
        placeholder={placeholder}
        spellCheck={false}
      />
      {error && (
        <div className="rcd-status failure" style={{ marginTop: '0.25rem' }}>
          {error}
        </div>
      )}
    </div>
  )
}

export default TextFileEditor
