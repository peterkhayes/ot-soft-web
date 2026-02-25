/**
 * Returns true if every key in `params` matches the corresponding value in `defaults`.
 */
export function isAtDefaults<T extends object>(params: T, defaults: T): boolean {
  return (Object.keys(defaults) as (keyof T)[]).every((k) => params[k] === defaults[k])
}

/**
 * Trigger a browser download of text content as a file.
 */
export function downloadTextFile(content: string, filename: string): void {
  const blob = new Blob([content], { type: 'text/plain;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
}

/**
 * Build an output filename by inserting a label before the extension.
 *
 * An optional `ext` parameter overrides the output extension (including the dot).
 *
 * Examples:
 *   makeOutputFilename('data.txt', 'Output')           → 'dataOutput.txt'
 *   makeOutputFilename('data.old.txt', 'Output')       → 'data.oldOutput.txt'
 *   makeOutputFilename('data', 'Output')               → 'dataOutput.txt'
 *   makeOutputFilename(null, 'Output')                 → 'Output.txt'
 *   makeOutputFilename('data.txt', 'Output', '.html')  → 'dataOutput.html'
 *   makeOutputFilename(null, 'Output', '.html')        → 'Output.html'
 */
export function makeOutputFilename(
  inputFilename: string | null,
  label: string,
  ext?: string,
): string {
  if (!inputFilename) {
    return label + (ext ?? '.txt')
  }
  const lastDot = inputFilename.lastIndexOf('.')
  if (lastDot > 0) {
    const base = inputFilename.substring(0, lastDot)
    const origExt = inputFilename.substring(lastDot)
    return base + label + (ext ?? origExt)
  }
  return inputFilename + label + (ext ?? '.txt')
}
