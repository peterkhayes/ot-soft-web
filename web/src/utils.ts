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
 * Examples:
 *   makeOutputFilename('data.txt', 'Output')    → 'dataOutput.txt'
 *   makeOutputFilename('data.old.txt', 'Output') → 'data.oldOutput.txt'
 *   makeOutputFilename('data', 'Output')         → 'dataOutput.txt'
 *   makeOutputFilename(null, 'Output')           → 'Output.txt'
 */
export function makeOutputFilename(inputFilename: string | null, label: string): string {
  if (!inputFilename) {
    return label + '.txt'
  }
  const lastDot = inputFilename.lastIndexOf('.')
  if (lastDot > 0) {
    return inputFilename.substring(0, lastDot) + label + inputFilename.substring(lastDot)
  }
  return inputFilename + label + '.txt'
}
