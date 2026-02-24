import { expect } from 'vitest'
import { render } from 'vitest-browser-react'
import { page } from '@vitest/browser/context'
import { DownloadProvider } from '../src/downloadContext'
import App from '../src/App'

export interface CapturedDownload {
  content: string
  filename: string
}

export function renderApp() {
  localStorage.clear()
  const downloads: CapturedDownload[] = []
  const onDownload = (content: string, filename: string) => downloads.push({ content, filename })
  render(
    <DownloadProvider value={onDownload}>
      <App />
    </DownloadProvider>
  )
  return downloads
}

export async function loadExample() {
  await expect.element(page.getByText('Load Example Tableau')).toBeVisible()
  await page.getByText('Load Example Tableau').click()
  await expect.element(page.getByText('Tableau Analysis')).toBeVisible()
}

export async function loadFile(filePath: string) {
  await expect.element(page.getByTestId('file-input')).toBeInTheDocument()
  await page.getByTestId('file-input').upload(filePath)
  await expect.element(page.getByText('Tableau Analysis')).toBeVisible()
}

/** Strip the date/time stamp from formatter output so snapshots are stable. */
export function normalizeOutput(content: string): string {
  return content.replace(/\d+-\d+-\d+, \d+:\d+ (am|pm)/gi, '<TIMESTAMP>')
}
