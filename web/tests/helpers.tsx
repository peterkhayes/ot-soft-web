import { page } from '@vitest/browser/context'
import { expect } from 'vitest'
import { render } from 'vitest-browser-react'

import App from '../src/App'
import { BlobDownloadProvider } from '../src/blobDownloadContext'
import { DownloadProvider } from '../src/downloadContext'

export interface CapturedDownload {
  content: string
  filename: string
}

export interface CapturedBlobDownload {
  blob: Blob
  filename: string
}

export interface AppDownloads {
  downloads: CapturedDownload[]
  blobDownloads: CapturedBlobDownload[]
}

export function renderApp(): AppDownloads {
  localStorage.clear()
  const downloads: CapturedDownload[] = []
  const blobDownloads: CapturedBlobDownload[] = []
  render(
    <DownloadProvider value={(content, filename) => downloads.push({ content, filename })}>
      <BlobDownloadProvider value={(blob, filename) => blobDownloads.push({ blob, filename })}>
        <App />
      </BlobDownloadProvider>
    </DownloadProvider>,
  )
  return { downloads, blobDownloads }
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
