import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { clickHasseDownload, loadExample, renderApp } from '../helpers'

test('Hasse diagram: appears after running RCD with FRed enabled', async () => {
  renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()
  await page.getByText('Run RCD Algorithm').click()

  // Results appear
  await expect
    .element(page.getByText('A ranking was found that generates the correct outputs'))
    .toBeVisible()

  // Hasse diagram section is shown
  await expect.element(page.getByText('Hasse Diagram')).toBeVisible()
})

test('Hasse diagram: SVG download contains correct content', async () => {
  const { blobDownloads } = renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()
  await page.getByText('Run RCD Algorithm').click()
  await expect.element(page.getByText('Hasse Diagram')).toBeVisible()

  await clickHasseDownload('Download SVG')

  expect(blobDownloads).toHaveLength(1)
  expect(blobDownloads[0].filename).toBe('TinyIllustrativeFileHasse.svg')
  expect(blobDownloads[0].blob.type).toBe('image/svg+xml')

  const text = await blobDownloads[0].blob.text()
  expect(text).toContain('<svg')
  // Constraint labels from the tiny example should appear in the diagram
  expect(text).toContain('*NoOns')
  expect(text).toContain('*Coda')
})

test('Hasse diagram: PNG download produces a PNG blob', async () => {
  const { blobDownloads } = renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()
  await page.getByText('Run RCD Algorithm').click()
  await expect.element(page.getByText('Hasse Diagram')).toBeVisible()

  await clickHasseDownload('Download PNG')

  // PNG rendering is async (canvas.toBlob), so wait for the download to appear
  await expect.poll(() => blobDownloads).toHaveLength(1)
  expect(blobDownloads[0].filename).toBe('TinyIllustrativeFileHasse.png')
  expect(blobDownloads[0].blob.type).toBe('image/png')
  expect(blobDownloads[0].blob.size).toBeGreaterThan(0)
})

test('Hasse diagram: hidden when FRed is disabled', async () => {
  renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()

  // Uncheck "Include ranking arguments"
  await page.getByRole('checkbox', { name: /Include ranking arguments/i }).click()

  await page.getByText('Run RCD Algorithm').click()
  await expect
    .element(page.getByText('A ranking was found that generates the correct outputs'))
    .toBeVisible()

  // No Hasse diagram
  await expect.element(page.getByText('Hasse Diagram')).not.toBeInTheDocument()
})
