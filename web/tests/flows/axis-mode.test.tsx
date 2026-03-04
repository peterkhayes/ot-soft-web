import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, normalizeOutput, renderApp } from '../helpers'

test('Axis mode: default is "Switch all" with transposed layout', async () => {
  renderApp()
  await loadExample()

  // "Switch all" radio should be checked by default
  await expect.element(page.getByRole('radio', { name: 'Switch all' })).toBeChecked()

  // Transposed layout: constraint abbreviations appear as body cells (not column headers)
  await expect.element(page.getByRole('cell', { name: '*NoOns' }).first()).toBeVisible()
})

test('Axis mode: switch to "Never switch" shows normal per-form layout', async () => {
  renderApp()
  await loadExample()

  // Verify we start in transposed layout
  await expect.element(page.getByRole('cell', { name: '*NoOns' }).first()).toBeVisible()

  // Switch to "Never switch"
  const neverSwitchRadio = page.getByRole('radio', { name: 'Never switch' })
  await expect.element(neverSwitchRadio).toBeVisible()
  await neverSwitchRadio.click()

  // After switching, transposed cells should disappear and constraint headers should appear
  // In normal layout, *NoOns appears as a column header <th>, not a body cell <td>
  await expect.element(page.getByText('*NoOns').first()).toBeVisible()
})

test('Axis mode: HTML download uses selected axis mode', async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Default is "Switch all" — run RCD and download HTML
  await page.getByText('Run RCD Algorithm').click()
  await expect
    .element(page.getByText('A ranking was found that generates the correct outputs'))
    .toBeVisible()

  await page.getByText('Download HTML').click()

  expect(downloads).toHaveLength(1)
  const html = normalizeOutput(downloads[0].content)
  // Transposed layout: winner marker and constraint abbreviations as rows
  expect(html).toContain('&#x261E;')
  expect(html).toContain('*NoOns')
  expect(html).toContain('*Coda')
})
