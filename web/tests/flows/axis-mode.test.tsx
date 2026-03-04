import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, normalizeOutput, renderApp } from '../helpers'

test('Axis mode: default is "Switch all" with transposed layout', async () => {
  renderApp()
  await loadExample()

  // "Switch all" radio should be checked by default
  const switchAllRadio = page.getByRole('radio', { name: 'Switch all' })
  await expect.element(switchAllRadio).toBeChecked()

  // Transposed layout shows constraint abbreviations in table body cells (constraint-label-cell)
  // Use .first() since the abbreviation appears once per form table
  await expect.element(page.getByRole('cell', { name: '*NoOns' }).first()).toBeVisible()

  // In transposed layout, there should be no "Input" column header
  // (the normal layout has "Input", "Candidate", "Freq" headers)
  await expect
    .element(page.getByRole('columnheader', { name: 'Input', exact: true }))
    .not.toBeInTheDocument()
})

test('Axis mode: switch to "Never switch" shows traditional layout', async () => {
  renderApp()
  await loadExample()

  // Verify transposed layout is showing (default)
  await expect.element(page.getByRole('cell', { name: '*NoOns' }).first()).toBeVisible()

  // Switch to "Never switch"
  await page.getByRole('radio', { name: 'Never switch' }).click()

  // Traditional layout has "Input", "Candidate", "Freq" column headers
  await expect.element(page.getByText('Input', { exact: true })).toBeVisible()
  await expect.element(page.getByText('Candidate', { exact: true })).toBeVisible()
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
  // Transposed layout: candidate names appear as column headers in the form table
  expect(html).toContain('&#x261E;')
  // Constraint abbreviations should appear as row headers
  expect(html).toContain('*NoOns')
  expect(html).toContain('*Coda')
})
