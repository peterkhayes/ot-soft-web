import { test, expect } from 'vitest'
import { page } from '@vitest/browser/context'
import { renderApp, loadExample } from '../helpers'

test('Hasse diagram: appears after running RCD with FRed enabled', async () => {
  renderApp()
  await loadExample()

  // Classical OT + FRed is selected by default; run RCD
  await page.getByText('Run RCD Algorithm').click()

  // Results appear
  await expect.element(page.getByText('A ranking was found that generates the correct outputs')).toBeVisible()

  // Hasse diagram section is shown
  await expect.element(page.getByText('Hasse Diagram')).toBeVisible()

  // Export buttons are enabled once SVG is rendered (only enabled after render)
  const svgBtn = page.getByRole('button', { name: /SVG/i })
  const pngBtn = page.getByRole('button', { name: /PNG/i })
  await expect.element(svgBtn).toBeEnabled()
  await expect.element(pngBtn).toBeEnabled()
})

test('Hasse diagram: hidden when FRed is disabled', async () => {
  renderApp()
  await loadExample()

  // Uncheck "Include ranking arguments"
  await page.getByRole('checkbox', { name: /Include ranking arguments/i }).click()

  await page.getByText('Run RCD Algorithm').click()
  await expect.element(page.getByText('A ranking was found that generates the correct outputs')).toBeVisible()

  // No Hasse diagram
  const hasse = page.getByText('Hasse Diagram')
  await expect.element(hasse).not.toBeInTheDocument()
})
