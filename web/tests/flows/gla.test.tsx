import { test, expect } from 'vitest'
import { page } from '@vitest/browser/context'
import { renderApp, loadExample } from '../helpers'

test('GLA (Stochastic OT): load example, run, see results, download', async () => {
  const downloads = renderApp()
  await loadExample()

  // Switch to Stochastic OT framework — radio input is display:none, click label text
  await page.getByText('Stochastic OT', { exact: true }).click()

  // Run GLA
  await page.getByText('Run GLA').click()

  // Assert structural results appear (no content snapshot — stochastic output)
  await expect.element(page.getByRole('heading', { name: 'Constraint Ranking Values' })).toBeVisible({ timeout: 15000 })
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
  await expect.element(page.getByRole('heading', { name: 'Matchup to Input Frequencies' })).toBeVisible()

  // Download button appears
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileGLA-StochasticOTOutput.txt')
  expect(downloads[0].content).toBeTruthy()
})

test('GLA (Online MaxEnt): switch mode, run, see results', async () => {
  renderApp() // downloads not needed for this test
  await loadExample()

  // Switch to Stochastic OT framework
  await page.getByText('Stochastic OT', { exact: true }).click()

  // Switch to Online MaxEnt mode inside GLA panel
  await page.getByRole('radio', { name: 'Online MaxEnt (weights)' }).click()

  await page.getByText('Run GLA').click()

  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible({ timeout: 15000 })
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
})
