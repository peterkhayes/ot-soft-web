import { test, expect } from 'vitest'
import { page } from '@vitest/browser/context'
import { renderApp, loadExample } from '../helpers'

test('NHG: load example, run, see results, download', async () => {
  const downloads = renderApp()
  await loadExample()

  // Switch to Noisy Harmonic Grammar framework — radio input is display:none, click label text
  await page.getByText('Noisy Harmonic Grammar', { exact: true }).click()

  // Run NHG
  await page.getByText('Run Noisy HG').click()

  // Assert structural results appear (no content snapshot — stochastic output)
  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible({ timeout: 15000 })
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
  await expect.element(page.getByRole('heading', { name: 'Matchup to Input Frequencies' })).toBeVisible()

  // Download button appears
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileNHGOutput.txt')
  expect(downloads[0].content).toBeTruthy()
})
