import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, renderApp } from '../helpers'

test('NHG: load example, run, see results, download', { timeout: 30000 }, async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Switch to Noisy Harmonic Grammar framework — radio input is display:none, click label text
  await page.getByText('Noisy Harmonic Grammar', { exact: true }).click()

  // Run NHG
  await page.getByText('Run Noisy HG').click()

  // Assert structural results appear (no content snapshot — stochastic output)
  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
  await expect
    .element(page.getByRole('heading', { name: 'Matchup to Input Frequencies' }))
    .toBeVisible()

  // Download button appears
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileNHGOutput.txt')
  expect(downloads[0].content).toBeTruthy()
})

test('NHG: custom learning schedule runs successfully', { timeout: 30000 }, async () => {
  renderApp()
  await loadExample()

  await page.getByText('Noisy Harmonic Grammar', { exact: true }).click()

  // Enable custom schedule by clicking the label text
  await page.getByText('Use custom learning schedule').click()

  // The textarea should appear with a default schedule
  await expect.element(page.getByRole('textbox')).toBeVisible()

  // Run with the default template schedule
  await page.getByText('Run Noisy HG').click()

  // Results should still appear
  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
})

test('NHG: load custom schedule from file', { timeout: 30000 }, async () => {
  renderApp()
  await loadExample()

  await page.getByText('Noisy Harmonic Grammar', { exact: true }).click()
  await page.getByText('Use custom learning schedule').click()

  // Upload a schedule file
  const content =
    'Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n' +
    '5000\t2\t2\t2\t2\n' +
    '5000\t0.2\t0.2\t2\t2'
  const file = new File([content], 'MySchedule.txt', { type: 'text/plain' })
  await page.getByTestId('nhg-schedule-file-input').upload(file)

  // Button label updates to show the filename
  await expect.element(page.getByText('Loaded: MySchedule.txt')).toBeVisible()

  // Run succeeds with the uploaded schedule
  await page.getByText('Run Noisy HG').click()
  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()
})
