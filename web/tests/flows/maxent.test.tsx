import { test, expect } from 'vitest'
import { page } from '@vitest/browser/context'
import { renderApp, loadExample } from '../helpers'

test('MaxEnt: load example, switch framework, run, see results, download', async () => {
  const downloads = renderApp()
  await loadExample()

  // Switch to Maximum Entropy framework
  await page.getByRole('radio', { name: 'Maximum Entropy' }).click()

  // Run MaxEnt
  await page.getByText('Run MaxEnt').click()

  // Assert results appear
  await expect.element(page.getByText('Constraint Weights')).toBeVisible()
  await expect.element(page.getByText('Log probability of data:')).toBeVisible()
  await expect.element(page.getByText('Predicted Probabilities')).toBeVisible()

  // Download button appears and works
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileMaxEntOutput.txt')
  expect(downloads[0].content).toMatchInlineSnapshot()
})
