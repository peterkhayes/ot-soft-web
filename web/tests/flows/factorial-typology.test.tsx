import { test, expect } from 'vitest'
import { page } from '@vitest/browser/context'
import { renderApp, loadExample } from '../helpers'

test('Factorial Typology: load example, run, see results, download', async () => {
  const downloads = renderApp()
  await loadExample()

  // Classical OT is the default framework; FT panel is below RCD
  // Run factorial typology
  await page.getByText('Run Factorial Typology').click()

  // Assert results appear
  await expect.element(page.getByText(/output patterns? found/)).toBeVisible()
  await expect.element(page.getByText('Output Patterns')).toBeVisible()
  await expect.element(page.getByText('List of Winners')).toBeVisible()

  // Download button appears and works
  await expect.element(page.getByText('Download Results')).toBeVisible({ timeout: 10000 })
  // There are two Download buttons (RCD and FT); click the one inside FT section
  const downloadButtons = page.getByText('Download Results')
  await downloadButtons.last().click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileFactorialTypology.txt')
  expect(downloads[0].content).toMatchInlineSnapshot()
})
