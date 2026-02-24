import { test, expect } from 'vitest'
import { page } from '@vitest/browser/context'
import path from 'path'
import { renderApp, loadExample, loadFile } from '../helpers'

test('RCD: load example, run, see results, download', async () => {
  const downloads = renderApp()
  await loadExample()

  // Classical OT is selected by default; run RCD
  await page.getByText('Run RCD Algorithm').click()

  // Assert results appear
  await expect.element(page.getByText('A ranking was found that generates the correct outputs')).toBeVisible()
  await expect.element(page.getByText('Stratum 1')).toBeVisible()

  // Download button appears and works
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileOutput.txt')
  expect(downloads[0].content).toMatchInlineSnapshot()
})

test('RCD: file upload flow', async () => {
  renderApp() // downloads not needed for this test
  await loadFile(path.resolve(__dirname, '../../../examples/tiny/input.txt'))
  await page.getByText('Run RCD Algorithm').click()
  await expect.element(page.getByText('A ranking was found that generates the correct outputs')).toBeVisible()
})
