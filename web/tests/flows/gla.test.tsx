import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, renderApp } from '../helpers'

test(
  'GLA (Stochastic OT): load example, run, see results, download',
  { timeout: 30000 },
  async () => {
    const { downloads } = renderApp()
    await loadExample()

    // Switch to Stochastic OT framework — radio input is display:none, click label text
    await page.getByText('Stochastic OT', { exact: true }).click()

    // Run GLA
    await page.getByText('Run GLA').click()

    // Assert structural results appear (no content snapshot — stochastic output)
    await expect
      .element(page.getByRole('heading', { name: 'Constraint Ranking Values' }))
      .toBeVisible()
    await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
    await expect
      .element(page.getByRole('heading', { name: 'Matchup to Input Frequencies' }))
      .toBeVisible()

    // Download button appears
    await expect.element(page.getByText('Download Results')).toBeVisible()
    await page.getByText('Download Results').click()

    expect(downloads).toHaveLength(1)
    expect(downloads[0].filename).toBe('TinyIllustrativeFileGLA-StochasticOTOutput.txt')
    expect(downloads[0].content).toBeTruthy()
  },
)

test('GLA (Stochastic OT): Hasse diagram appears after run', { timeout: 30000 }, async () => {
  renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()
  await page.getByText('Run GLA').click()

  // Hasse diagram section appears
  await expect.element(page.getByText('Hasse Diagram')).toBeVisible()

  // Export buttons are enabled once the SVG has rendered
  await expect.element(page.getByRole('button', { name: /SVG/i })).toBeEnabled()
  await expect.element(page.getByRole('button', { name: /PNG/i })).toBeEnabled()
})

test(
  'GLA (Stochastic OT): pairwise ranking probabilities table appears',
  { timeout: 30000 },
  async () => {
    renderApp()
    await loadExample()

    await page.getByText('Stochastic OT', { exact: true }).click()
    await page.getByText('Run GLA').click()

    // Pairwise ranking probabilities section appears
    await expect
      .element(page.getByRole('heading', { name: 'Pairwise Ranking Probabilities' }))
      .toBeVisible()
  },
)

test('GLA (Online MaxEnt): Hasse diagram not shown', { timeout: 30000 }, async () => {
  renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()
  await page.getByRole('radio', { name: 'Online MaxEnt (weights)' }).click()
  await page.getByText('Run GLA').click()

  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()
  // Hasse diagram should NOT appear in MaxEnt mode
  await expect.element(page.getByText('Hasse Diagram')).not.toBeInTheDocument()
  // Pairwise ranking probabilities should NOT appear in MaxEnt mode
  await expect.element(page.getByText('Pairwise Ranking Probabilities')).not.toBeInTheDocument()
})

test('GLA (Online MaxEnt): switch mode, run, see results', async () => {
  renderApp() // downloads not needed for this test
  await loadExample()

  // Switch to Stochastic OT framework
  await page.getByText('Stochastic OT', { exact: true }).click()

  // Switch to Online MaxEnt mode inside GLA panel
  await page.getByRole('radio', { name: 'Online MaxEnt (weights)' }).click()

  await page.getByText('Run GLA').click()

  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
})

test('GLA: custom learning schedule runs successfully', { timeout: 30000 }, async () => {
  renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()

  // Enable custom schedule by clicking the label text
  await page.getByText('Use custom learning schedule').click()

  // The textarea should appear with a default schedule
  await expect.element(page.getByRole('textbox')).toBeVisible()

  // Run with the default template schedule
  await page.getByText('Run GLA').click()

  // Results should still appear
  await expect
    .element(page.getByRole('heading', { name: 'Constraint Ranking Values' }))
    .toBeVisible()
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
})

test('GLA: multiple runs downloads CollateRuns.txt', { timeout: 60000 }, async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()

  // Click "Run 10 times & Download" (default is 10 runs)
  await page.getByText('Run 10 times & Download').click()

  // Download should appear with CollateRuns filename
  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toContain('CollateRuns')

  // Content should have G and O records
  const content = downloads[0].content
  expect(content).toMatch(/^G\t/m)
  expect(content).toMatch(/^O\t/m)
})

test('GLA: exact proportions runs successfully', { timeout: 30000 }, async () => {
  renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()

  // Enable exact proportions
  await page.getByText('Present data in exact proportions').click()

  await page.getByText('Run GLA').click()

  await expect
    .element(page.getByRole('heading', { name: 'Constraint Ranking Values' }))
    .toBeVisible()
  await expect.element(page.getByText('Log likelihood of data:')).toBeVisible()
})

test('GLA: load custom schedule from file', { timeout: 30000 }, async () => {
  renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()
  await page.getByText('Use custom learning schedule').click()

  // Upload a schedule file
  const content =
    'Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n' +
    '15000\t2\t2\t2\t2\n' +
    '15000\t0.2\t0.2\t2\t2'
  const file = new File([content], 'MySchedule.txt', { type: 'text/plain' })
  await page.getByTestId('gla-schedule-file-input').upload(file)

  // Button label updates to show the filename
  await expect.element(page.getByText('Loaded: MySchedule.txt')).toBeVisible()

  // Run succeeds with the uploaded schedule
  await page.getByText('Run GLA').click()
  await expect
    .element(page.getByRole('heading', { name: 'Constraint Ranking Values' }))
    .toBeVisible()
})

test('GLA: generate full history and download', { timeout: 30000 }, async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()

  // Enable full history generation
  await page.getByText('Generate full history (with input/output annotations)').click()

  await page.getByText('Run GLA').click()

  // Wait for results
  await expect
    .element(page.getByRole('heading', { name: 'Constraint Ranking Values' }))
    .toBeVisible()

  // Download Full History button should appear
  await expect.element(page.getByText('Download Full History')).toBeVisible()
  await page.getByText('Download Full History').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileFullHistory.txt')

  // Content should have Trial # header and data rows
  const lines = downloads[0].content.split('\n').filter((l: string) => l.length > 0)
  expect(lines[0]).toMatch(/^Trial #\t/)
  expect(lines[1]).toMatch(/^\(Initial\)/)
  expect(lines.length).toBeGreaterThan(2)
})

test('GLA: generate history and download', { timeout: 30000 }, async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Stochastic OT', { exact: true }).click()

  // Enable history generation
  await page.getByText('Generate history of ranking values').click()

  await page.getByText('Run GLA').click()

  // Wait for results
  await expect
    .element(page.getByRole('heading', { name: 'Constraint Ranking Values' }))
    .toBeVisible()

  // Download History button should appear
  await expect.element(page.getByText('Download History')).toBeVisible()
  await page.getByText('Download History').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileHistory.txt')

  // Content should be tab-separated with Trial header
  const lines = downloads[0].content.split('\n').filter((l: string) => l.length > 0)
  expect(lines[0]).toMatch(/^Trial\t/)
  expect(lines.length).toBeGreaterThan(1)
})

test(
  'GLA (Online MaxEnt): candidate probability history download',
  { timeout: 30000 },
  async () => {
    const { downloads } = renderApp()
    await loadExample()

    await page.getByText('Stochastic OT', { exact: true }).click()
    await page.getByRole('radio', { name: 'Online MaxEnt (weights)' }).click()

    // Enable candidate probability history
    await page.getByText('Generate history of candidate probabilities').click()

    await page.getByText('Run GLA').click()

    // Wait for results
    await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()

    // Download button should appear
    await expect.element(page.getByText('Download Candidate Probability History')).toBeVisible()
    await page.getByText('Download Candidate Probability History').click()

    expect(downloads).toHaveLength(1)
    expect(downloads[0].filename).toBe('TinyIllustrativeFileHistoryOfCandidateProbabilities.txt')

    // Content structure
    const lines = downloads[0].content.split('\n').filter((l: string) => l.length > 0)
    expect(lines[0]).toMatch(/^Trial #\t/)
    expect(lines[1]).toMatch(/^\(initial\)/)
    expect(lines.length).toBeGreaterThan(2)
  },
)
