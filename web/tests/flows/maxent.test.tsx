import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, normalizeOutput, renderApp } from '../helpers'

test('MaxEnt: load example, run, see results, download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Maximum Entropy is the default framework; run directly
  await page.getByText('Run MaxEnt').click()

  // Assert results appear
  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()
  await expect.element(page.getByText('Log probability of data:')).toBeVisible()
  await expect.element(page.getByRole('heading', { name: 'Predicted Probabilities' })).toBeVisible()

  // Download button appears and works
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileMaxEntOutput.txt')
  expect(normalizeOutput(downloads[0].content)).toMatchInlineSnapshot(`
    "Result of Applying Maximum Entropy to TinyIllustrativeFile.txt


    OTSoft 2.7, release date 2/1/2026

    <TIMESTAMP>


    For more detailed examination of results, please use a spreadsheet program to open the file 
    TabbedOutput.txt, located in the folder FilesForTinyIllustrativeFile.



    1. Constraints and weights


      20.725	*No Onset
      20.725	*Coda
       0.000	Max(t)
       0.000	Dep(?)


    2. Inputs, candidates, input frequencies, input proportions, predicted probabilities

    Inputs  Candidates  Input frequencies  Input proportions  Predicted probabilities
    a    0  0.000  0.000
      ?a  1  1.000  1.000
      a  0  0.000  0.000


    Inputs  Candidates  Input frequencies  Input proportions  Predicted probabilities
    tat    0  0.000  0.000
      ta  1  1.000  1.000
      tat  0  0.000  0.000


    Inputs  Candidates  Input frequencies  Input proportions  Predicted probabilities
    at    0  0.000  0.000
      ?a  1  1.000  1.000
      ?at  0  0.000  0.000
      a  0  0.000  0.000
      at  0  0.000  0.000


    Probability of data = -0.000000003991123878821702



    3. Weights Found

    20.725    20.725   *No Onset
    20.725    20.725   *Coda
    0.000     0.000   Max(t)
    0.000     0.000   Dep(?)

    4. Tableaux

    Input  Candidate  Harmony  exp(-H)  Predicted  Observed  *NoOns  *Coda  Max  Dep
                                                              20.725  20.725  0.000  0.000
      ?a  0.000  1.000  1.000  1.000                            *
      a  20.725  0.000  0.000  0.000       *                     

    Input  Candidate  Harmony  exp(-H)  Predicted  Observed  *NoOns  *Coda  Max  Dep
                                                              20.725  20.725  0.000  0.000
      ta  0.000  1.000  1.000  1.000                     *       
      tat  20.725  0.000  0.000  0.000              *              

    Input  Candidate  Harmony  exp(-H)  Predicted  Observed  *NoOns  *Coda  Max  Dep
                                                              20.725  20.725  0.000  0.000
      ?a  0.000  1.000  1.000  1.000                     *      *
      ?at  20.725  0.000  0.000  0.000              *             *
      a  20.725  0.000  0.000  0.000       *             *       
      at  41.451  0.000  0.000  0.000       *      *              

    Learning time:  0.000 minutes


    "
  `)
})

test('MaxEnt: generate history of weights and download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Enable history generation
  await page.getByText('Generate history of weights').click()

  // Run MaxEnt
  await page.getByText('Run MaxEnt').click()

  // Wait for results
  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()

  // Download History button should appear
  await expect.element(page.getByText('Download History')).toBeVisible()
  await page.getByText('Download History').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileHistoryOfWeights.txt')

  // MaxEnt history is deterministic — snapshot the content
  const content = downloads[0].content
  // Header should start with tab (leading tab) and contain constraint names
  const lines = content.split('\n').filter((l: string) => l.length > 0)
  expect(lines[0]).toMatch(/^\t/)
  // Should have header + initial (row 0) + 5 iterations = 7 lines
  expect(lines.length).toBe(7)
  // First data row should be iteration 0 with all zeros
  expect(lines[1]).toMatch(/^0\t/)
  // Last data row should be iteration 5
  expect(lines[6]).toMatch(/^5\t/)
})

test('MaxEnt: generate history of output probabilities and download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Enable output probability history
  await page.getByText('Generate history of output probabilities').click()

  // Run MaxEnt
  await page.getByText('Run MaxEnt').click()

  // Wait for results
  await expect.element(page.getByRole('heading', { name: 'Constraint Weights' })).toBeVisible()

  // Download Output Probability History button should appear
  await expect.element(page.getByText('Download Output Probability History')).toBeVisible()
  await page.getByText('Download Output Probability History').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileHistoryOfOutputProbabilities.txt')

  // MaxEnt output prob history is deterministic — check structure
  const content = downloads[0].content
  const lines = content.split('\n').filter((l: string) => l.length > 0)

  // Header should contain input form names
  expect(lines[0]).toContain('a')
  expect(lines[0]).toContain('tat')
  expect(lines[0]).toContain('at')

  // Should have header + 5 iteration rows = 6 lines
  expect(lines.length).toBe(6)

  // First data row should be iteration 1 (no initial row)
  expect(lines[1]).toMatch(/^1\t/)

  // Last data row should be iteration 5
  expect(lines[5]).toMatch(/^5\t/)
})
