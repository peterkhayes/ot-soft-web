import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, normalizeOutput, renderApp } from '../helpers'

test('MaxEnt: load example, switch framework, run, see results, download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Switch to Maximum Entropy framework — the radio input is display:none,
  // so click the visible label text instead
  await page.getByText('Maximum Entropy', { exact: true }).click()

  // Run MaxEnt
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
    "Results of Applying Maximum Entropy to TinyIllustrativeFile.txt


    <TIMESTAMP>

    OTSoft 2.7, release date 2/1/2026


    Parameters:
       Iterations: 100
       Weight minimum: 0
       Weight maximum: 50


    1. Constraint Weights

       *No Onset                                 21.416
       *Coda                                     21.416
       Max(t)                                    0.000
       Dep(?)                                    0.000

       Log probability of data: -0.000000


    2. Tableaux


    /a/:
             Obs%    Pred%  *NoOns  *Coda  Max  Dep
    >?a   100.0%   100.0%                        1
     a      0.0%     0.0%       1                 


    /tat/:
              Obs%    Pred%  *NoOns  *Coda  Max  Dep
    >ta    100.0%   100.0%                   1     
     tat     0.0%     0.0%              1          


    /at/:
              Obs%    Pred%  *NoOns  *Coda  Max  Dep
    >?a    100.0%   100.0%                   1    1
     ?at     0.0%     0.0%              1         1
     a       0.0%     0.0%       1           1     
     at      0.0%     0.0%       1      1          

    "
  `)
})
