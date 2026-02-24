import { test, expect } from 'vitest'
import { page } from '@vitest/browser/context'
import { renderApp, loadExample, normalizeOutput } from '../helpers'

test('Factorial Typology: load example, run, see results, download', async () => {
  const downloads = renderApp()
  await loadExample()

  // Classical OT is the default framework; FT panel is below RCD
  await page.getByText('Run Factorial Typology').click()

  // Assert results appear — ft-section-headers are <div>s not headings,
  // so use { exact: true } to avoid partial matches with the summary text
  await expect.element(page.getByText(/\d+ output patterns? found/)).toBeVisible()
  await expect.element(page.getByText('Output Patterns', { exact: true })).toBeVisible()
  await expect.element(page.getByText('List of Winners', { exact: true })).toBeVisible()

  // Download button appears and works (only FT has run, so there is exactly one)
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileFactorialTypology.txt')
  expect(normalizeOutput(downloads[0].content)).toMatchInlineSnapshot(`
    "Results of Factorial Typology Search for TinyIllustrativeFile.txt


    <TIMESTAMP>

    OTSoft 2.7, release date 2/1/2026
    Source file:  TinyIllustrativeFile.txt


    Constraints

       1. *No Onset                                 *NoOns
       2. *Coda                                     *Coda
       3. Max(t)                                    Max
       4. Dep(?)                                    Dep


    Summary Information

    With 4 constraints, the number of logically possible grammars is 24.
    There were 4 different output patterns.

    Forms marked as winners in the input file are marked with >.

           Output #1  Output #2  Output #3  Output #4
    /a/    >?a        >?a         a          a
    /tat/  >ta         tat       >ta         tat
    /at/   >?a         ?at        a          at



    List of Winners

    The following specifies for each candidate whether there is at least one ranking that derives it:

    /a/:
       >[?a          ]  yes
        [a           ]  yes

    /tat/:
       >[ta          ]  yes
        [tat         ]  yes

    /at/:
       >[?a          ]  yes
        [?at         ]  yes
        [a           ]  yes
        [at          ]  yes



    T-Orders

    The t-order is the set of implications in a factorial typology.

    If this input           has this output         then this input         has this output
    /at                  /  [?a                  ]  /a                   /  [?a]
    /at                  /  [?a                  ]  /tat                 /  [ta]
    /at                  /  [?at                 ]  /a                   /  [?a]
    /at                  /  [?at                 ]  /tat                 /  [tat]
    /at                  /  [a                   ]  /a                   /  [a]
    /at                  /  [a                   ]  /tat                 /  [ta]
    /at                  /  [at                  ]  /a                   /  [a]
    /at                  /  [at                  ]  /tat                 /  [tat]

    Nothing is implicated by these input-output pairs:

    Input                 Candidate
    /a                 /  [?a]
    /a                 /  [a]
    /tat               /  [ta]
    /tat               /  [tat]
    "
  `)
})
