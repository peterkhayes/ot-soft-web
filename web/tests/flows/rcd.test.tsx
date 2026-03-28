import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, normalizeOutput, renderApp } from '../helpers'

/** Build a valid a priori rankings file for a given set of constraint abbreviations with no rankings set. */
function emptyAprioriFile(abbrevs: string[]): string {
  const header = '\t' + abbrevs.join('\t')
  const rows = abbrevs.map((a) => a + '\t'.repeat(abbrevs.length))
  return [header, ...rows].join('\n')
}

test('RCD: load example, run, see results, download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()
  await page.getByText('Run RCD Algorithm').click()

  // Assert results appear
  await expect
    .element(page.getByText('A ranking was found that generates the correct outputs'))
    .toBeVisible()
  await expect.element(page.getByText('Stratum 1')).toBeVisible()

  // Download button appears and works
  await expect.element(page.getByText('Download Results')).toBeVisible()
  await page.getByText('Download Results').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileOutput.txt')
  expect(normalizeOutput(downloads[0].content)).toMatchInlineSnapshot(`
    "Results of Applying Recursive Constraint Demotion to TinyIllustrativeFile.txt


    <TIMESTAMP>

    OTSoft 2.7, release date 2/1/2026


    1. Result

    A ranking was found that generates the correct outputs.

       Stratum #1
          *No Onset                                 *NoOns
          *Coda                                     *Coda
       Stratum #2
          Max(t)                                    Max
          Dep(?)                                    Dep

    2. Tableaux


    /a/: 
        *NoOns¦*Coda|Max¦Dep
    >?a       ¦     |   ¦1  
     a    1!  ¦     |   ¦   


    /tat/: 
         *NoOns¦*Coda|Max¦Dep
    >ta        ¦     |1  ¦   
     tat       ¦ 1!  |   ¦   


    /at/: 
         *NoOns¦*Coda|Max¦Dep
    >?a        ¦     |1  ¦1  
     ?at       ¦ 1!  |   ¦1  
     a     1!  ¦     |1  ¦   
     at    1!  ¦ 1   |   ¦   


    3. Status of Proposed Constraints:  Necessary or Unnecessary

       *NoOns  Necessary
       *Coda   Necessary
       Max     Not necessary (but included to show Faithfulness violations
                  of a winning candidate)
       Dep     Not necessary (but included to show Faithfulness violations
                  of a winning candidate)

    A check has determined that the grammar will still work even if the 
    constraints marked above as unnecessary are removed en masse.


    4. Ranking Arguments, based on the Fusional Reduction Algorithm

    This run sought to obtain the Skeletal Basis, intended to keep each final ranking argument as pithy as possible.



    The final rankings obtained are as follows:

          *Coda >> Max
          *NoOns >> Dep


    5. Mini-Tableaux

    The following small tableaux may be useful in presenting ranking arguments. 
    They include all winner-rival comparisons in which there is just one 
    winner-preferring constraint and at least one loser-preferring constraint.  
    Constraints not violated by either candidate are omitted.


    /a/: 
        *NoOns|Dep
    >?a       |1  
     a    1   |   


    /tat/: 
         *Coda|Max
    >ta       |1  
     tat  1   |   


    /at/: 
         *Coda|Max¦Dep
    >?a       |1  ¦1  
     ?at  1   |   ¦1  


    /at/: 
        *NoOns|Max¦Dep
    >?a       |1  ¦1  
     a    1   |1  ¦   

    "
  `)
})

test('RCD: download HTML tableaux', async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Switch to "Never switch" axis mode so HTML uses the normal (non-transposed) layout
  await page.getByRole('radio', { name: 'Never switch' }).click()

  await page.getByText('Classical OT', { exact: true }).click()
  await page.getByText('Run RCD Algorithm').click()
  await expect
    .element(page.getByText('A ranking was found that generates the correct outputs'))
    .toBeVisible()

  await page.getByText('Download HTML').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileOutput.html')
  const html = normalizeOutput(downloads[0].content)
  // Verify it's a complete HTML document with styled tableaux
  expect(html).toContain('<!DOCTYPE html>')
  expect(html).toContain(
    '<title>OTSoft 2.7, release date 2/1/2026 TinyIllustrativeFile.txt</title>',
  )
  expect(html).toContain('<TIMESTAMP>')
  expect(html).toContain('1. Result')
  expect(html).toContain('A ranking was found that generates the correct outputs.')
  expect(html).toContain('2. Tableaux')
  expect(html).toContain('/a/')
  expect(html).toContain('/tat/')
  expect(html).toContain('/at/')
  // Shading classes present
  expect(html).toContain('class="cl4"')
  expect(html).toContain('class="cl8"')
  // Fatal violation markers
  expect(html).toContain('*!')
  // Winner marker
  expect(html).toContain('&#x261E;')
  expect(html).toContain('3. Status of Proposed Constraints')
  expect(html).toContain('4. Ranking Arguments')
  expect(html).toContain('5. Mini-Tableaux')
})

test('RCD: download sorted input file', async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()
  await page.getByText('Run RCD Algorithm').click()
  await expect
    .element(page.getByText('A ranking was found that generates the correct outputs'))
    .toBeVisible()

  // Sorted input download button appears after results
  await expect.element(page.getByText('Download Sorted Input')).toBeVisible()
  await page.getByText('Download Sorted Input').click()

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileSorted.txt')
  // Sorted input: constraints in stratum order, candidates sorted by harmony
  const content = downloads[0].content
  const lines = content.split('\n')
  // Header rows have constraints in stratum order
  expect(lines[0]).toContain('*No Onset\t*Coda\tMax(t)\tDep(?)')
  expect(lines[1]).toContain('*NoOns\t*Coda\tMax\tDep')
  // Winners are first for each form
  expect(lines[2]).toMatch(/^a\t\?a\t/)
  expect(lines[4]).toMatch(/^tat\tta\t/)
  expect(lines[6]).toMatch(/^at\t\?a\t/)
  // "at" form rivals sorted by harmony: ?at, a, at
  expect(lines[7]).toMatch(/^\t\?at\t/)
  expect(lines[8]).toMatch(/^\ta\t/)
  expect(lines[9]).toMatch(/^\tat\t/)
})

test('RCD: a priori rankings file upload', async () => {
  renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()

  // Enable a priori rankings checkbox
  await page.getByRole('checkbox', { name: 'Use a priori rankings' }).first().click()

  // Upload an empty apriori file (no rankings enforced — same result as no file)
  const abbrevs = ['*NoOns', '*Coda', 'Max', 'Dep']
  const content = emptyAprioriFile(abbrevs)
  const file = new File([content], 'APrioriRankings.txt', { type: 'text/plain' })
  await page.getByTestId('rcd-apriori-file-input').upload(file)

  // Filename appears in button label
  await expect.element(page.getByText('Loaded: APrioriRankings.txt')).toBeVisible()

  // RCD runs successfully with the apriori file
  await page.getByText('Run RCD Algorithm').click()
  await expect
    .element(page.getByText('A ranking was found that generates the correct outputs'))
    .toBeVisible()
})
