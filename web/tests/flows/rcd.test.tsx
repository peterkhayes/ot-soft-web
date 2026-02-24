import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { loadExample, normalizeOutput, renderApp } from '../helpers'

test('RCD: load example, run, see results, download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  // Classical OT is selected by default; run RCD
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
     a   1!   ¦     |   ¦   


    /tat/: 
         *NoOns¦*Coda|Max¦Dep
    >ta        ¦     |1  ¦   
     tat       ¦ 1!  |   ¦   


    /at/: 
         *NoOns¦*Coda|Max¦Dep
    >?a        ¦     |1  ¦1  
     ?at       ¦ 1!  |   ¦1  
     a    1!   ¦     |1  ¦   
     at   1!   ¦ 1   |   ¦   

    3. Status of Proposed Constraints:  Necessary or Unnecessary

       *NoOns  Necessary
       *Coda  Necessary
       Max  Not necessary (but included to show Faithfulness violations
                  of a winning candidate)
       Dep  Not necessary (but included to show Faithfulness violations
                  of a winning candidate)

    A check has determined that the grammar will still work even if the 
    constraints marked above as unnecessary are removed en masse.


    4. Ranking Arguments, based on the Fusional Reduction Algorithm

    This run sought to obtain the Skeletal Basis, intended to keep each final ranking argument as pithy as possible.



    Original set of ERCs:

       #       ERC   Evidence
       1       WeeL  for /a/,  ?a >> a, for /at/,  ?a >> a
       2       eWLe  for /tat/,  ta >> tat, for /at/,  ?a >> ?at
       3       WWLL  for /at/,  ?a >> at

    Recursive ranking search

       Recursive search has now reached this location in the search tree:  1

       Fusion of this ERC set is:  WWLL
       The following ERCs form the total information-loss residue:
          WeeL
          eWLe

       Fusion of total residue:  WWLL

       Skeletal basis of the fusion:  WWee
          WWee has no L's, so it cannot be retained in the Skeletal Basis.

       Recursive search has now reached this location in the search tree:  1, 1
       Current set of ERCs is based on constraint #1, *NoOns
       Working with the following ERC set:
          eWLe

       Fusion of this ERC set is:  eWLe
       The following ERCs form the total information-loss residue:
          (none)

       (The total information-loss residue is empty.)

       eWLe has a null residue and thus may be retained in the Skeletal Basis of ERCs.

       Recursive search has now reached this location in the search tree:  1, 2
       Current set of ERCs is based on constraint #2, *Coda
       Working with the following ERC set:
          WeeL

       Fusion of this ERC set is:  WeeL
       The following ERCs form the total information-loss residue:
          (none)

       (The total information-loss residue is empty.)

       WeeL has a null residue and thus may be retained in the Skeletal Basis of ERCs.


    Ranking argumentation:  Final result

    The following set of ERCs forms the Skeletal Basis for the ERC set as a whole, and thus encapsulates the available ranking information.

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
     a   1    |   


    /tat/: 
         *Coda|Max
    >ta       |1  
     tat  1   |   


    /at/: 
         *Coda|Max
    >?a       |1  
     ?at  1   |   


    /at/: 
        *NoOns|Dep
    >?a       |1  
     a   1    |   

    "
  `)
})
