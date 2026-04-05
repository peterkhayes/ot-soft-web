import { page } from '@vitest/browser/context'
import { expect, test } from 'vitest'

import { clickDownload, loadExample, normalizeOutput, renderApp } from '../helpers'

test('Factorial Typology: FTSum download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()

  // Enable the FTSum checkbox before running
  await page.getByText('Generate FTSum file').click()
  await page.getByText('Run Factorial Typology').click()
  await expect.element(page.getByText(/\d+ output patterns? found/)).toBeVisible()

  await clickDownload('Download FTSum')

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileFTSum.txt')
  expect(downloads[0].content).toMatchInlineSnapshot(`
    "/a/	/tat/	/at/
    ?a	ta	?a
    ?a	tat	?at
    a	ta	a
    a	tat	at
    "
  `)
})

test('Factorial Typology: CompactSum download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()

  // Enable the CompactSum checkbox before running
  await page.getByText('Generate CompactSum file').click()
  await page.getByText('Run Factorial Typology').click()
  await expect.element(page.getByText(/\d+ output patterns? found/)).toBeVisible()

  await clickDownload('Download CompactSum')

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileCompactSum.txt')
  expect(downloads[0].content).toMatchInlineSnapshot(`
    "2	?a	ta	
    3	?a	tat	?at	
    2	a	ta	
    3	a	tat	at	
    "
  `)
})

test('Factorial Typology: load example, run, see results, download', async () => {
  const { downloads } = renderApp()
  await loadExample()

  await page.getByText('Classical OT', { exact: true }).click()
  await page.getByText('Run Factorial Typology').click()

  // Assert results appear — ft-section-headers are <div>s not headings,
  // so use { exact: true } to avoid partial matches with the summary text
  await expect.element(page.getByText(/\d+ output patterns? found/)).toBeVisible()
  await expect.element(page.getByText('Output Patterns', { exact: true })).toBeVisible()
  await expect.element(page.getByText('List of Winners', { exact: true })).toBeVisible()

  // Download menu appears and works
  await clickDownload('Download Results')

  expect(downloads).toHaveLength(1)
  expect(downloads[0].filename).toBe('TinyIllustrativeFileFactorialTypology.txt')
  expect(normalizeOutput(downloads[0].content)).toMatchInlineSnapshot(`
    "Results of Factorial Typology Search

    <TIMESTAMP>

    OTSoft 2.7, release date 2/1/2026

    Source file:  TinyIllustrativeFile.txt



    1. Constraints

        Full Name  Abbr. 
    1.  *No Onset  *NoOns
    2.  *Coda      *Coda 
    3.  Max(t)     Max   
    4.  Dep(?)     Dep   

    All rankings were considered.



    2. Summary Information

    With 4 constraints, the number of logically possible grammars is 24.

    There were 4 different output patterns.

    Forms marked as winners in the input file are marked with >.

           Output #1  Output #2  Output #3  Output #4
    /a/    >?a        >?a         a          a
    /tat/  >ta         tat       >ta         tat
    /at/   >?a         ?at        a          at


    3. List of Winners

    The following specifies for each candidate whether there is at least one ranking that derives it:

    /a/:
       [?a]:           yes
       [a]:            yes
    /tat/:
       [ta]:           yes
       [tat]:          yes
    /at/:
       [?a]:           yes
       [?at]:          yes
       [a]:            yes
       [at]:           yes

    4. T-orders

    The t-order is the set of implications in a factorical typology.

    If this input  has this output  then this input  has this output
    /at/           [?a]             /a/              [?a]           
    /at/           [?a]             /tat/            [ta]           
    /at/           [?at]            /a/              [?a]           
    /at/           [?at]            /tat/            [tat]          
    /at/           [a]              /a/              [a]            
    /at/           [a]              /tat/            [ta]           
    /at/           [at]             /a/              [a]            
    /at/           [at]             /tat/            [tat]          

    Nothing is implicated by these input-output pairs:

    Input  Candidate
    a      ?a       
    a      a        
    tat    ta       
    tat    tat      


    "
  `)
})
