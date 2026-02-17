# OTSoft User Interface

This document maps the VB6 user interface elements to the algorithms and parameters they control. Use this as a guide for implementing the web interface.

## Main Window (`Main.frm`)

The main window has a title bar showing "OTSoft 2.7" and the currently loaded file path. The layout is organized into several visual regions:

```
┌──────────────────────────────────────────────────────────────────────┐
│ File  Edit  View  Print  Factorial Typology  A Priori Rankings      │
│                                              Hasse  Options  HTML   │
│                                              Help                   │
├──────────────┬────────────────┬───────────────┬──────────────────────┤
│              │                │               │                      │
│  [Rank btn]  │  (Progress     │  [FacType btn]│                      │
│              │   Window)      │               │                      │
├──────────────┼────────────────┼───────────────┼──────────────────────┤
│              │                │               │                      │
│  Choose      │  Ranking       │  Options for  │                      │
│  framework   │  argumentation │  crowded      │                      │
│              │                │  tableaux     │                      │
│  ○ ClassicOT │  ☑ Include     │               │                      │
│  ● MaxEnt    │    ranking     │  ○ Switch all │                      │
│  ○ NHG       │    arguments   │  ○ Switch     │                      │
│  ○ Stoch.OT  │  ☐ Use MIB    │    where      │                      │
│              │  ☑ Show details│    needed     │                      │
│              │  ☑ Minitableaux│  ● Never      │                      │
│              │                │    switch     │                      │
│              │  ☑ Diagnostics │               │                      │
│              │    if fails    │               │                      │
├──────────────┴────────────────┴───────────────┴──────────────────────┤
│  (Drag-and-drop target area for loading files)                      │
├─────────────────────────────────────────────────────────────────────┤
│              [View Results]                    [Exit]                │
└─────────────────────────────────────────────────────────────────────┘
```

### Primary Buttons

| Button | VB6 Name | Action |
|--------|----------|--------|
| **Rank [filename]** | `cmdRank` | Runs the selected ranking/learning algorithm on the loaded file. Dispatches based on framework radio button selection (see [Algorithm Dispatch](#algorithm-dispatch)). |
| **Factorial typology for [filename]** | `cmdFacType` | Runs factorial typology analysis. Calls `FactorialTypology.Main()`. |
| **View Results** | `cmdViewResults` | Opens the output file using the method selected in the View menu. |
| **Exit** | `cmdExit` | Exits the application. |

### Framework Selection

Radio buttons — only one can be selected at a time.

| Radio Button | VB6 Name | Default | Algorithm(s) Triggered |
|-------------|----------|---------|----------------------|
| Classical OT | `optConstraintDemotion` | No | RCD, or BCD/LFCD if selected in Options menu |
| Maximum Entropy | `optMaximumEntropy` | **Yes** | GLA form opens in MaxEnt mode |
| Noisy Harmonic Grammar | `optNoisyHarmonicGrammar` | No | NHG form opens |
| Stochastic OT | `optGLA` | No | GLA form opens in Stochastic OT mode |

Each framework has a "What is it?" info button (`cmdIdentifyCD`, `cmdIdentifyGLA`, `cmdIdentifyMaximumEntropy`, `cmdIdentifyNHG`) that shows a description dialog.

### Ranking Argumentation Frame

These control the FRed ranking argumentation algorithm that runs after categorical ranking.

| Checkbox | VB6 Name | Default | Effect |
|----------|----------|---------|--------|
| Include ranking arguments | `chkArguerOn` | Checked | Enables/disables FRed. When unchecked, hides the sub-options below. |
| Use Most Informative Basis | `chkMostInformativeBasis` | Unchecked | FRed mode: MIB (checked) or Skeletal Basis (unchecked) |
| Show details of argumentation | `chkDetailedArguments` | Checked | Include verbose FRed recursion tree in output |
| Include illustrative minitableaux | `chkMiniTableaux` | Checked | Generate mini-tableaux for pairwise ranking arguments |

### Options for Crowded Tableaux

Radio buttons controlling tableau axis orientation.

| Radio Button | VB6 Name | Default | Effect |
|-------------|----------|---------|--------|
| Switch axes for all tableaux | `optSwitchAll` | **Selected** | Constraints on vertical axis, candidates horizontal |
| Switch axes where needed | `optSwitchSomeAxes` | No | Auto-switch only when tableaux are too wide |
| Never switch axes | `optNeverSwitchAxes` | No | Traditional layout (constraints horizontal) |

### Other Main Window Controls

| Control | VB6 Name | Default | Effect |
|---------|----------|---------|--------|
| Diagnostics if ranking fails | `chkDiagnosticTableaux` | Checked | Show diagnostic tableaux when RCD/BCD/LFCD fails |
| Progress Window | `lblProgressWindow` | Empty | Shows algorithm progress messages during computation |

---

## Algorithm Dispatch

When the user clicks **Rank**, the `Rank()` subroutine dispatches based on framework selection:

```
If "Stochastic OT" selected:
    → Open GLA form (boersma.frm) in Stochastic OT mode

If "Maximum Entropy" selected:
    → Open GLA form (boersma.frm) in MaxEnt mode

If "Noisy Harmonic Grammar" selected:
    → Open NHG form (NoisyHarmonicGrammar.frm)

If "Classical OT" selected:
    → Check Options menu:
        If "Use Biased Constraint Demotion" checked:
            → Run BCD
        Else if "Use Low Faithfulness version of RCD" checked:
            → Run LFCD
        Else:
            → Run RCD
    → Then optionally run FRed (if "Include ranking arguments" checked)
    → Then print tableaux
```

---

## Menu Bar

### File Menu (`mnuFile`)

| Item | Shortcut | Action |
|------|----------|--------|
| Open | — | Shows instructions to drag-and-drop files onto the window |
| Reload new version | — | Reloads the currently loaded file (hidden until file loaded) |
| Save as .txt file | — | Saves input data as tab-delimited text |
| Save as Praat file | — | Exports to Praat `.OTGrammar` and `.PairDistribution` formats |
| Save as R file | — | Exports for R logistic regression analysis |
| Exit | Ctrl+X | Exits application |
| Open recent → | — | Submenu with 6 recently opened files |

### Edit Menu (`mnuEdit`)

| Item | Action |
|------|--------|
| Edit current file | Opens the loaded input file in the system editor (Excel for .xlsx) |

### View Menu (`mnuWordProcessorChoice`)

| Item | Action |
|------|--------|
| View result in OTSoft | Display output in the internal text viewer |
| View with your word processor | Open output file in system word processor |
| View result as web page | Open HTML output in web browser |
| Prepare for printing (MS Word) | Open quality output formatted for Word |
| View Hasse diagram of rankings | Display the ranking Hasse diagram (requires GraphViz) |
| Show how ranking was done | Open the `HowIRanked` log file |

### Print Menu (`mnuPrint`)

| Item | Action |
|------|--------|
| Prepare for quality printing (MS Word) | Same as View → Prepare for printing |
| Draft print (MS Word not needed) | Opens print options dialog (portrait/landscape, reduction %) |

### Factorial Typology Menu (`mnuFactoricalTypology`)

| Item | Default | Action |
|------|---------|--------|
| Include rankings in results | Unchecked | Add grammar listing for each output pattern |
| Include tableaux in results | Unchecked | Add tableaux (requires rankings checked) |
| Generate compact FT summary file | Unchecked | Write `FTSum.txt` |
| Compact file collapsing neutralized outputs | Unchecked | Write `CompactSum.txt` |
| View compact FT summary file | Hidden | Opens `FTSum.txt` (appears after generation) |
| View compact file collapsing neutralized outputs | Hidden | Opens `CompactSum.txt` (appears after generation) |

### A Priori Rankings Menu (`mnuAPrioriRankings`)

| Item | Default | Action |
|------|---------|--------|
| Rank constraints constrained by a priori rankings | Unchecked | Toggle a priori ranking enforcement |
| Make or edit a file containing a priori rankings | — | Creates template and opens in editor |
| Use strata to construct a priori ranking file | — | Saves current ranking result as a priori file |

### Hasse Menu (`mnuHasse`)

| Item | Action |
|------|--------|
| View Hasse diagram | Display Hasse diagram in viewer window |
| Edit the source file underlying Hasse diagram | Open `.dot` file in text editor |
| Replot Hasse diagram from altered source file | Re-run GraphViz to regenerate diagram |

### Options Menu (`mnuOptions`)

| Item | Default | Action |
|------|---------|--------|
| On ranking, save copy of input file, sorted by rank | Unchecked | Auto-save reordered input file after ranking |
| Print constraint names in small caps | Unchecked | Formatting option for output |
| Delete temporary files on exit | Unchecked | Clean up temp files when closing |
| Use the Low Faithfulness version of RCD | Unchecked | Switch Classical OT to LFCD algorithm |
| Use Biased Constraint Demotion | Unchecked | Switch Classical OT to BCD algorithm |
| BCD favors specific Faithfulness constraints | Hidden | Sub-option for specificity in BCD |
| Sort candidates in tableaux by harmony | **Checked** | Order candidates best-to-worst in tableaux |
| Restore default settings | — | Reset all options to defaults |

**Note**: "Low Faithfulness" and "Biased Constraint Demotion" are mutually exclusive — selecting one unchecks the other.

### HTML Menu (`mnuHTML`)

| Item | Action |
|------|--------|
| Options for HTML output | Opens HTML options dialog (shading darkness, custom color) |

### Help Menu (`mnuHelp`)

| Item | Action |
|------|--------|
| View manual as Adobe PDF file | Opens the PDF manual |
| About OTSoft | Shows version, credits, and citation info |

---

## GLA / Stochastic OT Form (`boersma.frm`)

Opens when the user selects Maximum Entropy or Stochastic OT and clicks Rank. This form has its own menu bar and parameter inputs.

### Framework Toggle

| Radio Button | VB6 Name | Default | Effect |
|-------------|----------|---------|--------|
| Compute ranking values for Stochastic OT | `optStochasticOT` | No | GLA with Gaussian noise + strict domination. Initial values default to 100. |
| Compute weights for MaxEnt | `optMaxEnt` | **Yes** | GLA with probability sampling. Initial values default to 0. |

### Parameter Inputs

| Input | VB6 Name | Default | Maps To |
|-------|----------|---------|---------|
| Number of times to go through forms | `txtNumberOfCycles` | `1000000` | Total learning iterations |
| Initial plasticity | `txtUpperPlasticity` | `2` | Starting learning rate |
| Final plasticity | `txtLowerPlasticity` | `.001` | Ending learning rate |
| Constraints ranked a priori must differ by | `txtValueThatImplementsAPrioriRankings` | `20` | Minimum numeric gap for a priori pairs |
| Number of times to test grammar | `txtTimesToTestGrammar` | `10000` | Evaluation trials after learning |

### Buttons

| Button | Action |
|--------|--------|
| Run | Starts the GLA/MaxEnt learning algorithm |
| Exit to main screen | Returns to main window |

### GLA Menus

#### Initial ranking values/weights

| Item | Default | Action |
|------|---------|--------|
| Use default initial values (all same) | **Checked** | All constraints start at 100 (Stochastic OT) or 0 (MaxEnt) |
| Use separate initial values for Markedness and Faithfulness | Hidden | Opens dialog to set separate initial values |
| Use fully customized initial values | Unchecked | Read from `ModelParameters.txt` |
| Use results of previous run as initial values | Unchecked | Read from `MostRecentRankingValues.txt` |
| Specify separate initial values... | — | Opens `frmInitialRankings` dialog |
| Edit file for fully customized initial values | — | Opens `ModelParameters.txt` in editor |

#### Learning schedule

| Item | Default | Action |
|------|---------|--------|
| Use the Magri update rule (stochastic OT) | Unchecked | Alternative update rule for Stochastic OT |
| Use custom learning schedule from file | Unchecked | Read schedule from `CustomLearningSchedule.txt` |
| Edit file with custom learning schedule | — | Opens schedule file in editor |

#### A Priori Rankings

| Item | Action |
|------|--------|
| Run GLA constrained by a priori rankings | Toggle a priori enforcement during learning |
| Make or edit a priori rankings file | Opens file in editor |

#### MaxEnt

| Item | Default | Action |
|------|---------|--------|
| Edit file with MaxEnt learning parameters | — | Opens `ModelParameters.txt` (for mu/sigma) |
| Run MaxEnt with Gaussian prior | Unchecked | Enable biased learning with mu/sigma parameters |
| Run the batch version of MaxEnt | Unchecked | Opens `MyMaxEnt.frm` instead of online GLA-MaxEnt |
| Allow weights to be negative | Unchecked | Remove floor of 0 on weights |

#### Options

| Item | Default | Action |
|------|---------|--------|
| Include tableaux in output file | Unchecked | Append full tableaux to results |
| Include pairwise ranking probabilities | Unchecked | Calculate P(Ci >> Cj) for all pairs |
| Print file with history of weights/ranking values | **Checked** | Write `History.txt` |
| Print same but more annotations | **Checked** | Write `FullHistory.txt` |
| Print history of candidate probabilities (MaxEnt only) | **Checked** | Write `HistoryOfCandidateProbabilities.txt` |
| Present data to GLA in exact proportions | Unchecked | Deterministic exemplar presentation (requires integer frequencies) |
| Multiple runs → 10 / 100 / 1000 runs | — | Run algorithm multiple times and collate results |

---

## Batch MaxEnt Form (`MyMaxEnt.frm`)

Opens from GLA form's menu: MaxEnt → "Run the batch version of MaxEnt".

### Parameter Inputs

| Input | VB6 Name | Default | Maps To |
|-------|----------|---------|---------|
| Number of iterations | `txtPrecision` | `5` | GIS iterations (more = more precise) |
| Report results to N decimal places | `txtDecimalPlaces` | `3` | Output precision |
| Minimum weight | `txtWeightMinimum` | `0` | Lower bound on learned weights |
| Maximum weight | `txtWeightMaximum` | `50` | Upper bound on learned weights |

### Buttons

| Button | Action |
|--------|--------|
| Run maxent | Starts batch MaxEnt optimization |
| Exit to main screen | Returns to main window |

### Menus

| Item | Default | Action |
|------|---------|--------|
| Include tableaux | **Checked** | Append tableaux to output |
| Generate file with history of weights | **Checked** | Write `HistoryOfWeights.txt` |
| Generate file with history of output probabilities | **Checked** | Write `HistoryOfOutputProbabilities.txt` |

---

## Noisy Harmonic Grammar Form (`NoisyHarmonicGrammar.frm`)

Opens when the user selects NHG and clicks Rank.

### Parameter Inputs

| Input | VB6 Name | Default | Maps To |
|-------|----------|---------|---------|
| Number of times to go through forms | `txtNumberOfCycles` | `5000` | Total learning iterations |
| Initial plasticity | `txtUpperPlasticity` | `2` | Starting learning rate |
| Final plasticity | `txtLowerPlasticity` | `0.002` | Ending learning rate |
| Constraints ranked a priori must differ by | `txtValueThatImplementsAPrioriRankings` | `20` | A priori enforcement gap |
| Number of times to test grammar | `txtTimesToTestGrammar` | `2000` | Evaluation trials after learning |

### Noise Configuration Checkboxes

These checkboxes select which of the 8 noise variants to use (see ALGORITHMS.md for details):

| Checkbox | VB6 Name | Default | Effect |
|----------|----------|---------|--------|
| Apply noise by tableau cell, not by constraint | `chkNoiseAppliesToTableauCells` | Unchecked | Per-cell noise vs per-constraint noise |
| Apply noise after multiplication of weights by violations | `chkNoiseIsAddedAfterMultiplication` | Unchecked | Post-multiply noise timing |
| Include noise even in cells with no violation | `chkNoiseForZeroCells` | Hidden | Add noise to zero-violation cells (only visible when post-multiply checked) |
| Allow constraint weights to go negative | `chkNegativeWeightsOK` | Unchecked | Remove non-negativity constraint on weights |
| Add noise to candidates, after harmony calculation | `chkLateNoise` | Unchecked | Single noise term on total harmony |
| Employ Exponential HNG | `chkExponentialNHG` | Unchecked | Use exp(weight) instead of raw weight |

### Buttons

| Button | Action |
|--------|--------|
| Run Noisy HG | Starts NHG learning |
| Exit to main screen | Returns to main window |

### NHG Menus

The menu structure mirrors the GLA form, with these differences:

- **Initial rankings**: Default is "all zero" (not 100)
- **Options → Use positive demi-Gaussians** (`mnuDemiGaussians`): Use half-Gaussian (positive only) noise
- **Options → Resolve ties by skipping trial** (`mnuResolveTiesBySkipping`): Skip update when candidates tie (default: pick randomly)

---

## Dialog Forms

### HTML Options (`frmHTMLOptions`)

Controls tableau shading in HTML output.

| Control | Default | Maps To |
|---------|---------|---------|
| Grayness slider (1-100) | `50` | Converted to HTML hex color for cell shading |
| HTML color code | empty | Direct hex color (overrides grayness) |
| Save and exit | — | Stores in global `gShadingColor` |

### Initial Rankings (`frmInitialRankings`)

Sets separate initial values for Markedness vs Faithfulness constraints.

| Control | Default | Maps To |
|---------|---------|---------|
| Markedness constraints value | `100` | `gCustomRankMark` |
| Faithfulness constraints value | `100` | `gCustomRankFaith` |

### Hasse Diagram Viewer (`frmHasse`)

Displays constraint ranking diagrams generated by GraphViz.

| Menu Item | Action |
|-----------|--------|
| Fit image to screen | Scale diagram to fit window |
| View image at original size | Show at full resolution |
| View image with MSPaint | Open `.gif` in external graphics program |

### Print Options (`frmPrinting`)

| Control | Default | Maps To |
|---------|---------|---------|
| Portrait / Landscape | Portrait | Page orientation |
| Reduction % | `80` | Print scaling |

### About (`frmAboutOTSoft`)

Shows version ("OTSoft 2.7"), credits (Bruce Hayes, Bruce Tesar, Kie Zuraw), and citation information. Has an "Acknowledgments" button listing contributors.

---

## File Loading

Files are loaded by **drag-and-drop** onto the main window (not via a file dialog). The application accepts:
- `.txt` — Tab-delimited text (primary format)
- `.xlsx` — Modern Excel workbooks
- `.in` — Legacy Ranker format

On load, the file is parsed by `DigestTheInputFile()`, which:
1. Detects the format from the file extension
2. Reads constraints, candidates, and violations
3. For categorical algorithms: determines winners from highest-frequency candidates
4. Updates button labels to show the filename
5. Stores the file path for output directory creation

---

## Output File Locations

All outputs are written to a folder created alongside the input file:

```
[InputFilePath]/FilesFor[FileName]/
    [FileName]DraftOutput.txt
    [FileName]QualityOutput.txt
    ResultsFor[FileName].htm
    HowIRanked[FileName].txt
    [FileName]Hasse.txt
    [FileName]Hasse.gif
    [FileName]FTSum.txt
    [FileName]CompactSum.txt
    [FileName]TOrder.txt
    History.txt / FullHistory.txt
    MostRecentRankingValues.txt / MostRecentWeights.txt
```

---

## Settings Persistence

User preferences are saved to `OTSoftRememberUserChoices.txt` and restored on next launch. This includes menu checkmark states and dialog values.
