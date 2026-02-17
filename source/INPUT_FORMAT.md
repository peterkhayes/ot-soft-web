# OTSoft Input File Format

This document describes the input file format accepted by OTSoft, based on the VB6 source code in `vb6/Main.frm`.

## Supported File Types

| Extension | Format | Notes |
|-----------|--------|-------|
| `.txt` | Tab-delimited text | Primary format, described below |
| `.xlsx` | Modern Excel | Read via Excel COM automation |
| `.xls` | Legacy Excel | **Not supported** (must convert to `.xlsx`) |
| `.in` | Traditional Ranker format | Legacy format from earlier software |

## Tab-Delimited Text Format

The primary input format is a tab-delimited text file. Blank lines are skipped.

### Structure

```
[Constraint Names]     [tab]  [tab]  [tab]  Con1Name  [tab]  Con2Name  [tab]  ...
[Abbreviations]        [tab]  [tab]  [tab]  Con1Abbr  [tab]  Con2Abbr  [tab]  ...
/input1/  [tab]  cand1  [tab]  freq  [tab]  viol1  [tab]  viol2  [tab]  ...
          [tab]  cand2  [tab]  freq  [tab]  viol1  [tab]  viol2  [tab]  ...
/input2/  [tab]  cand3  [tab]  freq  [tab]  viol1  [tab]  viol2  [tab]  ...
          [tab]  cand4  [tab]  freq  [tab]  viol1  [tab]  viol2  [tab]  ...
```

### Row Details

#### Row 1: Constraint Names (required)

- Columns 1-3 are ignored (can contain anything, typically left blank or used for labels)
- Column 4 onward: full constraint names, one per column
- The number of constraints is determined by this row

#### Row 2: Constraint Abbreviations (optional)

- Same column layout as Row 1
- If omitted, constraint names are used as abbreviations
- **Auto-detection**: The parser checks whether row 2 contains numeric values. If so, it treats row 2 as data (no abbreviation row). Otherwise, it treats it as abbreviations.
- Missing abbreviations at the end of the row are filled in by copying constraint names

#### Data Rows (Row 3+, or Row 2 if no abbreviation row)

| Column | Content | Notes |
|--------|---------|-------|
| 1 | Input form (e.g., `/tat/`) | Only present on the first candidate row for each input. Blank for subsequent candidates of the same input. |
| 2 | Candidate output (e.g., `ta`) | The surface form |
| 3 | Frequency | Numeric value. For categorical algorithms (RCD, BCD, LFCD), the highest-frequency candidate per input becomes the winner. For probabilistic algorithms, frequencies are used as training proportions. |
| 4+ | Violation counts | Integer violation counts for each constraint. Missing values default to 0. Can also contain structural descriptions (see below). |

### Example

From `TinyIllustrativeFile.txt`:

```
			*No Onset	*Coda	Max(t)	Dep(?)
			*NoOns	*Coda	Max	Dep
a	?a	1				1
	a		1
tat	ta	1			1
	tat			1
at	?a	1			1	1
	?at			1		1
	a		1		1
	at		1	1
```

This defines:
- 4 constraints: `*No Onset`, `*Coda`, `Max(t)`, `Dep(?)`
- 3 input forms: `a`, `tat`, `at`
- Input `a` has 2 candidates, `tat` has 2, `at` has 4
- Candidate `?a` for input `a` has frequency 1, making it the winner

### Winner Determination

For categorical algorithms (RCD, BCD, LFCD), the winner for each input is determined by:
1. The candidate with the highest frequency value
2. If multiple candidates share the highest frequency, the first one encountered is used

For probabilistic algorithms (GLA, MaxEnt, NHG), all candidates and their frequencies are used directly as training data proportions.

## Structural Descriptions

Instead of numeric violation counts, cells can contain **structural descriptions** — formal specifications that OTSoft uses to automatically compute violations. This is handled by `StructuralDescriptions.bas`.

Structural descriptions support:
- Markedness constraints: specified as phonological patterns that trigger violations
- Faithfulness constraints: require input-output correspondence; input and output must have equal length
- Natural class files: external files defining phonological classes, referenced in structural descriptions

When structural descriptions are present, OTSoft computes the violation counts automatically rather than reading them as integers.

## A Priori Rankings

A priori constraint rankings can be specified in a separate file. The system supports:

- **Template generation**: `PrintOutTemplateForAPrioriRankings()` creates a blank ranking template
- **Table format**: A matrix file where rows and columns are constraints, and cells indicate dominance relationships
- **Integration**: A priori rankings are converted to ERCs (Elementary Ranking Conditions) and added to the constraint set before running algorithms

The a priori ranking file format is a tab-delimited table:
- Row/column headers are constraint abbreviations
- Cell value `1` means the row constraint dominates the column constraint
- Cell value `0` or blank means no a priori relationship

## Constants and Limits

From `Module1.bas`:

| Constant | Value | Meaning |
|----------|-------|---------|
| Max constraints | ~1000 | Array dimension for constraints |
| Max rivals | ~1000 | Array dimension for rivals per input |
| Max forms | ~1000 | Array dimension for input forms |
