# `rcd_apriori` — Missing "A Priori Rankings" section + wrong section numbers + incomplete mass deletion message + wrong constraint necessity ordering

**Symptom:** VB6 output has "2. A Priori Rankings" between Result and Tableaux, shifting all subsequent section numbers up by 1. Our output skips this section entirely.

This is a multi-part fix:

## A. Missing "A Priori Rankings" section

VB6 `Main.frm:7004-7040` (`PrintOutTheAprioriRankings`) prints a table showing which constraints dominate others. The table uses `s.PrintTable` (`s.bas:244`) with `CenterCells=True`, which calls `PrintTextTable` (`s.bas:373`) using per-column widths and `CenteredFillout` (`s.bas:430`) for alignment (centered with leftward error for odd spacing). The VB6 array stores data as `Table(col, row)` (reversed indices, per `s.bas:264`), with axis-switching: `gAPrioriRankingsTable(i,j)=True` (i dominates j) places "yes" at `Table(j+1, i+1)` = row i, col j in the output. Our `apriori.rs` stores `table[i][j]=true` meaning i dominates j, so display should use `apriori[row][col]`. Note: `gla.rs:458` already has an a priori table formatter but uses `apriori[col][row]` (inverted) — that may be a bug in GLA (no golden file to test against yet).

## B. Dynamic section numbering

VB6 uses a global auto-incrementing counter `gLevel1HeadingNumber` (`Module1.bas:175`). Our code hardcodes section numbers (1, 2, 3, 4, 5) in `rcd.rs` and `fred.rs:463`. When a priori is present, sections shift: 1=Result, 2=A Priori Rankings, 3=Tableaux, 4=Status, 5=FRed, 6=Mini-Tableaux. Need a mutable counter, and `fred.rs:format_section4` must accept the section number as a parameter.

## C. Mass deletion message — missing "individual but not mass" case

VB6 `Main.frm:6347-6358` has two messages gated on `NumberOfDeletableConstraints >= 2`: if mass deletion succeeds, print success message; otherwise, print "although the grammar will still work with the removal of ANY ONE... will NOT work if they are removed en masse." `NumberOfDeletableConstraints` counts constraints where `TrulyNeeded=False` (`Main.frm:6183-6188`). Our code (`rcd.rs:463-470`) only handles the success case and emits `\n\n` for the failure case — it doesn't print the "individual but not mass" message.

## D. Constraint necessity ordering

VB6 `Main.frm:6287-6341` outputs constraints grouped by category in three separate loops: first all Necessary constraints (in original order), then all "Not necessary (but included to show Faithfulness violations)" (in order), then all "Not necessary" (in order). Our code (`rcd.rs:447`) iterates by constraint index, which produces a different order when categories are interleaved.
