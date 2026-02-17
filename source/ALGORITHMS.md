# OTSoft Algorithms

This document describes all algorithms implemented in OTSoft, based on the VB6 source code. OTSoft implements two categories of algorithms: **categorical ranking** algorithms (for Classical OT) and **probabilistic learning** algorithms (for Stochastic OT, Maximum Entropy, and Noisy Harmonic Grammar).

## Table of Contents

- [Categorical Ranking Algorithms](#categorical-ranking-algorithms)
  - [Recursive Constraint Demotion (RCD)](#recursive-constraint-demotion-rcd)
  - [Biased Constraint Demotion (BCD)](#biased-constraint-demotion-bcd)
  - [Low Faithfulness Constraint Demotion (LFCD)](#low-faithfulness-constraint-demotion-lfcd)
  - [FRed (Ranking Argumentation)](#fred-ranking-argumentation)
- [Probabilistic Learning Algorithms](#probabilistic-learning-algorithms)
  - [Gradual Learning Algorithm (GLA)](#gradual-learning-algorithm-gla)
  - [Maximum Entropy (Batch)](#maximum-entropy-batch)
  - [Noisy Harmonic Grammar (NHG)](#noisy-harmonic-grammar-nhg)
- [Factorial Typology](#factorial-typology)

---

## Categorical Ranking Algorithms

These algorithms find a stratified ranking of constraints that derives all observed winners. They operate on the Classical OT framework where constraints are strictly ordered.

### Common Concepts

**Strata**: Constraints are grouped into ranked strata. All constraints within a stratum are unordered relative to each other but outrank all constraints in lower strata.

**Demotability**: A constraint is *demotable* if it prefers at least one loser (rival) over the observed winner — i.e., the winner violates it more than the rival does.

**Still Informative**: A winner-rival pair is "still informative" if no already-ranked constraint rules out the rival. Once a higher-ranked constraint prefers the winner over a rival, that pair becomes uninformative.

**Termination conditions**:
- **Success**: All remaining unranked constraints are non-demotable → install them in the current stratum, done.
- **Failure**: All remaining unranked constraints are demotable → no valid ranking exists.
- **Continue**: Some demotable, some not → install non-demotable ones, continue to next stratum.

---

### Recursive Constraint Demotion (RCD)

**Source**: `RecursiveConstraintDemotion.bas`

The classic batch constraint demotion algorithm (Tesar & Smolensky 1993).

#### Algorithm

```
CurrentStratum = 0
Mark all constraints as unranked (Stratum[i] = 0)
Mark all winner-rival pairs as still informative

Loop:
    CurrentStratum += 1

    1. AVOID PREFERENCE FOR LOSERS
       For each still-informative winner-rival pair:
           For each unranked constraint C:
               If Winner violates C more than Rival:
                   Mark C as Demotable

    2. ENFORCE A PRIORI RANKINGS (if enabled)
       For each unranked constraint C1:
           For each C2 that C1 must dominate (per a priori table):
               Mark C2 as Demotable

    3. CHECK TERMINATION
       - If no constraints are demotable: SUCCESS (install all remaining)
       - If all constraints are demotable: FAILURE
       - Otherwise: Install non-demotable constraints in CurrentStratum

    4. UPDATE INFORMATIVENESS
       For each newly ranked constraint:
           For each winner-rival pair:
               If rival violates this constraint more than winner:
                   Mark pair as no longer informative

    Repeat loop
```

#### Output

- Returns: convergence status + number of strata + constraint-to-stratum mapping
- Writes `HowIRanked[filename].txt` documenting each step

---

### Biased Constraint Demotion (BCD)

**Source**: `BCD.bas`
**Author**: Programmed by Bruce Tesar
**Reference**: Prince & Tesar (2004)

BCD extends RCD with a bias toward ranking Faithfulness constraints low. It uses several heuristics to delay promoting Faithfulness constraints.

#### Algorithm

```
CurrentStratum = 0
Mark all constraints as unranked
Detect which constraints are Faithfulness vs Markedness

Loop:
    CurrentStratum += 1

    1. AVOID PREFERENCE FOR LOSERS (same as RCD)
       Mark demotable and active constraints.
       Active = prefers winner for at least one informative pair.

    2. FAITHFULNESS DELAY
       If any Markedness constraint is non-demotable:
           Install all non-demotable Markedness constraints
           Exclude ALL Faithfulness constraints from this stratum
           Skip to step 6

    3. AVOID THE INACTIVE
       Among Faithfulness constraints:
           If some are active: exclude all inactive ones
           If none are active: install all remaining (terminal case)

    4. FAVOR SPECIFICITY (if mnuSpecificBCD enabled)
       Compute violation subsets:
           C1 is a subset of C2 if:
             - For all forms/rivals: Violations(C1) ≤ Violations(C2)
             - C1 has at least one violation somewhere
       If C2's violations ⊆ C1's violations:
           Mark C1 as Subsetted (cannot join this stratum)

    5. FIND MINIMAL FAITHFULNESS SET
       Build candidate set: active, non-demotable, non-subsetted Faithfulness
       Try subsets of increasing size (1, 2, 3...):
           For each subset:
               Simulate ranking it
               Count how many Markedness constraints it "releases"
           Keep the subset that releases the most
           (Ties broken arbitrarily — first found wins; user warned)

    6. CHECK TERMINATION (same as RCD)

    7. UPDATE INFORMATIVENESS (same as RCD)

    Repeat loop
```

#### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `mnuSpecificBCD` | `False` | Enable specificity-based ranking priority |

#### Notes

- BCD does **not** support a priori rankings (shows warning if attempted)
- When ties occur during subset selection, a warning dialog is displayed

---

### Low Faithfulness Constraint Demotion (LFCD)

**Source**: `LowFaithfulnessConstraintDemotion.bas`
**Author**: Bruce Hayes

LFCD is Hayes's algorithm that also biases toward low Faithfulness rankings, using different heuristics than BCD. It introduces the concepts of "autonomy" (helper counting) and superset filtering.

#### Algorithm

```
CurrentStratum = 0
Compute violation subsets
Initialize NumberOfHelpers[i] = 10000 for all constraints

Loop:
    CurrentStratum += 1

    1. AVOID PREFERENCE FOR LOSERS (same as RCD)

    2. ENFORCE A PRIORI RANKINGS (if enabled)

    3. FAVOR MARKEDNESS
       If any Markedness constraint is non-demotable:
           Install it, exclude all Faithfulness
           Skip to step 7

    4. FAVOR ACTIVENESS
       For each Faithfulness constraint:
           Active = prefers winner for at least one non-superset rival
       "Non-superset" means: the rival does NOT have ≥ winner violations
       for every constraint (such rivals lose under any ranking)
       If some are active: exclude all inactive
       If none active: install all remaining non-demotable (terminal)

    5. FAVOR SPECIFICITY
       For each unranked, active, non-demotable Faithfulness C:
           If a more specific version exists and is unranked:
               Exclude the more general one

    6. FAVOR AUTONOMY
       For each still-informative, non-superset winner-rival pair:
           For each Faithfulness constraint C that prefers winner:
               Count "helpers" = other constraints also preferring winner
               (Superset Faithfulness constraints don't count as helpers)
               Record minimum helper count for C
       Find LowestNumberOfHelpers across all candidates
       Install only constraints with this lowest count
       Exclude constraints with more helpers

    7. CHECK TERMINATION (same as RCD)

    8. UPDATE INFORMATIVENESS (same as RCD)

    Repeat loop
```

#### Key Concept: Superset Rivals

A rival is a *superset* of the winner if `WinnerViolations(C) ≤ RivalViolations(C)` for ALL constraints. Such rivals lose under any ranking and are excluded from activeness and autonomy calculations.

---

### FRed (Ranking Argumentation)

**Source**: `Fred.bas`
**Reference**: Prince & Brasoveanu (2005), FRed = Fusional Reduction

FRed analyzes the ranking information in a dataset by computing a basis of Elementary Ranking Conditions (ERCs). It does not produce a single ranking but rather the set of ranking arguments supported by the data.

#### ERCs (Elementary Ranking Conditions)

Each winner-rival pair generates an ERC: a string over {W, L, e} where:
- **W**: constraint prefers winner (winner has fewer violations)
- **L**: constraint prefers loser/rival (winner has more violations)
- **e**: constraint is neutral (equal violations)

ERC status:
- **Valid**: has at least one W and one L
- **Uninformative**: has W's and e's only (rival is harmonically bounded)
- **Unsatisfiable**: has L's and e's only (winner cannot be derived)
- **Duplicate**: all e's (identical violation profiles)

#### Fusion Operation

ERCs are combined using fusion algebra:
- `e ⊕ X = X` (for any X)
- `W ⊕ W = W`
- `W ⊕ L = L`
- `L ⊕ X = L` (for any X)

Fusion summarizes the combined ranking requirement of a set of ERCs.

#### Algorithm

```
1. BUILD ORIGINAL ERC SET
   - Add a priori ranking ERCs (if any)
   - For each winner-rival pair, compute ERC
   - Filter: keep only Valid ERCs, report others
   - Deduplicate

2. RECURSIVE SEARCH (RecursiveRoutine)
   Input: a set of ERCs

   a. Compute Fusion of the ERC set
      - If fusion has L but no W: UNRANKABLE → store, return
      - If fusion has no L: TRIVIALLY SATISFIED → store, return

   b. Compute Total Information-Loss Residue
      For each constraint column:
          If column has only W's and e's (no L):
              Mark ERCs with 'e' in that column as "residue"
      Fuse all marked residue ERCs

   c. ENTAILMENT CHECK
      Skeletal Basis mode:
          SkeletalBasis = keep Fusion positions, replace with 'e'
              where residue has 'L'
          If result has no L's: entailed → discard
          Else: not entailed → add to final basis
      Most Informative Basis mode:
          Check if residue fusion entails the full fusion
          If yes: entailed → discard
          Else: not entailed → add to final basis

   d. FORM SEND-ON SETS (for recursion)
      For each constraint column:
          If column has no L's, at least one W, at least one e:
              Create subset = ERCs with 'e' in that column
              If subset is novel (not yet explored):
                  Recursively process this subset

3. REPORT RESULTS
   - List final basis ERCs
   - Convert to ranking statements:
     - Single W: "C_W >> {C_L1, C_L2, ...}"
     - Multiple W's: "At least one of {C_W1, ...} >> {C_L1, ...}"
```

#### Entailment

ERC1 *entails* ERC2 if, position by position:
- W entails only W
- e entails W and e
- L entails anything

#### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `chkMostInformativeBasis` | `True` | Use MIB (False = Skeletal Basis) |
| `chkDetailedArguments` | `False` | Include verbose recursive search tree |
| `chkMiniTableaux` | varies | Include mini-tableaux for pairwise arguments |

#### Hasse Diagram

FRed generates a GraphViz DOT file (`[filename]Hasse.txt`) for visualizing the ranking:
- Solid lines = certain rankings
- Dotted "or" lines = disjunctive rankings (multiple W's in an ERC)

---

## Probabilistic Learning Algorithms

These algorithms learn numeric weights or ranking values from frequency data. They are error-driven online learners (GLA, NHG) or batch optimizers (MaxEnt).

### Common Concepts

**Learning schedule**: All online algorithms use a multi-stage plasticity schedule. By default, 4 stages with geometrically interpolated plasticity values from initial (high) to final (low). A custom schedule can be loaded from `CustomLearningSchedule.txt` with columns: Trials, PlastMark, PlastFaith, NoiseMark, NoiseFaith.

**Exemplar selection**: Two modes:
- **Stochastic**: Random selection weighted by candidate frequencies
- **Exact proportions**: Deterministic presentation in exact input frequencies (requires integer frequencies). A shuffled array is created and processed sequentially.

**Initial values**: Can be:
- **AllSame**: Default values (100 for Stochastic OT, 0 for MaxEnt/NHG)
- **MarkednessFaithfulness**: Separate values for markedness vs faithfulness
- **FullyCustomized**: Read from `ModelParameters.txt`
- **ValuesFromPreviousRun**: Read from `MostRecentRankingValues.txt`

---

### Gradual Learning Algorithm (GLA)

**Source**: `boersma.frm`
**Reference**: Boersma (1997), Boersma & Hayes (2001)

The GLA operates in two modes: **Stochastic OT** and **online MaxEnt**.

#### Parameters

| Parameter | Default (Stochastic OT) | Default (MaxEnt) | Description |
|-----------|------------------------|-------------------|-------------|
| Number of cycles | 1,000,000 | 1,000,000 | Total training iterations |
| Initial plasticity | 2 | 2 | Starting learning rate |
| Final plasticity | 0.001 | 0.001 | Ending learning rate |
| Times to test grammar | 10,000 | 10,000 | Evaluation trials after learning |
| A priori ranking difference | 20 | 20 | Minimum numeric gap for a priori rankings |
| Initial ranking values | 100 (all) | 0 (all) | Starting constraint values |

#### Algorithm

```
For each learning stage (1 to 4):
    Set plasticity for this stage (geometric interpolation)

    For each cycle in this stage:
        1. SELECT EXEMPLAR from training data

        2. GENERATE A FORM using current grammar:

           Stochastic OT mode:
               Add Gaussian noise to each ranking value
               Sort constraints by noisy values (highest first)
               Evaluate candidates using strict domination (OT eval)
               Winner = candidate with best harmony

           MaxEnt mode:
               For each candidate:
                   H = Σ(weight[c] × violations[c])
                   P(candidate) = exp(-H) / Z,  where Z = Σ exp(-H)
               Sample a candidate according to probabilities

        3. COMPARE generated form with selected exemplar

        4. UPDATE WEIGHTS (if generated ≠ observed):
           For each constraint C:
               If winner_violations[C] ≠ loser_violations[C]:
                   delta = plasticity × (loser_viols - winner_viols)
                   ranking_value[C] += delta
                   (PlastMark for Markedness, PlastFaith for Faithfulness)

        5. ENFORCE A PRIORI RANKINGS (if active):
           Ensure a priori-ranked constraints differ by ≥ 20 units

        6. RECORD HISTORY (if requested)
```

#### Variants

| Option | Description |
|--------|-------------|
| Magri update rule | Alternative promotion calculation for Stochastic OT |
| Gaussian prior | Biased learning for MaxEnt (reads mu/sigma from `ModelParameters.txt`) |
| Negative weights OK | Allow weights below 0 (MaxEnt) |
| Multiple runs | Run 10/100/1000 times and collate results |
| Pairwise ranking probabilities | Calculate probability of each constraint pair ordering |

---

### Maximum Entropy (Batch)

**Source**: `MyMaxEnt.frm`
**Reference**: Goldwater & Johnson (2003), using Generalized Iterative Scaling (Goodman 2002)

Unlike the online GLA-MaxEnt, this is a batch optimizer that processes all data simultaneously per iteration.

#### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| Number of iterations | 5 | GIS iterations (higher = more precise) |
| Weight minimum | 0 | Lower bound on weights |
| Weight maximum | 50 | Upper bound on weights |
| Decimal places | 3 | Output precision |
| Sigma squared | 1 | Gaussian prior strength (hardcoded) |
| Using prior term | False | Enable Gaussian prior regularization |

#### Algorithm

```
Initialize all weights to 0

For each iteration (1 to N):

    1. CALCULATE PREDICTED PROPORTIONS
       For each input form:
           For each candidate:
               harmony = Σ(weight[c] × violations[c])
               eHarmony = exp(-harmony)
           Z = Σ(eHarmony) over all candidates
           predicted_proportion = eHarmony / Z

    2. CALCULATE EXPECTED VIOLATIONS
       For each constraint C:
           expected[C] = Σ over all inputs, candidates:
               total_freq[input] × predicted_proportion × violations[C]

    3. CALCULATE OBSERVED VIOLATIONS
       For each constraint C:
           observed[C] = Σ over all inputs, candidates:
               frequency[candidate] × violations[C]
       (Zeros replaced with 0.000000001 to avoid log(0))

    4. CALCULATE SLOWING FACTOR
       = max total violations across any single candidate
       (Required for GIS convergence)

    5. UPDATE WEIGHTS (Generalized Iterative Scaling)
       For each constraint C:
           Without prior:
               delta = log(observed[C] / expected[C]) / slowing_factor
               weight[C] -= delta

           With Gaussian prior (Newton's method):
               Solve: 0 = expected[C] × exp(delta × SF) +
                      (weight[C] + delta) / σ² - observed[C]
               weight[C] -= delta / 1000

       Enforce bounds: weight_min ≤ weight[C] ≤ weight_max

    6. REPORT PROGRESS (every 10,000 iterations)
       Log probability, gradient slope, objective function
```

#### Log Probability of Data

```
For each input:
    For each candidate:
        log_prob = log(exp(-harmony) / Z)
total = Σ(log_prob × frequency)
```

---

### Noisy Harmonic Grammar (NHG)

**Source**: `NoisyHarmonicGrammar.frm`
**Reference**: Boersma & Pater (for exponential variant)

NHG learns numeric weights where the winner is determined by harmony (weighted sum of violations) with added Gaussian noise.

#### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| Number of cycles | 5,000 | Total training iterations |
| Initial plasticity | 2 | Starting learning rate |
| Final plasticity | 0.002 | Ending learning rate |
| Times to test grammar | 2,000 | Evaluation trials after learning |
| Noise | 1.0 (normal) / 0.1 (exponential) | Standard deviation of Gaussian noise |
| Initial weights | 0 (all) | Starting constraint weights |

#### Noise Variants

NHG implements **8 different noise configurations** controlled by checkboxes:

| Option | Description |
|--------|-------------|
| Noise by constraint (default) | Same noise value applied to all cells in a constraint column |
| Noise by cell | Independent noise for each violation cell |
| Pre-multiplication noise (default) | Noise added to weight before multiplying by violations |
| Post-multiplication noise | Noise added to product of weight × violations |
| Noise for zero cells | Apply noise even when violation count is 0 |
| Late noise | Single noise term added after computing total harmony |
| Exponential NHG | Use `exp(weight)` instead of raw weight |
| Demi-Gaussians | Use positive-only half-Gaussian noise |

#### Algorithm

```
For each learning stage (1 to 4):
    Set plasticity for this stage

    For each cycle:
        1. SELECT EXEMPLAR

        2. GENERATE A FORM
           (See noise variant details below)
           Find candidate with lowest noisy harmony
           Handle ties: skip trial OR pick randomly

        3. UPDATE WEIGHTS (if generated ≠ observed):
           For each constraint C:
               If winner_viols[C] ≠ training_viols[C]:
                   weight[C] += plasticity × (winner_viols - training_viols)
                   (PlastMark for Markedness, PlastFaith for Faithfulness)

               If weight[C] < 0 AND NOT allow_negative:
                   weight[C] = 0  (unless exponential NHG)

        4. ENFORCE A PRIORI RANKINGS (if active)
```

#### Noise Variant Details

**Variant A (default)**: Pre-multiplication, by constraint
```
For each constraint C:
    noisy_weight[C] = weight[C] + noise × Gaussian()
    if noisy_weight < 0 and not allow_negative: noisy_weight = 0
For each candidate:
    harmony = Σ(noisy_weight[C] × violations[C])
```

**Variant A' (exponential)**: Same as A but `harmony = Σ(exp(noisy_weight[C]) × violations[C])`

**Variant B**: Pre-multiplication, by cell
```
For each candidate:
    For each constraint C:
        noisy_weight = weight[C] + Gaussian()
        harmony += noisy_weight × violations[C]
```

**Variant C**: Post-multiplication, noise for zero cells
```
For each candidate:
    For each constraint C:
        perturbation = Gaussian()
        if violations[C] = 0:
            harmony += perturbation  (noise even for zero)
        else:
            harmony += weight[C] × violations[C] + perturbation
```

**Variant D**: Post-multiplication, no noise for zero cells
```
Same as C, but zero-violation cells contribute 0 (no noise)
```

**Late noise**: Single noise term added to total harmony
```
For each candidate:
    harmony = Σ(weight[C] × violations[C]) + Gaussian()
```

#### Tie Handling

When multiple candidates have the same best harmony:
- **Skip** (`mnuResolveTiesBySkipping`): discard this trial, no weight update
- **Random** (default): pick randomly among tied candidates

---

## Factorial Typology

**Source**: `FactorialTypology.bas`

Computes the factorial typology: the set of all possible output patterns derivable by some ranking of the constraints.

### Algorithm

```
1. PREPROCESSING

   a. Install winners as candidates
      (Add winner to rival list as first candidate)

   b. Hunt for duplicate violations
      If two candidates have identical violations: offer to merge them

   c. Find possible outcomes (pre-filter)
      For each input form:
          For each candidate:
              Test if derivable using FastRCD
              If not derivable under any ranking: move to "permanent loser" list
      (This drastically reduces the search space)

2. INITIALIZE VALHALLA
   Valhalla = set of valid output patterns
   For first input form: Valhalla = all remaining candidates

3. INCREMENTAL CONSTRUCTION
   For each subsequent input form (2 to N):

       OldValhalla = current Valhalla

       For each pattern in OldValhalla:
           For each candidate of the new input form:
               Package test dataset:
                   - Winners = pattern's outputs for forms 1..N-1,
                     plus current candidate for form N
                   - Rivals = all other candidates
               Run FastRCD
               If succeeds: add extended pattern to NewValhalla

       Valhalla = NewValhalla

4. RESTORE permanent losers (for display)

5. OUTPUT GENERATION
```

### FastRCD

A streamlined version of RCD used internally for testing candidate derivability. Returns only True/False without documentation output. Also has a variant that respects a priori rankings (`FastRCDWithAPrioriRankings`).

### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| Include ranking in results | varies | Generate full listing with grammars per pattern |
| Include tableaux | varies | Add tableaux to full listing |
| FT Sum file | varies | Generate `FTSum.txt` |
| Compact FT file | varies | Generate `CompactSum.txt` |
| Include FRed arguments | varies | Run FRed for each output pattern |

### T-Order

The t-order (typological order) is the set of implications in the factorial typology:

```
For each input-output pair (I1, O1):
    For each other input I2:
        Check if I1→O1 uniquely determines I2's output
        If yes: record "If I1→O1 then I2→O2"

Also identifies:
    - Candidates that always win (trivially implied)
    - Candidates that never implicate anything
```
