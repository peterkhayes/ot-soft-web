# Conformance Test: VB6 Golden File Collection Checklist

This checklist guides you through running each test case in VB6 OTSoft on Windows and saving the output files for conformance testing.

## Setup

1. Open OTSoft 2.7 on Windows
2. Copy `examples/tiny/input.txt` to your Windows machine and rename it to **`TinyIllustrativeFile.txt`** (VB6 embeds the filename in output headers)
3. Copy `examples/tiny/apriori.txt` to your Windows machine (keep the name `apriori.txt`)

## Output collection

For each test case, save the **draft output** text file to the path shown. All paths are relative to the repository root.

### HTML output collection

Some test cases also require the **HTML tableau output** (the `.htm` file VB6 generates alongside the draft text). To collect these:

1. In VB6 OTSoft, ensure **HTML output** is enabled (check the menu option)
2. Set the shading grayness to the default value
3. Run the algorithm — VB6 writes a `.htm` file to the same output directory
4. Save the `.htm` file to the golden path shown (e.g., `conformance/golden/tiny/rcd_defaults.htm`)

The HTML conformance test compares tableaux **semantically** (cell content + CSS shading classes), not byte-for-byte, so minor formatting differences between VB6 versions are tolerated.

---

## RCD Tests

### tiny_rcd_defaults
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Recursive Constraint Demotion** (RCD)
- [ ] Use default settings: FRed enabled, Skeletal Basis, show details, mini-tableaux
- [ ] No a priori rankings
- [ ] Run and save output → `conformance/golden/tiny/rcd_defaults.txt`
- [ ] Also save the HTML output → `conformance/golden/tiny/rcd_defaults.htm`

**Note:** The text file is already collected. Verify it matches your VB6 version.

### tiny_rcd_no_fred
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **RCD**
- [ ] **Disable** FRed (uncheck "Ranking Arguments" / FRed option)
- [ ] No a priori rankings
- [ ] Run and save output → `conformance/golden/tiny/rcd_no_fred.txt`

### tiny_rcd_mib
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **RCD**
- [ ] Enable FRed with **Most Informative Basis** (MIB) instead of Skeletal Basis
- [ ] Show details, mini-tableaux enabled
- [ ] No a priori rankings
- [ ] Run and save output → `conformance/golden/tiny/rcd_mib.txt`

### tiny_rcd_apriori
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **RCD**
- [ ] Default FRed settings (Skeletal Basis)
- [ ] Load a priori rankings from `apriori.txt` (*NoOns >> Max)
- [ ] Run and save output → `conformance/golden/tiny/rcd_apriori.txt`

---

## BCD Tests

### tiny_bcd_defaults
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Biased Constraint Demotion** (BCD)
- [ ] Use default settings (non-specific)
- [ ] FRed enabled, Skeletal Basis, show details, mini-tableaux
- [ ] Run and save output → `conformance/golden/tiny/bcd_defaults.txt`
- [ ] Also save the HTML output → `conformance/golden/tiny/bcd_defaults.htm`

### tiny_bcd_specific
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **BCD**
- [ ] Enable **Favor Specificity** heuristic
- [ ] FRed enabled, Skeletal Basis, show details, mini-tableaux
- [ ] Run and save output → `conformance/golden/tiny/bcd_specific.txt`

---

## LFCD Tests

### tiny_lfcd_defaults
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Low Faithfulness Constraint Demotion** (LFCD)
- [ ] Default settings, FRed enabled, Skeletal Basis
- [ ] No a priori rankings
- [ ] Run and save output → `conformance/golden/tiny/lfcd_defaults.txt`
- [ ] Also save the HTML output → `conformance/golden/tiny/lfcd_defaults.htm`

### tiny_lfcd_apriori
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **LFCD**
- [ ] Default FRed settings (Skeletal Basis)
- [ ] Load a priori rankings from `apriori.txt` (*NoOns >> Max)
- [ ] Run and save output → `conformance/golden/tiny/lfcd_apriori.txt`

### tiny_lfcd_mib
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **LFCD**
- [ ] Enable FRed with **Most Informative Basis** (MIB)
- [ ] Show details, mini-tableaux enabled
- [ ] No a priori rankings
- [ ] Run and save output → `conformance/golden/tiny/lfcd_mib.txt`

---

## MaxEnt Tests

For all MaxEnt tests, set **Precision** (GIS iterations) to **5**.

### tiny_maxent_defaults
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Maximum Entropy**
- [ ] Set Precision to **5**
- [ ] Weight min = 0, weight max = 50
- [ ] Gaussian prior **disabled**
- [ ] Run and save output → `conformance/golden/tiny/maxent_defaults.txt`

### tiny_maxent_prior
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Maximum Entropy**
- [ ] Set Precision to **5**
- [ ] Weight min = 0, weight max = 50
- [ ] Enable **Gaussian prior** with sigma² = **1.0**
- [ ] Run and save output → `conformance/golden/tiny/maxent_prior.txt`

### tiny_maxent_sigma10
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Maximum Entropy**
- [ ] Set Precision to **5**
- [ ] Weight min = 0, weight max = 50
- [ ] Enable **Gaussian prior** with sigma² = **10.0**
- [ ] Run and save output → `conformance/golden/tiny/maxent_sigma10.txt`

---

## Factorial Typology Tests

### tiny_ft_defaults
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Factorial Typology**
- [ ] Default settings (no full listing)
- [ ] No a priori rankings
- [ ] Run and save output → `conformance/golden/tiny/ft_defaults.txt`

### tiny_ft_full_listing
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Factorial Typology**
- [ ] Enable **Full Listing** (detailed pattern derivations)
- [ ] No a priori rankings
- [ ] Run and save output → `conformance/golden/tiny/ft_full_listing.txt`

### tiny_ft_apriori
- [ ] Open `TinyIllustrativeFile.txt`
- [ ] Select **Factorial Typology**
- [ ] Default settings (no full listing)
- [ ] Load a priori rankings from `apriori.txt` (*NoOns >> Max)
- [ ] Run and save output → `conformance/golden/tiny/ft_apriori.txt`

---

## Ilokano Hiatus Resolution Tests

Use `examples/IlokanoHiatusResolution/input.txt`, renamed to **`IlokanoHiatusResolution.txt`** on Windows.

### ilokano_rcd_defaults (including HTML)
- [ ] Open `IlokanoHiatusResolution.txt`
- [ ] Select **RCD**, default settings (FRed enabled, Skeletal Basis, mini-tableaux)
- [ ] Run and save output → `conformance/golden/IlokanoHiatusResolution/rcd_defaults.txt`
- [ ] Also save the HTML output → `conformance/golden/IlokanoHiatusResolution/rcd_defaults.htm`
