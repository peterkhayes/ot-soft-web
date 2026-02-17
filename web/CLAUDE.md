# Web Interface

This directory contains the presentation layer for OT-Soft.

## Principles

### Presentation Logic Only

The Web code should contain:
- ✅ HTML structure
- ✅ CSS styling
- ✅ DOM manipulation
- ✅ User interaction handling
- ✅ Formatting data for display (e.g., building HTML tables from data structures)

The Web code should NOT contain:
- ❌ Parsing logic
- ❌ OT algorithms
- ❌ Data validation
- ❌ Computational logic

**All computational logic belongs in the Rust codebase.**

## Current Implementation

### Files

- `index.html` - Main page structure
- `style.css` - All styling, including Excel-like table appearance
- `main.js` - Application logic and data formatting
  - Loads WASM module
  - Handles file input
  - Calls Rust parsing functions
  - Formats Tableau objects into HTML tables
  - Displays results

### Interaction Flow

1. User loads a file or clicks "Load Tiny Example"
2. JavaScript calls `parse_tableau(text)` in Rust
3. Rust returns a `Tableau` object
4. JavaScript calls `formatTableauAsHTML(tableau)` to generate HTML
5. HTML is inserted into the DOM for display

This keeps Rust focused on data/logic and JavaScript focused on presentation.
