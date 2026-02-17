# Web Interface

This directory contains the presentation layer for OT-Soft, built with React + TypeScript + Vite.

## Principles

### Presentation Logic Only

The Web code should contain:
- React components and JSX
- CSS styling
- User interaction handling
- Formatting data for display

The Web code should NOT contain:
- Parsing logic
- OT algorithms
- Data validation
- Computational logic

**All computational logic belongs in the Rust codebase.**

## File Structure

```
web/
├── index.html              # Minimal shell (fonts + root div)
├── style.css               # All styling (unchanged from pre-React)
├── package.json            # npm dependencies
├── tsconfig.json           # TypeScript config
├── vite.config.ts          # Vite config (WASM + React plugins)
├── src/
│   ├── main.tsx            # Entry point (init React, import CSS)
│   ├── App.tsx             # Layout + state management + WASM init
│   ├── constants.ts        # TINY_EXAMPLE data
│   ├── vite-env.d.ts       # Vite type declarations
│   └── components/
│       ├── InputPanel.tsx   # File upload + example button
│       ├── TableauPanel.tsx # Tableau table display
│       └── RcdPanel.tsx     # RCD analysis + results + download
└── pkg/                    # WASM output (from wasm-pack, unchanged)
```

## Key Patterns

- WASM is initialized in `App.tsx` via dynamic import + `useEffect`
- WASM types from `pkg/ot_soft.d.ts` are used directly (type-safe)
- State lives in `App.tsx` and flows down as props
- `style.css` is imported globally in `main.tsx` — all CSS class names are unchanged

## Development

- `make dev` — Build WASM then start Vite dev server
- `make serve` — Start Vite dev server only (if WASM already built)
- `make web-build` — Production build of the web frontend

## Interaction Flow

1. User loads a file or clicks "Load Example Tableau"
2. `InputPanel` calls `parse_tableau(text)` from WASM
3. Result propagates via `onTableauLoaded` callback to `App.tsx` state
4. `TableauPanel` renders the `Tableau` object as a table
5. `RcdPanel` runs RCD analysis and displays strata
