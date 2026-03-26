---
status: done
---

# Add aria-labels to interactive SVG elements

Several interactive elements lack accessibility attributes:
- Download buttons in `HasseDiagram.tsx` have no `aria-label`
- Algorithm select in `RcdPanel.tsx` has no `aria-label`
- Decorative SVG icons across all panels lack `aria-hidden`
- Drag-and-drop area in `InputPanel.tsx` lacks `aria-label`
