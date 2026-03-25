# Extract shared DownloadButton component

Every panel (Rcd, MaxEnt, Gla, Nhg, FactorialTypology) has:
1. Identical download button SVG icon markup (duplicated 5+ times)
2. Near-identical `handleDownload` try/catch/alert error handling patterns (9 instances)

Extract a shared `DownloadButton` component and a `downloadWithErrorHandling` utility.
