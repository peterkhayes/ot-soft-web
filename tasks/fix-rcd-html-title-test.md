---
status: done
---

# Fix RCD HTML tableaux download test

The test `RCD: download HTML tableaux` in `web/tests/flows/rcd.test.tsx` fails because it expects the HTML title to be `OTSoft 2.7 TinyIllustrativeFile.txt` but the actual output includes the release date: `OTSoft 2.7, release date 2/1/2026 TinyIllustrativeFile.txt`. Update the test assertion to match the current output format.
