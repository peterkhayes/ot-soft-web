---
status: done
---

# Break up long formatting functions

Several functions exceed 200 lines and handle multiple distinct sections:

- `rcd.rs:format_output_with_options` (~265 lines)
- `rcd.rs:format_html_output_full` (~235 lines)
- `lfcd.rs:run_lfcd_with_apriori` (~325 lines)
- `bcd.rs:run_bcd` (~250 lines)
- `gla.rs:run_gla_with_schedule` (~405 lines)

Extract logical sub-sections into helper functions. This is a larger refactor and can be done incrementally.
