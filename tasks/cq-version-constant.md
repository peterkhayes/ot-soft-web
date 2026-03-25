# Extract version string into a constant

The version string `"OTSoft 2.7, release date 2/1/2026"` is duplicated across `lib.rs`, `rcd.rs`, `maxent.rs`, `nhg.rs`, `gla.rs`, and `factorial_typology.rs`. Extract it into a single constant defined once and referenced everywhere.
