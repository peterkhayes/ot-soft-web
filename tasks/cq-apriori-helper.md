# Extract apriori parsing helper in lib.rs

The pattern of parsing `apriori_text` and building the apriori table is repeated across multiple WASM entry points in `lib.rs`. Extract into a shared helper function.
