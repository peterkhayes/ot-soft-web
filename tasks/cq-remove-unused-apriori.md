# Remove unused _apriori parameter from apply_fred_options

`rcd.rs` line 224: `_apriori: &[Vec<bool>]` is unused (prefixed with `_`). Remove it from the function signature and all call sites in `lib.rs`.
