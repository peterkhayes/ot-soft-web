# Pass FredOptions struct to apply_fred_options

`apply_fred_options` in `rcd.rs` takes 4 individual boolean parameters. It's called 6 times in `lib.rs` each with `fred_opts.include_fred, fred_opts.use_mib, fred_opts.show_details, fred_opts.include_mini_tableaux`. Change the method to accept `&FredOptions` directly to reduce repetition and improve readability.
