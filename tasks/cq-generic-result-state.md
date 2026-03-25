# Standardize result state types with a generic

`RcdResultState | RcdErrorState`, `MaxEntResultState | MaxEntErrorState`, `GlaResultState | GlaErrorState`, `NhgResultState | NhgErrorState` all follow the same pattern but are separately defined. Create a generic `ResultState<T>` type or similar to reduce boilerplate.
