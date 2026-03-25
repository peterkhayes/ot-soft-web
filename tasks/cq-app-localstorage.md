# Use useLocalStorage hook consistently in App.tsx

`App.tsx` manually reads/writes localStorage for `framework`, `axisMode`, and `sortByHarmony` instead of using the existing `useLocalStorage` hook that other components already use. Migrate to the hook for consistency.
