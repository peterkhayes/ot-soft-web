// In Vitest browser mode setupFiles, lifecycle globals (beforeEach etc.) are
// injected automatically — do not import them from 'vitest'.
beforeEach(() => {
  localStorage.clear()
})
