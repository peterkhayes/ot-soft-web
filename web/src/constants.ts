// Tiny example tableau data (exact copy of examples/TinyIllustrativeFile/input.txt)
export const TINY_EXAMPLE = `\t\t\t*No Onset\t*Coda\tMax(t)\tDep(?)
\t\t\t*NoOns\t*Coda\tMax\tDep
a\t?a\t1\t\t\t\t1
\ta\t\t1
tat\tta\t1\t\t\t1
\ttat\t\t\t1
at\t?a\t1\t\t\t1\t1
\t?at\t\t\t1\t\t1
\ta\t\t1\t\t1
\tat\t\t1\t1
`

// Default 4-stage custom learning schedule template (matches Rust schedule::CUSTOM_SCHEDULE_TEMPLATE)
export const DEFAULT_SCHEDULE_TEMPLATE =
  'Trials\tPlastMark\tPlastFaith\tNoiseMark\tNoiseFaith\n' +
  '15000\t2\t2\t2\t2\n' +
  '15000\t0.2\t0.2\t2\t2\n' +
  '15000\t0.02\t0.02\t2\t2\n' +
  '15000\t0.002\t0.002\t2\t2'
