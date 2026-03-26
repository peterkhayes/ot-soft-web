/**
 * Generic discriminated union for algorithm result states.
 * Success branch: T with `error?: undefined`.
 * Error branch: `{ error: string }` with all T keys as `?: undefined`.
 * Discriminate with `result.error` — falsy means success.
 */
export type ResultState<T> =
  | (T & { error?: undefined })
  | ({ error: string } & { [K in keyof T]?: undefined })
