import { useState } from 'react'

/** Persist a single scalar value to localStorage with an optional validator. */
export function useLocalStorageValue<T>(
  key: string,
  defaultValue: T,
  isValid?: (v: unknown) => boolean,
): [T, (v: T) => void] {
  function read(): T {
    try {
      const raw = localStorage.getItem(key)
      if (raw === null) return defaultValue
      const parsed = JSON.parse(raw) as unknown
      if (typeof parsed !== typeof defaultValue) return defaultValue
      if (isValid && !isValid(parsed)) return defaultValue
      return parsed as T
    } catch {
      return defaultValue
    }
  }

  const [value, setValue] = useState<T>(read)

  function set(v: T) {
    setValue(v)
    try {
      localStorage.setItem(key, JSON.stringify(v))
    } catch {
      /* quota exceeded */
    }
  }

  return [value, set]
}

export function useLocalStorage<T extends object>(
  key: string,
  defaults: T,
): [T, (updater: Partial<T> | ((prev: T) => Partial<T>)) => void] {
  function read(): T {
    try {
      const raw = localStorage.getItem(key)
      if (raw === null) return defaults
      const parsed = JSON.parse(raw) as Record<string, unknown>
      if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) return defaults
      const result = { ...defaults }
      const d = defaults as Record<string, unknown>
      const r = result as Record<string, unknown>
      for (const k of Object.keys(defaults)) {
        if (k in parsed && typeof parsed[k] === typeof d[k]) {
          r[k] = parsed[k]
        }
      }
      return result
    } catch {
      return defaults
    }
  }

  const [value, setValue] = useState<T>(read)

  function set(updater: Partial<T> | ((prev: T) => Partial<T>)) {
    setValue((prev) => {
      const next = { ...prev, ...(typeof updater === 'function' ? updater(prev) : updater) }
      try {
        localStorage.setItem(key, JSON.stringify(next))
      } catch {
        /* quota exceeded */
      }
      return next
    })
  }

  return [value, set]
}
