import type { ResultState } from '../../types.ts'
import type { MaxEntDefaults } from '../../wasmDefaults.ts'

export interface MaxEntResultState {
  weights: { abbrev: string; fullName: string; weight: number }[]
  forms: {
    input: string
    candidates: {
      form: string
      obsPct: number
      predPct: number
      violations: number[]
    }[]
  }[]
  logProb: number
  history?: string
  outputProbHistory?: string
}

export type MaxEntState = ResultState<MaxEntResultState>

export type MaxEntParams = MaxEntDefaults

export type SetMaxEntParams = (update: Partial<MaxEntParams>) => void
