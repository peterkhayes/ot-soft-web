import type { ResultState } from '../../types.ts'
import type { GlaDefaults } from '../../wasmDefaults.ts'

export interface GlaResultState {
  values: { fullName: string; abbrev: string; value: number; active: boolean }[]
  forms: {
    input: string
    candidates: { form: string; obsPct: number; genPct: number }[]
  }[]
  logLikelihood: number
  maxentMode: boolean
  hasseDot?: string
  pairwiseData?: { headers: string[]; matrix: string[][] }
  history?: string
  fullHistory?: string
  candidateProbHistory?: string
}

export type GlaState = ResultState<GlaResultState>

export type GlaParams = GlaDefaults

export type SetGlaParams = (update: Partial<GlaParams>) => void
