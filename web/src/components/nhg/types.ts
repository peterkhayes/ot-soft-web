import type { ResultState } from '../../types.ts'
import type { NhgDefaults } from '../../wasmDefaults.ts'

export interface NhgResultState {
  weights: { fullName: string; abbrev: string; weight: number }[]
  forms: {
    input: string
    totalFreq: number
    candidates: {
      form: string
      frequency: number
      obsPct: number
      genCount: number
      genPct: number
    }[]
  }[]
  logLikelihood: number
  zeroPredictionWarning: boolean
  history?: string
  fullHistory?: string
}

export type NhgState = ResultState<NhgResultState>

export type NhgParams = NhgDefaults

export type SetNhgParams = (update: Partial<NhgParams>) => void
