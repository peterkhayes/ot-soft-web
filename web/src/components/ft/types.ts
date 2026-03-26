import type { ResultState } from '../../types.ts'
import type { FtDefaults } from '../../wasmDefaults.ts'

export interface FtCandidate {
  form: string
  isWinner: boolean
  derivable: boolean
}

export interface FtWinnersRow {
  formInput: string
  candidates: FtCandidate[]
}

export interface FtTOrderEntry {
  implicatorInput: string
  implicatorCandidate: string
  implicatedInput: string
  implicatedCandidate: string
}

export interface FtResultData {
  patternCount: number
  constraintCount: number
  formInputs: string[]
  patterns: { candidates: string[]; isWinner: boolean[] }[]
  winners: FtWinnersRow[]
  torder: FtTOrderEntry[]
  alwaysWinners: { input: string; candidate: string }[]
  nonImplicators: { input: string; candidate: string }[]
}

export type FtState = ResultState<{ data: FtResultData }>

export type FtParams = FtDefaults

export type SetFtParams = (update: Partial<FtParams>) => void
