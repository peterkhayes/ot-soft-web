import type { ResultState } from '../../types.ts'
import type { RcdDefaults } from '../../wasmDefaults.ts'

export interface StratumData {
  constraints: { abbrev: string; fullName: string }[]
}

export interface RcdResultState {
  success: boolean
  strata: StratumData[]
  tieWarning: boolean
  hasseDot?: string
  log?: string
}

export type RcdState = ResultState<RcdResultState>

export type Algorithm = 'rcd' | 'bcd' | 'bcd-specific' | 'lfcd'

export type RcdParams = RcdDefaults

export type SetRcdParams = (update: Partial<RcdParams>) => void

export const ALGORITHM_LABELS: Record<Algorithm, string> = {
  rcd: 'RCD',
  bcd: 'BCD',
  'bcd-specific': 'BCD (Specific)',
  lfcd: 'LFCD',
}

export const ALGORITHM_DESCRIPTIONS: Record<Algorithm, string> = {
  rcd: 'Recursive Constraint Demotion',
  bcd: 'Biased Constraint Demotion',
  'bcd-specific': 'Biased Constraint Demotion (favors specific faithfulness)',
  lfcd: 'Low Faithfulness Constraint Demotion',
}
