/**
 * Lazy-cached defaults read from Rust/WASM option constructors.
 *
 * Each getter constructs a WASM options object once, reads its default field
 * values, frees the WASM object, and caches the result. This ensures the web
 * frontend always stays in sync with the Rust defaults — no duplicated values.
 *
 * IMPORTANT: These getters must only be called after WASM is initialized
 * (i.e. inside component bodies that render after WASM loads).
 */

import {
  FredOptions,
  FtOptions,
  GlaOptions,
  MaxEntOptions,
  NhgOptions,
} from '../pkg/ot_soft.js'
import { DEFAULT_SCHEDULE_TEMPLATE } from './constants.ts'

// ─── MaxEnt ──────────────────────────────────────────────────────────────────

export interface MaxEntDefaults {
  iterations: number
  weightMin: number
  weightMax: number
  usePrior: boolean
  sigmaSquared: number
  generateHistory: boolean
  generateOutputProbHistory: boolean
  // UI-only (no Rust equivalent)
  sortByWeight: boolean
}

let _maxent: MaxEntDefaults | null = null
export function maxentDefaults(): MaxEntDefaults {
  if (!_maxent) {
    const d = new MaxEntOptions()
    _maxent = {
      iterations: d.iterations,
      weightMin: d.weight_min,
      weightMax: d.weight_max,
      usePrior: d.use_prior,
      sigmaSquared: d.sigma_squared,
      generateHistory: d.generate_history,
      generateOutputProbHistory: d.generate_output_prob_history,
      sortByWeight: true,
    }
    d.free()
  }
  return _maxent
}

// ─── GLA ─────────────────────────────────────────────────────────────────────

export interface GlaDefaults {
  maxentMode: boolean
  cycles: number
  initialPlasticity: number
  finalPlasticity: number
  testTrials: number
  negativeWeightsOk: boolean
  gaussianPrior: boolean
  sigma: number
  magriUpdateRule: boolean
  exactProportions: boolean
  aprioriText: string
  aprioriGap: number
  generateHistory: boolean
  generateFullHistory: boolean
  generateCandidateProbHistory: boolean
  // UI-only (no Rust equivalent)
  useCustomSchedule: boolean
  customSchedule: string
  multipleRunsCount: 10 | 100 | 1000
}

let _gla: GlaDefaults | null = null
export function glaDefaults(): GlaDefaults {
  if (!_gla) {
    const d = new GlaOptions()
    _gla = {
      maxentMode: d.maxent_mode,
      cycles: d.cycles,
      initialPlasticity: d.initial_plasticity,
      finalPlasticity: d.final_plasticity,
      testTrials: d.test_trials,
      negativeWeightsOk: d.negative_weights_ok,
      gaussianPrior: d.gaussian_prior,
      sigma: d.sigma,
      magriUpdateRule: d.magri_update_rule,
      exactProportions: d.exact_proportions,
      aprioriText: d.apriori_text,
      aprioriGap: d.apriori_gap,
      generateHistory: d.generate_history,
      generateFullHistory: d.generate_full_history,
      generateCandidateProbHistory: d.generate_candidate_prob_history,
      useCustomSchedule: false,
      customSchedule: DEFAULT_SCHEDULE_TEMPLATE,
      multipleRunsCount: 10,
    }
    d.free()
  }
  return _gla
}

// ─── NHG ─────────────────────────────────────────────────────────────────────

export interface NhgDefaults {
  cycles: number
  initialPlasticity: number
  finalPlasticity: number
  testTrials: number
  noiseByCell: boolean
  postMultNoise: boolean
  noiseForZeroCells: boolean
  lateNoise: boolean
  exponentialNhg: boolean
  demiGaussians: boolean
  negativeWeightsOk: boolean
  resolveTiesBySkipping: boolean
  exactProportions: boolean
  generateHistory: boolean
  generateFullHistory: boolean
  // UI-only (no Rust equivalent)
  useCustomSchedule: boolean
  customSchedule: string
}

let _nhg: NhgDefaults | null = null
export function nhgDefaults(): NhgDefaults {
  if (!_nhg) {
    const d = new NhgOptions()
    _nhg = {
      cycles: d.cycles,
      initialPlasticity: d.initial_plasticity,
      finalPlasticity: d.final_plasticity,
      testTrials: d.test_trials,
      noiseByCell: d.noise_by_cell,
      postMultNoise: d.post_mult_noise,
      noiseForZeroCells: d.noise_for_zero_cells,
      lateNoise: d.late_noise,
      exponentialNhg: d.exponential_nhg,
      demiGaussians: d.demi_gaussians,
      negativeWeightsOk: d.negative_weights_ok,
      resolveTiesBySkipping: d.resolve_ties_by_skipping,
      exactProportions: d.exact_proportions,
      generateHistory: d.generate_history,
      generateFullHistory: d.generate_full_history,
      useCustomSchedule: false,
      customSchedule: DEFAULT_SCHEDULE_TEMPLATE,
    }
    d.free()
  }
  return _nhg
}

// ─── RCD (FredOptions) ───────────────────────────────────────────────────────

export interface RcdDefaults {
  // UI-only (no Rust equivalent)
  algorithm: 'rcd' | 'bcd' | 'bcd-specific' | 'lfcd'
  // From Rust FredOptions
  includeFred: boolean
  useMib: boolean
  showDetails: boolean
  includeMiniTableaux: boolean
}

let _rcd: RcdDefaults | null = null
export function rcdDefaults(): RcdDefaults {
  if (!_rcd) {
    const d = new FredOptions()
    _rcd = {
      algorithm: 'rcd',
      includeFred: d.include_fred,
      useMib: d.use_mib,
      showDetails: d.show_details,
      includeMiniTableaux: d.include_mini_tableaux,
    }
    d.free()
  }
  return _rcd
}

// ─── Factorial Typology ──────────────────────────────────────────────────────

export interface FtDefaults {
  includeFullListing: boolean
  // UI-only (no Rust equivalent)
  includeFtsum: boolean
  includeCompactSum: boolean
}

let _ft: FtDefaults | null = null
export function ftDefaults(): FtDefaults {
  if (!_ft) {
    const d = new FtOptions()
    _ft = {
      includeFullListing: d.include_full_listing,
      includeFtsum: false,
      includeCompactSum: false,
    }
    d.free()
  }
  return _ft
}
