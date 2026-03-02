//! Shared Gaussian RNG for stochastic learners (GLA, NHG).
//!
//! Box-Muller transform matching the VB6 `Gaussian()` function.
//! Uses `getrandom` for a WASM-compatible (js) and native-compatible uniform source.

/// Uniform sample in [0, 1) from a cryptographic source.
pub fn getrandom_uniform() -> f64 {
    let mut bytes = [0u8; 8];
    getrandom::getrandom(&mut bytes).expect("getrandom failed");
    let n = u64::from_le_bytes(bytes);
    (n as f64) * (1.0 / 18_446_744_073_709_551_616.0_f64)
}

/// Controls how `gaussian()` maps its output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GaussianMode {
    /// Standard normal: values span (−∞, +∞).
    Standard,
    /// Demi-Gaussian (half-normal): returns `abs(value)`, always ≥ 0.
    /// Used by NHG when the "positive demi-Gaussians" option is enabled.
    DemiGaussian,
}

/// Gaussian RNG state. Caches the second value from Box-Muller.
pub struct Rng {
    stored: Option<f64>,
    mode: GaussianMode,
}

impl Rng {
    /// Create a new RNG with the given Gaussian mode.
    pub fn new(mode: GaussianMode) -> Self {
        Rng { stored: None, mode }
    }

    /// Uniform sample in [0, 1).
    pub fn uniform(&mut self) -> f64 {
        getrandom_uniform()
    }

    /// Standard normal deviate (σ=1). Scale by noise_sigma at call site.
    ///
    /// In `DemiGaussian` mode, returns `abs(value)`.
    pub fn gaussian(&mut self) -> f64 {
        let val = if let Some(stored) = self.stored.take() {
            stored
        } else {
            loop {
                let v1 = 2.0 * getrandom_uniform() - 1.0;
                let v2 = 2.0 * getrandom_uniform() - 1.0;
                let r = v1 * v1 + v2 * v2;
                if r > 0.0 && r < 1.0 {
                    let fac = (-2.0 * r.ln() / r).sqrt();
                    self.stored = Some(v1 * fac);
                    break v2 * fac;
                }
            }
        };
        match self.mode {
            GaussianMode::Standard => val,
            GaussianMode::DemiGaussian => val.abs(),
        }
    }
}
