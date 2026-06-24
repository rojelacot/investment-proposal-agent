// Monte Carlo projection of portfolio outcomes.
//
// A historical backtest answers "what happened." A Monte Carlo answers "what is
// the range of plausible futures" — by simulating many paths of annual returns
// drawn from a normal distribution parameterized by the portfolio's expected
// return and volatility. It produces a fan of percentile outcomes and, when a
// goal is supplied, the probability of reaching it.
//
// A seeded PRNG makes every run reproducible (and therefore testable). This is
// illustrative planning math, not a guarantee — consumers must disclaim it.

// mulberry32: tiny, fast, deterministic PRNG seeded from a 32-bit integer.
function makeRng(seed) {
  let a = seed >>> 0;
  return function () {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// Standard normal via Box–Muller, using a supplied uniform RNG.
function nextNormal(rng) {
  let u1 = rng();
  const u2 = rng();
  if (u1 < 1e-12) u1 = 1e-12; // avoid log(0)
  return Math.sqrt(-2 * Math.log(u1)) * Math.cos(2 * Math.PI * u2);
}

/** Percentile (0–1) of a numeric array via linear interpolation. */
export function percentile(sortedAsc, p) {
  const n = sortedAsc.length;
  if (n === 0) return null;
  if (n === 1) return sortedAsc[0];
  const idx = p * (n - 1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return sortedAsc[lo];
  return sortedAsc[lo] + (sortedAsc[hi] - sortedAsc[lo]) * (idx - lo);
}

/**
 * Run a Monte Carlo simulation of portfolio value over time.
 *
 * @param {object} p
 * @param {number}  p.initialValue        Starting portfolio value.
 * @param {number} [p.years=20]           Projection horizon.
 * @param {number} [p.expectedReturnPct=7] Mean annual return, %.
 * @param {number} [p.volatilityPct=12]   Annual volatility (stdev of returns), %.
 * @param {number} [p.annualContribution=0] Added each year (same units as value).
 * @param {number} [p.annualWithdrawal=0]  Withdrawn each year.
 * @param {number} [p.goalValue=null]      If set, success = ending value ≥ goal and never depleted.
 * @param {number} [p.simulations=2000]    Number of paths.
 * @param {number} [p.seed=12345]          PRNG seed for reproducibility.
 * @returns {{percentiles, successProbability, fan, endingValues, simulations, assumptions}}
 */
export function runMonteCarlo({
  initialValue,
  years = 20,
  expectedReturnPct = 7,
  volatilityPct = 12,
  annualContribution = 0,
  annualWithdrawal = 0,
  goalValue = null,
  simulations = 2000,
  seed = 12345,
} = {}) {
  const start = Math.max(Number(initialValue) || 0, 0);
  const n = Math.max(Math.floor(Number(years) || 0), 0);
  const mean = (Number(expectedReturnPct) || 0) / 100;
  const vol = Math.max(Number(volatilityPct) || 0, 0) / 100;
  const contrib = Number(annualContribution) || 0;
  const withdraw = Number(annualWithdrawal) || 0;
  const sims = Math.max(Math.floor(Number(simulations) || 0), 1);
  const goal = goalValue == null ? null : Number(goalValue);

  const assumptions = {
    initialValue: start, years: n, expectedReturnPct, volatilityPct,
    annualContribution: contrib, annualWithdrawal: withdraw, goalValue: goal, simulations: sims, seed,
  };

  const rng = makeRng(seed);

  // valuesByYear[y] holds every simulation's value at end of year y (y=0 is start).
  const valuesByYear = Array.from({ length: n + 1 }, () => new Array(sims));
  const endingValues = new Array(sims);
  let successes = 0;

  for (let s = 0; s < sims; s++) {
    let value = start;
    let depleted = false;
    valuesByYear[0][s] = value;
    for (let y = 1; y <= n; y++) {
      // Arithmetic-normal annual return, floored at -95% to keep a single draw
      // from sending the path implausibly negative.
      const r = Math.max(mean + vol * nextNormal(rng), -0.95);
      value = value * (1 + r) + contrib - withdraw;
      if (value <= 0) { value = 0; depleted = true; }
      valuesByYear[y][s] = value;
    }
    endingValues[s] = value;
    if (goal != null && value >= goal && !depleted) successes++;
  }

  // Percentile fan per year.
  const fan = valuesByYear.map((arr, year) => {
    const sorted = [...arr].sort((a, b) => a - b);
    return {
      year,
      p10: percentile(sorted, 0.10),
      p25: percentile(sorted, 0.25),
      p50: percentile(sorted, 0.50),
      p75: percentile(sorted, 0.75),
      p90: percentile(sorted, 0.90),
    };
  });

  const sortedEnding = [...endingValues].sort((a, b) => a - b);
  const percentiles = {
    p10: percentile(sortedEnding, 0.10),
    p25: percentile(sortedEnding, 0.25),
    p50: percentile(sortedEnding, 0.50),
    p75: percentile(sortedEnding, 0.75),
    p90: percentile(sortedEnding, 0.90),
  };

  return {
    percentiles,
    successProbability: goal == null ? null : successes / sims,
    fan,
    endingValues: sortedEnding,
    simulations: sims,
    assumptions,
  };
}
