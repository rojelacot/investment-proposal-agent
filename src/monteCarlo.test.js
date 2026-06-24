import { describe, it, expect } from "vitest";
import { runMonteCarlo, percentile } from "./monteCarlo.js";

describe("percentile", () => {
  it("interpolates linearly", () => {
    expect(percentile([10, 20, 30, 40], 0.5)).toBeCloseTo(25, 10);
    expect(percentile([10, 20, 30, 40], 0)).toBe(10);
    expect(percentile([10, 20, 30, 40], 1)).toBe(40);
  });
  it("handles trivial arrays", () => {
    expect(percentile([], 0.5)).toBeNull();
    expect(percentile([7], 0.9)).toBe(7);
  });
});

describe("runMonteCarlo — deterministic with zero volatility", () => {
  const res = runMonteCarlo({
    initialValue: 1000,
    years: 10,
    expectedReturnPct: 7,
    volatilityPct: 0, // no randomness → every path identical
    simulations: 50,
  });

  it("collapses all percentiles to the deterministic compounded value", () => {
    const expected = 1000 * 1.07 ** 10;
    for (const p of ["p10", "p25", "p50", "p75", "p90"]) {
      expect(res.percentiles[p]).toBeCloseTo(expected, 4);
    }
  });

  it("produces a fan with years+1 points starting at the initial value", () => {
    expect(res.fan).toHaveLength(11);
    expect(res.fan[0].p50).toBeCloseTo(1000, 6);
    expect(res.fan[10].p50).toBeCloseTo(1000 * 1.07 ** 10, 4);
  });
});

describe("runMonteCarlo — goal success probability", () => {
  it("is 1 when the deterministic outcome clears an easy goal", () => {
    const res = runMonteCarlo({
      initialValue: 1000, years: 10, expectedReturnPct: 7, volatilityPct: 0,
      goalValue: 1500, simulations: 100,
    });
    expect(res.successProbability).toBe(1);
  });

  it("is 0 when the deterministic outcome misses an impossible goal", () => {
    const res = runMonteCarlo({
      initialValue: 1000, years: 10, expectedReturnPct: 7, volatilityPct: 0,
      goalValue: 1e9, simulations: 100,
    });
    expect(res.successProbability).toBe(0);
  });

  it("is null when no goal is supplied", () => {
    const res = runMonteCarlo({ initialValue: 1000, years: 5, volatilityPct: 0 });
    expect(res.successProbability).toBeNull();
  });

  it("falls between 0 and 1 for a goal near the median, and rises with horizon", () => {
    const common = { initialValue: 1000, expectedReturnPct: 7, volatilityPct: 15, goalValue: 1500, simulations: 3000, seed: 42 };
    const short = runMonteCarlo({ ...common, years: 5 });
    const long = runMonteCarlo({ ...common, years: 20 });
    expect(short.successProbability).toBeGreaterThan(0);
    expect(short.successProbability).toBeLessThan(1);
    // More years of 7% drift → more likely to clear a fixed nominal goal.
    expect(long.successProbability).toBeGreaterThan(short.successProbability);
  });
});

describe("runMonteCarlo — reproducibility & ordering", () => {
  it("is reproducible for a fixed seed", () => {
    const a = runMonteCarlo({ initialValue: 1000, years: 15, volatilityPct: 12, simulations: 500, seed: 7 });
    const b = runMonteCarlo({ initialValue: 1000, years: 15, volatilityPct: 12, simulations: 500, seed: 7 });
    expect(a.percentiles.p50).toBe(b.percentiles.p50);
    expect(a.percentiles.p10).toBe(b.percentiles.p10);
  });

  it("keeps percentiles ordered p10 ≤ p50 ≤ p90", () => {
    const res = runMonteCarlo({ initialValue: 1000, years: 20, volatilityPct: 18, simulations: 2000, seed: 99 });
    expect(res.percentiles.p10).toBeLessThanOrEqual(res.percentiles.p50);
    expect(res.percentiles.p50).toBeLessThanOrEqual(res.percentiles.p90);
  });

  it("models contributions (more ending value than without)", () => {
    const base = runMonteCarlo({ initialValue: 1000, years: 10, volatilityPct: 0, expectedReturnPct: 5 });
    const withAdd = runMonteCarlo({ initialValue: 1000, years: 10, volatilityPct: 0, expectedReturnPct: 5, annualContribution: 100 });
    expect(withAdd.percentiles.p50).toBeGreaterThan(base.percentiles.p50);
  });
});
