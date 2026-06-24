import { describe, it, expect } from "vitest";
import { buildExitSchedule } from "./concentratedExit.js";

// The exit schedule drives a client-facing slide with real tax dollars, so the
// math is pinned with hand-computed expectations. Units are arbitrary ($M here).

describe("buildExitSchedule — basic glide path (no growth, no basis)", () => {
  const res = buildExitSchedule({
    stockPosition: 100,
    costBasisPct: 0,        // entire sale is gain
    investableAssets: 100,  // 100% concentrated to start
    ltcgRate: 20,
    growthRate: 0,
    annualReductionPct: 25, // sell 25/yr → 4 years
    targetConcentrationPct: 0,
  });

  it("sells the position over 4 equal years", () => {
    expect(res.rows).toHaveLength(4);
    expect(res.summary.totalSold).toBeCloseTo(100, 6);
    expect(res.summary.endingConcentrationPct).toBeCloseTo(0, 6);
  });

  it("taxes each tranche at the LTCG rate", () => {
    expect(res.rows[0].sold).toBeCloseTo(25, 6);
    expect(res.rows[0].realizedGain).toBeCloseTo(25, 6);
    expect(res.rows[0].tax).toBeCloseTo(5, 6);       // 25 * 20%
    expect(res.rows[0].netProceeds).toBeCloseTo(20, 6);
    expect(res.summary.totalTax).toBeCloseTo(20, 6); // 4 * 5
  });

  it("tracks the concentration glide path year by year", () => {
    expect(res.rows[0].concentrationPct).toBeCloseTo(78.9473, 3); // 75 / 95
    expect(res.rows[1].concentrationPct).toBeCloseTo(55.5556, 3); // 50 / 90
    expect(res.rows[2].concentrationPct).toBeCloseTo(29.4118, 3); // 25 / 85
    expect(res.rows[3].concentrationPct).toBeCloseTo(0, 6);
  });

  it("reports the all-at-once comparison", () => {
    expect(res.summary.immediateSaleTax).toBeCloseTo(20, 6); // 100 * 100% gain * 20%
  });
});

describe("buildExitSchedule — cost basis reduces the realized gain", () => {
  const res = buildExitSchedule({
    stockPosition: 100,
    costBasisPct: 40,       // $40 basis; only $60 is gain across the whole position
    investableAssets: 100,
    ltcgRate: 25,
    growthRate: 0,
    annualReductionPct: 50, // 2 years
    targetConcentrationPct: 0,
  });

  it("realizes gain net of the proportional basis", () => {
    expect(res.rows).toHaveLength(2);
    // Year 1: sell 50 of 100 → half the basis (20) comes off → gain 30, tax 7.5
    expect(res.rows[0].realizedGain).toBeCloseTo(30, 6);
    expect(res.rows[0].tax).toBeCloseTo(7.5, 6);
    expect(res.summary.totalTax).toBeCloseTo(15, 6); // matches selling all at once (no growth)
    expect(res.summary.immediateSaleTax).toBeCloseTo(15, 6); // 60 gain * 25%
  });
});

describe("buildExitSchedule — annual capital-gains budget caps the sale", () => {
  const res = buildExitSchedule({
    stockPosition: 100,
    costBasisPct: 0,         // gain fraction = 1, so budget maps 1:1 to sale size
    investableAssets: 100,
    ltcgRate: 20,
    growthRate: 0,
    annualReductionPct: 50,  // would sell 50/yr...
    annualGainsBudget: 10,   // ...but budget limits realized gain to 10/yr
    targetConcentrationPct: 0,
  });

  it("limits each year's sale to the gains budget", () => {
    expect(res.rows[0].sold).toBeCloseTo(10, 6);
    expect(res.rows[0].realizedGain).toBeCloseTo(10, 6);
    expect(res.rows[0].tax).toBeCloseTo(2, 6);
    expect(res.rows).toHaveLength(10); // 100 / 10
    expect(res.summary.totalTax).toBeCloseTo(20, 6);
  });
});

describe("buildExitSchedule — stops at the target concentration", () => {
  const res = buildExitSchedule({
    stockPosition: 50,
    costBasisPct: 0,
    investableAssets: 100,   // 50% concentrated, 50 already diversified
    ltcgRate: 0,             // no tax → isolate the glide path
    growthRate: 0,
    annualReductionPct: 20,  // sell 10/yr
    targetConcentrationPct: 10,
  });

  it("halts the year concentration reaches the target, position remaining", () => {
    expect(res.rows).toHaveLength(4);
    expect(res.rows[3].remainingPosition).toBeCloseTo(10, 6);
    expect(res.summary.endingConcentrationPct).toBeCloseTo(10, 6);
  });
});

describe("buildExitSchedule — guards", () => {
  it("returns nothing to do when already below the target concentration", () => {
    const res = buildExitSchedule({
      stockPosition: 5,
      costBasisPct: 10,
      investableAssets: 100, // 5% concentration
      targetConcentrationPct: 10,
    });
    expect(res.rows).toEqual([]);
    expect(res.summary.alreadyDiversified).toBe(true);
  });

  it("returns nothing for a zero position", () => {
    const res = buildExitSchedule({ stockPosition: 0, investableAssets: 100 });
    expect(res.rows).toEqual([]);
    expect(res.summary.yearsModeled).toBe(0);
  });

  it("treats a position with no other assets as 100% concentrated", () => {
    const res = buildExitSchedule({
      stockPosition: 100,
      costBasisPct: 0,
      investableAssets: 100,
      ltcgRate: 0,
      growthRate: 0,
      annualReductionPct: 50,
      targetConcentrationPct: 0,
    });
    expect(res.summary.startingConcentrationPct).toBeCloseTo(100, 6);
  });
});

describe("buildExitSchedule — growth defers rather than eliminates tax", () => {
  const res = buildExitSchedule({
    stockPosition: 100,
    costBasisPct: 20,
    investableAssets: 200,
    ltcgRate: 23.8,
    growthRate: 7.5,
    annualReductionPct: 20,
    targetConcentrationPct: 5,
  });

  it("produces a multi-year plan with non-negative deferral and honest totals", () => {
    expect(res.rows.length).toBeGreaterThan(1);
    expect(res.summary.taxDeferred).toBeGreaterThanOrEqual(0);
    // As the held shares keep appreciating, total tax paid over the plan can
    // exceed an immediate sale — the benefit is deferral + risk reduction.
    expect(res.summary.immediateSaleTax).toBeCloseTo(100 * 0.8 * 0.238, 6);
  });
});
