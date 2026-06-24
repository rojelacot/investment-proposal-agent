import { describe, it, expect } from "vitest";
import { projectFeeDrag, compareFeeDrag } from "./feeProjection.js";

describe("projectFeeDrag", () => {
  it("charges the fee on the grown balance each year", () => {
    // $1,000,000, 1% fee, 10% gross, 1 year:
    // grown = 1,100,000; fee = 11,000; balance = 1,089,000
    const r = projectFeeDrag({ portfolioValue: 1_000_000, annualFeePct: 1, years: 1, grossReturnPct: 10 });
    expect(r.rows).toHaveLength(1);
    expect(r.rows[0].feePaid).toBeCloseTo(11_000, 4);
    expect(r.endingBalance).toBeCloseTo(1_089_000, 4);
    expect(r.totalFees).toBeCloseTo(11_000, 4);
  });

  it("accumulates fees over multiple years", () => {
    const r = projectFeeDrag({ portfolioValue: 1_000_000, annualFeePct: 1, years: 3, grossReturnPct: 10 });
    expect(r.rows).toHaveLength(3);
    // cumulative fees strictly increase
    expect(r.rows[2].cumulativeFees).toBeGreaterThan(r.rows[0].cumulativeFees);
    expect(r.totalFees).toBeCloseTo(r.rows[2].cumulativeFees, 6);
  });

  it("with no fee, ending balance is pure compounding and total fees are zero", () => {
    const r = projectFeeDrag({ portfolioValue: 1000, annualFeePct: 0, years: 5, grossReturnPct: 7 });
    expect(r.totalFees).toBeCloseTo(0, 10);
    expect(r.endingBalance).toBeCloseTo(1000 * 1.07 ** 5, 6);
  });

  it("returns the starting value unchanged for a zero-year horizon", () => {
    const r = projectFeeDrag({ portfolioValue: 5000, annualFeePct: 1, years: 0 });
    expect(r.rows).toEqual([]);
    expect(r.endingBalance).toBe(5000);
  });
});

describe("compareFeeDrag", () => {
  it("reports positive savings when the proposed fee is lower", () => {
    const c = compareFeeDrag({
      portfolioValue: 1_000_000,
      currentFeePct: 1.2,
      proposedFeePct: 0.6,
      years: 20,
      grossReturnPct: 7,
    });
    expect(c.annualFeeReductionPct).toBeCloseTo(0.6, 6);
    expect(c.cumulativeFeeSavings).toBeGreaterThan(0);
    // Lower fees → more compounding → higher ending balance.
    expect(c.endingBalanceDifference).toBeGreaterThan(0);
    expect(c.proposed.totalFees).toBeLessThan(c.current.totalFees);
  });

  it("is symmetric: equal fees produce no difference", () => {
    const c = compareFeeDrag({
      portfolioValue: 500_000,
      currentFeePct: 0.8,
      proposedFeePct: 0.8,
      years: 15,
    });
    expect(c.cumulativeFeeSavings).toBeCloseTo(0, 6);
    expect(c.endingBalanceDifference).toBeCloseTo(0, 6);
  });
});
