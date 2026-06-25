import { describe, it, expect } from "vitest";
import { computeImpactScorecard } from "./impactScorecard.js";

// Mercer-like inputs ($M): $9M NVDA, 12% basis, LTCG 23.8%+13.3%=37.1%.
const baseData = {
  stockPosition: 9,
  costBasisPct: 12,
  federalTaxRate: 23.8,
  stateTaxRate: 13.3,
  immediateTax: 2.94,
  crtAllocation: 1.8,
  collarAllocation: 2.7,
  harvestingSleeve: 2.7,
  taxSavings: 0.25,
  stockPrice: 120,
  putStrike: 102, // 15% floor
};

describe("computeImpactScorecard", () => {
  it("returns zeros when no strategies are selected", () => {
    const r = computeImpactScorecard({ data: baseData, selectedStrategies: {} });
    expect(r.taxSaved).toBe(0);
    expect(r.downsideProtected).toBe(0);
    expect(r.immediateTax).toBeCloseTo(2.94, 6);
    expect(r.ltcgRatePct).toBeCloseTo(37.1, 6);
  });

  it("CRT saves cap-gains tax on the donated shares and removes them from crash risk", () => {
    const r = computeImpactScorecard({ data: baseData, selectedStrategies: { crt: true } });
    // 1.8 * (1-0.12) * 0.371 = 0.587...
    expect(r.perStrategy.crt.taxSaved).toBeCloseTo(1.8 * 0.88 * 0.371, 6);
    expect(r.perStrategy.crt.downsideProtected).toBeCloseTo(1.8 * 0.40, 6);
    expect(r.taxSaved).toBeCloseTo(r.perStrategy.crt.taxSaved, 6);
  });

  it("harvesting contributes tax savings only (no downside)", () => {
    const r = computeImpactScorecard({ data: baseData, selectedStrategies: { harvesting: true } });
    expect(r.perStrategy.harvesting.taxSaved).toBeCloseTo(0.25, 6);
    expect(r.perStrategy.harvesting.downsideProtected).toBe(0);
  });

  it("collar protects the drawdown beyond the put floor, no tax saving", () => {
    const r = computeImpactScorecard({ data: baseData, selectedStrategies: { collar: true } });
    // floor drop = 1 - 102/120 = 0.15; protected = 2.7 * (0.40 - 0.15) = 0.675
    expect(r.perStrategy.collar.downsideProtected).toBeCloseTo(2.7 * 0.25, 6);
    expect(r.perStrategy.collar.taxSaved).toBe(0);
  });

  it("diversification protects remaining single-stock capital (no double-count with CRT/collar)", () => {
    const r = computeImpactScorecard({
      data: baseData,
      selectedStrategies: { crt: true, collar: true, diversification: true },
    });
    // remaining = 9 - 1.8 - 2.7 = 4.5 ; protected = 4.5 * 0.40 = 1.8
    expect(r.perStrategy.diversification.downsideProtected).toBeCloseTo(4.5 * 0.40, 6);
    expect(r.perStrategy.diversification.taxSaved).toBe(0);
  });

  it("caps total downside protected at the whole position's severe-drawdown loss", () => {
    // All strategies on: CRT 1.8*.4 + collar .675 + diversification 4.5*.4 = 0.72+0.675+1.8 = 3.195
    // Cap = stockPosition * 0.40 = 3.6, so total stays under the cap here.
    const r = computeImpactScorecard({
      data: baseData,
      selectedStrategies: { crt: true, collar: true, diversification: true, harvesting: true },
    });
    expect(r.downsideProtected).toBeLessThanOrEqual(9 * 0.40 + 1e-9);
    expect(r.downsideProtected).toBeCloseTo(1.8 * 0.4 + 2.7 * 0.25 + 4.5 * 0.4, 6);
    expect(r.taxSaved).toBeCloseTo(1.8 * 0.88 * 0.371 + 0.25, 6);
  });

  it("derives immediateTax from basis when not supplied", () => {
    const r = computeImpactScorecard({
      data: { stockPosition: 10, costBasisPct: 0, federalTaxRate: 20, stateTaxRate: 0 },
      selectedStrategies: {},
    });
    expect(r.immediateTax).toBeCloseTo(10 * 1 * 0.20, 6); // 2.0
  });
});
