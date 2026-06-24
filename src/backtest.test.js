import { describe, it, expect } from "vitest";
import {
  toMonthlyReturns,
  buildTargetWeightMap,
  buildConcentratedWeightMap,
  weightedPortfolioReturns,
  computeTrailingAnnualReturns,
  averageAnnualReturn,
  weightedAverageAnnualReturn,
  summarizeReturns,
  buildHoldingsWeightMap,
} from "./backtest.js";

// These functions feed real numbers into client-facing proposals. The point of
// these tests is to lock in the math (CAGR, max drawdown, the 80%-coverage
// window, missing-data exclusion) so a future refactor can't silently change a
// figure that ends up in front of a client.

describe("toMonthlyReturns", () => {
  it("returns [] for empty or single-point series", () => {
    expect(toMonthlyReturns([])).toEqual([]);
    expect(toMonthlyReturns([{ date: "2020-01", close: 100 }])).toEqual([]);
    expect(toMonthlyReturns(null)).toEqual([]);
  });

  it("computes simple month-over-month returns", () => {
    const out = toMonthlyReturns([
      { date: "2020-01", close: 100 },
      { date: "2020-02", close: 110 },
      { date: "2020-03", close: 99 },
    ]);
    expect(out).toHaveLength(2);
    expect(out[0]).toEqual({ date: "2020-02", return: expect.closeTo(0.1, 10) });
    expect(out[1]).toEqual({ date: "2020-03", return: expect.closeTo(-0.1, 10) });
  });

  it("sorts unsorted input by date before computing", () => {
    const out = toMonthlyReturns([
      { date: "2020-02", close: 110 },
      { date: "2020-01", close: 100 },
    ]);
    expect(out).toEqual([{ date: "2020-02", return: expect.closeTo(0.1, 10) }]);
  });

  it("skips transitions involving non-positive closes", () => {
    const out = toMonthlyReturns([
      { date: "2020-01", close: 100 },
      { date: "2020-02", close: 0 },
      { date: "2020-03", close: 120 },
    ]);
    // 100->0 dropped (curr<=0), 0->120 dropped (prev<=0)
    expect(out).toEqual([]);
  });
});

describe("buildTargetWeightMap", () => {
  it("converts allocations to fractional weights with full coverage", () => {
    const { weights, coveragePct, excluded } = buildTargetWeightMap([
      { name: "A", ticker: "spy", alloc: 60, assetClass: "Equity" },
      { name: "B", ticker: "AGG", alloc: 40, assetClass: "Fixed Income" },
    ]);
    expect(weights).toEqual({ SPY: 0.6, AGG: 0.4 });
    expect(coveragePct).toBe(100);
    expect(excluded).toEqual([]);
  });

  it("substitutes a proxy for non-tradable tickers", () => {
    const { weights } = buildTargetWeightMap(
      [{ name: "Hedge", ticker: "SMA", alloc: 50, assetClass: "Alternatives" }],
      { Alternatives: "QAI" }
    );
    expect(weights.QAI).toBe(0.5);
    expect(weights.__proxied).toBeTruthy();
  });

  it("excludes non-tradable tickers with no proxy and reduces coverage", () => {
    const { weights, coveragePct, excluded } = buildTargetWeightMap([
      { name: "Public", ticker: "VTI", alloc: 70, assetClass: "Equity" },
      { name: "Private", ticker: "N/A", alloc: 30, assetClass: "Private" },
    ]);
    expect(weights).toEqual({ VTI: 0.7 });
    expect(coveragePct).toBeCloseTo(70, 6);
    expect(excluded).toEqual([{ name: "Private", assetClass: "Private", alloc: 30 }]);
  });

  it("skips zero/negative allocations and aggregates duplicate tickers", () => {
    const { weights } = buildTargetWeightMap([
      { name: "A", ticker: "SPY", alloc: 30, assetClass: "Equity" },
      { name: "B", ticker: "SPY", alloc: 20, assetClass: "Equity" },
      { name: "C", ticker: "AGG", alloc: 0, assetClass: "Fixed Income" },
    ]);
    expect(weights).toEqual({ SPY: 0.5 });
  });
});

describe("buildConcentratedWeightMap", () => {
  it("puts 100% in the concentrated ticker when conc=100", () => {
    expect(buildConcentratedWeightMap("AAPL", 100, "AOR")).toEqual({
      weights: { AAPL: 1 },
      coveragePct: 100,
    });
  });

  it("splits remainder into the benchmark proxy", () => {
    const { weights, coveragePct } = buildConcentratedWeightMap("AAPL", 60, "AOR");
    expect(weights.AAPL).toBeCloseTo(0.6, 10);
    expect(weights.AOR).toBeCloseTo(0.4, 10);
    expect(coveragePct).toBe(100);
  });

  it("reports partial coverage when there is no benchmark", () => {
    const { weights, coveragePct } = buildConcentratedWeightMap("AAPL", 60, "");
    expect(weights).toEqual({ AAPL: 0.6 });
    expect(coveragePct).toBe(60);
  });

  it("clamps concentration above 100% and rejects invalid input", () => {
    expect(buildConcentratedWeightMap("AAPL", 150, "AOR").weights.AAPL).toBe(1);
    expect(buildConcentratedWeightMap("", 50, "AOR")).toEqual({ weights: {}, coveragePct: 0 });
    expect(buildConcentratedWeightMap("AAPL", 0, "AOR")).toEqual({ weights: {}, coveragePct: 0 });
  });
});

describe("weightedPortfolioReturns", () => {
  it("returns a single ticker's series unchanged (weight renormalized to 1)", () => {
    const out = weightedPortfolioReturns(
      { SPY: 0.4 },
      { SPY: [{ date: "2020-02", return: 0.1 }, { date: "2020-03", return: -0.05 }] }
    );
    expect(out).toEqual([
      { date: "2020-02", return: expect.closeTo(0.1, 10) },
      { date: "2020-03", return: expect.closeTo(-0.05, 10) },
    ]);
  });

  it("blends two equal-weight tickers on shared dates", () => {
    const out = weightedPortfolioReturns(
      { A: 0.5, B: 0.5 },
      {
        A: [{ date: "2020-02", return: 0.1 }, { date: "2020-03", return: 0.2 }],
        B: [{ date: "2020-02", return: 0.0 }, { date: "2020-03", return: 0.0 }],
      }
    );
    expect(out[0].return).toBeCloseTo(0.05, 10);
    expect(out[1].return).toBeCloseTo(0.1, 10);
  });

  it("does NOT let a small short-history position collapse the window (80% rule)", () => {
    // B holds only 15% and has just 1 month of history. Window should start at A's
    // first date and exclude B entirely, rather than shrinking to B's 1 month.
    const out = weightedPortfolioReturns(
      { A: 0.85, B: 0.15 },
      {
        A: [
          { date: "2020-01", return: 0.01 },
          { date: "2020-02", return: 0.02 },
          { date: "2020-03", return: 0.03 },
        ],
        B: [{ date: "2020-03", return: 0.5 }],
      }
    );
    expect(out.map(o => o.date)).toEqual(["2020-01", "2020-02", "2020-03"]);
    expect(out[2].return).toBeCloseTo(0.03, 10); // pure A, B excluded
  });

  it("includes a large short-history position by shrinking the window (80% rule)", () => {
    // B now holds 30%, so excluding it would drop below 80% coverage. Window must
    // shrink to the dates where both have data.
    const out = weightedPortfolioReturns(
      { A: 0.7, B: 0.3 },
      {
        A: [
          { date: "2020-01", return: 0.01 },
          { date: "2020-02", return: 0.02 },
          { date: "2020-03", return: 0.03 },
        ],
        B: [
          { date: "2020-02", return: 0.1 },
          { date: "2020-03", return: 0.1 },
        ],
      }
    );
    expect(out.map(o => o.date)).toEqual(["2020-02", "2020-03"]);
    expect(out[0].return).toBeCloseTo(0.7 * 0.02 + 0.3 * 0.1, 10);
  });

  it("returns [] when no ticker has data", () => {
    expect(weightedPortfolioReturns({ A: 1 }, {})).toEqual([]);
  });
});

describe("computeTrailingAnnualReturns", () => {
  it("returns [] with fewer than 12 months", () => {
    expect(computeTrailingAnnualReturns([{ date: "2020-01", return: 0.01 }])).toEqual([]);
  });

  it("compounds a single 12-month block", () => {
    const months = Array.from({ length: 12 }, (_, i) => ({
      date: `2020-${String(i + 1).padStart(2, "0")}`,
      return: 0.01,
    }));
    const out = computeTrailingAnnualReturns(months);
    expect(out).toHaveLength(1);
    expect(out[0].return).toBeCloseTo(Math.pow(1.01, 12) - 1, 10);
    expect(out[0].endDate).toBe("2020-12");
  });

  it("splits 24 months into 2 non-overlapping blocks, oldest first, dropping leftovers", () => {
    // 25 months: oldest 1 month is a partial block and must be dropped.
    const months = Array.from({ length: 25 }, (_, i) => {
      const m = i % 12 + 1;
      const y = 2019 + Math.floor(i / 12);
      return { date: `${y}-${String(m).padStart(2, "0")}`, return: 0 };
    });
    const out = computeTrailingAnnualReturns(months, 10);
    expect(out).toHaveLength(2);
    out.forEach(b => expect(b.return).toBeCloseTo(0, 10));
  });

  it("respects maxYears cap", () => {
    const months = Array.from({ length: 60 }, () => ({ date: "x", return: 0 }));
    expect(computeTrailingAnnualReturns(months, 3)).toHaveLength(3);
  });
});

describe("averageAnnualReturn", () => {
  it("returns null without a full 12 months", () => {
    expect(averageAnnualReturn([{ date: "2020-01", return: 0.05 }])).toBeNull();
  });

  it("averages trailing annual returns", () => {
    // 24 months: first year +0% each month, second year structured so blocks differ.
    const months = [
      ...Array.from({ length: 12 }, () => ({ date: "a", return: 0 })),
      ...Array.from({ length: 12 }, () => ({ date: "b", return: 0.01 })),
    ];
    const res = averageAnnualReturn(months);
    expect(res.yearsUsed).toBe(2);
    const expected = (0 + (Math.pow(1.01, 12) - 1)) / 2;
    expect(res.value).toBeCloseTo(expected, 10);
  });
});

describe("weightedAverageAnnualReturn", () => {
  const year = r => Array.from({ length: 12 }, () => ({ date: "x", return: r }));

  it("weights each ticker's average and reports full coverage", () => {
    const res = weightedAverageAnnualReturn(
      { A: 0.5, B: 0.5 },
      { A: year(0), B: year(0.01) }
    );
    const bAnnual = Math.pow(1.01, 12) - 1;
    expect(res.value).toBeCloseTo(0.5 * 0 + 0.5 * bAnnual, 10);
    expect(res.coveragePct).toBe(100);
    expect(res.minYearsUsed).toBe(1);
  });

  it("renormalizes across tickers with enough history and reports partial coverage", () => {
    const res = weightedAverageAnnualReturn(
      { A: 0.5, B: 0.5 },
      { A: year(0.02), B: [{ date: "x", return: 0.5 }] } // B too short
    );
    const aAnnual = Math.pow(1.02, 12) - 1;
    expect(res.value).toBeCloseTo(aAnnual, 10); // A only, renormalized to 1
    expect(res.coveragePct).toBe(50);
  });

  it("returns null when no ticker has enough history", () => {
    expect(
      weightedAverageAnnualReturn({ A: 1 }, { A: [{ date: "x", return: 0.1 }] })
    ).toBeNull();
  });
});

describe("summarizeReturns", () => {
  it("returns nulls for an empty series", () => {
    const s = summarizeReturns([]);
    expect(s.months).toBe(0);
    expect(s.annualizedReturn).toBeNull();
    expect(s.maxDrawdown).toBeNull();
  });

  it("computes CAGR, cumulative return and max drawdown", () => {
    // +10%, -20%, +5%  ->  10000 -> 11000 -> 8800 -> 9240
    const s = summarizeReturns([
      { date: "2020-01", return: 0.1 },
      { date: "2020-02", return: -0.2 },
      { date: "2020-03", return: 0.05 },
    ]);
    expect(s.months).toBe(3);
    expect(s.startDate).toBe("2020-01");
    expect(s.endDate).toBe("2020-03");
    expect(s.cumulativeReturn).toBeCloseTo(9240 / 10000 - 1, 10);
    // peak 11000, trough 8800 -> dd = -0.2
    expect(s.maxDrawdown).toBeCloseTo(-0.2, 10);
    // CAGR = (1+cum)^(12/3) - 1
    const expectedCagr = Math.pow(9240 / 10000, 12 / 3) - 1;
    expect(s.annualizedReturn).toBeCloseTo(expectedCagr, 10);
    // growth series includes the implicit starting $10k point + one per month
    expect(s.growthSeries).toHaveLength(4);
    expect(s.growthSeries[0].value).toBe(10000);
    expect(s.growthSeries[3].value).toBeCloseTo(9240, 6);
  });

  it("tracks max drawdown across a new peak (uses running peak, not global)", () => {
    // up, down, recover above old peak, then a bigger drop from the new peak
    const s = summarizeReturns([
      { date: "2020-01", return: 0.5 },  // 15000 peak
      { date: "2020-02", return: -0.1 }, // 13500, dd=-0.1
      { date: "2020-03", return: 0.4 },  // 18900 new peak
      { date: "2020-04", return: -0.3 }, // 13230, dd=(13230-18900)/18900
    ]);
    expect(s.maxDrawdown).toBeCloseTo(-0.3, 10);
  });

  it("computes Sharpe and Sortino ratios from the risk-free hurdle", () => {
    // Steady +1%/mo: low volatility, no downside months → high Sharpe, null Sortino.
    const steady = Array.from({ length: 24 }, () => ({ date: "x", return: 0.01 }));
    const s = summarizeReturns(steady, 0.04);
    expect(s.sharpeRatio).not.toBeNull();
    expect(s.sharpeRatio).toBeGreaterThan(0);
    // No month falls below the monthly hurdle (0.01 > 0.04/12) → no downside risk.
    expect(s.downsideDeviation).toBeCloseTo(0, 10);
    expect(s.sortinoRatio).toBeNull();
  });

  it("Sortino penalizes only downside deviation", () => {
    const mixed = [
      { date: "a", return: 0.05 },
      { date: "b", return: -0.03 },
      { date: "c", return: 0.04 },
      { date: "d", return: -0.02 },
    ];
    const s = summarizeReturns(mixed, 0.04);
    expect(s.downsideDeviation).toBeGreaterThan(0);
    expect(s.sortinoRatio).not.toBeNull();
    // Downside deviation only counts the two negative months, so it's smaller
    // than total volatility → Sortino magnitude exceeds Sharpe magnitude here.
    expect(Math.abs(s.sortinoRatio)).toBeGreaterThan(Math.abs(s.sharpeRatio));
  });

  it("returns null risk-adjusted ratios for an empty series", () => {
    const s = summarizeReturns([]);
    expect(s.sharpeRatio).toBeNull();
    expect(s.sortinoRatio).toBeNull();
  });

  it("computes annualized volatility as monthly stdev * sqrt(12)", () => {
    const s = summarizeReturns([
      { date: "2020-01", return: 0.02 },
      { date: "2020-02", return: -0.02 },
      { date: "2020-03", return: 0.02 },
      { date: "2020-04", return: -0.02 },
    ]);
    const variance = (4 * 0.02 ** 2) / 3; // sample variance, n-1, mean is 0
    const expected = Math.sqrt(variance) * Math.sqrt(12);
    expect(s.annualizedVolatility).toBeCloseTo(expected, 10);
  });
});

describe("buildHoldingsWeightMap", () => {
  it("normalizes parsed holdings to fractional weights summing to 1", () => {
    const { weights, coveragePct } = buildHoldingsWeightMap([
      { ticker: "spy", pct: 75 },
      { ticker: "AGG", pct: 25 },
    ]);
    expect(weights.SPY).toBeCloseTo(0.75, 10);
    expect(weights.AGG).toBeCloseTo(0.25, 10);
    expect(coveragePct).toBe(100);
  });

  it("normalizes even when input does not sum to 100", () => {
    const { weights } = buildHoldingsWeightMap([
      { ticker: "A", pct: 30 },
      { ticker: "B", pct: 10 },
    ]);
    expect(weights.A).toBeCloseTo(0.75, 10);
    expect(weights.B).toBeCloseTo(0.25, 10);
  });

  it("aggregates duplicate tickers and ignores empty/zero rows", () => {
    const { weights } = buildHoldingsWeightMap([
      { ticker: "A", pct: 40 },
      { ticker: "A", pct: 40 },
      { ticker: "", pct: 20 },
      { ticker: "B", pct: 0 },
    ]);
    expect(weights).toEqual({ A: 1 });
  });

  it("returns empty map when nothing usable", () => {
    expect(buildHoldingsWeightMap([])).toEqual({ weights: {}, coveragePct: 0 });
  });
});
