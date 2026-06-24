import { describe, it, expect } from "vitest";
import { windowReturn, stressTest, DEFAULT_SCENARIOS } from "./stressTest.js";

const series = [
  { date: "2007-09", value: 100 },
  { date: "2007-10", value: 110 }, // window start
  { date: "2008-06", value: 90 },
  { date: "2009-02", value: 77 },  // window end
  { date: "2009-06", value: 95 },
];

describe("windowReturn", () => {
  it("computes cumulative return within an inclusive window", () => {
    // 110 → 77 over 2007-10..2009-02 = -30%
    expect(windowReturn(series, "2007-10", "2009-02")).toBeCloseTo(-0.3, 10);
  });

  it("ignores points outside the window", () => {
    // 2007-09 (100) and 2009-06 (95) are excluded.
    const r = windowReturn(series, "2007-10", "2009-02");
    expect(r).not.toBeCloseTo(77 / 100 - 1, 10);
  });

  it("returns null when fewer than two points fall in the window", () => {
    expect(windowReturn(series, "2010-01", "2010-12")).toBeNull();
    expect(windowReturn([{ date: "2008-01", value: 100 }], "2008-01", "2008-12")).toBeNull();
  });
});

describe("stressTest", () => {
  it("reports target and current per scenario, with coverage flags", () => {
    const target = [
      { date: "2020-01", value: 100 },
      { date: "2020-03", value: 70 }, // -30% COVID
    ];
    const current = [
      { date: "2020-01", value: 100 },
      { date: "2020-03", value: 60 }, // -40% COVID
    ];
    const res = stressTest(target, current);
    const covid = res.find(r => r.key === "covid2020");
    expect(covid.target).toBeCloseTo(-0.3, 10);
    expect(covid.current).toBeCloseTo(-0.4, 10);
    expect(covid.covered).toBe(true);
    // No data for 2008 → both null, not covered.
    const gfc = res.find(r => r.key === "gfc2008");
    expect(gfc.target).toBeNull();
    expect(gfc.covered).toBe(false);
  });

  it("handles a missing current series", () => {
    const target = [
      { date: "2022-01", value: 100 },
      { date: "2022-09", value: 80 },
    ];
    const res = stressTest(target, null);
    const bear = res.find(r => r.key === "bear2022");
    expect(bear.target).toBeCloseTo(-0.2, 10);
    expect(bear.current).toBeNull();
    expect(bear.covered).toBe(true);
  });

  it("returns one entry per default scenario", () => {
    expect(stressTest([], null)).toHaveLength(DEFAULT_SCENARIOS.length);
  });
});
