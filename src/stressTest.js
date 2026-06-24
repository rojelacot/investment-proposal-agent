// Pure helpers for historical stress testing.
//
// A backtest's headline CAGR hides how a portfolio behaved through the worst
// stretches. Stress testing re-slices the SAME real monthly history into known
// crisis windows (2008 GFC, 2020 COVID crash, 2022 bear) and reports the
// cumulative return over each — so a proposal can show resilience, not just
// average performance. Windows with no overlapping data are reported as null
// (skipped), never fabricated.

// Default crisis windows, inclusive of the months given ("YYYY-MM").
export const DEFAULT_SCENARIOS = [
  { key: "gfc2008",  label: "2008 Financial Crisis", start: "2007-10", end: "2009-02" },
  { key: "covid2020", label: "2020 COVID Crash",     start: "2020-01", end: "2020-03" },
  { key: "bear2022",  label: "2022 Bear Market",      start: "2022-01", end: "2022-09" },
];

/** Cumulative return of a growth series ([{date:"YYYY-MM", value}]) between
 *  `start` and `end` (inclusive). Returns null when the series doesn't span at
 *  least two points inside the window. */
export function windowReturn(growthSeries, start, end) {
  if (!Array.isArray(growthSeries) || growthSeries.length < 2) return null;
  const inWindow = growthSeries
    .filter(p => p && p.date >= start && p.date <= end && p.value > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
  if (inWindow.length < 2) return null;
  const first = inWindow[0].value;
  const last = inWindow[inWindow.length - 1].value;
  return last / first - 1;
}

/** Run a set of crisis scenarios against one or two growth series (target and
 *  optional current). Each result includes the per-portfolio return (or null if
 *  uncovered) and whether the scenario had any data at all. */
export function stressTest(targetSeries, currentSeries = null, scenarios = DEFAULT_SCENARIOS) {
  return scenarios.map(s => {
    const target = windowReturn(targetSeries, s.start, s.end);
    const current = currentSeries ? windowReturn(currentSeries, s.start, s.end) : null;
    return {
      key: s.key,
      label: s.label,
      start: s.start,
      end: s.end,
      target,
      current,
      covered: target != null || current != null,
    };
  });
}
