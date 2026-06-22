// Pure math helpers for backtesting "current portfolio" vs. "target (recommended)
// portfolio" using REAL historical monthly price data (fetched separately via
// /api/history/:ticker — see historyClient.js). Nothing in this file makes network
// calls or fabricates numbers; tickers with no available price history are excluded
// from the backtest and that exclusion is reported back, not silently absorbed.
//
// All "backtest" results here are historical/illustrative only. Every consumer of
// this module (modal UI, pptGenerator slide) must show a "past performance does not
// guarantee future results" disclaimer alongside any number computed here.

const NON_TRADABLE = new Set(["", "SMA", "N/A", "CUSTOM"]);

function isTradable(ticker) {
  return !!ticker && !NON_TRADABLE.has(String(ticker).trim().toUpperCase());
}

/** Converts a monthly close-price series [{date:"YYYY-MM", close}] into a monthly
 *  return series [{date, return}], where `return` is the simple month-over-month
 *  % change ending on `date`. */
export function toMonthlyReturns(series) {
  if (!Array.isArray(series) || series.length < 2) return [];
  const sorted = [...series].sort((a, b) => a.date.localeCompare(b.date));
  const out = [];
  for (let i = 1; i < sorted.length; i++) {
    const prev = sorted[i - 1].close;
    const curr = sorted[i].close;
    if (prev > 0 && curr > 0) out.push({ date: sorted[i].date, return: curr / prev - 1 });
  }
  return out;
}

/** Builds a { TICKER: weightFraction } map from a getFunds()-style fund list.
 *  Non-tradable tickers (SMA/N/A/CUSTOM — separately-managed accounts, private
 *  funds, hedge funds with no public price) are substituted with a proxy ticker
 *  from proxyMap[assetClass] when available; otherwise that fund's allocation is
 *  dropped from the backtest entirely (never silently treated as 0-risk cash).
 *  Returns the weight map plus how much of the original allocation is actually
 *  covered, so callers can disclose any gap. */
export function buildTargetWeightMap(funds, proxyMap = {}) {
  const weights = {};
  let totalAlloc = 0;
  let coveredAlloc = 0;
  const excluded = [];

  for (const f of funds) {
    const alloc = Number(f.alloc) || 0;
    if (alloc <= 0) continue;
    totalAlloc += alloc;

    let ticker = (f.ticker || "").trim().toUpperCase();
    let proxied = false;
    if (!isTradable(ticker)) {
      const proxy = (proxyMap[f.assetClass] || "").trim().toUpperCase();
      if (isTradable(proxy)) {
        ticker = proxy;
        proxied = true;
      } else {
        excluded.push({ name: f.name, assetClass: f.assetClass, alloc });
        continue;
      }
    }

    coveredAlloc += alloc;
    weights[ticker] = (weights[ticker] || 0) + alloc / 100;
    if (proxied) weights.__proxied = weights.__proxied || [];
  }

  return {
    weights,
    coveragePct: totalAlloc > 0 ? (coveredAlloc / totalAlloc) * 100 : 0,
    excluded,
  };
}

/** Builds the approximate "current portfolio" weight map for a client whose notes
 *  describe a concentrated single-stock position: concentrationPct in `ticker`,
 *  and the remainder in `benchmarkTicker` (a diversified proxy, e.g. a 60/40
 *  allocation ETF) standing in for "everything else" since the freeform notes
 *  don't give us a structured breakdown of the rest of the holdings. This is
 *  necessarily an approximation — callers should disclose the assumption. */
export function buildConcentratedWeightMap(ticker, concentrationPct, benchmarkTicker) {
  const t = (ticker || "").trim().toUpperCase();
  const b = (benchmarkTicker || "").trim().toUpperCase();
  const conc = Math.min(Math.max(Number(concentrationPct) || 0, 0), 100) / 100;
  if (!isTradable(t) || conc <= 0) return { weights: {}, coveragePct: 0 };

  const weights = { [t]: conc };
  let coveragePct = conc * 100;
  if (conc < 1 && isTradable(b)) {
    weights[b] = (weights[b] || 0) + (1 - conc);
    coveragePct = 100;
  }
  return { weights, coveragePct };
}

/** Combines per-ticker monthly return series into one weighted portfolio return
 *  series, using only the dates where every weighted ticker has data (so the
 *  comparison spans a single common, apples-to-apples window). Weights are
 *  renormalized across whatever tickers actually have data. */
export function weightedPortfolioReturns(weights, returnsByTicker) {
  const tickers = Object.keys(weights).filter(
    t => t !== "__proxied" && weights[t] > 0 && returnsByTicker[t]?.length
  );
  if (tickers.length === 0) return [];

  const byDate = {};
  for (const t of tickers) {
    for (const { date, return: r } of returnsByTicker[t]) {
      byDate[date] = byDate[date] || {};
      byDate[date][t] = r;
    }
  }

  const dates = Object.keys(byDate)
    .filter(d => tickers.every(t => byDate[d][t] != null))
    .sort();

  const totalWeight = tickers.reduce((s, t) => s + weights[t], 0) || 1;

  return dates.map(date => ({
    date,
    return: tickers.reduce((sum, t) => sum + (weights[t] / totalWeight) * byDate[date][t], 0),
  }));
}

/** Splits a monthly return series into up to `maxYears` trailing, non-overlapping
 *  12-month blocks, walking backward from the most recent month, and compounds
 *  each block into a single annual return. Using rolling 12-month blocks (rather
 *  than calendar years) avoids distortion from a partial in-progress year and
 *  works regardless of what month a proposal happens to be run in. Any leftover
 *  partial block at the oldest end (<12 months) is dropped — every value here is
 *  a real, fully-realized 12-month return, never an extrapolated partial year. */
export function computeTrailingAnnualReturns(monthlyReturns, maxYears = 10) {
  if (!Array.isArray(monthlyReturns) || monthlyReturns.length < 12) return [];
  const sorted = [...monthlyReturns].sort((a, b) => a.date.localeCompare(b.date));
  const out = [];
  let end = sorted.length;
  while (end - 12 >= 0 && out.length < maxYears) {
    const block = sorted.slice(end - 12, end);
    const compounded = block.reduce((acc, m) => acc * (1 + m.return), 1) - 1;
    out.push({ endDate: block[block.length - 1].date, return: compounded });
    end -= 12;
  }
  return out.reverse(); // chronological order, oldest first
}

/** A single ticker's "annualized return" as the simple average of its last
 *  (up to) `maxYears` trailing 12-month returns — not a CAGR/compounding figure.
 *  Returns null when there isn't even one full 12-month period of real price
 *  history, rather than guessing. */
export function averageAnnualReturn(monthlyReturns, maxYears = 10) {
  const years = computeTrailingAnnualReturns(monthlyReturns, maxYears);
  if (years.length === 0) return null;
  const value = years.reduce((s, y) => s + y.return, 0) / years.length;
  return { value, yearsUsed: years.length, years };
}

/** Portfolio-level version: averages each held ticker's own trailing-annual-return
 *  average, then combines them weighted by allocation (renormalized across only
 *  the tickers that actually have enough history — same "never silently treat
 *  missing data as 0" rule as the rest of this module). `coveragePct` reports how
 *  much of the input weight map was actually represented in the result. */
export function weightedAverageAnnualReturn(weights, returnsByTicker, maxYears = 10) {
  const tickers = Object.keys(weights).filter(t => t !== "__proxied" && weights[t] > 0);
  const totalInputWeight = tickers.reduce((s, t) => s + weights[t], 0);
  if (totalInputWeight === 0) return null;

  const perTicker = [];
  let coveredWeight = 0;
  for (const t of tickers) {
    const avg = averageAnnualReturn(returnsByTicker[t], maxYears);
    if (avg) {
      perTicker.push({ ticker: t, weight: weights[t], ...avg });
      coveredWeight += weights[t];
    }
  }
  if (coveredWeight === 0) return null;

  const value = perTicker.reduce((s, p) => s + (p.weight / coveredWeight) * p.value, 0);
  const minYearsUsed = Math.min(...perTicker.map(p => p.yearsUsed));
  return {
    value,
    coveragePct: (coveredWeight / totalInputWeight) * 100,
    minYearsUsed,
    perTicker,
  };
}

/** Summary stats for a monthly return series: growth of a hypothetical $10,000,
 *  CAGR ("annualizedReturn"), annualized volatility (stdev of monthly returns,
 *  scaled by sqrt(12)), and max drawdown. Returns nulls when there isn't enough
 *  overlapping history to compute a meaningful number — callers should treat
 *  null as "not enough data," never as zero. */
export function summarizeReturns(monthlyReturns) {
  if (!monthlyReturns || monthlyReturns.length === 0) {
    return {
      months: 0, startDate: null, endDate: null,
      cumulativeReturn: null, annualizedReturn: null,
      annualizedVolatility: null, maxDrawdown: null, growthSeries: [],
    };
  }

  let value = 10000;
  let peak = value;
  let maxDrawdown = 0;
  const growthSeries = [{ date: monthlyReturns[0].date, value }];
  for (const { date, return: r } of monthlyReturns) {
    value *= 1 + r;
    peak = Math.max(peak, value);
    maxDrawdown = Math.min(maxDrawdown, (value - peak) / peak);
    growthSeries.push({ date, value });
  }

  const months = monthlyReturns.length;
  const cumulativeReturn = value / 10000 - 1;
  const years = months / 12;
  const annualizedReturn = years > 0 ? Math.pow(1 + cumulativeReturn, 1 / years) - 1 : null;

  const mean = monthlyReturns.reduce((s, { return: r }) => s + r, 0) / months;
  const variance =
    monthlyReturns.reduce((s, { return: r }) => s + (r - mean) ** 2, 0) / Math.max(months - 1, 1);
  const annualizedVolatility = Math.sqrt(variance) * Math.sqrt(12);

  return {
    months,
    startDate: monthlyReturns[0].date,
    endDate: monthlyReturns[monthlyReturns.length - 1].date,
    cumulativeReturn,
    annualizedReturn,
    annualizedVolatility,
    maxDrawdown,
    growthSeries,
  };
}
