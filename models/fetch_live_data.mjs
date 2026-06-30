// Live Step-4 data pull. Replaces the seeded/calibrated series with REAL data.
//
// Sources:
//   - Yahoo Finance via the yahoo-finance2 client (same data as the Python
//     `yfinance` package; that package is not installed here and direct Yahoo
//     access 429s this IP, while the Node client handles cookies/crumb and is
//     already a project dependency). Proxies: SPY (US equity), AGG (IG bonds),
//     VNQ+DBC 50/50 (diversified alternatives).
//   - FRED CSV (no key required): DGS3MO (cash / risk-free), CPIAUCSL (CPI).
//
// Cross-reference rule (CLAUDE.md Step 4): each return is cross-checked against a
// documented second source; if the two diverge >15%, the LOWER value is adopted
// and flagged. Everything is logged to data/source_log.json.
//
// Outputs: data/cma.json, data/historical_returns.csv, data/source_log.json
import { createRequire } from "module";
import { writeFileSync } from "fs";
import path from "path";
import { fileURLToPath } from "url";

const require = createRequire(import.meta.url);
const { default: YF } = require("yahoo-finance2");
const yf = new YF({ suppressNotices: ["ripHistorical"], validation: { logErrors: false, logOptionsErrors: false } });

const HERE = path.dirname(fileURLToPath(import.meta.url));
const DATA = path.join(HERE, "..", "data");
const RETRIEVAL = new Date().toISOString().slice(0, 10);

// Documented secondary reference returns (real, long-run published figures).
const SECONDARY = {
  us_equity:    { value: 0.100, name: "Damodaran NYU Stern, S&P 500 geometric 1928-2024",
                  url: "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histretSP.html" },
  bonds:        { value: 0.048, name: "Damodaran NYU Stern, 10-yr T.Bond geometric 1928-2024",
                  url: "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histretSP.html" },
  alternatives: { value: 0.057, name: "FTSE NAREIT (REIT) + S&P GSCI (commodity) long-run blend",
                  url: "https://www.reit.com/data-research/reit-indexes/annual-index-values-returns" },
};

async function monthlyCloses(ticker) {
  const r = await yf.chart(ticker, { period1: new Date("1990-01-01"), interval: "1mo" });
  const out = [];
  for (const q of r.quotes) {
    const px = q.adjclose ?? q.close;
    if (px != null && px > 0 && q.date) {
      out.push({ ym: new Date(q.date).toISOString().slice(0, 7), px: Number(px) });
    }
  }
  // de-dup by month (keep last), sort
  const map = new Map(out.map((o) => [o.ym, o.px]));
  return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([ym, px]) => ({ ym, px }));
}

function monthlyReturns(series) {
  const out = [];
  for (let i = 1; i < series.length; i++) {
    out.push({ ym: series[i].ym, r: series[i].px / series[i - 1].px - 1 });
  }
  return out;
}

const mean = (a) => a.reduce((s, x) => s + x, 0) / a.length;
const std = (a) => {
  const m = mean(a);
  return Math.sqrt(a.reduce((s, x) => s + (x - m) ** 2, 0) / (a.length - 1));
};

function annualizedGeo(returns) {
  const growth = returns.reduce((g, x) => g * (1 + x), 1);
  return Math.pow(growth, 12 / returns.length) - 1;
}

function pearson(a, b) {
  const ma = mean(a), mb = mean(b);
  let num = 0, da = 0, db = 0;
  for (let i = 0; i < a.length; i++) { num += (a[i] - ma) * (b[i] - mb); da += (a[i] - ma) ** 2; db += (b[i] - mb) ** 2; }
  return num / Math.sqrt(da * db);
}

async function fredCsv(id) {
  const res = await fetch(`https://fred.stlouisfed.org/graph/fredgraph.csv?id=${id}`);
  if (!res.ok) throw new Error(`FRED ${id} HTTP ${res.status}`);
  const text = await res.text();
  const rows = text.trim().split("\n").slice(1).map((l) => l.split(","));
  // header is "observation_date,<ID>" or "DATE,<ID>"
  return rows
    .map(([date, val]) => ({ date: date.trim(), val: parseFloat(val) }))
    .filter((o) => Number.isFinite(o.val));
}

async function main() {
  console.log("Pulling Yahoo Finance monthly closes…");
  const [spy, agg, vnq, dbc] = await Promise.all(
    ["SPY", "AGG", "VNQ", "DBC"].map(monthlyCloses)
  );
  for (const [t, s] of [["SPY", spy], ["AGG", agg], ["VNQ", vnq], ["DBC", dbc]]) {
    console.log(`  ${t}: ${s.length} months ${s[0].ym}..${s[s.length - 1].ym}`);
  }

  const rSpy = monthlyReturns(spy);
  const rAgg = monthlyReturns(agg);
  const rVnq = monthlyReturns(vnq);
  const rDbc = monthlyReturns(dbc);

  // Alternatives = 50/50 VNQ + DBC on overlapping months.
  const dbcMap = new Map(rDbc.map((o) => [o.ym, o.r]));
  const rAlt = rVnq.filter((o) => dbcMap.has(o.ym)).map((o) => ({ ym: o.ym, r: 0.5 * o.r + 0.5 * dbcMap.get(o.ym) }));

  const primary = {
    us_equity:    { ret: annualizedGeo(rSpy.map((o) => o.r)), vol: std(rSpy.map((o) => o.r)) * Math.sqrt(12),
                    proxy: "SPY", win: `${rSpy[0].ym}..${rSpy[rSpy.length - 1].ym}` },
    bonds:        { ret: annualizedGeo(rAgg.map((o) => o.r)), vol: std(rAgg.map((o) => o.r)) * Math.sqrt(12),
                    proxy: "AGG", win: `${rAgg[0].ym}..${rAgg[rAgg.length - 1].ym}` },
    alternatives: { ret: annualizedGeo(rAlt.map((o) => o.r)), vol: std(rAlt.map((o) => o.r)) * Math.sqrt(12),
                    proxy: "VNQ+DBC 50/50", win: `${rAlt[0].ym}..${rAlt[rAlt.length - 1].ym}` },
  };

  console.log("Pulling FRED (DGS3MO, CPIAUCSL)…");
  const [dgs3mo, cpi] = await Promise.all([fredCsv("DGS3MO"), fredCsv("CPIAUCSL")]);
  const cashYield = mean(dgs3mo.slice(-252).map((o) => o.val)) / 100; // trailing ~1yr avg, to decimal
  const cpiYoY = cpi.at(-1).val / cpi.at(-13).val - 1;

  // Apply cross-reference / >15% divergence rule.
  const sourceLog = [];
  const adopted = {};
  for (const k of ["us_equity", "bonds", "alternatives"]) {
    const p = primary[k].ret, s = SECONDARY[k].value;
    const diverge = Math.abs(p - s) / Math.min(p, s) > 0.15;
    const useVal = diverge ? Math.min(p, s) : p;
    adopted[k] = { ret: useVal, vol: primary[k].vol, diverge };
    sourceLog.push({
      field: `${k} expected_return (primary, trailing realized)`,
      value: p.toFixed(4),
      source_name: `Yahoo Finance via yahoo-finance2 — ${primary[k].proxy}`,
      source_url: "https://finance.yahoo.com/",
      date_range: primary[k].win, retrieval_date: RETRIEVAL,
    });
    sourceLog.push({
      field: `${k} expected_return (cross-reference)`,
      value: s.toFixed(4),
      source_name: SECONDARY[k].name, source_url: SECONDARY[k].url,
      date_range: "long-run published", retrieval_date: RETRIEVAL,
      note: diverge ? `Diverges >15% from primary; adopted LOWER (${useVal.toFixed(4)}) and flagged.`
                    : `Within 15%; adopted primary trailing figure (${useVal.toFixed(4)}).`,
    });
    sourceLog.push({
      field: `${k} volatility (annualized monthly)`,
      value: primary[k].vol.toFixed(4),
      source_name: `Yahoo Finance via yahoo-finance2 — ${primary[k].proxy}`,
      source_url: "https://finance.yahoo.com/",
      date_range: primary[k].win, retrieval_date: RETRIEVAL,
    });
  }
  sourceLog.push({
    field: "cash expected_return / risk_free_rate", value: cashYield.toFixed(4),
    source_name: "FRED 3-Month Treasury Constant Maturity (DGS3MO), trailing ~1yr avg",
    source_url: "https://fred.stlouisfed.org/series/DGS3MO",
    date_range: `${dgs3mo.slice(-252)[0].date}..${dgs3mo.at(-1).date}`, retrieval_date: RETRIEVAL,
  });
  sourceLog.push({
    field: "inflation_cpi (YoY)", value: cpiYoY.toFixed(4),
    source_name: "FRED CPI All Urban Consumers (CPIAUCSL), 12-month change",
    source_url: "https://fred.stlouisfed.org/series/CPIAUCSL",
    date_range: `${cpi.at(-13).date}..${cpi.at(-1).date}`, retrieval_date: RETRIEVAL,
  });

  // Correlations from overlapping monthly returns (equity/bonds/alts); cash ~ uncorrelated.
  const common = rAlt.map((o) => o.ym).filter((ym) => {
    return rSpy.find((x) => x.ym === ym) && rAgg.find((x) => x.ym === ym);
  });
  const pick = (arr) => common.map((ym) => arr.find((x) => x.ym === ym).r);
  const eS = pick(rSpy), bS = pick(rAgg), aS = rAlt.filter((o) => common.includes(o.ym)).map((o) => o.r);
  const corr = {
    us_equity:    { us_equity: 1, bonds: pearson(eS, bS), alternatives: pearson(eS, aS), cash: 0.0 },
    bonds:        { us_equity: pearson(bS, eS), bonds: 1, alternatives: pearson(bS, aS), cash: 0.10 },
    alternatives: { us_equity: pearson(aS, eS), bonds: pearson(aS, bS), alternatives: 1, cash: 0.0 },
    cash:         { us_equity: 0.0, bonds: 0.10, alternatives: 0.0, cash: 1 },
  };
  const round = (x) => Math.round(x * 100) / 100;
  for (const a of Object.keys(corr)) for (const b of Object.keys(corr[a])) corr[a][b] = round(corr[a][b]);

  const cma = {
    as_of: RETRIEVAL,
    description: "LIVE Step-4 assumptions. Per-asset returns are trailing realized geometric returns from Yahoo Finance proxies (longest available window), cross-referenced against published long-run sources; where divergence exceeded 15% the LOWER figure was adopted. Volatilities are annualized from monthly returns; correlations are sample correlations of monthly returns. Cash/risk-free and CPI from FRED.",
    risk_free_rate: round(cashYield * 100) / 100,
    inflation_cpi: round(cpiYoY * 100) / 100,
    assets: {
      us_equity:    { label: "US Large-Cap Equity (SPY)", expected_return: round(adopted.us_equity.ret * 100) / 100, volatility: round(adopted.us_equity.vol * 100) / 100, divergence_flag: adopted.us_equity.diverge },
      bonds:        { label: "US Investment-Grade Bonds (AGG)", expected_return: round(adopted.bonds.ret * 100) / 100, volatility: round(adopted.bonds.vol * 100) / 100, divergence_flag: adopted.bonds.diverge },
      alternatives: { label: "Diversified Alternatives (VNQ+DBC 50/50)", expected_return: round(adopted.alternatives.ret * 100) / 100, volatility: round(adopted.alternatives.vol * 100) / 100, divergence_flag: adopted.alternatives.diverge },
      cash:         { label: "Cash / 3-Month T-Bills (FRED DGS3MO)", expected_return: round(cashYield * 100) / 100, volatility: 0.01, divergence_flag: false },
    },
    correlations: corr,
  };
  writeFileSync(path.join(DATA, "cma.json"), JSON.stringify(cma, null, 2) + "\n");
  writeFileSync(path.join(DATA, "source_log.json"), JSON.stringify(sourceLog, null, 2) + "\n");

  // historical_returns.csv: real calendar-year returns over the common full-year window.
  const annual = (returns) => {
    const byYear = new Map();
    for (const { ym, r } of returns) {
      const y = ym.slice(0, 4);
      byYear.set(y, (byYear.get(y) ?? { g: 1, n: 0 }));
      const e = byYear.get(y); e.g *= 1 + r; e.n += 1;
    }
    return byYear;
  };
  const yEq = annual(rSpy), yBo = annual(rAgg), yAl = annual(rAlt);
  const cashByYear = new Map();
  for (const { date, val } of dgs3mo) {
    const y = date.slice(0, 4);
    cashByYear.set(y, (cashByYear.get(y) ?? []));
    cashByYear.get(y).push(val / 100);
  }
  // Only full (12-month) years present in all equity/bond/alt series.
  const years = [...yAl.keys()].filter((y) =>
    yEq.has(y) && yBo.has(y) && yAl.get(y).n === 12 && yEq.get(y).n === 12 && yBo.get(y).n === 12 && cashByYear.has(y)
  ).sort();
  const lines = ["year,us_equity,bonds,alternatives,cash"];
  for (const y of years) {
    const cash = cashByYear.get(y).reduce((s, x) => s + x, 0) / cashByYear.get(y).length;
    lines.push([y, (yEq.get(y).g - 1).toFixed(6), (yBo.get(y).g - 1).toFixed(6),
                (yAl.get(y).g - 1).toFixed(6), cash.toFixed(6)].join(","));
  }
  writeFileSync(path.join(DATA, "historical_returns.csv"), lines.join("\n") + "\n");

  console.log(`\nAdopted CMA:`, JSON.stringify(cma.assets, null, 2));
  console.log(`risk_free=${cma.risk_free_rate} inflation=${cma.inflation_cpi}`);
  console.log(`historical_returns.csv: ${years.length} full years ${years[0]}..${years.at(-1)}`);
}

main().catch((e) => { console.error("FETCH FAILED:", e); process.exit(1); });
