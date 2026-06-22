import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { createRequire } from "module";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5174;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const FINNHUB_API_KEY = process.env.FINNHUB_API_KEY;

// ── Yahoo Finance via yahoo-finance2 library ─────────────────────────────────
// Root cause of the 429 problem: Node's bare fetch() lacks the session cookies
// and browser fingerprint that Yahoo requires. yahoo-finance2 handles all of
// that internally (cookie jar, crumb, proper headers) and is the standard
// community solution for server-side Yahoo data access.
const require = createRequire(import.meta.url);
const { default: YahooFinanceClass } = require("yahoo-finance2");
const yahooFinance = new YahooFinanceClass({
  suppressNotices: ["ripHistorical"],
  validation: { logErrors: false, logOptionsErrors: false },
});

// ── Previous close via Finnhub (primary) then Yahoo (fallback) ───────────────
async function fetchPreviousClose(ticker) {
  // 1. Finnhub — fast, reliable, free with an API key
  if (FINNHUB_API_KEY) {
    try {
      const r = await fetch(
        `https://finnhub.io/api/v1/quote?symbol=${encodeURIComponent(ticker)}&token=${FINNHUB_API_KEY}`
      );
      if (r.ok) {
        const d = await r.json();
        const price = Number(d.pc) || Number(d.c);
        console.log(`[fetchPreviousClose] ${ticker} Finnhub pc:${d.pc} c:${d.c}`);
        if (price > 0) return price;
      }
    } catch (e) {
      console.warn(`[fetchPreviousClose] Finnhub error: ${e.message}`);
    }
  }

  // 2. Yahoo Finance via library (handles cookies/crumb automatically)
  try {
    const q = await yahooFinance.quote(ticker);
    const price = Number(q.regularMarketPrice || q.regularMarketPreviousClose || 0);
    console.log(`[fetchPreviousClose] ${ticker} Yahoo price:${price}`);
    if (price > 0) return price;
  } catch (e) {
    console.warn(`[fetchPreviousClose] Yahoo error: ${e.message}`);
  }

  return null;
}

// ── /api/quote/:ticker ───────────────────────────────────────────────────────

app.get("/api/quote/:ticker", async (req, res) => {
  try {
    const ticker = String(req.params.ticker || "").trim().toUpperCase();
    if (!ticker) return res.status(400).json({ error: "Missing ticker." });

    const price = await fetchPreviousClose(ticker);
    if (price) {
      return res.json({
        ticker,
        currentPrice: price,
        previousClose: price,
        timestamp: new Date().toISOString(),
        source: "Yahoo Finance",
      });
    }

    return res.status(502).json({ error: "Could not fetch price from any source." });
  } catch (err) {
    res.status(500).json({ error: err.message || "Unknown quote error." });
  }
});

// ── /api/options/:ticker ─────────────────────────────────────────────────────

app.get("/api/options/:ticker", async (req, res) => {
  try {
    const ticker = String(req.params.ticker || "").trim().toUpperCase();
    if (!ticker) return res.status(400).json({ error: "Missing ticker." });

    const previousClose = await fetchPreviousClose(ticker);
    if (!previousClose) return res.status(502).json({ error: "Could not fetch price." });

    const putStrike  = +(previousClose * 0.85).toFixed(2);
    const callStrike = +(previousClose * 1.15).toFixed(2);

    res.json({
      ticker,
      currentPrice: previousClose,
      previousClose,
      asOfDate: new Date().toISOString().slice(0, 10),
      expirationLabel: "~12 months (illustrative)",
      put:  { strike: putStrike,  premium: null, impliedVolatility: null },
      call: { strike: callStrike, premium: null, impliedVolatility: null },
      netCost: null,
      netCostPerShare: "See advisor",
      source: "Yahoo Finance (prior close)",
      priceOnly: true,
    });
  } catch (err) {
    res.status(500).json({ error: err.message || "Unknown options error." });
  }
});

// ── /api/history/:ticker ─────────────────────────────────────────────────────
// Monthly close history for backtesting. Real market data only — no synthetic
// or fabricated prices. Returns null on failure so callers can skip that ticker.

const historyCache = new Map(); // ticker:range -> { ts, data }
const HISTORY_CACHE_MS = 1000 * 60 * 60 * 6; // 6h

// Map the range string to how many years of history to request (adding 1y buffer
// so even the most recent 10 trailing-12-month periods are fully covered).
const RANGE_YEARS = { "10y": 11, "5y": 6, "2y": 3, "1y": 2 };

app.get("/api/history/:ticker", async (req, res) => {
  try {
    const ticker = String(req.params.ticker || "").trim().toUpperCase();
    if (!ticker) return res.status(400).json({ error: "Missing ticker." });
    const range = String(req.query.range || "10y");

    const cacheKey = `${ticker}:${range}`;
    const cached = historyCache.get(cacheKey);
    if (cached && Date.now() - cached.ts < HISTORY_CACHE_MS) {
      return res.json(cached.data);
    }

    const years = RANGE_YEARS[range] || 11;
    const period1 = new Date();
    period1.setFullYear(period1.getFullYear() - years);

    console.log(`[history] fetching ${ticker} (${range}, back ${years}y from ${period1.toISOString().slice(0,7)})`);

    let result;
    try {
      result = await yahooFinance.chart(ticker, { period1, interval: "1mo" });
    } catch (e) {
      console.warn(`[history] ${ticker} -> yahoo-finance2 threw: ${e.message}`);
      return res.status(500).json({ error: e.message || "Yahoo Finance error.", ticker });
    }

    const quotes = result?.quotes;
    if (!quotes?.length) {
      console.warn(`[history] ${ticker} -> no quotes in response`);
      return res.status(404).json({ error: `No history available for ${ticker}.`, ticker });
    }

    const series = quotes
      .filter(q => q.adjclose != null || q.close != null)
      .map(q => ({
        date: (q.date instanceof Date ? q.date : new Date(q.date))
          .toISOString().slice(0, 7),           // "YYYY-MM"
        close: Number(q.adjclose ?? q.close),
      }))
      .filter(q => q.close > 0);

    if (series.length < 2) {
      return res.status(404).json({ error: `Insufficient history for ${ticker}.`, ticker });
    }

    console.log(`[history] ${ticker} -> ${series.length} monthly closes (${series[0].date} – ${series[series.length-1].date})`);

    const payload = { ticker, range, series, source: "Yahoo Finance" };
    historyCache.set(cacheKey, { ts: Date.now(), data: payload });
    res.json(payload);
  } catch (err) {
    console.error(`[history] ${req.params.ticker} -> ${err.message}`);
    res.status(500).json({ error: err.message || "Unknown history error.", ticker: req.params.ticker });
  }
});

// Optional: serve built app from dist
app.use(express.static(path.join(__dirname, "dist")));

app.use((req, res) => {
  res.sendFile(path.join(__dirname, "dist", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Quote server running at http://localhost:${PORT}`);
});
