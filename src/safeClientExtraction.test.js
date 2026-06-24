import { describe, it, expect } from "vitest";
import { parseCSVRow, parseHoldingsFromText } from "./safeClientExtraction.js";

// The holdings parser turns an uploaded statement (CSV/TSV/PDF text) into the
// {ticker, pct} weights that drive the "current portfolio" backtest. A parsing
// bug here silently feeds wrong allocations into a client proposal, so these
// tests pin down the RFC 4180 quoting fix and both extraction strategies.

describe("parseCSVRow (RFC 4180)", () => {
  it("splits a plain row", () => {
    expect(parseCSVRow("AAPL,Apple Inc.,35.2%")).toEqual(["AAPL", "Apple Inc.", "35.2%"]);
  });

  it("keeps commas inside double-quoted fields together (the headline fix)", () => {
    // Without quote awareness this would become ["NVDA", '"$46', '500', '000"'].
    expect(parseCSVRow('"NVDA","$46,500,000"')).toEqual(["NVDA", "$46,500,000"]);
  });

  it("handles escaped double-quotes inside a quoted field", () => {
    expect(parseCSVRow('"a","b,c","d""e"')).toEqual(["a", "b,c", 'd"e']);
  });

  it("trims whitespace around cells and preserves trailing empties", () => {
    expect(parseCSVRow("AAPL ,  60% , ")).toEqual(["AAPL", "60%", ""]);
  });
});

describe("parseHoldingsFromText — short / empty input", () => {
  it("returns [] for text under 50 chars", () => {
    expect(parseHoldingsFromText("AAPL 50%")).toEqual([]);
    expect(parseHoldingsFromText("")).toEqual([]);
    expect(parseHoldingsFromText(null)).toEqual([]);
  });

  it("returns [] when only one holding can be found", () => {
    const text = "Portfolio statement for review. Only one position present: AAPL 100%.";
    expect(parseHoldingsFromText(text)).toEqual([]);
  });
});

describe("parseHoldingsFromText — Strategy 1: structured CSV with header", () => {
  it("parses quoted market-value columns and normalizes to percentages", () => {
    const csv = [
      "Symbol,Market Value",
      '"NVDA","$46,500,000"',
      '"AAPL","$23,250,000"',
    ].join("\n");
    const out = parseHoldingsFromText(csv);
    expect(out).toHaveLength(2);
    const byTicker = Object.fromEntries(out.map(h => [h.ticker, h.pct]));
    expect(byTicker.NVDA).toBeCloseTo(66.6667, 3);
    expect(byTicker.AAPL).toBeCloseTo(33.3333, 3);
  });

  it("parses an explicit percent/weight column directly", () => {
    // Inputs must clear the parser's 50-char minimum guard, as real statements do.
    const csv = [
      "Ticker,Name,Weight",
      "AAPL,Apple Inc,50%",
      "MSFT,Microsoft Corp,30%",
      "GOOG,Alphabet Inc,20%",
    ].join("\n");
    const out = parseHoldingsFromText(csv);
    expect(Object.fromEntries(out.map(h => [h.ticker, h.pct]))).toEqual({
      AAPL: expect.closeTo(50, 6),
      MSFT: expect.closeTo(30, 6),
      GOOG: expect.closeTo(20, 6),
    });
  });

  it("supports tab-separated values", () => {
    const tsv = ["Symbol\tMarket Value", "NVDA\t$30,000,000", "AAPL\t$10,000,000"].join("\n");
    const out = parseHoldingsFromText(tsv);
    expect(Object.fromEntries(out.map(h => [h.ticker, h.pct]))).toEqual({
      NVDA: expect.closeTo(75, 6),
      AAPL: expect.closeTo(25, 6),
    });
  });

  it("normalizes share-class tickers to Yahoo format (BRK.B -> BRK-B)", () => {
    const csv = [
      "Symbol,Market Value",
      "BRK.B,1000000",
      "AAPL,1000000",
      "MSFT,1000000",
    ].join("\n");
    const out = parseHoldingsFromText(csv);
    expect(out.map(h => h.ticker).sort()).toEqual(["AAPL", "BRK-B", "MSFT"]);
  });

  it("skips non-ticker rows like TOTAL / CASH", () => {
    const csv = [
      "Symbol,Market Value",
      "AAPL,60",
      "MSFT,40",
      "TOTAL,100",
      "CASH,5",
    ].join("\n");
    const out = parseHoldingsFromText(csv);
    expect(out.map(h => h.ticker).sort()).toEqual(["AAPL", "MSFT"]);
  });

  it("deduplicates repeated tickers by summing", () => {
    const csv = [
      "Symbol,Market Value",
      "AAPL,3000000",
      "AAPL,3000000",
      "MSFT,4000000",
    ].join("\n");
    const out = parseHoldingsFromText(csv);
    const byTicker = Object.fromEntries(out.map(h => [h.ticker, h.pct]));
    expect(byTicker.AAPL).toBeCloseTo(60, 6);
    expect(byTicker.MSFT).toBeCloseTo(40, 6);
  });
});

describe("parseHoldingsFromText — Strategy 2: free-form text / PDF", () => {
  it("extracts ticker + percent pairs from prose lines", () => {
    const text = [
      "Portfolio Holdings Report as of Q1 2026 — review copy",
      "AAPL Apple Inc. 35.5%",
      "MSFT Microsoft Corp 24.5%",
      "Cash and equivalents 40%",
    ].join("\n");
    const out = parseHoldingsFromText(text);
    // Cash line has no valid ticker, so it's dropped; remaining two renormalize.
    expect(out.map(h => h.ticker).sort()).toEqual(["AAPL", "MSFT"]);
    const byTicker = Object.fromEntries(out.map(h => [h.ticker, h.pct]));
    expect(byTicker.AAPL).toBeCloseTo((35.5 / 60) * 100, 4);
    expect(byTicker.MSFT).toBeCloseTo((24.5 / 60) * 100, 4);
  });

  it("ignores out-of-range percentages (>99 or <0.1)", () => {
    const text = [
      "Account performance summary for the trailing period under review",
      "AAPL returned 120% over five years",
      "MSFT returned 0.05% last month",
      "VTI 50%",
      "AGG 50%",
    ].join("\n");
    const out = parseHoldingsFromText(text);
    expect(out.map(h => h.ticker).sort()).toEqual(["AGG", "VTI"]);
  });
});
