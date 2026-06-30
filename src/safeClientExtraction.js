function cleanNumber(value) {
  return Number(String(value || "").replace(/,/g, ""));
}

function toMillions(value, unit = "M") {
  const n = cleanNumber(value);
  if (!Number.isFinite(n)) return null;

  const u = String(unit || "M").toLowerCase();

  if (u === "b" || u === "billion") return n * 1000;
  if (u === "k" || u === "thousand") return n / 1000;
  if (u === "m" || u === "mm" || u === "million") return n;

  // No unit:
  // 86,000,000 should become 86
  // 34,200,000 should become 34.2
  // 86 should stay 86 because users often mean $86M in these notes
  if (n >= 1_000_000) return n / 1_000_000;

  return n;
}

function normalizeMillionValue(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;

  // If old extractor returned raw dollars, convert to millions.
  if (n >= 1_000_000) return n / 1_000_000;

  return n;
}

function moneyRegex() {
  return String.raw`\$?\s*([\d,.]+)\s*(MM|M|million|B|billion|K|thousand)?`;
}

function findMoney(text, patterns) {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return toMillions(match[1], match[2]);
    }
  }
  return null;
}

function findPercent(text, patterns) {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const n = cleanNumber(match[1]);
      if (Number.isFinite(n)) return n;
    }
  }
  return null;
}

// IRS Table S approximation: remainder factor at 5% payout by single-life age
const CRT_DEDUCTION_TABLE = [
  [45, 0.215], [50, 0.245], [55, 0.278], [60, 0.314],
  [65, 0.356], [70, 0.404], [75, 0.458], [80, 0.520], [85, 0.584],
];

function crtDeductionFactor(age) {
  const a = Math.max(45, Math.min(85, age || 65));
  for (let i = 0; i < CRT_DEDUCTION_TABLE.length - 1; i++) {
    const [a0, f0] = CRT_DEDUCTION_TABLE[i];
    const [a1, f1] = CRT_DEDUCTION_TABLE[i + 1];
    if (a >= a0 && a <= a1) return f0 + (f1 - f0) * (a - a0) / (a1 - a0);
  }
  return 0.356;
}

function extractClientAge(notes) {
  const currentYear = new Date().getFullYear();

  const agePatterns = [
    /\bage[d]?\s*[:-]?\s*(\d{2})\b/i,
    /(\d{2})\s+years?\s+old/i,
    /client\s+is\s+(\d{2})\s+years?/i,
    /(\d{2})\s*[-–]\s*year[-\s]old/i,
  ];
  for (const p of agePatterns) {
    const m = notes.match(p);
    if (m) { const a = parseInt(m[1]); if (a >= 35 && a <= 90) return a; }
  }

  const dobPatterns = [
    /born\s+(?:in\s+)?(\d{4})/i,
    /\bDOB\b[:\s]+\d{1,2}\/\d{1,2}\/(\d{4})/i,
    /date\s+of\s+birth[:\s]+\d{1,2}\/\d{1,2}\/(\d{4})/i,
  ];
  for (const p of dobPatterns) {
    const m = notes.match(p);
    if (m) { const yr = parseInt(m[1]); if (yr >= 1930 && yr <= currentYear - 35) return currentYear - yr; }
  }

  return null;
}

function extractTicker(notes) {
  const upper = notes.toUpperCase();

  const directPatterns = [
    /\bticker\s*[:-]?\s*([A-Z]{1,5})\b/i,
    /\bsymbol\s*[:-]?\s*([A-Z]{1,5})\b/i,
    /\b([A-Z]{2,5})\s+(?:position|stock|shares|holding|holdings)\b/i,
    /\b(?:position|stock|shares|holding|holdings)\s+(?:in|of)?\s*([A-Z]{2,5})\b/i,
    /\$[\d,.]+\s*(?:M|MM|million|B|billion|K|thousand)?\s+([A-Z]{2,5})\b/i,
  ];

  for (const pattern of directPatterns) {
    const match = notes.match(pattern);
    if (match?.[1]) {
      const ticker = match[1].toUpperCase();
      const banned = ["THE", "AND", "FOR", "WITH", "HAS", "TAX", "IRA", "AUM"];
      if (!banned.includes(ticker)) return ticker;
    }
  }

  const companyMap = [
    ["NVIDIA", "NVDA"],
    ["TESLA", "TSLA"],
    ["APPLE", "AAPL"],
    ["MICROSOFT", "MSFT"],
    ["AMAZON", "AMZN"],
    ["GOOGLE", "GOOGL"],
    ["ALPHABET", "GOOGL"],
    ["META", "META"],
    ["FACEBOOK", "META"],
    ["BROADCOM", "AVGO"],
    ["AMD", "AMD"],
    ["NETFLIX", "NFLX"],
    ["SALESFORCE", "CRM"],
    ["ORACLE", "ORCL"],
  ];

  for (const [company, ticker] of companyMap) {
    if (upper.includes(company)) return ticker;
  }

  return "";
}

function extractStockPosition(text, ticker) {
  const tickerEscaped = ticker ? ticker.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") : "";

  // Position values must have an explicit money scale.
  // This prevents accidentally pulling 5 from "5% payout rate."
  const scaledMoney = String.raw`\$?\s*([\d,.]+)\s*(MM|M|million|B|billion)`;

  const cleanedText = text
    .replace(/(?:stock|share)\s+price[^.\n]*/gi, " ")
    .replace(/current\s+(?:stock|share)\s+price[^.\n]*/gi, " ")
    .replace(/trades?\s+at[^.\n]*/gi, " ")
    .replace(/trading\s+(?:at|price)[^.\n]*/gi, " ")
    .replace(/payout\s+rate[^.\n]*/gi, " ")
    .replace(/tax\s+rate[^.\n]*/gi, " ")
    .replace(/cost\s+basis[^.\n]*/gi, " ");

  const patterns = [];

  if (tickerEscaped) {
    patterns.push(
      // "$46.5M NVDA position"
      new RegExp(`${scaledMoney}\\s+(?:in\\s+)?\\b${tickerEscaped}\\b\\s*(?:concentrated\\s+)?(?:position|holding|holdings|shares|stock)?`, "i"),

      // "$46.5M in NVDA"
      new RegExp(`${scaledMoney}\\s+in\\s+\\b${tickerEscaped}\\b`, "i"),

      // "NVDA position of $46.5M"
      new RegExp(`\\b${tickerEscaped}\\b\\s+(?:concentrated\\s+)?(?:position|holding|holdings|shares|stock)\\s*(?:of|worth|valued\\s+at|approximately|around|is|=|:)\\s*${scaledMoney}`, "i"),

      // "position in NVDA of $46.5M"
      new RegExp(`(?:concentrated\\s+)?(?:position|holding|holdings|shares|stock)\\s+(?:in|of)\\s+\\b${tickerEscaped}\\b\\s*(?:of|worth|valued\\s+at|approximately|around|is|=|:)\\s*${scaledMoney}`, "i"),

      // "NVDA: $46.5M" or "NVDA - $46.5M"
      new RegExp(`\\b${tickerEscaped}\\b\\s*[:\\-–—]\\s*${scaledMoney}`, "i")
    );
  }

  patterns.push(
    // "concentrated position of $46.5M"
    new RegExp(`(?:concentrated\\s+position|concentrated\\s+stock|single[-\\s]?stock\\s+position|stock\\s+position|equity\\s+position)\\s*(?:of|worth|valued\\s+at|approximately|around|is|=|:)\\s*${scaledMoney}`, "i"),

    // "$46.5M concentrated position"
    new RegExp(`${scaledMoney}\\s+(?:concentrated\\s+position|concentrated\\s+stock|single[-\\s]?stock\\s+position|stock\\s+position|equity\\s+position)`, "i")
  );

  for (const pattern of patterns) {
    const match = cleanedText.match(pattern);
    if (match) {
      const value = toMillions(match[1], match[2]);

      // Guardrail: concentrated stock position should not be a tiny number accidentally pulled from a rate.
      if (value && value >= 0.1) {
        return value;
      }
    }
  }

  return null;
}

function extractSharePrice(text) {
  const pricePatterns = [
    /(?:stock|share)\s+price[^$0-9]{0,40}\$?\s*([\d,.]+)/i,
    /current\s+(?:stock|share)\s+price[^$0-9]{0,40}\$?\s*([\d,.]+)/i,
    /trades?\s+at[^$0-9]{0,40}\$?\s*([\d,.]+)/i,
    /trading\s+(?:at|price)[^$0-9]{0,40}\$?\s*([\d,.]+)/i,
    /reference\s+price[^$0-9]{0,40}\$?\s*([\d,.]+)/i,
  ];

  for (const pattern of pricePatterns) {
    const match = text.match(pattern);
    if (match) {
      const n = cleanNumber(match[1]);
      if (Number.isFinite(n)) return n;
    }
  }

  return null;
}

function extractMoneyFields(notes, ticker) {
  const text = notes.replace(/\n+/g, " ");
  const m = moneyRegex();

  const netWorth = findMoney(text, [
    new RegExp(`(?:total\\s+)?net\\s+worth[^$0-9]{0,60}${m}`, "i"),
    new RegExp(`worth[^$0-9]{0,40}${m}`, "i"),
  ]);

  const investableAssets = findMoney(text, [
    new RegExp(`investable\\s+assets[^$0-9]{0,60}${m}`, "i"),
    new RegExp(`managed\\s+assets[^$0-9]{0,60}${m}`, "i"),
    new RegExp(`assets\\s+under\\s+management[^$0-9]{0,60}${m}`, "i"),
    new RegExp(`\\bAUM\\b[^$0-9]{0,60}${m}`, "i"),
    new RegExp(`portfolio\\s+(?:value|assets|size)[^$0-9]{0,60}${m}`, "i"),
  ]);

  const stockPosition = extractStockPosition(text, ticker);

  const annualIncome = findMoney(text, [
    new RegExp(`annual\\s+income[^$0-9]{0,60}${m}`, "i"),
    new RegExp(`income[^$0-9]{0,40}${m}`, "i"),
  ]);

  return { netWorth, investableAssets, stockPosition, annualIncome };
}

export function safeExtractClientData(notes, clientName, oldExtractClientData) {
  const oldData = oldExtractClientData(notes, clientName);
  const ticker = extractTicker(notes);
  const clientAge = extractClientAge(notes) ?? oldData.clientAge ?? 65;

  const money = extractMoneyFields(notes, ticker);

  const costBasisPct =
    findPercent(notes, [
      /cost\s+basis[^0-9]{0,40}([\d,.]+)\s*%/i,
      /\bbasis[^0-9]{0,40}([\d,.]+)\s*%/i,
    ]) ?? oldData.costBasisPct ?? 15;

  const totalTaxRate =
    findPercent(notes, [
      /combined\s+tax\s+rate[^0-9]{0,40}([\d,.]+)\s*%/i,
      /tax\s+rate[^0-9]{0,40}([\d,.]+)\s*%/i,
      /capital\s+gains\s+tax[^0-9]{0,40}([\d,.]+)\s*%/i,
    ]) ?? oldData.totalTaxRate ?? oldData.taxRate ?? 37.1;

  const stockPrice = extractSharePrice(notes) ?? oldData.stockPrice ?? 200;

  const cleaned = {
    ...oldData,

    ticker: ticker || oldData.ticker || "",
    clientAge,

    // Use explicitly extracted values first.
    netWorth: money.netWorth ?? normalizeMillionValue(oldData.netWorth) ?? 0,
    investableAssets: money.investableAssets ?? normalizeMillionValue(oldData.investableAssets) ?? 0,
    stockPosition: money.stockPosition ?? normalizeMillionValue(oldData.stockPosition) ?? 0,
    income: money.annualIncome ?? normalizeMillionValue(oldData.income) ?? 0,

    costBasisPct,
    totalTaxRate,
    taxRate: totalTaxRate,
    stockPrice,
  };

  // If old extractor guessed a position but notes clearly say no concentrated stock, zero it.
  if (/no concentrated stock|no concentration|not concentrated|no single stock/i.test(notes)) {
    cleaned.stockPosition = 0;
    cleaned.ticker = "";
  }

  // If there is no stock position at all, ensure ticker is blank too.
  if (!cleaned.stockPosition) {
    cleaned.ticker = "";
  }

  cleaned.concentration =
    cleaned.investableAssets > 0 && cleaned.stockPosition > 0
      ? Math.min(Math.max((cleaned.stockPosition / cleaned.investableAssets) * 100, 0), 100)
      : 0;

  cleaned.costBasis = cleaned.stockPosition * (cleaned.costBasisPct / 100);
  cleaned.embeddedGain = Math.max(cleaned.stockPosition - cleaned.costBasis, 0);
  cleaned.immediateTax = cleaned.embeddedGain * (cleaned.totalTaxRate / 100);

  // CRT size: 10–20% of stock position based on charitable intent signals in notes.
  // Strong signals (very charitable) → 20%; moderate → 15%; no signals → 10%.
  const highCharity = /\b(very\s+charitable|deeply\s+charitable|philanthropi|foundation|donor.advised|charity\s+is\s+a\s+priority|charitably\s+inclined|major\s+giver|legacy\s+giving|significant\s+giving|loves?\s+to\s+give|passion.*giv|giv.*passion)\b/i.test(notes);
  const someCharity = /\b(charitable|philanthrop|giving|donate|donating|nonprofit|non.profit|church|synagogue|temple|cause|legacy|estate\s+plan)\b/i.test(notes);
  const crtPct = highCharity ? 0.20 : someCharity ? 0.15 : 0.10;
  cleaned.crtAllocation = cleaned.stockPosition * crtPct;
  cleaned.harvestingSleeve = cleaned.stockPosition * 0.30;
  cleaned.collarAllocation = cleaned.stockPosition * 0.30;

  cleaned.crtPayoutRate = 5; // fixed 5% payout rate
  cleaned.crtIncome = cleaned.crtAllocation * (cleaned.crtPayoutRate / 100);

  // CRT deduction: IRS Table S, 5% payout, single life — factor varies by client age
  const crtDeduction = cleaned.crtAllocation * crtDeductionFactor(cleaned.clientAge);
  cleaned.charitableDeductionLow  = crtDeduction;
  cleaned.charitableDeductionHigh = crtDeduction;

  // 130/30 generates ~25% of sleeve in harvestable losses annually (long-side + short-side combined)
  cleaned.annualHarvestLosses = cleaned.harvestingSleeve * 0.25;

  // Always use LTCG federal rate (20% + 3.8% NIIT = 23.8%) — not ordinary income rate
  cleaned.federalTaxRate = 23.8;
  cleaned.stateTaxRate = Math.max(cleaned.totalTaxRate - cleaned.federalTaxRate, 0);

  cleaned.federalTaxSavings = cleaned.annualHarvestLosses * (cleaned.federalTaxRate / 100);
  cleaned.stateTaxSavings   = cleaned.annualHarvestLosses * (cleaned.stateTaxRate   / 100);
  cleaned.taxSavings        = cleaned.federalTaxSavings + cleaned.stateTaxSavings;

  cleaned.putStrike = cleaned.stockPrice * 0.85;
  cleaned.callStrike = cleaned.stockPrice * 1.19;

  cleaned.putFloorValue  = cleaned.collarAllocation * 0.85;
  cleaned.callCapValue   = cleaned.collarAllocation * 1.19;

  // After-CRT concentration: both the stock position AND investable assets shrink by the CRT amount
  const postCrtStock  = cleaned.stockPosition - cleaned.crtAllocation;
  const postCrtAssets = cleaned.investableAssets - cleaned.crtAllocation;
  cleaned.afterCrtConcentration =
    postCrtAssets > 0
      ? Math.min(Math.max((postCrtStock / postCrtAssets) * 100, 0), 100)
      : 0;

  cleaned.drawdown40Impact = cleaned.stockPosition * 0.4;

  // Parse actual holdings from uploaded document text (if present)
  const holdings = parseHoldingsFromText(notes);
  if (holdings.length >= 2) cleaned.currentHoldings = holdings;

  return cleaned;
}

// ─── Holdings parser ──────────────────────────────────────────────────────────
// Extracts {ticker, pct} pairs from portfolio report text (PDF, XLSX, CSV, DOCX).
// Works on both structured CSV exports ("AAPL,Apple Inc.,35.2%") and the messier
// text that pdfjsLib produces from scanned/rendered PDF statements.

const HOLDING_IGNORE = new Set([
  "ETF","IRA","LLC","LTD","INC","USD","AUM","YTD","NAV","TBD","SMA","REIT",
  "TOTAL","CASH","PAGE","DATE","TYPE","FUND","BOND","NOTE","CORP","INTL",
  "CLASS","UNIT","ACCOUNT","VALUE","PRICE","SHARES","MARKET","SYMBOL","NAME",
  "DESCRIPTION","ALLOC","WEIGHT","AND","THE","FOR","NOT","ALL","NEW","MKT",
  "EST","YR","QTR","JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP",
  "OCT","NOV","DEC","USA","AM","PM","NET","TAX","GAIN","LOSS","COST","BASIS",
  "UNREALIZED","REALIZED","ACCRUED","INCOME","DIVIDEND","INTEREST","FEE",
]);

function isValidTicker(t) {
  // Accept plain tickers (AAPL) and share-class variants (BRK.B, BF.A)
  return !!t && /^[A-Z]{1,5}(\.[A-Z])?$/.test(t) && !HOLDING_IGNORE.has(t.split(".")[0]);
}

/** Normalize ticker for Yahoo Finance: BRK.B → BRK-B */
function normalizeTicker(t) {
  return (t || "").replace(/\./g, "-");
}

function normalizeHoldings(raw) {
  // raw = [{ticker, value}] or [{ticker, pct}]
  if (raw.length < 2) return [];
  // Deduplicate — sum values/pcts for same ticker
  const deduped = new Map();
  for (const h of raw) {
    const key = h.ticker;
    const existing = deduped.get(key) || { ticker: key, value: 0, pct: 0 };
    if (h.value !== undefined) existing.value += h.value;
    if (h.pct   !== undefined) existing.pct   += h.pct;
    deduped.set(key, existing);
  }
  const result = [...deduped.values()];
  // Convert values → pct if needed
  const useValue = result[0].value > 0;
  const total = result.reduce((s, h) => s + (useValue ? h.value : h.pct), 0);
  if (total <= 0) return [];
  const normalized = result.map(h => ({
    ticker: h.ticker,
    pct: ((useValue ? h.value : h.pct) / total) * 100,
  })).filter(h => h.pct > 0.01);
  return normalized.length >= 2 ? normalized : [];
}

/** Splits a single CSV line respecting RFC 4180 double-quoting.
 *  "NVDA","$46,500,000" → ["NVDA", "$46,500,000"] instead of ["NVDA", '"$46', '500', '000"'] */
export function parseCSVRow(line) {
  const cells = [];
  let cur = "";
  let inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuote && line[i + 1] === '"') { cur += '"'; i++; } // escaped quote
      else { inQuote = !inQuote; }
    } else if (ch === "," && !inQuote) {
      cells.push(cur.trim());
      cur = "";
    } else {
      cur += ch;
    }
  }
  cells.push(cur.trim());
  return cells;
}

export function parseHoldingsFromText(text) {
  if (!text || text.length < 50) return [];

  const lines = text.split(/[\n\r]+/).map(l => l.trim()).filter(Boolean);

  // ── Strategy 1: Structured CSV/TSV with a header row ─────────────────────
  // Detects header row containing "Symbol"/"Ticker" + "Market Value"/"%" columns.
  for (let hi = 0; hi < Math.min(lines.length, 15); hi++) {
    const isTab = lines[hi].includes("\t");
    // For TSV use simple split; for CSV use the RFC 4180-aware parser
    const splitRow = (line) =>
      isTab ? line.split("\t").map(c => c.trim()) : parseCSVRow(line);

    const cols = splitRow(lines[hi]).map(c => c.replace(/^"|"$/g, "").toLowerCase());

    const symIdx = cols.findIndex(c => /^(symbol|ticker|sym|tick)$/.test(c));
    if (symIdx < 0) continue;

    const valIdx = cols.findIndex(c => /market.?value|mkt.?val|market val|\bvalue\b|\bbalance\b|\bholdings\b|\bamount\b/i.test(c));
    const pctIdx = cols.findIndex(c => /^(%|pct|percent|weight|alloc|allocation)$/.test(c));
    if (valIdx < 0 && pctIdx < 0) continue;

    const raw = [];
    for (let ri = hi + 1; ri < lines.length; ri++) {
      const row = splitRow(lines[ri]);
      if (row.length <= Math.max(symIdx, valIdx >= 0 ? valIdx : 0, pctIdx >= 0 ? pctIdx : 0)) continue;
      const ticker = normalizeTicker((row[symIdx] || "").replace(/^"|"$/g, "").trim().toUpperCase());
      if (!isValidTicker(ticker.replace(/-/g, "."))) continue; // validate pre-normalization form

      if (pctIdx >= 0) {
        const pct = parseFloat((row[pctIdx] || "").replace(/[%,\s"]/g, ""));
        if (isFinite(pct) && pct > 0) raw.push({ ticker, pct });
      } else {
        const value = parseFloat((row[valIdx] || "").replace(/[$,\s"]/g, ""));
        if (isFinite(value) && value > 0) raw.push({ ticker, value });
      }
    }

    const result = normalizeHoldings(raw);
    if (result.length >= 2) return result;
  }

  // ── Strategy 2: Line-by-line percentage scan (free-form text / PDF) ──────
  const raw = [];
  for (const line of lines) {
    const pctMatch = line.match(/([\d,]+\.?\d*)\s*%/);
    if (!pctMatch) continue;
    const pct = parseFloat(pctMatch[1].replace(/,/g, ""));
    if (!isFinite(pct) || pct < 0.1 || pct > 99) continue;

    let ticker = null;
    // CSV: first field
    const csvM = line.match(/^([A-Z]{1,5})[,\t]/);
    if (csvM && isValidTicker(csvM[1])) ticker = csvM[1];
    // Start of line: TICKER followed by a word char
    if (!ticker) {
      const startM = line.match(/^([A-Z]{2,5})\s+\S/);
      if (startM && isValidTicker(startM[1])) ticker = startM[1];
    }
    // Any uppercase word before the %
    if (!ticker) {
      const before = line.slice(0, line.indexOf(pctMatch[0]));
      for (const m of [...before.matchAll(/\b([A-Z]{2,5})\b/g)]) {
        if (isValidTicker(m[1])) { ticker = m[1]; break; }
      }
    }
    if (ticker) raw.push({ ticker, pct });
  }

  return normalizeHoldings(raw);
}
