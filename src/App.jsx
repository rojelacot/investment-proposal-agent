import { useState, useEffect, useRef } from "react";
import {
  Document,
  Packer,
  Paragraph,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
} from "docx";
import { detectMissingInfo } from "./clarifyingQuestions";
import MissingInfoPanel from "./MissingInfoPanel";
import { recomputeReviewedData, getProposalQuality } from "./proposalUpgrades";
import FileUploadBox from "./FileUploadBox";
import { safeExtractClientData, parseHoldingsFromText } from "./safeClientExtraction";
import { fmtM, fmtK, pct } from "./formatters";
import { generatePowerPoint } from "./pptGenerator";
import ProposalPreviewModal from "./ProposalPreviewModal";
import { getFunds } from "./portfolioData";
import { getCheapestPassiveTicker } from "./fundAlternatives";
import {
  toMonthlyReturns,
  buildTargetWeightMap,
  buildConcentratedWeightMap,
  buildHoldingsWeightMap,
  weightedPortfolioReturns,
  summarizeReturns,
  weightedAverageAnnualReturn,
} from "./backtest";
import "./App.css";

export default function App() {
  const [clientName, setClientName] = useState("");
  const [notes, setNotes] = useState("");
  const [status, setStatus] = useState("");
  const [proposal, setProposal] = useState(null);
  const [clarifyingQuestions, setClarifyingQuestions] = useState([]);
  const [clarificationAnswers, setClarificationAnswers] = useState({});
  const [reviewData, setReviewData] = useState(null);
  const [previewModal, setPreviewModal] = useState(null);

  const [selectedProposalModules, setSelectedProposalModules] = useState({
    executiveSummary: true,
    clientProfileGoals: true,
    financialPicture: true, // planning-scope coverage slide
    capabilities: true, // firm capabilities / "how we help" slide
    taxOnTheTable: true, // hero value slide: tax saved + downside protected
    recommendedInvestmentApproach: false, // opt-in: portfolio strategy + risk/allocation slides
    riskManagementOverview: false,
    goalsTimeline: false,
    liquidityNeedsReview: false,
    taxPlanningOverview: false,
    incomeExpenseSnapshot: false,
    retirementPlanning: false,
    legacyWealthTransfer: false,
    estatePlanningReview: false,
    restrictionsImplementationNotes: false,
    implementationTimeline: false,
    nextSteps: true,
    // Analytics & projection slides — opt-in: only render when checked.
    feeDragAnalysis: false,
    monteCarloProjection: false,
    stressTestAnalysis: false,
  });

  const [selectedServices, setSelectedServices] = useState({
    retirementPlanning: false,
    financialPlanning: false,
    legacyTransfer: false,
    estatePlanning: false,
    riskOverview: false,
  });

  const [selectedRiskProfile, setSelectedRiskProfile] = useState("");
  const [useRecommendedApproach, setUseRecommendedApproach] = useState(false);
  const [firmName, setFirmName] = useState("");
  const [advisorName, setAdvisorName] = useState("");
  const [liveStrategyAllocations, setLiveStrategyAllocations] = useState(null);

  useEffect(() => {
    fetch("/api/strategies")
      .then(r => r.ok ? r.json() : null)
      .then(data => { if (data) setLiveStrategyAllocations(data); })
      .catch(() => {}); // silently fall back to portfolioData.js if unavailable
  }, []);

  const [collarOptions, setCollarOptions] = useState(null);
  const [collarOptionsLoading, setCollarOptionsLoading] = useState(false);
  const [collarOptionsError, setCollarOptionsError] = useState("");

  // Holdings scanned from uploaded documents (used as current portfolio in backtest)
  const [scannedHoldings, setScannedHoldings] = useState([]);

  // Current-vs-target portfolio backtest (real historical data, see backtest.js).
  const [backtestResult, setBacktestResult] = useState(null);
  const [backtestLoading, setBacktestLoading] = useState(false);
  const [backtestError, setBacktestError] = useState("");

  const riskProfileOptions = [
    { key: "conservative", label: "Conservative", mix: "20/80" },
    { key: "moderatelyConservative", label: "Moderately Conservative", mix: "30/70" },
    { key: "conservativePlus", label: "Conservative Plus", mix: "40/60" },
    { key: "balanced", label: "Balanced", mix: "50/50" },
    { key: "balancedPlus", label: "Balanced Plus", mix: "60/40" },
    { key: "growth", label: "Growth", mix: "70/30" },
    { key: "growthPlus", label: "Growth Plus", mix: "80/20" },
    { key: "aggressive", label: "Aggressive", mix: "100/0" },
  ];


const [selectedPortfolioStrategies, setSelectedPortfolioStrategies] = useState({
    corePrivate: false,
    selectLiquidity: false,
    traditional: false,
    focusedB: false,
    selectLiquidityUsBias: false,
    traditionalUsBias: false,
  });
  const [qualityReport, setQualityReport] = useState(null);
  const [clientType, setClientType] = useState("Concentrated stock executive");
  const [selectedStrategies, setSelectedStrategies] = useState({
    crt: true,
    harvesting: true,
    collar: true,
    diversification: true,
    muniBonds: false,
    donorAdvisedFund: false,
    exchangeFund: false,
    estatePlanning: false,
  });

  // ── Session persistence ───────────────────────────────────────────────────
  // Auto-save the working session to localStorage so a page refresh (or an
  // accidental tab close) doesn't lose an in-progress proposal. We restore once
  // on mount, then save whenever the tracked inputs/selections change.
  const SESSION_KEY = "ipa.session.v1";
  const hydratedRef = useRef(false);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(SESSION_KEY);
      if (raw) {
        const s = JSON.parse(raw);
        if (s.firmName != null) setFirmName(s.firmName);
        if (s.advisorName != null) setAdvisorName(s.advisorName);
        if (s.clientName != null) setClientName(s.clientName);
        if (s.clientType != null) setClientType(s.clientType);
        if (s.notes != null) setNotes(s.notes);
        if (s.reviewData) setReviewData(s.reviewData);
        if (s.qualityReport) setQualityReport(s.qualityReport);
        if (s.selectedStrategies) setSelectedStrategies(s.selectedStrategies);
        if (s.selectedPortfolioStrategies) setSelectedPortfolioStrategies(s.selectedPortfolioStrategies);
        if (s.selectedProposalModules) setSelectedProposalModules(s.selectedProposalModules);
        if (s.selectedRiskProfile != null) setSelectedRiskProfile(s.selectedRiskProfile);
        if (typeof s.useRecommendedApproach === "boolean") setUseRecommendedApproach(s.useRecommendedApproach);
        if (Array.isArray(s.scannedHoldings)) setScannedHoldings(s.scannedHoldings);
      }
    } catch { /* corrupt/unavailable storage — start fresh */ }
    hydratedRef.current = true;
  }, []);

  useEffect(() => {
    if (!hydratedRef.current) return;
    try {
      localStorage.setItem(SESSION_KEY, JSON.stringify({
        firmName, advisorName, clientName, clientType, notes,
        reviewData, qualityReport, selectedStrategies,
        selectedPortfolioStrategies, selectedProposalModules,
        selectedRiskProfile, useRecommendedApproach, scannedHoldings,
        savedAt: new Date().toISOString(),
      }));
    } catch { /* quota or serialization issue — non-fatal */ }
  }, [
    firmName, advisorName, clientName, clientType, notes,
    reviewData, qualityReport, selectedStrategies,
    selectedPortfolioStrategies, selectedProposalModules,
    selectedRiskProfile, useRecommendedApproach, scannedHoldings,
  ]);


  async function fetchPreviousClosePrice(ticker) {
    const cleanTicker = String(ticker || "").trim().toUpperCase();

    if (!cleanTicker) return null;

    try {
      const res = await fetch(`/api/quote/${encodeURIComponent(cleanTicker)}`);

      if (!res.ok) return null;

      const data = await res.json();

      const previousClose = Number(data.previousClose);
      const currentPrice = Number(data.currentPrice);
      const price = Number.isFinite(previousClose) && previousClose > 0 ? previousClose : currentPrice;

      if (!Number.isFinite(price) || price <= 0) return null;

      return {
        ticker: cleanTicker,
        close: price,
        currentPrice,
        previousClose: price,
        date: data.timestamp,
        source: data.source || "Local quote route",
      };
    } catch (err) {
      console.warn("Could not fetch previous close from local quote route:", err);
      return null;
    }
  }

  async function fetchCollarOptions(ticker) {
    const cleanTicker = String(ticker || "").trim().toUpperCase();
    if (!cleanTicker) return;

    setCollarOptionsLoading(true);
    setCollarOptionsError("");
    setCollarOptions(null);

    try {
      const res = await fetch(`/api/options/${encodeURIComponent(cleanTicker)}`);
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || `Options fetch failed (${res.status})`);
      }
      const data = await res.json();
      setCollarOptions(data);
    } catch (err) {
      console.warn("Collar options fetch failed:", err);
      // A TypeError from fetch() means the request never reached the backend at
      // all (Express server.js on :5174 not running, or the app was opened
      // without the Vite /api proxy) — surface that as a clear, advisor-friendly
      // message instead of the raw browser string "Failed to fetch". A plain
      // Error carries a real reason from the backend, so keep its message.
      const message =
        err instanceof TypeError
          ? "Live options service not reachable"
          : err.message || "Could not load live options data.";
      setCollarOptionsError(message);
    } finally {
      setCollarOptionsLoading(false);
    }
  }

  function isUsableTicker(ticker) {
    const t = String(ticker || "").trim().toUpperCase();
    return /^[A-Z]{1,5}$/.test(t) && !["STOCK", "SHARE", "SHARES", "CLIENT"].includes(t);
  }

  function extractDirectTickerFromNotes(notes = "") {
    const text = String(notes || "");

    const patterns = [
      /\b(?:ticker|symbol)\s*[:\-]?\s*([A-Z]{1,5})\b/i,
      /\(([A-Z]{1,5})\)/,
      /\b([A-Z]{1,5})\b\s+(?:stock|shares|position|holding|holdings|concentration)\b/i,
      /(?:stock|shares|position|holding|holdings|concentration)\s+(?:in|of)\s+\b([A-Z]{1,5})\b/i,
    ];

    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match?.[1] && isUsableTicker(match[1])) return match[1].toUpperCase();
    }

    return "";
  }

  function extractCompanyNameFromNotes(notes = "") {
    const text = String(notes || "");

    const knownCompanies = [
      "Nvidia",
      "Tesla",
      "Apple",
      "Microsoft",
      "Amazon",
      "Alphabet",
      "Google",
      "Meta",
      "Netflix",
      "AMD",
      "Broadcom",
      "Salesforce",
      "Oracle",
      "Palantir",
      "Costco",
      "Walmart",
      "JPMorgan",
      "Goldman Sachs",
      "Morgan Stanley",
      "Visa",
      "Mastercard",
      "Coca-Cola",
      "Pepsi",
    ];

    const lower = text.toLowerCase();

    for (const company of knownCompanies) {
      if (lower.includes(company.toLowerCase())) return company;
    }

    const patterns = [
      /(?:concentrated|large|single[-\s]?stock)?\s*(?:position|holding|holdings|shares|stock)\s+(?:in|of)\s+([A-Z][A-Za-z&.\-\s]{2,45})/i,
      /\$?\s*[\d,.]+\s*(?:M|MM|million|B|billion)\s+(?:in|of)\s+([A-Z][A-Za-z&.\-\s]{2,45})/i,
      /([A-Z][A-Za-z&.\-\s]{2,45})\s+(?:stock|shares|position|holding|holdings|concentration)/i,
    ];

    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match?.[1]) {
        return match[1]
          .replace(/\b(position|holding|holdings|shares|stock|concentration|worth|valued|approximately|around)\b/gi, "")
          .trim();
      }
    }

    return "";
  }

  async function resolveTickerFromNotes(notes = "", existingTicker = "") {
    if (isUsableTicker(existingTicker)) return existingTicker.trim().toUpperCase();

    const directTicker = extractDirectTickerFromNotes(notes);
    if (isUsableTicker(directTicker)) return directTicker;

    const companyName = extractCompanyNameFromNotes(notes);
    if (!companyName) return "";

    try {
      const res = await fetch(`/api/symbol-search/${encodeURIComponent(companyName)}`);
      if (!res.ok) return "";

      const data = await res.json();
      const ticker = String(data.ticker || "").trim().toUpperCase();

      return isUsableTicker(ticker) ? ticker : "";
    } catch (err) {
      console.warn("Could not resolve ticker from notes:", err);
      return "";
    }
  }

  async function applyMarketQuoteToProposalData(rawData, notes = "") {
    let cleaned = applyConcentratedStockFromNotes(rawData, notes);

    const resolvedTicker = await resolveTickerFromNotes(notes, cleaned.ticker);

    if (resolvedTicker) {
      cleaned.ticker = resolvedTicker;
    }

    if (!isUsableTicker(cleaned.ticker)) {
      return forceCollarFromStockPrice(normalizeMillionsForReview(cleaned));
    }

    const quote = await fetchPreviousClosePrice(cleaned.ticker);

    if (quote?.close) {
      cleaned = forceCollarFromStockPrice(
        normalizeMillionsForReview({
          ...cleaned,
          ticker: cleaned.ticker,
          stockPrice: quote.close,
          priorCloseDate: quote.date,
          priceSource: quote.source || "Local quote route",
        })
      );
    } else {
      cleaned = forceCollarFromStockPrice(normalizeMillionsForReview(cleaned));
    }

    return cleaned;
  }

  function downloadBlob(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);
  }

  function normalizeClientNotes(raw) {
    let text = raw;

    // Convert full dollar amounts like $50,000,000 into $50M
    text = text.replace(/\$([\d,]+)(?!\s*(M|million|B|billion))/gi, (match, num) => {
      const clean = parseFloat(num.replaceAll(",", ""));
      if (!Number.isFinite(clean)) return match;
      if (clean >= 1000000000) return `$${(clean / 1000000000).toFixed(2)}B`;
      if (clean >= 1000000) return `$${(clean / 1000000).toFixed(2)}M`;
      return match;
    });

    // Convert plain comma numbers like 54,000,000 into $54M
    text = text.replace(/(?<![$\d])([1-9]\d{0,2}(?:,\d{3}){2,})(?!\d)/g, (match, num) => {
      const clean = parseFloat(num.replaceAll(",", ""));
      if (!Number.isFinite(clean)) return match;
      if (clean >= 1000000000) return `$${(clean / 1000000000).toFixed(2)}B`;
      if (clean >= 1000000) return `$${(clean / 1000000).toFixed(2)}M`;
      return match;
    });

    // Convert "50 million dollars" into "$50M"
    text = text.replace(/([\d.]+)\s*million dollars/gi, "$1M");
    text = text.replace(/([\d.]+)\s*billion dollars/gi, "$1B");

    return text;
  }

  function toMillions(value, unit) {
    const n = parseFloat(value);
    if (!unit) return n;
    const u = unit.toLowerCase();
    if (u.startsWith("b")) return n * 1000;
    return n;
  }

  // Format helpers imported from ./formatters

  function extractMoneyAfter(text, patterns) {
    for (const pattern of patterns) {
      const m = text.match(pattern);
      if (m) return toMillions(m[1], m[2]);
    }
    return null;
  }

  // IRS Table S approximation: remainder factor at 5% payout by single-life age
  const CRT_DEDUCTION_TABLE_APP = [
    [45, 0.215], [50, 0.245], [55, 0.278], [60, 0.314],
    [65, 0.356], [70, 0.404], [75, 0.458], [80, 0.520], [85, 0.584],
  ];
  function crtDeductionFactorApp(age) {
    const a = Math.max(45, Math.min(85, age || 65));
    for (let i = 0; i < CRT_DEDUCTION_TABLE_APP.length - 1; i++) {
      const [a0, f0] = CRT_DEDUCTION_TABLE_APP[i];
      const [a1, f1] = CRT_DEDUCTION_TABLE_APP[i + 1];
      if (a >= a0 && a <= a1) return f0 + (f1 - f0) * (a - a0) / (a1 - a0);
    }
    return 0.356;
  }

  function extractClientData(text, name) {
    text = normalizeClientNotes(text);

    const companyMap = [
      ["tesla", "TSLA"],
      ["apple", "AAPL"],
      ["microsoft", "MSFT"],
      ["amazon", "AMZN"],
      ["google", "GOOGL"],
      ["alphabet", "GOOGL"],
      ["meta", "META"],
      ["facebook", "META"],
      ["nvidia", "NVDA"],
      ["netflix", "NFLX"],
      ["salesforce", "CRM"],
      ["broadcom", "AVGO"],
      ["amd", "AMD"],
      ["oracle", "ORCL"],
      ["adobe", "ADBE"],
      ["intel", "INTC"],
      ["shopify", "SHOP"],
      ["snowflake", "SNOW"],
      ["palantir", "PLTR"],
    ];

    const tickerMatch = text.match(/\b(NVDA|AAPL|MSFT|TSLA|AMZN|GOOGL|GOOG|META|CRM|AVGO|AMD|NFLX|ORCL|ADBE|INTC|SHOP|SNOW|PLTR)\b/i);
    const companyMatch = companyMap.find(([company]) => text.toLowerCase().includes(company));
    const ticker = tickerMatch ? tickerMatch[1].toUpperCase() : companyMatch ? companyMatch[1] : "";

    // Extract client age for age-based CRT deduction
    const currentYear = new Date().getFullYear();
    let clientAge = 65; // default
    const ageMatch = text.match(/\bage[d]?\s*[:\-]?\s*(\d{2})\b/i)
      || text.match(/(\d{2})\s+years?\s+old/i)
      || text.match(/client\s+is\s+(\d{2})\s+years?/i);
    if (ageMatch) { const a = parseInt(ageMatch[1]); if (a >= 35 && a <= 90) clientAge = a; }
    else {
      const dobMatch = text.match(/born\s+(?:in\s+)?(\d{4})/i)
        || text.match(/\bDOB\b[:\s]+\d{1,2}\/\d{1,2}\/(\d{4})/i);
      if (dobMatch) { const yr = parseInt(dobMatch[1]); if (yr >= 1930 && yr <= currentYear - 35) clientAge = currentYear - yr; }
    }

    const allMoney = [...text.matchAll(/\$?\s*([\d.]+)\s*(M|million|B|billion)/gi)].map((m) =>
      toMillions(m[1], m[2])
    );

    let netWorth =
      extractMoneyAfter(text, [
        /net worth[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
        /total net worth[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
      ]) ?? allMoney[0] ?? 0;

    let investableAssets =
      extractMoneyAfter(text, [
        /investable assets[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
        /investable[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
        /portfolio[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
      ]) ?? null;

    const tickerRegex = new RegExp(`${ticker}[^\\n$]*\\$?\\s*([\\d.]+)\\s*(M|million|B|billion)`, "i");

    let stockPosition =
      extractMoneyAfter(text, [
        tickerRegex,
        /stock position[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
        /concentrated stock position[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
        /concentrated position[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
        /position[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
      ]) ?? null;

    let income =
      extractMoneyAfter(text, [
        /annual income[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
        /income[^$\d]*\$?\s*([\d.]+)\s*(M|million|B|billion)/i,
      ]) ?? 0;

    if (!stockPosition && allMoney.length >= 1) stockPosition = allMoney[allMoney.length - 1];
    if (!investableAssets && allMoney.length >= 2) investableAssets = allMoney[1];
    if (!netWorth && allMoney.length >= 1) netWorth = allMoney[0];

    const percentages = [...text.matchAll(/([\d.]+)%/g)].map((m) => parseFloat(m[1]));

    const concentrationMatch =
      text.match(/portfolio concentration[^%\d]*([\d.]+)%/i) ||
      text.match(/concentration[^%\d]*([\d.]+)%/i);

    let concentration = concentrationMatch ? parseFloat(concentrationMatch[1]) : null;

    // If investable assets are missing or clearly too small, infer a more reasonable value
    if ((!investableAssets || investableAssets < stockPosition) && stockPosition) {
      if (concentration && concentration > 0 && concentration <= 100) {
        investableAssets = stockPosition / (concentration / 100);
      } else if (netWorth && netWorth > stockPosition) {
        investableAssets = Math.max(stockPosition / 0.65, netWorth * 0.63);
      } else {
        investableAssets = stockPosition / 0.65;
      }
    }

    // Last fallback
    if (!investableAssets && stockPosition) investableAssets = stockPosition / 0.65;
    // Do NOT fabricate a stockPosition when none is present.
    if (!netWorth && investableAssets) netWorth = investableAssets * 1.6;
    if (netWorth < investableAssets) netWorth = investableAssets * 1.6;

    concentration = stockPosition > 0 && investableAssets > 0
      ? Math.min(Math.max((stockPosition / investableAssets) * 100, 0), 100)
      : 0;

    const totalTaxMatch =
      text.match(/tax rate[^%\d]*([\d.]+)%/i) ||
      text.match(/total tax[^%\d]*([\d.]+)%/i) ||
      text.match(/combined tax[^%\d]*([\d.]+)%/i);

    const federalTaxMatch =
      text.match(/federal tax[^%\d]*([\d.]+)%/i) ||
      text.match(/federal capital gains[^%\d]*([\d.]+)%/i);

    const stateTaxMatch =
      text.match(/state tax[^%\d]*([\d.]+)%/i) ||
      text.match(/california tax[^%\d]*([\d.]+)%/i);

    let totalTaxRate = totalTaxMatch ? parseFloat(totalTaxMatch[1]) : null;
    let federalTaxRate = federalTaxMatch ? parseFloat(federalTaxMatch[1]) : null;
    let stateTaxRate = stateTaxMatch ? parseFloat(stateTaxMatch[1]) : null;

    if (federalTaxRate == null && stateTaxRate == null && totalTaxRate != null) {
      federalTaxRate = Math.min(23.8, totalTaxRate * 0.65);
      stateTaxRate = Math.max(totalTaxRate - federalTaxRate, 0);
    }

    if (federalTaxRate == null) federalTaxRate = 23.8;
    if (stateTaxRate == null) stateTaxRate = 13.3;
    if (totalTaxRate == null) totalTaxRate = federalTaxRate + stateTaxRate;

    const basisMatch =
      text.match(/cost basis[^%\d]*([\d.]+)%/i) ||
      text.match(/basis[^%\d]*([\d.]+)%/i);

    const costBasisPct = basisMatch ? parseFloat(basisMatch[1]) : 15;

    const priceMatch =
      text.match(/current share price[^$\d]*\$?\s*([\d.]+)/i) ||
      text.match(/share price[^$\d]*\$?\s*([\d.]+)/i) ||
      text.match(/stock price[^$\d]*\$?\s*([\d.]+)/i) ||
      text.match(/trading price[^$\d]*\$?\s*([\d.]+)/i) ||
      text.match(/trades at[^$\d]*\$?\s*([\d.]+)/i);

    let stockPrice = priceMatch ? parseFloat(priceMatch[1]) : NaN;

    // If no price found in notes, leave as 0 — forceCollarFromStockPrice will apply
    // FALLBACK_MARKET_PRICES, and downloadPowerPoint will override with live data.
    if (!Number.isFinite(stockPrice) || stockPrice <= 10) {
      stockPrice = 0;
    }

    const crtAllocation = stockPosition * 0.30;
    const harvestingSleeve = stockPosition * 0.30;
    const collarAllocation = stockPosition * 0.30;

    const costBasis = stockPosition * (costBasisPct / 100);
    const embeddedGain = Math.max(stockPosition - costBasis, 0);
    const immediateTax = embeddedGain * (totalTaxRate / 100);

    const crtPayoutRate = 5;
    const crtIncome = crtAllocation * (crtPayoutRate / 100);

    // CRT deduction: IRS Table S, 5% payout, single life — factor varies by client age
    const crtDeduction = crtAllocation * crtDeductionFactorApp(clientAge);
    const charitableDeductionLow  = crtDeduction;
    const charitableDeductionHigh = crtDeduction;

    // 130/30: ~25% annual harvest rate (long + short book combined)
    const annualHarvestLosses = harvestingSleeve * 0.25;
    // Always use LTCG federal rate (23.8%)
    const adjFederalTaxRate = 23.8;
    const adjStateTaxRate   = Math.max(totalTaxRate - adjFederalTaxRate, 0);
    const federalTaxSavings = annualHarvestLosses * (adjFederalTaxRate / 100);
    const stateTaxSavings   = annualHarvestLosses * (adjStateTaxRate   / 100);
    const taxSavings        = federalTaxSavings + stateTaxSavings;

    const putStrike    = stockPrice * 0.85;
    const callStrike   = stockPrice * 1.19;
    const putFloorValue  = collarAllocation * 0.85;
    const callCapValue   = collarAllocation * 1.19;

    // After-CRT: both stock position and investable assets shrink by CRT amount
    const postCrtStock  = stockPosition - crtAllocation;
    const postCrtAssets = investableAssets - crtAllocation;
    const afterCrtConcentration = postCrtAssets > 0
      ? Math.min(Math.max((postCrtStock / postCrtAssets) * 100, 0), 100)
      : 0;

    const drawdown40Impact = stockPosition * 0.40;

    return {
      name,
      ticker,
      clientAge,
      netWorth,
      investableAssets,
      stockPosition,
      income,
      concentration,
      totalTaxRate,
      taxRate: totalTaxRate,
      federalTaxRate: adjFederalTaxRate,
      stateTaxRate: adjStateTaxRate,
      costBasisPct,
      costBasis,
      embeddedGain,
      immediateTax,
      stockPrice,
      crtAllocation,
      harvestingSleeve,
      collarAllocation,
      crtPayoutRate,
      crtIncome,
      charitableDeductionLow,
      charitableDeductionHigh,
      annualHarvestLosses,
      federalTaxSavings,
      stateTaxSavings,
      taxSavings,
      putStrike,
      callStrike,
      putFloorValue,
      callCapValue,
      afterCrtConcentration,
      drawdown40Impact,
    };
  }


  // STRATEGY_CHECKBOX_OUTPUT_LOGIC_APPLIED
  function getSelectedProposalModuleLabels() {
    const labels = {
      executiveSummary: "Executive Summary",
      clientProfileGoals: "Client Profile & Goals",
      recommendedInvestmentApproach: "Recommended Investment Approach",
      riskManagementOverview: "Risk Management Overview",
      goalsTimeline: "Goals Timeline",
      liquidityNeedsReview: "Liquidity Needs Review",
      taxPlanningOverview: "Tax Planning Overview",
      incomeExpenseSnapshot: "Income & Expense Snapshot",
      retirementPlanning: "Retirement Planning",
      legacyWealthTransfer: "Legacy & Wealth Transfer",
      estatePlanningReview: "Estate Planning Review",
      restrictionsImplementationNotes: "Restrictions & Implementation Notes",
      implementationTimeline: "Implementation Timeline",
      nextSteps: "Planning Focus",
    };

    return Object.keys(selectedProposalModules)
      .filter((key) => selectedProposalModules[key])
      .map((key) => labels[key]);
  }

  function getSelectedServiceLabels() {
    const labels = {
      retirementPlanning: "Retirement Planning",
      financialPlanning: "Financial Planning",
      legacyTransfer: "Legacy / Transfer of Wealth",
      estatePlanning: "Estate Planning",
      riskOverview: "Risk Overview",
    };

    return Object.keys(selectedServices)
      .filter((key) => selectedServices[key])
      .map((key) => labels[key]);
  }

  function extractRiskNumberFromText(text = "") {
    const raw = String(text || "");
    const patterns = [
      /risk\s*number[^0-9]{0,20}(\d{1,2})/i,
      /risk\s*score[^0-9]{0,20}(\d{1,2})/i,
      /risk\s*tolerance[^0-9]{0,20}(\d{1,2})/i,
      /\bRQ\b[^0-9]{0,20}(\d{1,2})/i,
    ];

    for (const pattern of patterns) {
      const match = raw.match(pattern);
      if (match?.[1]) {
        const value = Number(match[1]);
        if (value >= 1 && value <= 99) return value;
      }
    }

    return null;
  }

  function riskProfileFromRiskNumber(riskNumber) {
    if (!Number.isFinite(riskNumber)) return "balanced";
    if (riskNumber <= 24) return "conservative";
    if (riskNumber <= 34) return "moderatelyConservative";
    if (riskNumber <= 44) return "conservativePlus";
    if (riskNumber <= 54) return "balanced";
    if (riskNumber <= 64) return "balancedPlus";
    if (riskNumber <= 74) return "growth";
    if (riskNumber <= 84) return "growthPlus";
    return "aggressive";
  }

  function riskProfileIndex(key) {
    return riskProfileOptions.findIndex((option) => option.key === key);
  }

  function shiftRiskProfile(key, shift) {
    const index = riskProfileIndex(key);
    if (index < 0) return key;

    const nextIndex = Math.max(0, Math.min(riskProfileOptions.length - 1, index + shift));
    return riskProfileOptions[nextIndex].key;
  }

  function getRiskProfileDisplay(key) {
    const found = riskProfileOptions.find((option) => option.key === key);
    return found ? `${found.label} — ${found.mix}` : "Balanced — 50/50";
  }

  function toggleRiskProfile(key) {
    setSelectedRiskProfile((prev) => (prev === key ? "" : key));
  }

  function hasSelectedRiskProfile() {
    return !!selectedRiskProfile;
  }

  function inferRiskProfileKeyFromRecommendation(text = "") {
    const t = String(text || "").toLowerCase();

    // Highest confidence exact matches.
    if (t.includes("balanced plus") || /\b60\s*\/\s*40\b/.test(t)) return "balancedPlus";
    if (t.includes("balanced") || /\b50\s*\/\s*50\b/.test(t)) return "balanced";
    if (t.includes("growth plus") || /\b80\s*\/\s*20\b/.test(t)) return "growthPlus";
    if (t.includes("growth") || /\b70\s*\/\s*30\b/.test(t)) return "growth";
    if (t.includes("conservative plus") || /\b40\s*\/\s*60\b/.test(t)) return "conservativePlus";
    if (t.includes("moderately conservative") || /\b30\s*\/\s*70\b/.test(t)) return "moderatelyConservative";
    if (t.includes("conservative") || /\b20\s*\/\s*80\b/.test(t)) return "conservative";

    // Only use aggressive if explicitly recommended, not if it appears in a menu.
    if (/\brecommend(?:ed|s|ation)?[^.\\n]{0,80}\baggressive\b/.test(t) || /\b100\s*\/\s*0\b/.test(t)) {
      return "aggressive";
    }

    return "";
  }





  function inferPortfolioStrategyKeyFromRecommendation(text = "") {
    const t = String(text || "").toLowerCase();

    // Priority rule:
    // Select Liquidity must win over Core Private if both appear anywhere.
    // This prevents generic/core menu text from overriding the actual recommendation.
    if (t.includes("select liquidity") && t.includes("us bias")) return "selectLiquidityUsBias";
    if (t.includes("select liquidity")) return "selectLiquidity";

    if (t.includes("traditional") && t.includes("us bias")) return "traditionalUsBias";
    if (t.includes("focused b") || t.includes("focused-b")) return "focusedB";

    // Core Private only if it appears as the explicit recommended strategy.
    if (
      t.includes("recommended strategy: core private") ||
      t.includes("recommend core private") ||
      t.includes("recommended approach: core private") ||
      t.includes("model strategy: core private")
    ) {
      return "corePrivate";
    }

    if (
      t.includes("recommended strategy: traditional") ||
      t.includes("recommend traditional") ||
      t.includes("recommended approach: traditional") ||
      t.includes("model strategy: traditional")
    ) {
      return "traditional";
    }

    return "";
  }







  function makeOnlyPortfolioStrategySelection(key) {
    return {
      corePrivate: key === "corePrivate",
      selectLiquidity: key === "selectLiquidity",
      traditional: key === "traditional",
      focusedB: key === "focusedB",
      selectLiquidityUsBias: key === "selectLiquidityUsBias",
      traditionalUsBias: key === "traditionalUsBias",
    };
  }





  useEffect(() => {
    if (!reviewData) return;

    // Fix ticker if extractor accidentally picked strategy words like CORE.
    const cleanTicker = resolveClientTickerFromNotes(notes, reviewData?.ticker);
    if (cleanTicker && cleanTicker !== reviewData?.ticker) {
      setReviewData((prev) => ({
        ...prev,
        ticker: cleanTicker,
      }));
    }

    // Auto-fetch live options when collar strategy is selected and ticker is valid.
    if (selectedStrategies?.collar && isUsableTicker(reviewData.ticker)) {
      fetchCollarOptions(reviewData.ticker);
    }
  }, [reviewData?.ticker, selectedStrategies?.collar]);

  function getSelectedRiskProfileLabel() {
    return getRiskProfileDisplay(selectedRiskProfile);
  }

  function recommendInvestmentApproach(data = {}, notes = "") {
    const text = String(notes || "").toLowerCase();
    const riskNumber = extractRiskNumberFromText(notes);

    const investableAssets = Number(data.investableAssets || 0);
    const concentration = Number(data.concentration || 0);
    const age = Number(data.age || data.clientAge || 0);

    const liquidityNeed =
      text.includes("withdraw") ||
      text.includes("distribution") ||
      text.includes("liquidity") ||
      text.includes("cash need") ||
      text.includes("home") ||
      text.includes("house") ||
      text.includes("mortgage") ||
      text.includes("build") ||
      text.includes("college") ||
      text.includes("tuition") ||
      text.includes("income need") ||
      text.includes("retirement income");

    const retirementSoon =
      text.includes("retire") ||
      text.includes("retirement") ||
      text.includes("stopping work") ||
      text.includes("stop working") ||
      text.includes("part time");

    const alternativesMentioned =
      text.includes("alternatives") ||
      text.includes("private equity") ||
      text.includes("private credit") ||
      text.includes("hedge") ||
      text.includes("real estate") ||
      text.includes("alts");

    const avoidAlternatives =
      text.includes("no alternatives") ||
      text.includes("liquid only") ||
      text.includes("simplicity") ||
      text.includes("simple portfolio");

    const usBias =
      text.includes("us bias") ||
      text.includes("u.s. bias") ||
      text.includes("domestic bias") ||
      text.includes("prefer us") ||
      text.includes("prefer u.s.") ||
      text.includes("home bias");

    const focused =
      text.includes("focused b") ||
      text.includes("no us core") ||
      text.includes("no u.s. core") ||
      text.includes("growth and value tilt") ||
      text.includes("value and growth tilt");

    let riskProfile = riskProfileFromRiskNumber(riskNumber);

    if (!riskNumber) {
      if (text.includes("aggressive") || text.includes("maximum growth")) riskProfile = "aggressive";
      else if (text.includes("growth plus")) riskProfile = "growthPlus";
      else if (text.includes("growth")) riskProfile = "growth";
      else if (text.includes("balanced plus")) riskProfile = "balancedPlus";
      else if (text.includes("balanced")) riskProfile = "balanced";
      else if (text.includes("moderately conservative")) riskProfile = "moderatelyConservative";
      else if (text.includes("conservative plus")) riskProfile = "conservativePlus";
      else if (text.includes("conservative") || text.includes("capital preservation")) riskProfile = "conservative";
      else if (age >= 70) riskProfile = "conservativePlus";
      else if (age >= 60) riskProfile = "balanced";
      else if (age >= 50) riskProfile = "balancedPlus";
      else if (age >= 40) riskProfile = "growth";
      else if (age > 0 && age < 40) riskProfile = "growthPlus";
      else if (concentration >= 60) riskProfile = "balanced";
      else if (concentration >= 40) riskProfile = "balancedPlus";
    }

    // Reduce risk one notch when risk tolerance is growth-oriented but planning constraints require liquidity/income.
    if ((liquidityNeed || retirementSoon) && ["growthPlus", "aggressive"].includes(riskProfile)) {
      riskProfile = shiftRiskProfile(riskProfile, -1);
    }

    // Reduce one notch for high concentration because the total household risk is already elevated.
    if (concentration >= 60 && ["growthPlus", "aggressive"].includes(riskProfile)) {
      riskProfile = shiftRiskProfile(riskProfile, -1);
    }

    let portfolioKey = "traditional";
    let modelReason = "Traditional is the default when the proposal emphasizes liquidity, simplicity, and no dedicated alternative allocation.";

    if (focused) {
      portfolioKey = "focusedB";
      modelReason = "Focused B fits when the client wants no U.S. Core equity and a larger U.S. Growth/Value tilt.";
    } else if (usBias && investableAssets >= 10 && !avoidAlternatives) {
      portfolioKey = "selectLiquidityUsBias";
      modelReason = "Select Liquidity — U.S. Bias fits when the client wants a domestic tilt while maintaining liquid alternative exposure.";
    } else if (usBias) {
      portfolioKey = "traditionalUsBias";
      modelReason = "Traditional — U.S. Bias fits when the client wants a domestic tilt and a simpler no-alternatives allocation.";
    } else if (investableAssets >= 10 && !liquidityNeed && !avoidAlternatives) {
      portfolioKey = "corePrivate";
      modelReason = "Core Private fits larger clients where illiquid alternatives may be appropriate.";
    } else if ((alternativesMentioned || investableAssets < 10) && !avoidAlternatives) {
      portfolioKey = "selectLiquidity";
      modelReason = "Select Liquidity fits clients who may benefit from alternative exposure but need more accessible or liquid implementation.";
    } else {
      portfolioKey = "traditional";
      modelReason = "Traditional fits clients prioritizing simplicity, liquidity, and no alternatives.";
    }

    const riskReasonParts = [];
    if (riskNumber) riskReasonParts.push(`Risk Number ${riskNumber} maps to ${getRiskProfileDisplay(riskProfile)}.`);
    if (liquidityNeed) riskReasonParts.push("Near-term liquidity or withdrawal needs support avoiding an overly aggressive allocation.");
    if (retirementSoon) riskReasonParts.push("Retirement or income planning needs support balancing growth with stability.");
    if (concentration >= 40) riskReasonParts.push("High single-stock concentration increases household risk and supports a diversified risk profile.");
    if (!riskReasonParts.length) riskReasonParts.push("The suggested allocation is based on the available time horizon, risk language, and planning objectives.");

    return {
      portfolioKey,
      riskProfile,
      riskNumber,
      portfolioLabel: {
        corePrivate: "Core Private",
        selectLiquidity: "Select Liquidity",
        traditional: "Traditional",
        focusedB: "Focused B",
        selectLiquidityUsBias: "Select Liquidity — U.S. Bias",
        traditionalUsBias: "Traditional — U.S. Bias",
      }[portfolioKey],
      riskLabel: getRiskProfileDisplay(riskProfile),
      modelReason,
      riskReason: riskReasonParts.join(" "),
    };
  }

  function applyInvestmentApproachRecommendationForData(dataForRecommendation) {
    if (!dataForRecommendation) return;

    const rec = recommendInvestmentApproach(dataForRecommendation, notes);

    setSelectedPortfolioStrategies({
      corePrivate: rec.portfolioKey === "corePrivate",
      selectLiquidity: rec.portfolioKey === "selectLiquidity",
      traditional: rec.portfolioKey === "traditional",
      focusedB: rec.portfolioKey === "focusedB",
      selectLiquidityUsBias: rec.portfolioKey === "selectLiquidityUsBias",
      traditionalUsBias: rec.portfolioKey === "traditionalUsBias",
    });

    setSelectedRiskProfile(rec.riskProfile);
  }

  function applyInvestmentApproachRecommendation() {
    applyInvestmentApproachRecommendationForData(reviewData);
  }

  function isBadProposalTicker(value) {
    const t = String(value || "").trim().toUpperCase();

    if (!t) return true;

    const bad = new Set([
      "CORE",
      "SELECT",
      "LIQUIDITY",
      "TRADITIONAL",
      "FOCUSED",
      "FOCUSEDB",
      "BALANCED",
      "PLUS",
      "GROWTH",
      "VALUE",
      "EQUITY",
      "EQUITIES",
      "FIXED",
      "INCOME",
      "ALTERNATIVES",
      "ALTERNATIVE",
      "MUNI",
      "MUNIS",
      "BONDS",
      "CASH",
      "STRATEGY",
      "PORTFOLIO",
      "PRIVATE",
      "PUBLIC",
      "US",
      "USA",
      "INTL",
      "INTERNATIONAL",
      "EMERGING",
      "MARKETS",
      "ETF",
      "ETFS",
      "INDEX",
      "CLIENT",
      "REVIEW",
      "NA",
      "N/A",
    ]);

    if (bad.has(t)) return true;

    // Must look like a real public ticker.
    // Allows BRK.B / BRK-B style tickers.
    if (!/^[A-Z]{1,5}([.-][A-Z])?$/.test(t)) return true;

    return false;
  }


function resolveClientTickerFromNotes(notesText = "", fallbackTicker = "") {
    const raw = String(notesText || "");
    const upperRaw = raw.toUpperCase();

    const normalize = (value) => String(value || "").trim().toUpperCase();

    const accept = (value) => {
      const ticker = normalize(value);
      return !isBadProposalTicker(ticker) ? ticker : "";
    };

    // 1. Explicit ticker fields should win.
    const explicitPatterns = [
      /(?:ticker|symbol|stock ticker|public ticker)\s*[:\-]\s*([A-Z]{1,5}(?:[.\-][A-Z])?)/i,
      /(?:concentrated position|single[-\s]?stock position|stock position)\s*[:\-]\s*([A-Z]{1,5}(?:[.\-][A-Z])?)/i,
      /(?:NVIDIA|Nvidia)\s*(?:Corp|Corporation)?\s*\(?\s*(NVDA)\s*\)?/i,
      /(?:Apple)\s*(?:Inc)?\s*\(?\s*(AAPL)\s*\)?/i,
      /(?:Microsoft)\s*(?:Corp|Corporation)?\s*\(?\s*(MSFT)\s*\)?/i,
      /(?:Tesla)\s*(?:Inc)?\s*\(?\s*(TSLA)\s*\)?/i,
      /(?:Amazon)\s*(?:\.com)?\s*\(?\s*(AMZN)\s*\)?/i,
      /(?:Meta)\s*(?:Platforms)?\s*\(?\s*(META)\s*\)?/i,
      /(?:Alphabet|Google)\s*\(?\s*(GOOGL|GOOG)\s*\)?/i,
    ];

    for (const pattern of explicitPatterns) {
      const match = raw.match(pattern);
      const ticker = accept(match?.[1]);
      if (ticker) return ticker;
    }

    // 2. Company-name mapping. This prevents "CORE" from being chosen from allocation rows.
    const companyMap = [
      ["NVIDIA", "NVDA"],
      ["NVIDIA CORP", "NVDA"],
      ["APPLE", "AAPL"],
      ["MICROSOFT", "MSFT"],
      ["TESLA", "TSLA"],
      ["AMAZON", "AMZN"],
      ["META", "META"],
      ["ALPHABET", "GOOGL"],
      ["GOOGLE", "GOOGL"],
    ];

    for (const [name, ticker] of companyMap) {
      if (upperRaw.includes(name)) return ticker;
    }

    // 3. Fallback ticker only if it is not a strategy/allocation word.
    const fallback = accept(fallbackTicker);
    if (fallback) return fallback;

    return "";
  }


function cleanProposalTicker(value, notesText = "") {
    const resolved = resolveClientTickerFromNotes(notesText, value);
    return resolved || "";
  }


function getSelectedPortfolioStrategyLabels() {
    const labels = {
      corePrivate: "Core Private",
      selectLiquidity: "Select Liquidity",
      traditional: "Traditional",
      focusedB: "Focused B",
      selectLiquidityUsBias: "Select Liquidity (US Bias)",
      traditionalUsBias: "Traditional (US Bias)",
    };

    return Object.keys(selectedPortfolioStrategies)
      .filter((key) => selectedPortfolioStrategies[key])
      .map((key) => labels[key]);
  }

  function getSelectedStrategyLabels() {
    const labels = {
      crt: "Charitable Remainder Trust",
      harvesting: "Leveraged Tax-Loss Harvesting",
      collar: "Option Collar",
    };

    return Object.keys(selectedStrategies)
      .filter((key) => selectedStrategies[key])
      .map((key) => labels[key]);
  }

  function toggleStrategy(key) {
    setSelectedStrategies((prev) => ({
      ...prev,
      [key]: !prev[key],
    }));
  }

  function getAssumptionsList(data) {
    const assumptions = [
      `Client type selected: ${clientType}.`,
      `Strategies selected: ${getSelectedStrategyLabels().join(", ") || "None selected"}.`,
      `CRT allocation is modeled at 32% of the concentrated stock position.`,
      `Tax-loss harvesting sleeve is modeled at 29% of the concentrated stock position.`,
      `Option collar allocation is modeled at 26% of the concentrated stock position.`,
      `CRT payout rate is assumed at ${data?.crtPayoutRate || 5}%.`,
      `Harvested losses are estimated at 20% of the harvesting sleeve.`,
      `Put strike is modeled at 85% of the reference share price.`,
      `Call strike is modeled at 119% of the reference share price.`,
      `Figures are planning estimates and should be confirmed with tax lots, current market pricing, CPA review, and legal review.`,
    ];

    if (!notes.toLowerCase().includes("cost basis")) {
      assumptions.unshift("Cost basis was not clearly provided in the notes, so a default assumption may have been used.");
    }

    if (!/tax rate|combined tax|federal tax|state tax|capital gains tax|california tax/i.test(notes)) {
      assumptions.unshift("Tax rate was not clearly provided in the notes, so default federal/state assumptions may have been used.");
    }

    if (!/share price|stock price|trades at|trading price/i.test(notes)) {
      assumptions.unshift("Current stock price was not clearly provided, so a default reference price may have been used.");
    }

    return assumptions;
  }

  function getCurrentStep() {
    if (proposal) return 4;
    if (reviewData) return 3;
    if (clarifyingQuestions.length > 0) return 2;
    return 1;
  }

  function goBack() {
    const step = getCurrentStep();
    if (step === 4) {
      setProposal(null);
      setStatus("");
    } else if (step === 3) {
      setReviewData(null);
      setQualityReport(null);
      setStatus("");
    } else if (step === 2) {
      setClarifyingQuestions([]);
      setClarificationAnswers({});
      setStatus("");
    }
  }



  function addExtractedDocumentText(extractedText) {
    setNotes((prev) => {
      const divider = prev.trim() ? "\n\n" : "";
      return `${prev}${divider}Uploaded Source Documents:\n${extractedText}`;
    });

    // Try to scan holdings from the uploaded document immediately
    const found = parseHoldingsFromText(extractedText);
    if (found.length >= 2) {
      setScannedHoldings(found);
      setStatus(`Document uploaded — found ${found.length} holdings. Run Agent to continue.`);
    } else {
      setStatus("Document uploaded and added to client notes.");
    }
  }


  // Last-resort fallbacks — only used if live quote AND prior-close API both fail.
  // These are approximate and will be overridden by live data when the server is running.
  const FALLBACK_MARKET_PRICES = {
    NVDA: 205,
    TSLA: 340,
    AAPL: 210,
    MSFT: 470,
    AMZN: 230,
    GOOGL: 185,
    META: 650,
  };

  function forceCollarFromStockPrice(data) {
    const cleaned = { ...data };
    const ticker = String(cleaned.ticker || "").trim().toUpperCase();

    let price = Number(cleaned.stockPrice || cleaned.currentPrice || 0);

    // If the app extracted a stale/default price like $120 for NVDA, override it.
    // The app should use prior close when available; this fallback prevents bad collar math.
    const fallback = FALLBACK_MARKET_PRICES[ticker];
    if (fallback && (!Number.isFinite(price) || price <= 0 || Math.abs(price - fallback) / fallback > 0.15)) {
      price = fallback;
      cleaned.priceSource = cleaned.priceSource || "Fallback market reference price";
      cleaned.priorCloseDate = cleaned.priorCloseDate || "latest available";
    }

    if (Number.isFinite(price) && price > 0) {
      cleaned.stockPrice = price;
      cleaned.putStrike = price * 0.85;
      cleaned.callStrike = price * 1.19;
      cleaned.priceUsedForCollar = price;
    }

    return cleaned;
  }

  function parseMoneyToMillions(rawNumber, rawUnit = "") {
    const n = Number(String(rawNumber || "").replace(/,/g, ""));
    if (!Number.isFinite(n)) return 0;

    const unit = String(rawUnit || "").toLowerCase();

    if (unit === "b" || unit === "bn" || unit === "billion") return n * 1000;
    if (unit === "m" || unit === "mm" || unit === "million") return n;
    if (unit === "k" || unit === "thousand") return n / 1000;

    // App stores money values in millions.
    if (n >= 1000000) return n / 1000000;
    if (n > 0 && n < 100000) return n;

    return n;
  }

  function detectTickerFromNotes(notes = "") {
    const text = String(notes || "");
    const upper = text.toUpperCase();

    const known = {
      NVIDIA: "NVDA",
      TESLA: "TSLA",
      APPLE: "AAPL",
      MICROSOFT: "MSFT",
      AMAZON: "AMZN",
      GOOGLE: "GOOGL",
      ALPHABET: "GOOGL",
      META: "META",
      FACEBOOK: "META",
      NETFLIX: "NFLX",
      AMD: "AMD",
      BROADCOM: "AVGO",
    };

    for (const [company, ticker] of Object.entries(known)) {
      if (upper.includes(company)) return ticker;
    }

    const patterns = [
      /\b(?:ticker|symbol)\s*[:\-]?\s*([A-Z]{2,5})\b/i,
      /\b([A-Z]{2,5})\b\s+(?:stock|shares|position|holding|holdings|concentration)\b/i,
      /(?:stock|shares|position|holding|holdings|concentration)\s+(?:in|of)\s+\b([A-Z]{2,5})\b/i,
    ];

    for (const pattern of patterns) {
      const m = text.match(pattern);
      if (m?.[1]) return m[1].toUpperCase();
    }

    return "";
  }

  function extractConcentratedPositionFromNotes(notes = "", ticker = "") {
    const text = String(notes || "");
    const t = String(ticker || "").trim().toUpperCase();
    const escapedTicker = t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

    const money = String.raw`\$?\s*([\d,.]+)\s*(B|BN|billion|M|MM|million)`;

    const patterns = [];

    if (t) {
      patterns.push(
        new RegExp(`${money}\\s+(?:of\\s+|in\\s+)?\\b${escapedTicker}\\b\\s*(?:stock|shares|position|holding|holdings|concentration)?`, "i"),
        new RegExp(`\\b${escapedTicker}\\b\\s*(?:stock|shares|position|holding|holdings|concentration)?\\s*(?:of|worth|valued at|approximately|around|=|:)\\s*${money}`, "i"),
        new RegExp(`(?:concentrated\\s+)?(?:position|holding|holdings|shares|stock)\\s+(?:in|of)\\s+\\b${escapedTicker}\\b\\s*(?:of|worth|valued at|approximately|around|=|:)\\s*${money}`, "i")
      );
    }

    patterns.push(
      new RegExp(`${money}\\s+(?:concentrated\\s+)?(?:stock|position|holding|holdings|shares|concentration)`, "i"),
      new RegExp(`(?:concentrated\\s+)?(?:stock|position|holding|holdings|shares|concentration)\\s*(?:of|worth|valued at|approximately|around|=|:)\\s*${money}`, "i")
    );

    for (const pattern of patterns) {
      const m = text.match(pattern);
      if (m) return parseMoneyToMillions(m[1], m[2]);
    }

    return 0;
  }

  function applyConcentratedStockFromNotes(data, notes = "") {
    const cleaned = { ...data };

    const tickerFromNotes = detectTickerFromNotes(notes);
    const ticker = String(cleaned.ticker || tickerFromNotes || "").toUpperCase();

    if (ticker) cleaned.ticker = ticker;

    const extractedPosition = extractConcentratedPositionFromNotes(notes, ticker);

    // Only override if notes clearly provide a concentrated stock value.
    if (extractedPosition > 0) {
      cleaned.stockPosition = extractedPosition;
      cleaned.concentratedPosition = extractedPosition;
    }

    cleaned.netWorth = parseMoneyToMillions(cleaned.netWorth);
    cleaned.investableAssets = parseMoneyToMillions(cleaned.investableAssets);
    cleaned.stockPosition = parseMoneyToMillions(cleaned.stockPosition);

    if (cleaned.investableAssets > 0 && cleaned.stockPosition > 0) {
      cleaned.concentration = Math.min(
        Math.max((cleaned.stockPosition / cleaned.investableAssets) * 100, 0),
        100
      );
    }

    return forceCollarFromStockPrice(normalizeMillionsForReview(cleaned));
  }

  function normalizeMillionsForReview(data) {
    const normalize = (value) => {
      const n = Number(value || 0);
      if (!Number.isFinite(n)) return 0;
      return n >= 1000000 ? n / 1000000 : n;
    };

    const cleaned = {
      ...data,
      netWorth: normalize(data.netWorth),
      investableAssets: normalize(data.investableAssets),
      stockPosition: normalize(data.stockPosition),
      income: normalize(data.income),
    };

    cleaned.concentration =
      cleaned.investableAssets > 0 && cleaned.stockPosition > 0
        ? Math.min(Math.max((cleaned.stockPosition / cleaned.investableAssets) * 100, 0), 100)
        : 0;

    return cleaned;
  }


  function runAgent() {
    if (!notes.trim()) {
      alert("Paste client notes first.");
      return;
    }

    const missing = detectMissingInfo(notes);

    if (missing.length > 0 && !notes.includes("Additional Clarifications:")) {
      setClarifyingQuestions(missing);
      setClarificationAnswers({});
      setReviewData(null);
      setProposal(null);
      setStatus("Missing information found. Answer the clarifying questions below.");
      return;
    }

    const name = clientName.trim() || "Client Household";
    let data = applyConcentratedStockFromNotes(safeExtractClientData(notes, name, extractClientData), notes);
    const quality = getProposalQuality(notes, data);

    setQualityReport(quality);
    setClarifyingQuestions([]);
    setClarificationAnswers({});
    {
      const normalized = normalizeMillionsForReview(data);

      // Apply recommendation immediately so dropdowns are pre-filled.
      applyInvestmentApproachRecommendationForData(normalized);
      setReviewData(normalized);

      // Then refresh with live collar pricing and re-apply.
      fetchPreviousClosePrice(normalized.ticker).then((priorClose) => {
        if (priorClose?.close) {
          const updated = forceCollarFromStockPrice(
            normalizeMillionsForReview({
              ...normalized,
              stockPrice: priorClose.close,
              priorCloseDate: priorClose.date,
              priceSource: priorClose.source,
            })
          );

          applyInvestmentApproachRecommendationForData(updated);
          setReviewData(updated);
        }
      });
    }
    setProposal(null);
    setStatus("Review the extracted data below. Edit anything wrong, then confirm and generate.");
  }

  function updateClarificationAnswer(key, value) {
    setClarificationAnswers((prev) => ({
      ...prev,
      [key]: value,
    }));
  }

  function addClarificationsToNotes() {
    const added = clarifyingQuestions
      .map((item) => `${item.label}: ${clarificationAnswers[item.key] || "Not provided"}`)
      .join("\n");

    setNotes(notes + "\n\nAdditional Clarifications:\n" + added);
    setClarifyingQuestions([]);
    setClarificationAnswers({});
    setStatus("Clarifications added. Click Run Agent again to review extracted data.");
  }

  function generateWithAssumptions() {
    const name = clientName.trim() || "Client Household";
    let data = applyConcentratedStockFromNotes(safeExtractClientData(notes, name, extractClientData), notes);
    const quality = getProposalQuality(notes, data);

    setQualityReport(quality);
    setClarifyingQuestions([]);
    setClarificationAnswers({});
    {
      const normalized = normalizeMillionsForReview(data);

      // Apply recommendation immediately so dropdowns are pre-filled.
      applyInvestmentApproachRecommendationForData(normalized);
      setReviewData(normalized);

      // Then refresh with live collar pricing and re-apply.
      fetchPreviousClosePrice(normalized.ticker).then((priorClose) => {
        if (priorClose?.close) {
          const updated = forceCollarFromStockPrice(
            normalizeMillionsForReview({
              ...normalized,
              stockPrice: priorClose.close,
              priorCloseDate: priorClose.date,
              priceSource: priorClose.source,
            })
          );

          applyInvestmentApproachRecommendationForData(updated);
          setReviewData(updated);
        }
      });
    }
    setProposal(null);
    setStatus("Review the extracted data below. Assumptions were used for missing details.");
  }

  function updateReviewData(key, value) {
    setReviewData((prev) => ({
      ...prev,
      [key]: key === "ticker" ? value : Number(value),
    }));
  }

  // Alternatives (Core Private / Select Liquidity) require $10M+ net worth.
  // Below that, auto-fallback to Traditional (no alternatives sleeve).
  const ALT_NET_WORTH_THRESHOLD_M = 10;
  const ALT_MODEL_KEYS = new Set(["corePrivate", "selectLiquidity", "selectLiquidityUsBias"]);
  const ALT_PORTFOLIO_LABELS = {
    corePrivate:           "Core Private",
    selectLiquidity:       "Select Liquidity",
    selectLiquidityUsBias: "Select Liquidity — U.S. Bias",
  };

  // Re-checks the qualification gate against the current portfolio selection
  // and a given net worth. Flips selectedPortfolioStrategies to Traditional
  // and returns fallback info if the client doesn't qualify; otherwise null.
  function applyAltQualificationGate(netWorthM) {
    const activeKey = Object.keys(selectedPortfolioStrategies).find(k => selectedPortfolioStrategies[k]) || "";
    if (!ALT_MODEL_KEYS.has(activeKey) || (netWorthM || 0) >= ALT_NET_WORTH_THRESHOLD_M) return null;

    setSelectedPortfolioStrategies({
      corePrivate: false,
      selectLiquidity: false,
      traditional: true,
      focusedB: false,
      selectLiquidityUsBias: false,
      traditionalUsBias: false,
    });
    return { from: ALT_PORTFOLIO_LABELS[activeKey] || activeKey, to: "Traditional" };
  }

  function confirmAndGenerate() {
    if (!reviewData) return;
    const name = clientName.trim() || "Client Household";
    const finalData = recomputeReviewedData(normalizeMillionsForReview(reviewData));

    let portfolioKey = Object.keys(selectedPortfolioStrategies).find(k => selectedPortfolioStrategies[k]) || "";
    const PORTFOLIO_LABELS = {
      corePrivate:           "Core Private",
      selectLiquidity:       "Select Liquidity",
      selectLiquidityUsBias: "Select Liquidity — U.S. Bias",
      traditional:           "Traditional",
      traditionalUsBias:     "Traditional — U.S. Bias",
      focusedB:              "Focused B",
    };

    // Qualification gate: alternatives need net worth >= $10M
    const altFallback = applyAltQualificationGate(finalData.netWorth);
    if (altFallback) portfolioKey = "traditional";

    const riskOpt = riskProfileOptions.find(o => o.key === selectedRiskProfile);

    setPreviewModal({
      name,
      finalData,
      portfolioKey,
      portfolioModelLabel: PORTFOLIO_LABELS[portfolioKey] || "",
      riskProfileKey:      selectedRiskProfile || "",
      riskProfileLabel:    riskOpt ? `${riskOpt.label} — ${riskOpt.mix}` : "",
      altFallback,
    });
  }

  function proceedFromPreview(editedData) {
    if (!previewModal) return;
    const { name } = previewModal;
    const finalData = editedData || previewModal.finalData;

    // Re-check the qualification gate in case net worth was edited inside the
    // preview modal (e.g. lowered below $10M after CorePrivate/SelectLiquidity
    // was originally selected).
    const lateFallback = applyAltQualificationGate(finalData.netWorth);
    if (lateFallback) {
      setStatus(`Net worth below $10M — switched portfolio model from ${lateFallback.from} to ${lateFallback.to} (alternatives require $10M+ net worth).`);
    }

    setProposal({
      name,
      notes,
      data: finalData,
      clientType,
      selectedStrategies,
      selectedStrategyLabels: getSelectedStrategyLabels(),
      assumptions: getAssumptionsList(finalData),
      // Snapshot of the historical backtest shown in the preview modal (see
      // runPortfolioBacktest), carried forward so the PPTX Transition Analysis
      // slide can show the same numbers without refetching.
      backtest: backtestResult,
    });
    setPreviewModal(null);
    setReviewData(null);
    if (!lateFallback) setStatus("Proposal generated. Download the Word document or PowerPoint.");
  }

  function closePreviewModal() {
    setPreviewModal(null);
  }

  // ── Portfolio Transition Backtest (current vs. target portfolio) ────────────
  // Uses REAL historical monthly prices via /api/history/:ticker — never fabricated
  // numbers. "Current portfolio" is approximated from the same freeform-notes
  // extraction already used elsewhere in the app (data.ticker + data.concentration
  // for a concentrated stock position), with the untracked remainder stood in by a
  // diversified allocation ETF matched to the client's selected risk profile. Any
  // line that can't be tied to a public ticker is excluded (never zero-weighted) and
  // reported back via coveragePct/excluded so the gap can be disclosed in the UI.

  // Diversified "everything else" proxies (iShares Core Allocation ETFs), picked by
  // closest stock/bond mix to the client's selected risk profile.
  const BENCHMARK_BY_STOCK_WEIGHT = [
    { ticker: "AOK", stock: 30 },  // Conservative   (~30/70)
    { ticker: "AOM", stock: 40 },  // Moderate       (~40/60)
    { ticker: "AOR", stock: 60 },  // Growth         (~60/40)
    { ticker: "AOA", stock: 80 },  // Aggressive     (~80/20)
  ];
  const STOCK_WEIGHT_BY_RISK_PROFILE = {
    conservative: 20, moderatelyConservative: 30, conservativePlus: 40,
    balanced: 50, balancedPlus: 60, growth: 70, growthPlus: 80, aggressive: 100,
  };
  function pickDiversifiedBenchmarkTicker(riskProfileKey) {
    const targetStock = STOCK_WEIGHT_BY_RISK_PROFILE[riskProfileKey] ?? 60;
    let best = BENCHMARK_BY_STOCK_WEIGHT[0];
    for (const b of BENCHMARK_BY_STOCK_WEIGHT) {
      if (Math.abs(b.stock - targetStock) < Math.abs(best.stock - targetStock)) best = b;
    }
    return best.ticker;
  }

  // Published expense ratios for the iShares Core Allocation ETFs used as the
  // "everything else" proxy — approximate, for illustrative fee comparison only.
  const BENCHMARK_FEE_BY_TICKER = { AOK: 0.15, AOM: 0.15, AOR: 0.15, AOA: 0.15 };

  async function fetchTickerHistory(ticker, range = "10y") {
    try {
      const res = await fetch(`/api/history/${encodeURIComponent(ticker)}?range=${range}`);
      if (!res.ok) {
        // Surface *why* — e.g. "Failed to fetch" / network error here usually means
        // the Express backend (server.js, port 5174) isn't running, while a JSON
        // error body means the backend IS reachable but the upstream price-history
        // provider rejected the request. Distinguishing these is the difference
        // between "start the backend" and "the data provider is down/rate-limited."
        let detail = `HTTP ${res.status}`;
        try {
          const errJson = await res.json();
          if (errJson?.error) detail = errJson.error;
        } catch { /* body wasn't JSON — keep the HTTP status as the detail */ }
        console.warn(`[backtest] History fetch failed for ${ticker}: ${detail}`);
        return null;
      }
      const json = await res.json();
      return Array.isArray(json?.series) ? json.series : null;
    } catch (err) {
      // Fetch threw before getting a response at all — almost always means the
      // backend at localhost:5174 isn't running or isn't reachable via the Vite
      // proxy (e.g. the app was opened via `vite preview`/static dist instead of
      // `npm run dev`, which is the only mode that proxies /api to the backend).
      console.warn(`[backtest] History fetch threw for ${ticker} (is "npm run dev" / server.js running?):`, err);
      return null;
    }
  }

  async function runPortfolioBacktest({ portfolioKey, riskProfileKey, data }) {
    setBacktestLoading(true);
    setBacktestError("");
    setBacktestResult(null);
    try {
      const funds = getFunds(portfolioKey, riskProfileKey); // alloc > 0 only

      // Proxy non-tradable (SMA/N/A/CUSTOM) fund lines with the cheapest tradable
      // passive option from the same asset class, where one exists.
      const proxyMap = {};
      for (const f of funds) {
        if (proxyMap[f.assetClass] != null) continue;
        proxyMap[f.assetClass] = getCheapestPassiveTicker(f.assetClass) || "";
      }

      const targetMap = buildTargetWeightMap(funds, proxyMap);
      const benchmarkTicker = pickDiversifiedBenchmarkTicker(riskProfileKey);

      // Weighted average fee (%) of the recommended target portfolio, from the
      // model's own published fund fees — independent of price history, so this
      // is always available even when a backtest can't be run.
      const targetWeightedFeePct = funds.reduce(
        (s, f) => s + ((Number(f.alloc) || 0) / 100) * (Number(f.fee) || 0), 0
      );

      // Prefer the live-scanned holdings (from upload) over whatever safeExtractClientData found
      const effectiveHoldings = scannedHoldings.length >= 2 ? scannedHoldings : (data?.currentHoldings || []);
      const hasActualHoldings = effectiveHoldings.length >= 2;
      const hasConcentration  = isUsableTicker(data?.ticker) && Number(data?.concentration) > 0;
      const currentMap = hasActualHoldings
        ? buildHoldingsWeightMap(effectiveHoldings)
        : hasConcentration
        ? buildConcentratedWeightMap(data.ticker, data.concentration, benchmarkTicker)
        : null;
      const concPct = hasConcentration ? Math.min(Math.max(Number(data.concentration) || 0, 0), 100) : null;
      // Expense ratios for common ETFs / mutual funds (as % — e.g. 0.03 = 3 bps).
      // Individual stocks have 0% ongoing fee. Anything not in this map defaults to 0.
      const HOLDING_FEE_MAP = {
        VOO:0.03,VTI:0.03,VXUS:0.07,BND:0.03,BNDX:0.07,
        IVV:0.03,IWM:0.19,QQQ:0.20,SPY:0.0945,
        VO:0.04,VB:0.05,VEA:0.05,VWO:0.08,VIG:0.06,VNQ:0.12,VYM:0.06,VUG:0.04,VTV:0.04,
        VTEB:0.05,MUB:0.05,SUB:0.07,SCHP:0.03,
        QUAL:0.15,USMV:0.15,MTUM:0.15,HDV:0.08,
        VTSAX:0.04,VFIAX:0.04,VIGAX:0.05,VTIAX:0.11,VBTLX:0.05,VMFXX:0.11,VWUAX:0.26,
        AGG:0.03,LQD:0.14,HYG:0.48,TLT:0.15,IEF:0.15,SHV:0.15,BIL:0.14,
        GLD:0.40,IAU:0.25,SLV:0.50,
        ARTKX:0.63,APHIX:0.96,DODFX:0.63,ODVIX:0.88,ARSIX:1.18,POLIX:0.80,
        JHQDX:0.60,RAPIX:0.81,CCLFX:2.32,BPMAX:3.29,GRIFX:1.72,
        // individual stocks → 0 (not in map, falls back to 0 below)
      };
      const currentWeightedFeePct = hasActualHoldings
        ? (() => {
            // Weight-average fees across holdings; stocks not in the map → 0% fee
            const h = effectiveHoldings;
            const total = h.reduce((s, x) => s + (Number(x.pct) || 0), 0);
            if (total <= 0) return null;
            const blended = h.reduce((s, x) => {
              const fee = HOLDING_FEE_MAP[(x.ticker || "").toUpperCase()] ?? 0;
              return s + fee * (Number(x.pct) || 0) / total;
            }, 0);
            return blended;
          })()
        : hasConcentration
        ? (1 - concPct / 100) * (BENCHMARK_FEE_BY_TICKER[benchmarkTicker] ?? 0)
        : null;

      const allTickers = new Set([
        ...Object.keys(targetMap.weights).filter(t => t !== "__proxied"),
        ...(currentMap ? Object.keys(currentMap.weights) : []),
      ]);

      if (allTickers.size === 0) {
        setBacktestResult({
          target: null,
          current: null,
          benchmarkTicker,
          note: "Not enough publicly-tradable tickers in this model to run a historical backtest.",
        });
        return;
      }

      const returnsByTicker = {};
      await Promise.all([...allTickers].map(async (t) => {
        const series = await fetchTickerHistory(t);
        if (series) returnsByTicker[t] = toMonthlyReturns(series);
      }));

      const targetReturns = weightedPortfolioReturns(targetMap.weights, returnsByTicker);
      const targetSummary = summarizeReturns(targetReturns);
      // "Annualized Return" shown to the client is the weighted average of each
      // holding's own trailing 12-month returns over (up to) the last 10 years —
      // not the CAGR of the combined series — per the firm's preferred presentation.
      const targetAvgAnnualReturn = weightedAverageAnnualReturn(targetMap.weights, returnsByTicker);

      let currentSummary = null;
      let currentAvgAnnualReturn = null;
      if (currentMap) {
        const currentReturns = weightedPortfolioReturns(currentMap.weights, returnsByTicker);
        console.log("[backtest] current portfolio: window", currentReturns[0]?.date, "→", currentReturns[currentReturns.length-1]?.date, `(${currentReturns.length} months)`);
        currentSummary = summarizeReturns(currentReturns);
        console.log("[backtest] current maxDrawdown:", (currentSummary?.maxDrawdown * 100).toFixed(1) + "%", "| annualized:", (currentSummary?.annualizedReturn * 100).toFixed(1) + "%");
        currentAvgAnnualReturn = weightedAverageAnnualReturn(currentMap.weights, returnsByTicker);
      }

      setBacktestResult({
        target: {
          summary: targetSummary,
          avgAnnualReturn: targetAvgAnnualReturn,
          coveragePct: targetMap.coveragePct,
          excluded: targetMap.excluded,
          weightedFeePct: targetWeightedFeePct,
        },
        current: currentMap
          ? {
              summary: currentSummary,
              avgAnnualReturn: currentAvgAnnualReturn,
              coveragePct: currentMap.coveragePct,
              ticker: data.ticker,
              concentration: data.concentration,
              weightedFeePct: currentWeightedFeePct,
              fromUploadedHoldings: hasActualHoldings,
              holdingCount: hasActualHoldings ? effectiveHoldings.length : null,
            }
          : null,
        benchmarkTicker,
        benchmarkFeePct: BENCHMARK_FEE_BY_TICKER[benchmarkTicker] ?? null,
        missingTickers: [...allTickers].filter(t => !returnsByTicker[t]),
      });
    } catch (err) {
      console.warn("Portfolio backtest failed:", err);
      setBacktestError(err.message || "Could not run portfolio backtest.");
    } finally {
      setBacktestLoading(false);
    }
  }

  // Run the backtest once when the preview modal opens for a new proposal.
  useEffect(() => {
    if (!previewModal) {
      setBacktestResult(null);
      setBacktestLoading(false);
      setBacktestError("");
      return;
    }
    runPortfolioBacktest({
      portfolioKey: previewModal.portfolioKey,
      riskProfileKey: previewModal.riskProfileKey,
      data: previewModal.finalData,
    });
  }, [previewModal]);

  function cancelReview() {
    setReviewData(null);
    setQualityReport(null);
    setStatus("Review canceled. Edit the notes and run the agent again.");
  }

  function tableRow(label, value) {
    return new TableRow({
      children: [
        new TableCell({ children: [new Paragraph(label)] }),
        new TableCell({ children: [new Paragraph(value)] }),
      ],
    });
  }

  async function downloadWord() {
    if (!proposal) {
      alert("Run the agent first.");
      return;
    }

    try {
      setStatus("Creating concise 1–2 page Word memo...");

      const { name, notes, data } = proposal;
      const assumptions = proposal.assumptions || getAssumptionsList(data);
      const strategyLabels = proposal.selectedStrategyLabels || getSelectedStrategyLabels();

      const concentrationPctForPpt =
        Number(data.concentration) > 1
          ? Number(data.concentration)
          : Number(data.concentration || 0) * 100;

      const isHighConcentrationClient =
        concentrationPctForPpt >= 25 &&
        Number(data.stockPosition || 0) > 0 &&
        !!data.ticker;

      const hasConcentratedStockStrategy =
        !!selectedStrategies?.crt ||
        !!selectedStrategies?.harvesting ||
        !!selectedStrategies?.collar;

      const includeConcentratedStockSlides =
        isHighConcentrationClient && hasConcentratedStockStrategy;

      // Normal-client branch removed. Using concentrated-stock deck only.


      const bullet = (txt) =>
        new Paragraph({
          text: txt,
          bullet: { level: 0 },
          spacing: { after: 80 },
        });

      const smallSpace = () => new Paragraph({ text: "", spacing: { after: 60 } });

      const section = (txt) =>
        new Paragraph({
          text: txt,
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 120, after: 80 },
        });

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: 720,
                  right: 720,
                  bottom: 720,
                  left: 720,
                },
              },
            },
            children: [
              new Paragraph({
                text: `${name} Investment Proposal Memo`,
                heading: HeadingLevel.TITLE,
              }),
              new Paragraph("Prepared by Investment Proposal Project"),
              new Paragraph("Confidential — For discussion purposes only"),
              smallSpace(),

              section("Executive Summary"),
              new Paragraph(
                `The client type is ${proposal.clientType || clientType}. The client holds a concentrated ${data.ticker} position of approximately ${fmtM(data.stockPosition)}, representing about ${pct(data.concentration)} of investable assets. The recommended plan includes: ${strategyLabels.join(", ")}. Together, these tools are designed to reduce single-stock risk, manage taxes, protect downside, and create a more controlled transition out of the position.`
              ),

              section("Key Client Numbers"),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                  tableRow("Total Net Worth", fmtM(data.netWorth)),
                  tableRow("Investable Assets", fmtM(data.investableAssets)),
                  tableRow(`${data.ticker} Position`, fmtM(data.stockPosition)),
                  tableRow("Current Concentration", pct(data.concentration)),
                  tableRow("Estimated Embedded Gain", fmtM(data.embeddedGain)),
                  tableRow("Estimated Tax If Fully Sold Today", fmtM(data.immediateTax)),
                  tableRow("Annual Income", data.income ? fmtM(data.income) : "Not provided"),
                ],
              }),

              section("Recommended Strategy"),
              bullet(`Charitable Remainder Trust: contribute approximately ${fmtM(data.crtAllocation)} of appreciated ${data.ticker} shares.`),
              bullet(`130/30 Leveraged Tax-Loss Harvesting: allocate approximately ${fmtM(data.harvestingSleeve)} to create tax losses that may offset future gains.`),
              bullet(`Option Collar: protect approximately ${fmtM(data.collarAllocation)} of remaining shares with a put floor and call cap.`),

              section("1. Charitable Remainder Trust"),
              bullet(`Contribution amount: ${fmtM(data.crtAllocation)} of appreciated ${data.ticker} shares.`),
              bullet(`Estimated annual payout: ${fmtK(data.crtIncome)} per year at a ${pct(data.crtPayoutRate)} payout rate.`),
              bullet(`Estimated charitable deduction: ${fmtM(data.charitableDeductionLow)}–${fmtM(data.charitableDeductionHigh)}.`),
              bullet(`Planning purpose: reduce concentration, create income, diversify inside the trust, and support charitable goals.`),

              ...(selectedStrategies.harvesting
                ? [
                    section("2. 130/30 Leveraged Tax-Loss Harvesting"),
                    bullet(`Harvesting sleeve: ${fmtM(data.harvestingSleeve)}.`),
                    bullet(`Estimated first-year harvested losses: ${fmtM(data.annualHarvestLosses)}.`),
                    bullet(`Estimated tax savings: ${fmtK(data.federalTaxSavings)} federal + ${fmtK(data.stateTaxSavings)} state = ${fmtK(data.taxSavings)} total.`),
                  ]
                : []),
              bullet(`Planning purpose: create tax losses that may offset future gains from staged ${data.ticker} sales.`),

              section("3. Option Collar"),
              bullet(`Protected position: ${fmtM(data.collarAllocation)}.`),
              bullet(`Reference price: $${data.stockPrice.toFixed(2)} per share.`),
              bullet(`Illustrative put floor: $${data.putStrike.toFixed(2)}. This helps limit downside below that level.`),
              bullet(`Illustrative call cap: $${data.callStrike.toFixed(2)}. This caps upside above that level in exchange for protection.`),
              bullet(`Planning purpose: protect remaining shares without triggering an immediate taxable sale.`),

              section("Combined Impact"),
              bullet(`Current concentration: ${pct(data.concentration)}.`),
              bullet(`Estimated concentration after CRT: ${pct(data.afterCrtConcentration)}.`),
              bullet(`Estimated 40% stock drawdown risk before strategy: ${fmtM(data.drawdown40Impact)}.`),
              bullet(`The CRT reduces exposure, the harvesting sleeve creates tax flexibility, and the collar protects the remaining position during the transition.`),

              section("Next Steps"),
              bullet("Confirm exact stock value, cost basis, tax lots, and trading restrictions."),
              bullet("Review CRT structure with the estate attorney and CPA."),
              bullet("Design and fund the 130/30 harvesting sleeve."),
              bullet("Price the option collar using current market options data."),
              bullet("Coordinate implementation with the advisor, CPA, attorney, and options desk."),

              section("Assumptions Used"),
              ...assumptions.map((item) => bullet(item)),

              section("Important Notes"),
              new Paragraph(
                "Figures are estimates based on the client-provided information in the notes. This memo is for discussion purposes only and is not legal, tax, or investment advice."
              ),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      downloadBlob(blob, `${name.replaceAll(" ", "_")}_Concise_Investment_Proposal_Memo.docx`);
      setStatus("Concise 1–2 page Word memo downloaded.");
    } catch (error) {
      console.error(error);
      alert("Word document failed. Open Inspect → Console to see the error.");
      setStatus("Word document failed.");
    }
  }



  async function downloadPowerPoint() {
    if (!proposal) { alert("Run the agent first."); return; }
    try {
      setStatus("Creating PowerPoint...");

      const { name, notes } = proposal;
      let data = { ...proposal.data };

      const resolvedTicker = await resolveTickerFromNotes(notes, data.ticker);
      if (resolvedTicker) data.ticker = resolvedTicker;

      // A stock price already on the proposal — whether originally fetched or
      // manually edited by the advisor in the preview modal — is authoritative.
      // Only fetch a fresh one if there genuinely isn't a usable price yet, so
      // edits made under "All Calculated Numbers" aren't silently clobbered.
      const hasUsablePrice = Number(data.stockPrice) > 0;

      if (!hasUsablePrice) {
        // data.stockPrice below is always the prior trading day's close — this
        // app has no live/intraday quote source, by design, so the collar is
        // priced consistently with the rest of the proposal.
        const quote = await fetchPreviousClosePrice(data.ticker);
        if (quote?.close) {
          data = {
            ...data,
            stockPrice: quote.close,
            putStrike: quote.close * 0.85,
            callStrike: quote.close * 1.19,
            priorCloseDate: quote.date,
            priceSource: quote.source,
          };
        }
      }

      // Overlay illustrative options desk data (strikes/premium estimates) when
      // available. IMPORTANT: the reference stock price always stays whatever is
      // already on `data` (either the advisor's edited value or the prior close
      // fetched above) — it is never replaced by anything from collarOptions.
      if (hasUsablePrice ? (data.putStrike == null || data.callStrike == null) : (collarOptions?.put?.strike || collarOptions?.call?.strike)) {
        const refPrice = data.stockPrice;
        const collarPutStrike  = collarOptions?.put?.strike  ?? refPrice * 0.85;
        const collarCallStrike = collarOptions?.call?.strike ?? refPrice * 1.19;
        data = {
          ...data,
          putStrike: hasUsablePrice && data.putStrike != null ? data.putStrike : collarPutStrike,
          callStrike: hasUsablePrice && data.callStrike != null ? data.callStrike : collarCallStrike,
        };
      }
      if (collarOptions?.put?.strike || collarOptions?.call?.strike) {
        const refPrice = data.stockPrice;
        data = {
          ...data,
          putFloorValue:  hasUsablePrice && data.putFloorValue != null ? data.putFloorValue : data.collarAllocation * (data.putStrike  / refPrice),
          callCapValue:   hasUsablePrice && data.callCapValue  != null ? data.callCapValue  : data.collarAllocation * (data.callStrike / refPrice),
          collarPutPremium:  collarOptions.put?.premium  ?? null,
          collarCallPremium: collarOptions.call?.premium ?? null,
          collarNetCost:     collarOptions.netCost ?? null,
          collarNetCostLabel: collarOptions.netCostPerShare ?? "See advisor",
          collarExpiration:  collarOptions.expirationLabel ?? "",
          collarImpliedVol:  collarOptions.put?.impliedVolatility ?? null,
          // No live options-chain feed exists in this app — always prior close.
          collarLive: false,
          priceSource: hasUsablePrice ? data.priceSource : (collarOptions.source || data.priceSource),
          priorCloseDate: hasUsablePrice ? data.priorCloseDate : (collarOptions.asOfDate || data.priorCloseDate),
        };
      }

      const isHighConcentrationClient =
        Number(data.concentration || 0) >= 20 || Number(data.stockPosition || 0) >= 1;
      const hasConcentratedStockStrategy =
        !!selectedStrategies?.crt ||
        !!selectedStrategies?.harvesting ||
        !!selectedStrategies?.collar;
      const includeConcentratedStockSlidesPpt = isHighConcentrationClient && hasConcentratedStockStrategy;

      const riskNumberFromNotes = extractRiskNumberFromText(notes);

      const blob = await generatePowerPoint({
        name,
        notes,
        data,
        selectedStrategies,
        selectedProposalModules,
        selectedPortfolioStrategies,
        selectedRiskProfile,
        selectedServices,
        // Portfolio strategy, risk profile/allocation, and the transition slides
        // only render when the "Recommended Investment Approach" box is checked.
        includePortfolioStrategySlides: selectedProposalModules.recommendedInvestmentApproach === true,
        includeConcentratedStockSlides: includeConcentratedStockSlidesPpt,
        clientType,
        strategyLabels: getSelectedStrategyLabels(),
        assumptions: getAssumptionsList(data),
        riskNumber: riskNumberFromNotes,
        fundSwaps: data.fundSwaps || {},
        backtest: proposal.backtest || null,
        firmName: firmName.trim() || "Meridian Wealth Partners",
        advisorName: advisorName.trim(),
        liveStrategyAllocations,
      });

      downloadBlob(blob, `${name.replaceAll(" ", "_")}_Investment_Proposal.pptx`);
      setStatus("PowerPoint downloaded.");
    } catch (error) {
      console.error("PowerPoint generation error:", error);
      alert("PowerPoint failed: " + (error?.message || error));
      setStatus("PowerPoint failed.");
    }
  }


  return (
    <div className={`page ${reviewData || proposal || clarifyingQuestions.length > 0 ? "review-mode" : ""}`}>
      <header className="topbar">
        {getCurrentStep() > 1 && (
          <button onClick={goBack} className="back-btn">
            ← Back
          </button>
        )}
        <h1>Investment Proposal Project</h1>
        <div className="badge">AI-POWERED</div>
      </header>

      <main className="container">
        {!reviewData && !proposal && clarifyingQuestions.length === 0 && (
        <section className="card firm-settings-card">
          <h2>Firm Settings</h2>
          <div className="firm-settings-grid">
            <div className="firm-settings-field">
              <label>Firm Name</label>
              <input
                placeholder="e.g. Meridian Wealth Partners"
                value={firmName}
                onChange={(e) => setFirmName(e.target.value)}
              />
            </div>
            <div className="firm-settings-field">
              <label>Advisor Name</label>
              <input
                placeholder="e.g. Jane Smith"
                value={advisorName}
                onChange={(e) => setAdvisorName(e.target.value)}
              />
            </div>
          </div>
        </section>
        )}

        {!reviewData && !proposal && clarifyingQuestions.length === 0 && (
        <section className="card">
          <h2>Paste Client Notes</h2>
          <p className="intro">
            Paste client notes, financials, concentrated stock details, tax concerns, goals, and legacy objectives.
          </p>

          <FileUploadBox onTextExtracted={addExtractedDocumentText} />


          <input
            placeholder="Client name, e.g. Mercer Household"
            value={clientName}
            onChange={(e) => setClientName(e.target.value)}
          />

          <textarea
            placeholder={`Paste your client notes here...

Example:
Client has $50M net worth, $30M investable assets, $18M AAPL position, 60% concentration, $4M annual income, 37% tax rate, 12% cost basis, current share price $185. Wants to reduce concentration, preserve upside, and support charitable giving.`}
            value={notes}
            onChange={(e) => setNotes(e.target.value)}
          />

          <button onClick={runAgent}>Run Agent</button>
        </section>
        )}

        <div className="output-panel">
          <div className="empty-state">
          </div>

          {status && <p className="status">{status}</p>}

          <MissingInfoPanel
            questions={clarifyingQuestions}
            answers={clarificationAnswers}
            onAnswerChange={updateClarificationAnswer}
            onAddAnswers={addClarificationsToNotes}
            onGenerateAnyway={generateWithAssumptions}
          />

          {qualityReport && reviewData && (
            <div className="proposal-quality-panel">
              <div className="quality-score-circle">
                {qualityReport.score}%
              </div>

              <div className="quality-score-content">
                <h3>Proposal Quality Score</h3>
                <p>
                  This score estimates how complete the client notes are before generation.
                </p>

                <div className="quality-columns">
                  <div>
                    <h4>Strong</h4>
                    <ul>
                      {qualityReport.strengths.map((item, index) => (
                        <li key={index}>✓ {item}</li>
                      ))}
                    </ul>
                  </div>

                  <div>
                    <h4>Needs Review</h4>
                    <ul>
                      {qualityReport.gaps.map((item, index) => (
                        <li key={index}>⚠ {item}</li>
                      ))}
                    </ul>
                  </div>
                </div>
              </div>
            </div>
          )}

          {scannedHoldings.length >= 2 && (
            <div className="strategy-review-box strategy-review-box-clean">
              <h4>✓ Holdings Scanned — {scannedHoldings.length} positions found</h4>
              <p>These will be used as the current portfolio in the backtest comparison.</p>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "6px 10px", marginTop: 6 }}>
                {scannedHoldings.map(h => (
                  <span key={h.ticker} style={{ fontSize: 11, background: "#F0F4FF", border: "1px solid #C0D4F5", borderRadius: 6, padding: "3px 9px", color: "var(--navy)", fontWeight: 600 }}>
                    {h.ticker} <span style={{ fontWeight: 400, color: "#6b7a99" }}>{h.pct.toFixed(1)}%</span>
                  </span>
                ))}
              </div>
              <div style={{ marginTop: 10, fontSize: 11, color: "#9bacc8" }}>
                Not right?{" "}
                <label style={{ cursor: "pointer", color: "var(--blue)", textDecoration: "underline", fontWeight: 600 }}>
                  Upload a simple holdings Excel instead
                  <input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={async (e) => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    const { default: XLSX } = await import("xlsx");
                    const ab = await file.arrayBuffer();
                    const wb = XLSX.read(ab, { type: "array" });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    const csv = XLSX.utils.sheet_to_csv(ws);
                    const found = parseHoldingsFromText(csv);
                    if (found.length >= 2) setScannedHoldings(found);
                    e.target.value = "";
                  }} />
                </label>
                {" "}(two columns: Ticker, %)
              </div>
            </div>
          )}

          {reviewData && (
            <div className="review-panel">
              <h3>Review Extracted Data</h3>
              <p>Edit anything wrong before generating the final proposal.</p>

              <div className="review-grid">
                <label>
                  Ticker / Company
                  <input
                    value={reviewData.ticker || ""}
                    onChange={(e) => updateReviewData("ticker", e.target.value)}
                  />
                </label>

                <label>
                  Net Worth ($M)
                  <input
                    type="number"
                    value={reviewData.netWorth || 0}
                    onChange={(e) => updateReviewData("netWorth", e.target.value)}
                  />
                </label>

                <label>
                  Investable Assets ($M)
                  <input
                    type="number"
                    value={reviewData.investableAssets || 0}
                    onChange={(e) => updateReviewData("investableAssets", e.target.value)}
                  />
                </label>

                <label>
                  Concentrated Position ($M)
                  <input
                    type="number"
                    value={reviewData.stockPosition || 0}
                    onChange={(e) => updateReviewData("stockPosition", e.target.value)}
                  />
                </label>

                <label>
                  Cost Basis (%)
                  <input
                    type="number"
                    value={reviewData.costBasisPct || 0}
                    onChange={(e) => updateReviewData("costBasisPct", e.target.value)}
                  />
                </label>

                <label>
                  Tax Rate (%)
                  <input
                    type="number"
                    value={reviewData.taxRate || reviewData.totalTaxRate || 0}
                    onChange={(e) => updateReviewData("taxRate", e.target.value)}
                  />
                </label>
              </div>

              <div className="strategy-review-box strategy-review-box-clean proposal-modules-box">
                <h4>Slide Selection</h4>
                <p>Choose which slides to include in the PowerPoint deck.</p>

                <div className="module-section-title">Core Slides</div>
                <div className="slide-selection-grid">
                  {[
                    { key: "executiveSummary", label: "Executive Summary", slides: "1 slide", desc: "Key metrics and strategy overview." },
                    { key: "clientProfileGoals", label: "Client Profile & Goals", slides: "1 slide", desc: "Background, goals, and planning timeline." },
                    { key: "financialPicture", label: "Your Financial Picture", slides: "1 slide", desc: "Planning-scope coverage: cash, tax, real estate, estate." },
                    { key: "capabilities", label: "How We Help (Capabilities)", slides: "1 slide", desc: "Three-pillar firm capabilities / services overview." },
                    { key: "taxOnTheTable", label: "Tax You're Leaving on the Table", slides: "1 slide", desc: "Hero slide: tax saved + downside protected by the plan." },
                    { key: "recommendedInvestmentApproach", label: "Recommended Investment Approach", slides: "2 slides", desc: "Portfolio strategy table and allocation chart." },
                    { key: "nextSteps", label: "Next Steps", slides: "1 slide", desc: "Action plan and implementation checklist." },
                  ].map((option) => (
                    <label key={option.key} className={`slide-card ${selectedProposalModules[option.key] ? "slide-card--selected" : ""}`}>
                      <input
                        type="checkbox"
                        checked={!!selectedProposalModules[option.key]}
                        onChange={(e) =>
                          setSelectedProposalModules((prev) => ({ ...prev, [option.key]: e.target.checked }))
                        }
                      />
                      <div className="slide-card-body">
                        <div className="slide-card-top">
                          <span className="slide-card-label">{option.label}</span>
                          <span className="slide-card-badge">{option.slides}</span>
                        </div>
                        <span className="slide-card-desc">{option.desc}</span>
                      </div>
                    </label>
                  ))}
                </div>

                <div className="module-section-title" style={{ marginTop: "1.25rem" }}>Planning Module Slides</div>
                <div className="slide-selection-grid">
                  {[
                    { key: "riskManagementOverview", label: "Risk Assessment Overview", slides: "1 slide", desc: "Risk spectrum and downside/upside scenario comparison." },
                    { key: "estatePlanningReview", label: "Estate Planning", slides: "2 slides", desc: "Estate breakdown and trust transfer flow visual." },
                  ].map((option) => (
                    <label key={option.key} className={`slide-card ${selectedProposalModules[option.key] ? "slide-card--selected" : ""}`}>
                      <input
                        type="checkbox"
                        checked={!!selectedProposalModules[option.key]}
                        onChange={(e) => {
                          setSelectedProposalModules((prev) => ({ ...prev, [option.key]: e.target.checked }));
                          if (option.key === "riskManagementOverview") setSelectedServices((prev) => ({ ...prev, riskOverview: e.target.checked }));
                          if (option.key === "retirementPlanning") setSelectedServices((prev) => ({ ...prev, retirementPlanning: e.target.checked }));
                          if (option.key === "legacyWealthTransfer") setSelectedServices((prev) => ({ ...prev, legacyTransfer: e.target.checked }));
                          if (option.key === "estatePlanningReview") setSelectedServices((prev) => ({ ...prev, estatePlanning: e.target.checked }));
                        }}
                      />
                      <div className="slide-card-body">
                        <div className="slide-card-top">
                          <span className="slide-card-label">{option.label}</span>
                          <span className="slide-card-badge">{option.slides}</span>
                        </div>
                        <span className="slide-card-desc">{option.desc}</span>
                      </div>
                    </label>
                  ))}
                </div>

                <div className="module-section-title" style={{ marginTop: "1.25rem" }}>Analytics &amp; Projection Slides</div>
                <p className="module-section-hint">Require a historical backtest (a concentrated position or uploaded holdings) to render.</p>
                <div className="slide-selection-grid">
                  {[
                    { key: "feeDragAnalysis", label: "Fee Drag Analysis", slides: "1 slide", desc: "20-year cumulative cost of fees: current vs. proposed weighted fee." },
                    { key: "monteCarloProjection", label: "Monte Carlo Projection", slides: "1 slide", desc: "Range of 20-year outcomes with a percentile fan and goal probability." },
                    { key: "stressTestAnalysis", label: "Stress Test", slides: "1 slide", desc: "Returns through the 2008, 2020, and 2022 market drawdowns." },
                  ].map((option) => (
                    <label key={option.key} className={`slide-card ${selectedProposalModules[option.key] ? "slide-card--selected" : ""}`}>
                      <input
                        type="checkbox"
                        checked={!!selectedProposalModules[option.key]}
                        onChange={(e) => setSelectedProposalModules((prev) => ({ ...prev, [option.key]: e.target.checked }))}
                      />
                      <div className="slide-card-body">
                        <div className="slide-card-top">
                          <span className="slide-card-label">{option.label}</span>
                          <span className="slide-card-badge">{option.slides}</span>
                        </div>
                        <span className="slide-card-desc">{option.desc}</span>
                      </div>
                    </label>
                  ))}
                </div>
              </div>

              <div className="strategy-review-box strategy-review-box-clean">
                <h4>For Clients with High Concentrated Stock</h4>
                <p>Select the concentrated-stock strategies to include in the final proposal.</p>

                <div className="slide-selection-grid">
                  {[
                    {
                      key: "crt",
                      label: "Charitable Remainder Trust",
                      slides: "2 slides",
                      desc: "CRT flow diagram and sell-now vs. CRT comparison table.",
                    },
                    {
                      key: "harvesting",
                      label: "Leveraged Tax-Loss Harvesting",
                      slides: "2 slides",
                      desc: "130/30 exposure chart and annual tax savings breakdown.",
                    },
                    {
                      key: "collar",
                      label: "Options Collar",
                      slides: "1 slide",
                      desc: "Payoff chart with put floor, call cap, and protected range.",
                    },
                    {
                      key: "diversification",
                      label: "Phased Diversification Schedule",
                      slides: "1 slide",
                      desc: "Year-by-year exit plan with capital-gains budget and declining concentration.",
                    },
                  ].map((option) => (
                    <label key={option.key} className={`slide-card ${selectedStrategies[option.key] ? "slide-card--selected" : ""}`}>
                      <input
                        type="checkbox"
                        checked={!!selectedStrategies[option.key]}
                        onChange={(e) => {
                          setSelectedStrategies((prev) => ({ ...prev, [option.key]: e.target.checked }));
                          if (option.key === "collar" && e.target.checked && reviewData?.ticker && isUsableTicker(reviewData.ticker)) {
                            fetchCollarOptions(reviewData.ticker);
                          }
                        }}
                      />
                      <div className="slide-card-body">
                        <div className="slide-card-top">
                          <span className="slide-card-label">{option.label}</span>
                          <span className="slide-card-badge">{option.slides}</span>
                        </div>
                        <span className="slide-card-desc">{option.desc}</span>
                      </div>
                    </label>
                  ))}
                </div>

                {selectedStrategies?.collar && (
                  <div className="collar-options-status">
                    {collarOptionsLoading && <span className="collar-status-loading">⏳ Fetching live options data…</span>}
                    {collarOptionsError && <span className="collar-status-error">⚠ {collarOptionsError} — estimated strikes will be used.</span>}
                    {collarOptions && !collarOptionsLoading && (
                      <span className="collar-status-live">
                        ✓ Live options loaded ({collarOptions.expirationLabel}) —{" "}
                        Put ${collarOptions.put.strike?.toFixed(2)} · Call ${collarOptions.call.strike?.toFixed(2)} · {collarOptions.netCostPerShare}
                        <button className="collar-refresh-btn" onClick={() => fetchCollarOptions(reviewData?.ticker)}>↻ Refresh</button>
                      </span>
                    )}
                  </div>
                )}
              </div>

              {reviewData && (() => {
                const recommendation = recommendInvestmentApproach(reviewData, notes);

                return (
                  <div className="recommendation-card">
                    <div className="recommendation-header">
                      <h4>Recommended Investment Approach</h4>
                      <button type="button" className="secondary-button" onClick={applyInvestmentApproachRecommendation}>
                        Apply Recommendation
                      </button>
                    </div>

                    <div className="recommendation-grid">
                      <div>
                        <span className="recommendation-label">Suggested Strategy Model</span>
                        <strong>{recommendation.portfolioLabel}</strong>
                        <p>{recommendation.modelReason}</p>
                      </div>

                      <div>
                        <span className="recommendation-label">Suggested Risk Profile</span>
                        <strong>{recommendation.riskLabel}</strong>
                        <p>{recommendation.riskReason}</p>
                      </div>
                    </div>
                  </div>
                );
              })()}

              <div className="strategy-review-box strategy-review-box-clean">
                <label className="use-recommended-checkbox">
                  <input
                    type="checkbox"
                    checked={useRecommendedApproach}
                    onChange={(e) => {
                      const checked = e.target.checked;
                      setUseRecommendedApproach(checked);
                      if (checked) {
                        applyInvestmentApproachRecommendationForData(reviewData);
                      } else {
                        setSelectedPortfolioStrategies({ corePrivate: false, selectLiquidity: false, traditional: false, focusedB: false, selectLiquidityUsBias: false, traditionalUsBias: false });
                        setSelectedRiskProfile("");
                      }
                    }}
                  />
                </label>
                <div className="dropdown-pair">
                  <div className="dropdown-field">
                    <label className="dropdown-label">Portfolio Strategy</label>
                    <select
                      className="dropdown-select"
                      value={Object.keys(selectedPortfolioStrategies).find(k => selectedPortfolioStrategies[k]) || ""}
                      onChange={(e) => {
                        const key = e.target.value;
                        setSelectedPortfolioStrategies({
                          corePrivate: key === "corePrivate",
                          selectLiquidity: key === "selectLiquidity",
                          traditional: key === "traditional",
                          focusedB: key === "focusedB",
                          selectLiquidityUsBias: key === "selectLiquidityUsBias",
                          traditionalUsBias: key === "traditionalUsBias",
                        });
                      }}
                    >
                      <option value="">— Select strategy —</option>
                      <option value="corePrivate">Core Private</option>
                      <option value="selectLiquidity">Select Liquidity</option>
                      <option value="selectLiquidityUsBias">Select Liquidity (US Bias)</option>
                      <option value="traditional">Traditional</option>
                      <option value="traditionalUsBias">Traditional (US Bias)</option>
                      <option value="focusedB">Focused B</option>
                    </select>
                  </div>

                  <div className="dropdown-field">
                    <label className="dropdown-label">Risk Profile / Allocation</label>
                    <select
                      className="dropdown-select"
                      value={selectedRiskProfile || ""}
                      onChange={(e) => setSelectedRiskProfile(e.target.value || "")}
                    >
                      <option value="">— Select risk profile —</option>
                      {riskProfileOptions.map((opt) => (
                        <option key={opt.key} value={opt.key}>
                          {opt.label} — {opt.mix}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              <div className="review-actions">
                <button onClick={confirmAndGenerate}>Confirm and Generate</button>
                <button className="secondary-button" onClick={cancelReview}>Cancel Review</button>
              </div>
            </div>
          )}

          {proposal && (
            <div>
              <button onClick={downloadWord}>Download Word Document</button>
              <button onClick={downloadPowerPoint}>Download PowerPoint</button>
            </div>
          )}
        </div>

</main>

      {previewModal && <ProposalPreviewModal
        data={previewModal.finalData}
        name={previewModal.name}
        selectedStrategies={selectedStrategies}
        portfolioModel={previewModal.portfolioKey}
        riskProfile={previewModal.riskProfileKey}
        portfolioLabel={previewModal.portfolioModelLabel}
        riskLabel={previewModal.riskProfileLabel}
        altFallback={previewModal.altFallback}
        backtestResult={backtestResult}
        backtestLoading={backtestLoading}
        backtestError={backtestError}
        onConfirm={proceedFromPreview}
        onBack={closePreviewModal}
      />}
    </div>
  );
}
