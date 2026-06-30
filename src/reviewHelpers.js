export function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

export function recomputeDerivedData(input) {
  const data = { ...input };

  data.netWorth = toNumber(data.netWorth);
  data.investableAssets = toNumber(data.investableAssets);
  data.stockPosition = toNumber(data.stockPosition);
  data.income = toNumber(data.income);
  data.totalTaxRate = toNumber(data.totalTaxRate || data.taxRate, 37.1);
  data.taxRate = data.totalTaxRate;
  data.federalTaxRate = toNumber(data.federalTaxRate, 23.8);
  data.stateTaxRate = toNumber(data.stateTaxRate, Math.max(data.totalTaxRate - data.federalTaxRate, 0));
  data.costBasisPct = toNumber(data.costBasisPct, 15);
  data.stockPrice = toNumber(data.stockPrice, 200);

  if (!data.investableAssets && data.stockPosition) {
    data.investableAssets = data.stockPosition / 0.65;
  }

  if (!data.netWorth && data.investableAssets) {
    data.netWorth = data.investableAssets * 1.25;
  }

  if (data.netWorth && data.investableAssets && data.netWorth < data.investableAssets) {
    data.netWorth = data.investableAssets * 1.25;
  }

  data.concentration =
    data.investableAssets > 0
      ? Math.min(Math.max((data.stockPosition / data.investableAssets) * 100, 0), 100)
      : 0;

  data.crtAllocation = data.stockPosition * 0.32;
  data.harvestingSleeve = data.stockPosition * 0.29;
  data.collarAllocation = data.stockPosition * 0.26;

  data.costBasis = data.stockPosition * (data.costBasisPct / 100);
  data.embeddedGain = Math.max(data.stockPosition - data.costBasis, 0);
  data.immediateTax = data.embeddedGain * (data.totalTaxRate / 100);

  data.crtPayoutRate = 5;
  data.crtIncome = data.crtAllocation * (data.crtPayoutRate / 100);
  data.charitableDeductionLow = data.crtAllocation * 0.30;
  data.charitableDeductionHigh = data.crtAllocation * 0.37;

  data.annualHarvestLosses = data.harvestingSleeve * 0.20;
  data.federalTaxSavings = data.annualHarvestLosses * (data.federalTaxRate / 100);
  data.stateTaxSavings = data.annualHarvestLosses * (data.stateTaxRate / 100);
  data.taxSavings = data.federalTaxSavings + data.stateTaxSavings;

  data.putStrike = data.stockPrice * 0.85;
  data.callStrike = data.stockPrice * 1.19;
  data.putFloorValue = data.collarAllocation * 0.85;
  data.callCapValue = data.collarAllocation * 1.19;

  data.afterCrtConcentration =
    data.investableAssets > 0
      ? Math.min(Math.max(((data.stockPosition - data.crtAllocation) / data.investableAssets) * 100, 0), 100)
      : Math.min(data.concentration * 0.68, 100);

  data.drawdown40Impact = data.stockPosition * 0.40;

  return data;
}

export function analyzeClientReview(notes, data) {

  const checks = {
    ticker:
      data.ticker &&
      data.ticker !== "Concentrated Stock" &&
      /\b(nvda|nvidia|aapl|apple|msft|microsoft|tsla|tesla|amzn|amazon|googl|google|meta|crm|avgo|amd|nflx|orcl|adbe|intc|shop|snow|pltr)\b/i.test(notes),

    netWorth: /\b(net worth|total net worth|balance sheet|worth approximately)\b/i.test(notes),

    investableAssets:
      /\b(investable assets|aum|portfolio value|managed assets|assets under management|proposed assets)\b/i.test(notes),

    stockPosition:
      /\b(concentrated position|stock position|position|shares|holding|holdings|nvda|nvidia|aapl|apple|msft|tsla|tesla)\b/i.test(notes) &&
      /\$?\s*[\d.]+\s*(m|million|b|billion)/i.test(notes),

    costBasis: /\b(cost basis|basis|low basis)\b/i.test(notes),

    taxRate: /\b(tax rate|combined tax|federal tax|state tax|capital gains tax|california tax)\b/i.test(notes),

    goals:
      /\b(goal|goals|objective|objectives|wants|needs|priority|retirement|income|estate|tax|charitable|liquidity|legacy|diversify|reduce concentration)\b/i.test(notes),

    riskTolerance:
      /\b(risk tolerance|risk profile|conservative|moderate|balanced|growth|aggressive)\b/i.test(notes),

    timeHorizon:
      /\b(time horizon|horizon|long-term|short-term|years|retirement in|within)\b/i.test(notes),

    restrictions:
      /\b(restriction|restrictions|exclude|avoid|do not buy|esg|no single position|liquidity need|preference|preferences)\b/i.test(notes),
  };

  const fields = [
    {
      key: "ticker",
      label: "Ticker / Company",
      type: "text",
      confidence: checks.ticker ? "High" : "Assumed",
      note: checks.ticker ? "Detected from notes." : "Not clearly provided; review this carefully.",
    },
    {
      key: "netWorth",
      label: "Net Worth ($M)",
      type: "number",
      confidence: checks.netWorth ? "High" : "Assumed",
      note: checks.netWorth ? "Labeled in notes." : "Estimated or inferred.",
    },
    {
      key: "investableAssets",
      label: "Investable Assets ($M)",
      type: "number",
      confidence: checks.investableAssets ? "High" : "Needs Review",
      note: checks.investableAssets ? "Labeled in notes." : "Important field. Confirm before generating.",
    },
    {
      key: "stockPosition",
      label: "Concentrated Position ($M)",
      type: "number",
      confidence: checks.stockPosition ? "High" : "Needs Review",
      note: checks.stockPosition ? "Detected from notes." : "Important field. Confirm before generating.",
    },
    {
      key: "costBasisPct",
      label: "Cost Basis (%)",
      type: "number",
      confidence: checks.costBasis ? "High" : "Assumed",
      note: checks.costBasis ? "Detected from notes." : "Default assumption used.",
    },
    {
      key: "totalTaxRate",
      label: "Combined Tax Rate (%)",
      type: "number",
      confidence: checks.taxRate ? "High" : "Assumed",
      note: checks.taxRate ? "Detected from notes." : "Default tax assumption used.",
    },
    {
      key: "stockPrice",
      label: "Stock Price ($)",
      type: "number",
      confidence: /\b(share price|stock price|trades at|trading price)\b/i.test(notes) ? "High" : "Assumed",
      note: /\b(share price|stock price|trades at|trading price)\b/i.test(notes)
        ? "Detected from notes."
        : "Default price used if not provided.",
    },
    {
      key: "income",
      label: "Annual Income ($M)",
      type: "number",
      confidence: /\b(income|annual income)\b/i.test(notes) ? "High" : "Optional",
      note: /\b(income|annual income)\b/i.test(notes) ? "Detected from notes." : "Optional; can be left as 0.",
    },
  ];

  const blockingQuestions = [];

  if (!checks.ticker) {
    blockingQuestions.push({
      key: "ticker",
      label: "Ticker / Company",
      question: "What is the concentrated stock ticker or company name?",
    });
  }

  if (!checks.stockPosition) {
    blockingQuestions.push({
      key: "stockPosition",
      label: "Concentrated Position",
      question: "What is the approximate value of the concentrated stock position?",
    });
  }

  if (!checks.investableAssets) {
    blockingQuestions.push({
      key: "investableAssets",
      label: "Investable Assets",
      question: "What are the client's approximate investable assets or portfolio value?",
    });
  }

  if (!checks.goals) {
    blockingQuestions.push({
      key: "goals",
      label: "Client Goals",
      question: "What are the client’s main goals: diversification, tax efficiency, income, estate planning, charitable giving, liquidity, or something else?",
    });
  }

  const assumptions = [];

  if (!checks.costBasis) assumptions.push("Cost basis was not clearly provided, so the app used a default cost basis assumption.");
  if (!checks.taxRate) assumptions.push("Tax rate was not clearly provided, so the app used a default combined tax-rate assumption.");
  if (!checks.riskTolerance) assumptions.push("Risk tolerance was not clearly provided and should be confirmed during review.");
  if (!checks.timeHorizon) assumptions.push("Time horizon was not clearly provided and should be confirmed before implementation.");
  if (!checks.restrictions) assumptions.push("Investment restrictions/preferences were not clearly provided.");
  if (!/\b(share price|stock price|trades at|trading price)\b/i.test(notes)) assumptions.push("Stock price was not clearly provided, so a default share-price assumption may have been used.");

  return {
    fields,
    assumptions,
    blockingQuestions,
  };
}
