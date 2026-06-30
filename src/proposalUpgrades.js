export function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
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

export function recomputeReviewedData(input) {
  const data = { ...input };

  data.netWorth = toNumber(data.netWorth);
  data.investableAssets = toNumber(data.investableAssets);
  data.stockPosition = toNumber(data.stockPosition);
  data.income = toNumber(data.income);
  data.costBasisPct = toNumber(data.costBasisPct, 15);
  data.taxRate = toNumber(data.taxRate || data.totalTaxRate, 37.1);
  data.totalTaxRate = data.taxRate;
  data.federalTaxRate = toNumber(data.federalTaxRate, 23.8);
  data.stateTaxRate = toNumber(data.stateTaxRate, Math.max(data.taxRate - data.federalTaxRate, 0));
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

  // Use stored pct if available; default to 15 (range is 10–20% based on charitable intent)
  data.crtPct        = toNumber(data.crtPct,        15);
  data.harvestingPct = toNumber(data.harvestingPct, 30);
  data.collarPct     = toNumber(data.collarPct,     30);

  // Derive dollar amounts; CRT has $1M minimum (values stored in millions)
  const rawCrt = data.stockPosition * (data.crtPct / 100);
  data.crtAllocation    = data.stockPosition > 0 ? Math.max(rawCrt, 1) : 0;
  // Sync pct back so displayed % always matches actual $ (handles $1M floor)
  if (data.stockPosition > 0) data.crtPct = (data.crtAllocation / data.stockPosition) * 100;

  data.harvestingSleeve = data.stockPosition * (data.harvestingPct / 100);
  data.collarAllocation = data.stockPosition * (data.collarPct     / 100);

  data.costBasis = data.stockPosition * (data.costBasisPct / 100);
  data.embeddedGain = Math.max(data.stockPosition - data.costBasis, 0);
  data.immediateTax = data.embeddedGain * (data.taxRate / 100);

  // IRC §664: CRT payout rate must be 5%–50%. Default 5%, but preserve user edits.
  data.crtPayoutRate = Math.min(50, Math.max(5, toNumber(data.crtPayoutRate, 5)));
  data.clientAge = toNumber(data.clientAge, 65);
  data.crtIncome = data.crtAllocation * (data.crtPayoutRate / 100);

  // CRT deduction: IRS Table S factor, single life, varies by client age.
  // Table is built at a 5% payout rate; deduction is illustrative if the payout
  // rate is adjusted away from 5% (a true Table S lookup would also vary by rate).
  const crtDeduction = data.crtAllocation * crtDeductionFactor(data.clientAge);
  data.charitableDeductionLow  = crtDeduction;
  data.charitableDeductionHigh = crtDeduction;

  // 130/30: ~25% annual harvest rate
  data.annualHarvestLosses = data.harvestingSleeve * 0.25;
  data.federalTaxSavings = data.annualHarvestLosses * (data.federalTaxRate / 100);
  data.stateTaxSavings = data.annualHarvestLosses * (data.stateTaxRate / 100);
  data.taxSavings = data.federalTaxSavings + data.stateTaxSavings;

  data.putStrike = data.stockPrice * 0.85;
  data.callStrike = data.stockPrice * 1.19;
  data.putFloorValue = data.collarAllocation * 0.85;
  data.callCapValue = data.collarAllocation * 1.19;

  // After-CRT: both stock position AND investable assets shrink by CRT amount
  const postCrtStock  = data.stockPosition - data.crtAllocation;
  const postCrtAssets = data.investableAssets - data.crtAllocation;
  data.afterCrtConcentration =
    postCrtAssets > 0
      ? Math.min(Math.max((postCrtStock / postCrtAssets) * 100, 0), 100)
      : 0;

  data.drawdown40Impact = data.stockPosition * 0.4;

  return data;
}

export function getProposalQuality(notes, data) {

  const checks = [
    {
      label: "Ticker/company provided",
      passed: data.ticker && data.ticker !== "Concentrated Stock",
      importance: 12,
    },
    {
      label: "Concentrated position value provided",
      passed:
        Number(data.stockPosition || 0) > 0 &&
        (
          /\b(position|stock position|concentrated position|concentrated stock|holding|holdings|shares|equity position|single stock|ticker)\b/i.test(notes) ||
          /\b(NVDA|TSLA|AAPL|MSFT|AMZN|GOOGL|META|CRM|AVGO|AMD|NFLX|ORCL|ADBE|INTC|SHOP|SNOW|PLTR)\b/i.test(notes) ||
          /\b(nvidia|tesla|apple|microsoft|amazon|google|alphabet|meta)\b/i.test(notes)
        ),
      importance: 16,
    },
    {
      label: "Investable assets / portfolio value provided",
      passed:
        data.investableAssets > 0 &&
        /\b(investable assets|portfolio value|aum|managed assets|assets under management|portfolio)\b/i.test(notes),
      importance: 16,
    },
    {
      label: "Net worth provided",
      passed: data.netWorth > 0 && /\b(net worth|total net worth|balance sheet|worth approximately)\b/i.test(notes),
      importance: 10,
    },
    {
      label: "Client goals provided",
      passed:
        /\b(goal|goals|objective|objectives|wants|needs|priority|retirement|income|estate|tax|charitable|liquidity|legacy|diversify)\b/i.test(notes),
      importance: 14,
    },
    {
      label: "Risk tolerance provided",
      passed: /\b(risk tolerance|risk profile|conservative|moderate|balanced|growth|aggressive)\b/i.test(notes),
      importance: 8,
    },
    {
      label: "Time horizon provided",
      passed: /\b(time horizon|horizon|long-term|short-term|years|retirement in|within)\b/i.test(notes),
      importance: 8,
    },
    {
      label: "Tax rate or state tax details provided",
      passed: /\b(tax rate|combined tax|federal tax|state tax|capital gains tax|california tax)\b/i.test(notes),
      importance: 8,
    },
    {
      label: "Cost basis provided",
      passed: /\b(cost basis|basis|low basis)\b/i.test(notes),
      importance: 5,
    },
    {
      label: "Restrictions/preferences provided",
      passed: /\b(restriction|restrictions|exclude|avoid|do not buy|esg|no single position|preference|preferences)\b/i.test(notes),
      importance: 3,
    },
  ];

  const maxScore = checks.reduce((sum, item) => sum + item.importance, 0);
  const earnedScore = checks
    .filter((item) => item.passed)
    .reduce((sum, item) => sum + item.importance, 0);

  const score = Math.round((earnedScore / maxScore) * 100);

  return {
    score,
    strengths: checks.filter((item) => item.passed).map((item) => item.label),
    gaps: checks.filter((item) => !item.passed).map((item) => item.label),
  };
}
