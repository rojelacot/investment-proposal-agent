export const requiredFields = [
  {
    key: "clientName",
    label: "Client / household name",
    question: "What is the client's full name or household name?"
  },
  {
    key: "clientAge",
    label: "Client age",
    question: "What is the client's age? (Used to calculate the CRT charitable deduction based on IRS actuarial tables.)"
  },
  {
    key: "netWorth",
    label: "Net worth",
    question: "What is the client's approximate net worth?"
  },
  {
    key: "managedAssets",
    label: "Managed / investable assets",
    question: "How much in assets are being managed or proposed?"
  },
  {
    key: "currentHoldings",
    label: "Current holdings",
    question: "What does the current portfolio look like? Include stocks, bonds, cash, real estate, alternatives, or concentrated positions."
  },
  {
    key: "goals",
    label: "Client goals",
    question: "What are the client's main goals? For example: retirement, tax efficiency, income, estate planning, liquidity, or charitable giving."
  },
  {
    key: "riskTolerance",
    label: "Risk tolerance",
    question: "What is the client's risk tolerance: conservative, moderate, growth, aggressive, or something else?"
  },
  {
    key: "timeHorizon",
    label: "Time horizon",
    question: "What is the client's investment time horizon?"
  },
  {
    key: "taxSituation",
    label: "Tax situation",
    question: "Are there tax considerations, such as high income, unrealized gains, low cost basis stock, business sale, charitable goals, or tax-loss harvesting?"
  },
  {
    key: "restrictions",
    label: "Restrictions / preferences",
    question: "Are there any investment restrictions or preferences? For example: do not buy META, exclude private prisons, ESG preferences, liquidity needs, or no single position above 5%."
  },
  {
    key: "strategies",
    label: "Recommended strategies",
    question: "Which strategies should be considered? For example: CRT, tax-loss harvesting, option collar, diversification, muni bonds, or estate planning."
  },
  {
    key: "riskNumber",
    label: "Client risk score (Nitrogen / Riskalyze)",
    question: "Does the client have a Nitrogen or Riskalyze risk number (1–99)? If so, what is it? If not, leave blank."
  }
];

export function detectMissingInfo(notes) {
  const text = notes.toLowerCase();

  // Risk number
  const hasRiskNumber =
    /risk\s*number[^0-9]{0,20}\d{1,2}/i.test(notes) ||
    /risk\s*score[^0-9]{0,20}\d{1,2}/i.test(notes) ||
    /\bRQ\b[^0-9]{0,20}\d{1,2}/i.test(notes);

  // Age: must have an actual number, not just the word "age"
  const hasClientAge =
    /\bage[d]?\s*[:\-]?\s*\d{2}\b/i.test(notes) ||
    /\d{2}\s+years?\s+old/i.test(notes) ||
    /client\s+is\s+\d{2}\s+years?/i.test(notes) ||
    /\d{2}\s*[-–]\s*year[-\s]old/i.test(notes) ||
    /born\s+(?:in\s+)?\d{4}/i.test(notes) ||
    /\bDOB\b[:\s]+\d/i.test(notes) ||
    /date\s+of\s+birth[:\s]+\d/i.test(notes);

  // Cost basis: must have an actual percentage number
  const hasCostBasisPct =
    /cost\s+basis[^%\d]{0,40}\d{1,3}\s*%/i.test(notes) ||
    /\bbasis[^%\d]{0,40}\d{1,3}\s*%/i.test(notes) ||
    /\d{1,3}\s*%\s+(?:cost\s+)?basis/i.test(notes);

  // Tax rate: must have an actual percentage number
  const hasTotalTaxRate =
    /(?:combined|total|capital\s+gains?)\s+tax\s+rate[^%\d]{0,40}\d{1,3}\s*%/i.test(notes) ||
    /tax\s+rate[^%\d]{0,40}\d{1,3}\s*%/i.test(notes) ||
    /\d{1,3}\s*%\s+(?:combined\s+)?tax\s+rate/i.test(notes);

  // Stock price: must have an actual dollar amount that looks like a share price
  const hasStockPrice =
    /(?:stock|share|current)\s+price[^$\d]{0,30}\$?\s*\d{2,}/i.test(notes) ||
    /trades?\s+at[^$\d]{0,30}\$?\s*\d{2,}/i.test(notes) ||
    /trading\s+(?:at|price)[^$\d]{0,30}\$?\s*\d{2,}/i.test(notes) ||
    /reference\s+price[^$\d]{0,30}\$?\s*\d{2,}/i.test(notes);

  // Only ask for stock-specific fields if there's a concentrated position mentioned
  const hasConcentratedStock =
    /\b(stock|position|holding|holdings|shares|equity|ticker|nvda|tsla|aapl|msft|amzn|googl|meta)\b/i.test(notes) &&
    /\$?\s*[\d,.]+\s*(M|million|B|billion)/i.test(notes);

  const checks = {
    clientName:    /\b(client|household|family|mr\.|mrs\.|ms\.|dr\.|couple)\b/i.test(notes),
    clientAge:     hasClientAge,
    netWorth:      /\b(net worth|nw|worth|total assets|balance sheet)\b/i.test(notes),
    managedAssets: /\b(managed assets|aum|investable assets|portfolio value|assets under management|proposed assets)\b/i.test(notes),
    currentHoldings: /\b(holding|holdings|position|stock|equity|bond|cash|real estate|alternative|nvidia|nvda|tesla|portfolio)\b/i.test(notes),
    costBasisPct:  !hasConcentratedStock || hasCostBasisPct,  // only required if there's a concentrated stock
    totalTaxRate:  hasTotalTaxRate,
    stockPrice:    !hasConcentratedStock || hasStockPrice,     // only required if there's a concentrated stock
    goals:         /\b(goal|goals|objective|objectives|wants|needs|priority|retirement|income|estate|tax|charitable|liquidity|legacy)\b/i.test(notes),
    riskTolerance: /\b(risk tolerance|risk profile|conservative|moderate|balanced|growth|aggressive)\b/i.test(notes),
    timeHorizon:   /\b(time horizon|horizon|years|long-term|short-term|retirement in|within)\b/i.test(notes),
    taxSituation:  /\b(tax|capital gains|unrealized gain|basis|cost basis|income|charitable|crt|tax-loss|harvesting|state tax)\b/i.test(notes),
    restrictions:  /\b(restriction|restrictions|exclude|do not buy|avoid|esg|no single position|liquidity need|preference|preferences)\b/i.test(notes),
    strategies:    /\b(crt|charitable remainder trust|tax-loss harvesting|harvesting|option collar|collar|diversification|muni|municipal|estate planning|trust)\b/i.test(notes),
    riskNumber:    hasRiskNumber,
  };

  return requiredFields.filter(field => !checks[field.key]);
}
