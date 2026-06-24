// Pure math for a phased concentrated-position exit ("diversification schedule").
//
// A client holding a large single-stock position rarely wants to sell it all at
// once — that triggers the entire embedded capital gain in one tax year and can
// push them into the top bracket. The standard advisor solution is a multi-year
// glide path: sell a tranche each year, ideally within an annual capital-gains
// budget, and reinvest the after-tax proceeds into a diversified portfolio until
// the position is no longer a concentration risk.
//
// This module models that plan honestly:
//   * Cost basis is tracked in DOLLARS and reduced proportionally as shares are
//     sold, so the realized gain grows correctly as the stock appreciates (a
//     fixed "basis %" would understate the gain over time).
//   * Both the remaining position AND the diversified sleeve grow at a market
//     rate, so the concentration glide path reflects real portfolio dynamics.
//   * Tax is applied at the long-term capital-gains rate on the realized gain.
//
// Everything here is illustrative — consumers must show a "not tax advice /
// past performance" disclaimer alongside any figure produced here.

/**
 * Build a year-by-year concentrated-position exit schedule.
 *
 * @param {object}  p
 * @param {number}  p.stockPosition        Current $ value of the concentrated position.
 * @param {number}  p.costBasisPct         Cost basis as a % of the CURRENT value (0–100).
 * @param {number}  p.investableAssets     Total investable assets (position + everything else).
 * @param {number} [p.ltcgRate=23.8]       Effective long-term cap-gains rate, % (fed + state + NIIT).
 * @param {number} [p.growthRate=7.5]      Assumed annual market growth, %.
 * @param {number} [p.annualReductionPct=20] Target % of the ORIGINAL position to sell each year.
 * @param {number} [p.annualGainsBudget=null] Optional $ cap on realized gains per year (null = none).
 * @param {number} [p.targetConcentrationPct=10] Stop once concentration falls to/below this %.
 * @param {number} [p.maxYears=12]         Hard cap on the number of years modeled.
 * @returns {{rows: Array, summary: object, assumptions: object}}
 */
export function buildExitSchedule({
  stockPosition,
  costBasisPct,
  investableAssets,
  ltcgRate = 23.8,
  growthRate = 7.5,
  annualReductionPct = 20,
  annualGainsBudget = null,
  targetConcentrationPct = 10,
  maxYears = 12,
} = {}) {
  const pos0 = Number(stockPosition) || 0;
  const invest0 = Number(investableAssets) || 0;
  const basisPct = Math.min(Math.max(Number(costBasisPct) || 0, 0), 100);
  const g = (Number(growthRate) || 0) / 100;
  const taxRate = Math.max(Number(ltcgRate) || 0, 0) / 100;
  const reductionPct = Math.max(Number(annualReductionPct) || 0, 0);
  const budget = annualGainsBudget == null ? null : Math.max(Number(annualGainsBudget) || 0, 0);
  const target = Math.max(Number(targetConcentrationPct) || 0, 0);
  const cap = Math.max(Math.floor(Number(maxYears) || 0), 0);

  const assumptions = {
    ltcgRate, growthRate, annualReductionPct, annualGainsBudget,
    targetConcentrationPct, maxYears, costBasisPct: basisPct,
  };

  const startingConcentrationPct =
    pos0 > 0 && invest0 > 0 ? Math.min((pos0 / invest0) * 100, 100) : pos0 > 0 ? 100 : 0;

  // Selling the whole position today: the full embedded gain is taxed at once.
  const immediateGain = pos0 * (1 - basisPct / 100);
  const immediateSaleTax = immediateGain * taxRate;

  const emptySummary = {
    startingConcentrationPct,
    endingConcentrationPct: startingConcentrationPct,
    yearsModeled: 0,
    totalTax: 0,
    totalSold: 0,
    totalReinvested: 0,
    immediateSaleTax,
    taxDeferred: 0,
    alreadyDiversified: startingConcentrationPct <= target,
  };

  // Nothing to do: no position, or already at/below the target concentration.
  if (pos0 <= 0 || reductionPct <= 0 || startingConcentrationPct <= target) {
    return { rows: [], summary: emptySummary, assumptions };
  }

  let posValue = pos0;
  let basisDollars = pos0 * (basisPct / 100);
  let diversifiedValue = Math.max(invest0 - pos0, 0);

  const plannedAnnualSale = pos0 * (reductionPct / 100);

  const rows = [];
  let cumulativeTax = 0;
  let cumulativeSold = 0;
  let cumulativeReinvested = 0;

  for (let year = 1; year <= cap; year++) {
    // 1. Market growth (basis dollars don't change with appreciation).
    posValue *= 1 + g;
    diversifiedValue *= 1 + g;
    const beginningPosition = posValue;

    // 2. Plan this year's sale, capped by what's left and by the gains budget.
    let sale = Math.min(plannedAnnualSale, posValue);
    if (budget != null && posValue > 0) {
      const gainFraction = 1 - basisDollars / posValue; // share of a sale that is gain
      const maxSaleByBudget = gainFraction > 0 ? budget / gainFraction : sale;
      sale = Math.min(sale, maxSaleByBudget);
    }
    sale = Math.max(sale, 0);

    // 3. Realize gain on the sold fraction; pay tax; reinvest the rest.
    const fractionSold = posValue > 0 ? sale / posValue : 0;
    const basisSold = basisDollars * fractionSold;
    const realizedGain = sale - basisSold;
    const tax = realizedGain * taxRate;
    const netProceeds = sale - tax;

    posValue -= sale;
    basisDollars -= basisSold;
    diversifiedValue += netProceeds;

    cumulativeTax += tax;
    cumulativeSold += sale;
    cumulativeReinvested += netProceeds;

    const totalValue = posValue + diversifiedValue;
    const concentrationPct = totalValue > 0 ? (posValue / totalValue) * 100 : 0;

    rows.push({
      year,
      beginningPosition,
      sold: sale,
      realizedGain,
      tax,
      netProceeds,
      remainingPosition: posValue,
      diversifiedValue,
      totalValue,
      concentrationPct,
      cumulativeTax,
      cumulativeSold,
    });

    // Stop once we've reached the target concentration or sold out.
    if (concentrationPct <= target || posValue <= 0.0001) break;
  }

  const last = rows[rows.length - 1];
  const summary = {
    startingConcentrationPct,
    endingConcentrationPct: last ? last.concentrationPct : startingConcentrationPct,
    yearsModeled: rows.length,
    totalTax: cumulativeTax,
    totalSold: cumulativeSold,
    totalReinvested: cumulativeReinvested,
    immediateSaleTax,
    // Honest framing: the phased plan doesn't necessarily pay LESS total tax —
    // its value is spreading the liability over time (deferral) while cutting
    // single-stock risk. taxDeferred = what an all-at-once sale would owe today
    // beyond what the first year of the plan realizes.
    taxDeferred: Math.max(immediateSaleTax - (rows[0]?.tax || 0), 0),
    alreadyDiversified: false,
  };

  return { rows, summary, assumptions };
}
