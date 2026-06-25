// Pure math for the proposal's two hero metrics: TAX SAVED and DOWNSIDE
// PROTECTED. These quantify the value of the selected strategies in the two
// terms clients care about most, and feed both the "Tax You're Leaving on the
// Table" hero slide and the per-strategy scorecard tags.
//
// Everything here is illustrative and intentionally conservative; consumers must
// show a "not tax advice / past performance" disclaimer. All dollar inputs and
// outputs are in the same units the rest of the app uses ($M).
//
// Definitions (documented so they're defensible in front of a client):
//  • Long-term capital-gains rate = federal (incl. NIIT) + state.
//  • A "severe drawdown" is modeled at 40% for single-stock crash risk.
//  • CRT: donating appreciated shares avoids the capital-gains tax on the
//    donated amount, and removes that amount from single-stock crash risk.
//  • Tax-loss harvesting: first-year tax savings already computed upstream.
//  • Collar: caps losses below the put floor, so it protects the portion of a
//    severe drawdown that falls beyond the floor.
//  • Diversification: the remaining single-stock capital (beyond CRT/collar)
//    is moved out of crash risk over time. It DEFERS rather than saves tax, so
//    it contributes to downside protected but not to tax saved.

const SEVERE_DRAWDOWN = 0.40;

/**
 * @param {object} p
 * @param {object} p.data               Cleaned client data (stockPosition, costBasisPct, …).
 * @param {object} p.selectedStrategies Map of { crt, harvesting, collar, diversification }.
 * @returns {{taxSaved, downsideProtected, immediateTax, ltcgRatePct, perStrategy}}
 */
export function computeImpactScorecard({ data = {}, selectedStrategies = {} } = {}) {
  const n = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);

  const stockPosition = Math.max(n(data.stockPosition), 0);
  const costBasisPct = Math.min(Math.max(n(data.costBasisPct), 0), 100);
  const gainFraction = 1 - costBasisPct / 100;
  const ltcgRatePct = (n(data.federalTaxRate) || 23.8) + n(data.stateTaxRate);
  const ltcg = ltcgRatePct / 100;

  const crtAllocation = Math.max(n(data.crtAllocation), 0);
  const collarAllocation = Math.max(n(data.collarAllocation), 0);
  const stockPrice = Math.max(n(data.stockPrice), 0);
  const putStrike = Math.max(n(data.putStrike), 0);
  const floorDropPct = stockPrice > 0 ? Math.max(0, 1 - putStrike / stockPrice) : 0.15;

  // Tax if the whole position were sold outright today (the "naive sale" baseline).
  const immediateTax =
    n(data.immediateTax) || stockPosition * gainFraction * ltcg;

  const perStrategy = {
    crt: { taxSaved: 0, downsideProtected: 0 },
    harvesting: { taxSaved: 0, downsideProtected: 0 },
    collar: { taxSaved: 0, downsideProtected: 0 },
    diversification: { taxSaved: 0, downsideProtected: 0 },
  };

  if (selectedStrategies.crt) {
    perStrategy.crt.taxSaved = crtAllocation * gainFraction * ltcg;
    perStrategy.crt.downsideProtected = crtAllocation * SEVERE_DRAWDOWN;
  }
  if (selectedStrategies.harvesting) {
    // First-year harvesting tax savings, computed upstream.
    perStrategy.harvesting.taxSaved = Math.max(n(data.taxSavings), 0);
  }
  if (selectedStrategies.collar) {
    // The collar caps the part of a severe drawdown that falls below the floor.
    perStrategy.collar.downsideProtected =
      collarAllocation * Math.max(0, SEVERE_DRAWDOWN - floorDropPct);
  }
  if (selectedStrategies.diversification) {
    // Remaining single-stock capital (beyond what CRT/collar already cover)
    // that gets diversified out of crash risk. Clamped so we never claim to
    // protect more than the position is worth.
    const alreadyCovered = (selectedStrategies.crt ? crtAllocation : 0) +
      (selectedStrategies.collar ? collarAllocation : 0);
    const remaining = Math.max(0, stockPosition - alreadyCovered);
    perStrategy.diversification.downsideProtected = remaining * SEVERE_DRAWDOWN;
  }

  const taxSaved = Object.values(perStrategy).reduce((s, x) => s + x.taxSaved, 0);
  // Total downside protected can't exceed the whole position's severe-drawdown loss.
  const downsideProtected = Math.min(
    Object.values(perStrategy).reduce((s, x) => s + x.downsideProtected, 0),
    stockPosition * SEVERE_DRAWDOWN
  );

  return { taxSaved, downsideProtected, immediateTax, ltcgRatePct, perStrategy };
}
