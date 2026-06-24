// Pure math for the long-horizon cost of investment fees ("fee drag").
//
// A 1%-vs-0.4% fee difference looks tiny on a fact sheet but compounds into a
// large dollar figure over a multi-decade relationship, because every dollar
// paid in fees is also a dollar that never compounds. This module projects the
// cumulative dollar cost of fees and the gap between a current and a proposed
// fee schedule, so a proposal can show the real lifetime impact.
//
// Convention: each year the balance grows at the gross return, then the annual
// fee is charged on the grown balance. Illustrative only — actual fees, returns,
// and timing vary; consumers must show a disclaimer.

/**
 * Project fee drag for a single fee schedule.
 *
 * @param {object} p
 * @param {number}  p.portfolioValue   Starting portfolio value.
 * @param {number}  p.annualFeePct     All-in annual fee, % (e.g. 0.85).
 * @param {number} [p.years=20]        Projection horizon.
 * @param {number} [p.grossReturnPct=7] Assumed gross annual return, %.
 * @returns {{rows: Array, endingBalance: number, totalFees: number}}
 */
export function projectFeeDrag({
  portfolioValue,
  annualFeePct,
  years = 20,
  grossReturnPct = 7,
} = {}) {
  const start = Math.max(Number(portfolioValue) || 0, 0);
  const fee = Math.max(Number(annualFeePct) || 0, 0) / 100;
  const gross = (Number(grossReturnPct) || 0) / 100;
  const n = Math.max(Math.floor(Number(years) || 0), 0);

  const rows = [];
  let balance = start;
  let totalFees = 0;

  for (let year = 1; year <= n; year++) {
    const grown = balance * (1 + gross);
    const feePaid = grown * fee;
    balance = grown - feePaid;
    totalFees += feePaid;
    rows.push({ year, balance, feePaid, cumulativeFees: totalFees });
  }

  return { rows, endingBalance: rows.length ? balance : start, totalFees };
}

/**
 * Compare a current vs. proposed fee schedule on the same portfolio, returning
 * both projections plus the cumulative fee savings and the ending-balance
 * difference attributable to the lower fee (the compounding benefit).
 *
 * @param {object} p
 * @param {number}  p.portfolioValue
 * @param {number}  p.currentFeePct
 * @param {number}  p.proposedFeePct
 * @param {number} [p.years=20]
 * @param {number} [p.grossReturnPct=7]
 */
export function compareFeeDrag({
  portfolioValue,
  currentFeePct,
  proposedFeePct,
  years = 20,
  grossReturnPct = 7,
} = {}) {
  const current = projectFeeDrag({ portfolioValue, annualFeePct: currentFeePct, years, grossReturnPct });
  const proposed = projectFeeDrag({ portfolioValue, annualFeePct: proposedFeePct, years, grossReturnPct });

  return {
    current,
    proposed,
    years,
    // Positive when the proposed schedule is cheaper.
    cumulativeFeeSavings: current.totalFees - proposed.totalFees,
    endingBalanceDifference: proposed.endingBalance - current.endingBalance,
    annualFeeReductionPct: (Number(currentFeePct) || 0) - (Number(proposedFeePct) || 0),
  };
}
