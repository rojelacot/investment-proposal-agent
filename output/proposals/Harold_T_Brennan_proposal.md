# Investment Proposal — Harold T. Brennan

**Prepared:** 2026-06-30
**Investment Horizon:** 9 years
**Portfolio Size:** $1,800,000
**Risk Profile:** CONSERVATIVE

---

## 1. Client Summary

Harold T. Brennan, age 71, holds an investable portfolio of $1,800,000
with a 9-year horizon and a self-described low risk tolerance.

**Goals**
- Preserve capital and avoid large drawdowns in retirement
- Generate stable income to supplement Social Security and a pension
- Maintain a reserve for potential long-term care needs

**Restrictions**
- No individual equities — pooled funds / ETFs only
- No high-yield (below-investment-grade) bonds
- Maintain at least 10% portfolio liquidity

---

## 2. Risk Profile Classification

**CONSERVATIVE.** Driven by: risk_tolerance == "low".

---

## 3. Recommended Allocation

| Asset Class | Weight |
|---|---:|
| US Large-Cap Equity | 25% |
| Investment-Grade Bonds | 60% |
| Diversified Alternatives | 5% |
| Cash / Short-Term | 10% |
| **Total** | **100%** |

This is the CONSERVATIVE base allocation, retained after confirming it satisfies the client's
restrictions (liquidity floor 10%, no restricted asset classes held).

---

## 4. Model Methodology

All material figures are produced by Python scripts (the source of truth); none are hand-computed.

- **Assumptions (Step 4):** LIVE Step-4 assumptions. Per-asset returns are trailing realized geometric returns from Yahoo Finance proxies (longest available window), cross-referenced against published long-run sources; where divergence exceeded 15% the LOWER figure was adopted. Volatilities are annualized from monthly returns; correlations are sample correlations of monthly returns. Cash/risk-free and CPI from FRED. Stored in `data/cma.json`,
  every figure cross-referenced and logged to `data/source_log.json`.
- **Backtest** — `models/backtest.py`, window 2007–2025
  (19 years).
- **Monte Carlo** — `models/monte_carlo.py`, 20,000 paths over 9 years,
  drawing from the *same* `cma.json` assumptions as the backtest.
- **Risk metrics** — `models/risk_metrics.py`.
- **DCF** — not run: the allocation holds no individual equity positions.

**Validation (Step 6):** independent CAGR recompute matched
the script; MC median is within the plausible band; backtest and
Monte Carlo confirmed to share inputs. Observation: backtest realized mean 0.0540 vs expected 0.0560 (+-0.14sigma).

---

## 5. Backtest Results

| Metric | Value |
|---|---:|
| Portfolio CAGR | 5.21% |
| Annualized volatility | 6.30% |
| Best year | 2019: +14.16% |
| Worst year | 2022: -12.56% |
| Positive years | 78.9% |
| Growth multiple | 2.62x |

The growth multiple reflects one realized historical path and is **not** a forecast.

---

## 6. Monte Carlo Projection

20,000 simulated paths, 9-year horizon, start $1,800,000.
Results are probability ranges, not guarantees.

| Percentile | Terminal Value |
|---|---:|
| P5 | $2,209,339 |
| P10 | $2,348,763 |
| P25 | $2,598,302 |
| P50 | $2,901,062 |
| P75 | $3,238,594 |
| P90 | $3,569,541 |
| P95 | $3,782,229 |
| Mean | $2,936,135 |

- Median annualized return: **≈ 5.4%**.
- **≈ 73% probability** of preserving real purchasing power
  (ending ≥ $2,622,448, the start grown at 4.3% CPI).
- **≈ 9% probability** of at least doubling the portfolio.

---

## 7. Risk Metrics

| Metric | Value |
|---|---:|
| Sharpe ratio | 0.24 |
| Sortino ratio | 0.15 |
| Maximum drawdown | -12.56% |
| 95% 1-yr Value-at-Risk | 6.63% |
| Downside deviation | 10.14% |
| Risk-free rate | 3.90% |

---

## 8. Data Sources

From `data/source_log.json`:

| Field | Value | Source | URL |
|---|---|---|---|
| us_equity expected_return (primary, trailing realized) | 0.1082 | Yahoo Finance via yahoo-finance2 — SPY | https://finance.yahoo.com/ |
| us_equity expected_return (cross-reference) | 0.1000 | Damodaran NYU Stern, S&P 500 geometric 1928-2024 | https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histretSP.html |
| us_equity volatility (annualized monthly) | 0.1498 | Yahoo Finance via yahoo-finance2 — SPY | https://finance.yahoo.com/ |
| bonds expected_return (primary, trailing realized) | 0.0309 | Yahoo Finance via yahoo-finance2 — AGG | https://finance.yahoo.com/ |
| bonds expected_return (cross-reference) | 0.0480 | Damodaran NYU Stern, 10-yr T.Bond geometric 1928-2024 | https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/histretSP.html |
| bonds volatility (annualized monthly) | 0.0451 | Yahoo Finance via yahoo-finance2 — AGG | https://finance.yahoo.com/ |
| alternatives expected_return (primary, trailing realized) | 0.0484 | Yahoo Finance via yahoo-finance2 — VNQ+DBC 50/50 | https://finance.yahoo.com/ |
| alternatives expected_return (cross-reference) | 0.0570 | FTSE NAREIT (REIT) + S&P GSCI (commodity) long-run blend | https://www.reit.com/data-research/reit-indexes/annual-index-values-returns |
| alternatives volatility (annualized monthly) | 0.1644 | Yahoo Finance via yahoo-finance2 — VNQ+DBC 50/50 | https://finance.yahoo.com/ |
| cash expected_return / risk_free_rate | 0.0390 | FRED 3-Month Treasury Constant Maturity (DGS3MO), trailing ~1yr avg | https://fred.stlouisfed.org/series/DGS3MO |
| inflation_cpi (YoY) | 0.0427 | FRED CPI All Urban Consumers (CPIAUCSL), 12-month change | https://fred.stlouisfed.org/series/CPIAUCSL |

---

## 9. Model Limitations & Disclosures

- Monte Carlo assumes the adopted means, volatilities, and correlations persist; real markets exhibit
  fat tails and regime shifts not captured by normally distributed annual returns.
- The backtest is not predictive of future returns; it is a single realized path.
- All probabilities are model estimates, not guarantees.
- This is a sample proposal generated by the pipeline for illustration; it is not personalized advice.

---

*Generated by the Investment Proposal Agent pipeline following CLAUDE.md Steps 1–10.*
