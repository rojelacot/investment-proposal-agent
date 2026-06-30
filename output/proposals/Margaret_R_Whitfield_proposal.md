# Investment Proposal — Margaret R. Whitfield

**Prepared:** 2026-06-30
**Investment Horizon:** 12 years
**Portfolio Size:** $2,500,000
**Risk Profile:** MODERATE

---

## 1. Client Summary

Margaret R. Whitfield, age 58, holds an investable portfolio of $2,500,000
with a 12-year horizon and a self-described moderate risk tolerance.

**Goals**
- Fund retirement income beginning at age 65
- Preserve real (inflation-adjusted) purchasing power over the horizon
- Leave a meaningful legacy for two adult children

**Restrictions**
- No tobacco-producing companies
- No direct cryptocurrency holdings
- Maintain at least 10% portfolio liquidity

---

## 2. Risk Profile Classification

**MODERATE.** Driven by: risk_tolerance == "moderate"; age 58; horizon 12 (residual case).

---

## 3. Recommended Allocation

| Asset Class | Weight |
|---|---:|
| US Large-Cap Equity | 45% |
| Investment-Grade Bonds | 30% |
| Diversified Alternatives | 15% |
| Cash / Short-Term | 10% |
| **Total** | **100%** |

This is the MODERATE base allocation, retained after confirming it satisfies the client's
restrictions (liquidity floor 10%, no restricted asset classes held).

---

## 4. Model Methodology

All material figures are produced by Python scripts (the source of truth); none are hand-computed.

- **Assumptions (Step 4):** LIVE Step-4 assumptions. Per-asset returns are trailing realized geometric returns from Yahoo Finance proxies (longest available window), cross-referenced against published long-run sources; where divergence exceeded 15% the LOWER figure was adopted. Volatilities are annualized from monthly returns; correlations are sample correlations of monthly returns. Cash/risk-free and CPI from FRED. Stored in `data/cma.json`,
  every figure cross-referenced and logged to `data/source_log.json`.
- **Backtest** — `models/backtest.py`, window 2007–2025
  (19 years).
- **Monte Carlo** — `models/monte_carlo.py`, 20,000 paths over 12 years,
  drawing from the *same* `cma.json` assumptions as the backtest.
- **Risk metrics** — `models/risk_metrics.py`.
- **DCF** — not run: the allocation holds no individual equity positions.

**Validation (Step 6):** independent CAGR recompute matched
the script; MC median is within the plausible band; backtest and
Monte Carlo confirmed to share inputs. Observation: backtest realized mean 0.0736 vs expected 0.0765 (+-0.13sigma).

---

## 5. Backtest Results

| Metric | Value |
|---|---:|
| Portfolio CAGR | 6.86% |
| Annualized volatility | 10.19% |
| Best year | 2019: +19.89% |
| Worst year | 2008: -19.01% |
| Positive years | 78.9% |
| Growth multiple | 3.52x |

The growth multiple reflects one realized historical path and is **not** a forecast.

---

## 6. Monte Carlo Projection

20,000 simulated paths, 12-year horizon, start $2,500,000.
Results are probability ranges, not guarantees.

| Percentile | Terminal Value |
|---|---:|
| P5 | $3,488,895 |
| P10 | $3,924,025 |
| P25 | $4,731,836 |
| P50 | $5,801,990 |
| P75 | $7,101,993 |
| P90 | $8,494,211 |
| P95 | $9,407,376 |
| Mean | $6,048,224 |

- Median annualized return: **≈ 7.3%**.
- **≈ 87% probability** of preserving real purchasing power
  (ending ≥ $4,129,072, the start grown at 4.3% CPI).
- **≈ 69% probability** of at least doubling the portfolio.

---

## 7. Risk Metrics

| Metric | Value |
|---|---:|
| Sharpe ratio | 0.34 |
| Sortino ratio | 0.26 |
| Maximum drawdown | -19.01% |
| 95% 1-yr Value-at-Risk | 13.43% |
| Downside deviation | 13.27% |
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
