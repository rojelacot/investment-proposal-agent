# Investment Proposal Agent — Validated Loop

Execute these steps in order. State which step you're on before proceeding.
Never skip a step. Never compute material numbers yourself — always run a script.

---

## Project Structure

```
investment-proposals/
├── CLAUDE.md
├── clients/
│   └── client_A.json
├── models/
│   ├── monte_carlo.py
│   ├── dcf.py
│   ├── risk_metrics.py        ← Sharpe, Sortino, vol
│   └── backtest.py
├── data/
│   ├── source_log.json        ← every data point + source + retrieval date
│   └── historical_returns.csv ← cached asset class returns
├── template/
│   └── proposal_template.md
└── output/
    └── proposals/
```

---

## STEP 1 — Load Client Data

Read `/clients/{filename}.json`.

---

## STEP 2 — Validate Data

Required fields: `name`, `age`, `risk_tolerance`, `investment_horizon`, `portfolio_size`, `goals`, `restrictions`.

Missing field → **STOP**. Report exactly what is missing. Do not proceed.

---

## STEP 3 — Classify Risk Profile

- **CONSERVATIVE** — `risk_tolerance == "low"` OR (`age > 60` AND `horizon < 5`)
- **AGGRESSIVE** — `risk_tolerance == "high"` AND `horizon > 10` AND `age < 50`
- **MODERATE** — everything else

State the profile and which fields drove the decision.

---

## STEP 4 — Source Market & Historical Data

For every assumption used downstream (expected returns, volatility, correlations):

**Pull via:**
- `yfinance` — price history, dividend yields, expense ratios
- `FRED` — risk-free rate, inflation (CPI), Treasury yields
- `Stooq` — secondary price check if yfinance data looks anomalous

**Expected return & volatility assumptions:**
- Primary: trailing historical returns from yfinance (min 10-yr window, or longest available)
- Cross-check: Damodaran NYU Stern data (https://pages.stern.nyu.edu/~adamodar/) or Vanguard capital markets assumptions

If two reputable sources diverge > 15% on the same figure:
- Use the **lower (more conservative)** figure
- Flag the divergence explicitly in the proposal's methodology section

**Log every data pull to `/data/source_log.json`:**
```json
{
  "field": "",
  "value": "",
  "source_name": "",
  "source_url": "",
  "date_range": "",
  "retrieval_date": ""
}
```

Approved sources: Vanguard research, SPDR/State Street fact sheets, Fama-French data library, FRED, Morningstar, SEC filings, fund prospectuses, Damodaran NYU.

**Never use a single source for a return/volatility assumption** — cross-reference at least 2.

---

## STEP 5 — Run Quantitative Models (Python = source of truth)

Run in this order, passing Step 4 data as inputs:

1. `models/backtest.py` — historical performance of proposed allocation, 1990–present (or longest available window)
2. `models/monte_carlo.py` — 10,000+ path simulation over client's `investment_horizon`
3. `models/dcf.py` — only if proposal includes individual equity positions
4. `models/risk_metrics.py` — Sharpe, Sortino, max drawdown, VaR (95%)

Capture raw script output exactly as produced (stdout/JSON). **Do not paraphrase numbers.**

---

## STEP 6 — Validate Model Output

Run every time. No exceptions.

- [ ] Manually recompute one intermediate calculation independent of the script (e.g., hand-check CAGR for one asset over the backtest period) — confirms the script logic is not broken
- [ ] Monte Carlo median/mean terminal value is within plausible bounds for the asset mix (cross-check against Step 4 historical data)
- [ ] Backtest and Monte Carlo draw from the **same** Step 4 return assumptions — no mismatched inputs
- [ ] All four allocation validation checks pass for this client's specific mix (not just a generic template)

Any failure → **re-run the script, do not proceed to Step 7**. Report the discrepancy and likely cause.

---

## STEP 7 — Generate Portfolio Allocation

Base models (adjust for client restrictions and Step 5 results):

| Profile      | Allocation                                              |
|--------------|---------------------------------------------------------|
| CONSERVATIVE | 60% Bonds / 25% Large-Cap Equity / 10% Cash / 5% Alts  |
| MODERATE     | 45% Equity / 30% Bonds / 15% Alternatives / 10% Cash   |
| AGGRESSIVE   | 75% Equity / 15% Alternatives / 5% Bonds / 5% Cash     |

---

## STEP 8 — Self-Check Allocation

- [ ] Sums to 100%
- [ ] No restricted assets included
- [ ] Matches risk profile from Step 3
- [ ] Consistent with Monte Carlo / backtest results from Step 5 (e.g., do not recommend aggressive growth language if simulation shows high probability of shortfall)

Fail any check → **redo Step 7**.

---

## STEP 9 — Write Proposal

Fill `/template/proposal_template.md` completely. Include:

- **Data Sources** section — every source from `source_log.json`
- **Model Methodology** section — what was run, what assumptions, what time periods, which script names
- Monte Carlo results stated as **probability ranges, not single-point guarantees**
  - ✅ "70% probability of reaching $X by year 10"
  - ❌ "will reach $X"

---

## STEP 10 — Save Output

- `/output/proposals/{client_name}_proposal.md` — the final proposal
- `/output/proposals/{client_name}_model_run.json` — raw Step 5 output for audit

---

## Hard Rules

- Never invent a data source or URL
- Never state a model result with more precision than the model supports
- Never skip the Step 4 cross-reference requirement
- Always disclose model limitations in the final proposal:
  - Monte Carlo assumes historical patterns persist
  - Backtest is not predictive of future returns
