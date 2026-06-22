# Investment Proposal Agent — How It Works

---

## What It Does
You paste in a client brief (free text), the app reads it, runs the strategy math, and builds a branded PowerPoint proposal. Takes about 30 seconds start to finish.

---

## Step-by-Step Flow

1. **Type your client notes** — paste in anything you know: name, age, stock position, goals, tax situation, etc. Plain English works fine.
2. **App reads the notes** — it automatically pulls out the key numbers (position size, age, tax rate, cost basis, etc.) using pattern recognition.
3. **Clarifying questions appear** — if anything critical is missing, the app asks follow-up questions before you proceed.
4. **Review panel populates** — all extracted data appears in editable fields so you can correct anything it got wrong.
5. **Select your strategies** — choose which strategies to include: CRT, Tax-Loss Harvesting, Option Collar, Estate Planning, etc.
6. **Pick a portfolio model + risk profile** — the app suggests one based on the notes; you can override.
7. **Preview screen** — before generating, you see every calculated number and the full fund lineup. Click any number to edit it.
8. **Generate** — one click produces the finished PowerPoint.

---

## Proposal Quality Score

- A live score (0–100%) grades your notes as you type.
- Rewards: having a ticker, position size, investable assets, goals, risk tolerance, time horizon, tax context.
- It's **informational only** — a low score doesn't block you from generating.

---

## How the Math Works

**Every strategy is sized at 30% of the stock position:**
- CRT gets 30%
- Harvesting sleeve gets 30%
- Option collar gets 30%
- Up to 90% of the position can be addressed; the rest stays as unmanaged stock.

**CRT assumptions:**
- Payout rate is fixed at **5%** (IRS minimum for a valid CRT).
- The charitable deduction is calculated using **IRS actuarial tables by the client's age** — the older the client, the higher the deduction factor (ranges from ~22% at age 45 to ~58% at age 85).
- Annual income from the CRT = contribution × 5%.

**Harvesting sleeve (130/30 long/short):**
- Assumes **25% of the sleeve** generates harvestable losses each year.
- Tax savings = those losses × the client's combined tax rate.

**Option collar:**
- Put floor at **−15%** below current stock price.
- Call cap at **+19%** above current stock price.

---

## Defaults (Used When Info Is Missing)

These kick in silently if a number wasn't provided. You can always override them in the preview modal.

| Field | Default |
|---|---|
| Cost basis | 15% |
| Combined tax rate | 37.1% |
| Federal LTCG rate | 23.8% (20% + 3.8% NIIT) |
| Stock price | $200 (auto-fetched live if ticker is known) |
| Client age | 65 |

---

## Portfolio Models

| Model | Account Type | Alternatives? |
|---|---|---|
| **Core Private** | Qualified | Yes — includes private equity (Partners Group) and hedge funds (North Rock) |
| **Select Liquidity** | Taxable | Yes — interval funds only (real estate, private credit, hedged equity, real assets) |
| **Traditional** | Taxable | No — public stocks and bonds only |
| **Focused B** | Taxable | No — same as Traditional but no US Core equity (heavier value/growth tilt) |

- The app suggests a model and risk profile automatically based on your notes.
- Risk profiles range from Conservative (20/80) to Aggressive (100/0).
- Partners Group and North Rock only appear in **Core Private** — they're always 0% in the other models.

---

## Fund Data

- All fund names, allocations, tickers, and fees come directly from **Strategies 2026Q2.xlsx**.
- The allocation slides in the proposal pull the exact funds and percentages for the chosen model and risk profile.
- The preview modal shows the full fund lineup — funds with 0% at the selected risk profile appear greyed out so you can see the complete universe.

---

## Guardrails

- If there's no stock position, all concentrated-stock strategies (CRT, harvesting, collar) are automatically zeroed out and those slides are skipped.
- If the model is Traditional or Focused B, the Alternatives slice is hidden from the donut chart (those are no-alternatives portfolios).
- Net worth can't be less than investable assets — if it is, the app corrects it automatically.
- Stock price is auto-fetched live when you open the preview. If the API fails, you can type it in manually.
- CRT deduction age is clamped between 45 and 85 (the range of the IRS table).
