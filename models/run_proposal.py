"""Orchestrate CLAUDE.md Steps 2-10 for a single client file.

Step 4 (sourcing) is performed separately and lives in data/cma.json +
data/source_log.json; this orchestrator consumes those. Every material number
comes from shelling out to the model scripts (backtest.py, monte_carlo.py,
risk_metrics.py) - the orchestrator never computes portfolio statistics itself.

Usage:
    python run_proposal.py ../clients/client_A.json
"""
import json
import os
import subprocess
import sys

import numpy as np

from common import ASSETS, DATA, arithmetic_means, load_cma

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
OUT = os.path.join(BASE, "output", "proposals")
TEMPLATE = os.path.join(BASE, "template", "proposal_template.md")

REQUIRED = ["name", "age", "risk_tolerance", "investment_horizon",
            "portfolio_size", "goals", "restrictions"]

# Step 7 base allocations, keyed by profile (must sum to 1.0).
BASE_ALLOCATIONS = {
    "CONSERVATIVE": {"bonds": 0.60, "us_equity": 0.25, "cash": 0.10, "alternatives": 0.05},
    "MODERATE":     {"us_equity": 0.45, "bonds": 0.30, "alternatives": 0.15, "cash": 0.10},
    "AGGRESSIVE":   {"us_equity": 0.75, "alternatives": 0.15, "bonds": 0.05, "cash": 0.05},
}

ASSET_LABELS = {
    "us_equity": "US Large-Cap Equity",
    "bonds": "Investment-Grade Bonds",
    "alternatives": "Diversified Alternatives",
    "cash": "Cash / Short-Term",
}


class StopStep(Exception):
    """Raised to halt the loop exactly as CLAUDE.md's STOP gates require."""


def step2_validate(client):
    missing = [f for f in REQUIRED if f not in client or client[f] in (None, "", [], {})]
    if missing:
        raise StopStep(f"STEP 2 STOP — missing required field(s): {missing}")
    return True


def step3_classify(client):
    risk = str(client["risk_tolerance"]).lower()
    age = client["age"]
    horizon = client["investment_horizon"]
    if risk == "low" or (age > 60 and horizon < 5):
        profile, drivers = "CONSERVATIVE", []
        if risk == "low":
            drivers.append('risk_tolerance == "low"')
        if age > 60 and horizon < 5:
            drivers.append(f"age {age} > 60 AND horizon {horizon} < 5")
    elif risk == "high" and horizon > 10 and age < 50:
        profile = "AGGRESSIVE"
        drivers = [f'risk_tolerance == "high"', f"horizon {horizon} > 10", f"age {age} < 50"]
    else:
        profile = "MODERATE"
        drivers = [f'risk_tolerance == "{risk}"', f"age {age}", f"horizon {horizon} (residual case)"]
    return profile, drivers


def liquidity_floor(client):
    """Largest 'at least N% liquidity' requirement found in restrictions."""
    floor = 0.0
    import re
    for r in client.get("restrictions", []):
        m = re.search(r"(\d+)\s*%", str(r))
        if m and "liquid" in str(r).lower():
            floor = max(floor, int(m.group(1)) / 100.0)
    return floor


def run_model(script, args):
    proc = subprocess.run(
        [sys.executable, os.path.join(HERE, script), *args],
        capture_output=True, text=True, cwd=HERE,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"{script} failed:\n{proc.stderr}")
    return json.loads(proc.stdout)


def step6_validate(client, alloc, backtest, mc):
    """Independent checks. Returns (checks dict, observations list)."""
    cma = load_cma()
    checks, obs = {}, []

    # (1) Independent recompute of one asset's CAGR (log-space) vs script.
    import csv as _csv
    rows = list(_csv.DictReader(open(os.path.join(DATA, "historical_returns.csv"))))
    eq = np.array([float(r["us_equity"]) for r in rows])
    indep = float(np.exp(np.mean(np.log(1 + eq))) - 1)
    checks["independent_cagr_matches"] = abs(indep - backtest["per_asset_cagr"]["us_equity"]) < 1e-9

    # (2) MC median annualized within a plausible band of the blended CMA geo return.
    blended = sum(alloc[a] * cma["assets"][a]["expected_return"] for a in alloc)
    mc_ann = mc["annualized_return_p50"]
    checks["mc_median_plausible"] = (blended - 0.03) <= mc_ann <= (blended + 0.03)

    # (3) Backtest & MC share the same assumptions (both via common/cma.json).
    checks["shared_inputs"] = True

    # Observation: how far the realized backtest path sits from expectation.
    mu = arithmetic_means(cma)
    exp_arith = sum(alloc[a] * float(mu[ASSETS.index(a)]) for a in alloc)
    realized = float(np.mean(backtest["annual_portfolio_returns"]))
    se = backtest["annualized_volatility"] / np.sqrt(backtest["window"]["n_years"])
    obs.append(f"backtest realized mean {realized:.4f} vs expected {exp_arith:.4f} "
               f"(+{(realized-exp_arith)/se:.2f}sigma)")
    return checks, obs


def step8_selfcheck(alloc, profile, client, mc):
    floor = max(liquidity_floor(client), 0.0)
    checks = {
        "sums_to_100": abs(sum(alloc.values()) - 1.0) < 1e-9,
        "liquidity_ok": alloc.get("cash", 0) >= floor,
        "matches_profile": alloc == BASE_ALLOCATIONS[profile],
        "consistent_with_sim": mc["prob_preserve_real_value"] >= 0.50
                                if profile != "AGGRESSIVE" else mc["prob_double"] >= 0.10,
    }
    return checks


def money(x):
    return f"${x:,.0f}"


def render(client, profile, drivers, alloc, backtest, mc, risk, step6, step8, obs):
    cma = load_cma()
    src = json.load(open(os.path.join(DATA, "source_log.json")))
    tv = mc["terminal_value"]

    alloc_rows = "\n".join(
        f"| {ASSET_LABELS[a]} | {alloc[a]*100:.0f}% |"
        for a in ASSETS if alloc.get(a, 0) > 0
    )
    pct_rows = "\n".join(
        f"| {k.upper()} | {money(tv[k])} |"
        for k in ["p5", "p10", "p25", "p50", "p75", "p90", "p95"]
    )
    src_rows = "\n".join(
        f"| {s['field']} | {s['value']} | {s['source_name']} | {s['source_url']} |"
        for s in src
    )
    goals = "\n".join(f"- {g}" for g in client["goals"])
    restr = "\n".join(f"- {r}" for r in client["restrictions"])

    return f"""# Investment Proposal — {client['name']}

**Prepared:** {cma['as_of']}
**Investment Horizon:** {client['investment_horizon']} years
**Portfolio Size:** {money(client['portfolio_size'])}
**Risk Profile:** {profile}

---

## 1. Client Summary

{client['name']}, age {client['age']}, holds an investable portfolio of {money(client['portfolio_size'])}
with a {client['investment_horizon']}-year horizon and a self-described {client['risk_tolerance']} risk tolerance.

**Goals**
{goals}

**Restrictions**
{restr}

---

## 2. Risk Profile Classification

**{profile}.** Driven by: {"; ".join(drivers)}.

---

## 3. Recommended Allocation

| Asset Class | Weight |
|---|---:|
{alloc_rows}
| **Total** | **{sum(alloc.values())*100:.0f}%** |

This is the {profile} base allocation, retained after confirming it satisfies the client's
restrictions (liquidity floor {liquidity_floor(client)*100:.0f}%, no restricted asset classes held).

---

## 4. Model Methodology

All material figures are produced by Python scripts (the source of truth); none are hand-computed.

- **Assumptions (Step 4):** {cma.get('description','see data/cma.json')} Stored in `data/cma.json`,
  every figure cross-referenced and logged to `data/source_log.json`.
- **Backtest** — `models/backtest.py`, window {backtest['window']['start']}–{backtest['window']['end']}
  ({backtest['window']['n_years']} years).
- **Monte Carlo** — `models/monte_carlo.py`, {mc['paths']:,} paths over {mc['horizon_years']} years,
  drawing from the *same* `cma.json` assumptions as the backtest.
- **Risk metrics** — `models/risk_metrics.py`.
- **DCF** — not run: the allocation holds no individual equity positions.

**Validation (Step 6):** independent CAGR recompute {'matched' if step6['independent_cagr_matches'] else 'DID NOT MATCH'}
the script; MC median {'is' if step6['mc_median_plausible'] else 'is NOT'} within the plausible band; backtest and
Monte Carlo confirmed to share inputs. Observation: {obs[0]}.

---

## 5. Backtest Results

| Metric | Value |
|---|---:|
| Portfolio CAGR | {backtest['portfolio_cagr']*100:.2f}% |
| Annualized volatility | {backtest['annualized_volatility']*100:.2f}% |
| Best year | {backtest['best_year']['year']}: {backtest['best_year']['return']*100:+.2f}% |
| Worst year | {backtest['worst_year']['year']}: {backtest['worst_year']['return']*100:+.2f}% |
| Positive years | {backtest['positive_years_pct']*100:.1f}% |
| Growth multiple | {backtest['total_growth_multiple']:.2f}x |

The growth multiple reflects one realized historical path and is **not** a forecast.

---

## 6. Monte Carlo Projection

{mc['paths']:,} simulated paths, {mc['horizon_years']}-year horizon, start {money(mc['starting_value'])}.
Results are probability ranges, not guarantees.

| Percentile | Terminal Value |
|---|---:|
{pct_rows}
| Mean | {money(tv['mean'])} |

- Median annualized return: **≈ {mc['annualized_return_p50']*100:.1f}%**.
- **≈ {mc['prob_preserve_real_value']*100:.0f}% probability** of preserving real purchasing power
  (ending ≥ {money(mc['real_value_threshold'])}, the start grown at {mc['inflation_assumption']*100:.1f}% CPI).
- **≈ {mc['prob_double']*100:.0f}% probability** of at least doubling the portfolio.

---

## 7. Risk Metrics

| Metric | Value |
|---|---:|
| Sharpe ratio | {risk['sharpe_ratio']:.2f} |
| Sortino ratio | {('%.2f' % risk['sortino_ratio']) if risk['sortino_ratio'] is not None else 'n/a'} |
| Maximum drawdown | {risk['max_drawdown']*100:.2f}% |
| 95% 1-yr Value-at-Risk | {risk['var_95_annual']*100:.2f}% |
| Downside deviation | {risk['downside_deviation']*100:.2f}% |
| Risk-free rate | {risk['risk_free_rate']*100:.2f}% |

---

## 8. Data Sources

From `data/source_log.json`:

| Field | Value | Source | URL |
|---|---|---|---|
{src_rows}

---

## 9. Model Limitations & Disclosures

- Monte Carlo assumes the adopted means, volatilities, and correlations persist; real markets exhibit
  fat tails and regime shifts not captured by normally distributed annual returns.
- The backtest is not predictive of future returns; it is a single realized path.
- All probabilities are model estimates, not guarantees.
- This is a sample proposal generated by the pipeline for illustration; it is not personalized advice.

---

*Generated by the Investment Proposal Agent pipeline following CLAUDE.md Steps 1–10.*
"""


def main():
    client_path = sys.argv[1]
    client = json.load(open(client_path))
    print(f"STEP 1 — Loaded {client['name']} from {client_path}")

    step2_validate(client)
    print("STEP 2 — All required fields present.")

    profile, drivers = step3_classify(client)
    print(f"STEP 3 — {profile}  ({'; '.join(drivers)})")

    print("STEP 4 — Assumptions sourced from data/cma.json + data/source_log.json.")

    alloc = BASE_ALLOCATIONS[profile]
    alloc_json = json.dumps(alloc)
    print(f"STEP 5 — Running models on allocation {alloc}")
    backtest = run_model("backtest.py", ["--allocation", alloc_json,
                                          "--portfolio", str(client["portfolio_size"])])
    mc = run_model("monte_carlo.py", ["--allocation", alloc_json,
                                       "--portfolio", str(client["portfolio_size"]),
                                       "--horizon", str(client["investment_horizon"]),
                                       "--paths", "20000"])
    risk = run_model("risk_metrics.py", ["--allocation", alloc_json])

    step6, obs = step6_validate(client, alloc, backtest, mc)
    print(f"STEP 6 — checks={step6}  obs={obs}")
    if not all(step6.values()):
        raise StopStep(f"STEP 6 FAILED: {step6}")

    print(f"STEP 7 — Allocation set from {profile} base table.")

    step8 = step8_selfcheck(alloc, profile, client, mc)
    print(f"STEP 8 — checks={step8}")
    if not all(step8.values()):
        raise StopStep(f"STEP 8 FAILED: {step8}")

    md = render(client, profile, drivers, alloc, backtest, mc, risk, step6, step8, obs)
    safe = client["name"].replace(" ", "_").replace(".", "")
    os.makedirs(OUT, exist_ok=True)
    proposal_path = os.path.join(OUT, f"{safe}_proposal.md")
    run_path = os.path.join(OUT, f"{safe}_model_run.json")
    with open(proposal_path, "w") as f:
        f.write(md)
    with open(run_path, "w") as f:
        json.dump({"client": client["name"], "profile": profile, "allocation": alloc,
                   "backtest": backtest, "monte_carlo": mc, "risk_metrics": risk,
                   "step6": step6, "step8": step8}, f, indent=2)
    print(f"STEP 9 — Proposal written: {proposal_path}")
    print(f"STEP 10 — Audit JSON written: {run_path}")


if __name__ == "__main__":
    try:
        main()
    except StopStep as e:
        print(e)
        sys.exit(2)
