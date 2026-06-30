"""STEP 5.1 - Historical performance of the proposed allocation.

Reads the cached annual return series and a proposed allocation, then reports
realized portfolio statistics over the full window. Numbers are emitted as JSON
on stdout; nothing here is paraphrased downstream.

Usage:
    python backtest.py --allocation '{"us_equity":0.45,"bonds":0.30,"alternatives":0.15,"cash":0.10}' \
                       --portfolio 2500000
"""
import argparse
import json

import numpy as np

from common import ASSETS, load_returns, parse_allocation, weights_vector


def cagr(returns):
    n = len(returns)
    growth = np.prod(1.0 + returns)
    return growth ** (1.0 / n) - 1.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--allocation", required=True)
    ap.add_argument("--portfolio", type=float, default=1.0)
    args = ap.parse_args()

    alloc = parse_allocation(args.allocation)
    w = weights_vector(alloc)
    years, R = load_returns()

    port_returns = R @ w  # annual portfolio returns
    cumulative = np.cumprod(1.0 + port_returns)
    ending_value = args.portfolio * cumulative[-1]

    # Per-asset CAGR is exposed so Step 6 can hand-check one of them.
    per_asset_cagr = {a: float(cagr(R[:, i])) for i, a in enumerate(ASSETS)}

    out = {
        "model": "backtest",
        "window": {"start": years[0], "end": years[-1], "n_years": len(years)},
        "allocation": alloc,
        "portfolio_cagr": float(cagr(port_returns)),
        "annualized_volatility": float(np.std(port_returns, ddof=1)),
        "best_year": {"year": int(years[int(np.argmax(port_returns))]),
                       "return": float(np.max(port_returns))},
        "worst_year": {"year": int(years[int(np.argmin(port_returns))]),
                        "return": float(np.min(port_returns))},
        "positive_years_pct": float(np.mean(port_returns > 0)),
        "starting_value": args.portfolio,
        "ending_value": float(ending_value),
        "total_growth_multiple": float(cumulative[-1]),
        "per_asset_cagr": per_asset_cagr,
        "annual_portfolio_returns": [float(x) for x in port_returns],
    }
    print(json.dumps(out, indent=2))


if __name__ == "__main__":
    main()
