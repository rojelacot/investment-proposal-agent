"""STEP 5.4 - Risk metrics for the proposed allocation.

Sharpe, Sortino, max drawdown, and 95% VaR computed from the same cached
portfolio return series the backtest uses. Risk-free rate comes from cma.json.

Usage:
    python risk_metrics.py --allocation '{...}'
"""
import argparse
import json

import numpy as np

from common import load_cma, load_returns, parse_allocation, weights_vector


def max_drawdown(cumulative):
    running_max = np.maximum.accumulate(cumulative)
    drawdowns = cumulative / running_max - 1.0
    return float(np.min(drawdowns))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--allocation", required=True)
    args = ap.parse_args()

    cma = load_cma()
    rf = cma["risk_free_rate"]
    alloc = parse_allocation(args.allocation)
    w = weights_vector(alloc)
    _, R = load_returns()

    r = R @ w
    mean, vol = np.mean(r), np.std(r, ddof=1)
    excess = r - rf
    downside = r[r < rf]
    downside_dev = np.sqrt(np.mean((downside - rf) ** 2)) if downside.size else 0.0

    cumulative = np.cumprod(1.0 + r)
    var95 = -float(np.percentile(r, 5))  # historical 1-yr VaR at 95%

    out = {
        "model": "risk_metrics",
        "allocation": alloc,
        "risk_free_rate": rf,
        "mean_annual_return": float(mean),
        "annualized_volatility": float(vol),
        "sharpe_ratio": float(np.mean(excess) / vol),
        "sortino_ratio": float(np.mean(excess) / downside_dev) if downside_dev else None,
        "max_drawdown": max_drawdown(cumulative),
        "var_95_annual": var95,
        "downside_deviation": float(downside_dev),
    }
    print(json.dumps(out, indent=2))


if __name__ == "__main__":
    main()
