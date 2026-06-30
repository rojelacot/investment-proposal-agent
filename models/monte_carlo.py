"""STEP 5.2 - Monte Carlo simulation over the client's horizon.

Draws from the SAME cma.json assumptions used to build the backtest series
(arithmetic means + covariance), so the two models are consistent by
construction. Reports terminal-value distribution and probability ranges -
never single-point guarantees.

Usage:
    python monte_carlo.py --allocation '{...}' --portfolio 2500000 --horizon 12 --paths 20000
"""
import argparse
import json

import numpy as np

from common import (arithmetic_means, cov_matrix, load_cma, parse_allocation,
                    weights_vector)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--allocation", required=True)
    ap.add_argument("--portfolio", type=float, required=True)
    ap.add_argument("--horizon", type=int, required=True)
    ap.add_argument("--paths", type=int, default=20000)
    ap.add_argument("--seed", type=int, default=20260630)
    args = ap.parse_args()

    cma = load_cma()
    alloc = parse_allocation(args.allocation)
    w = weights_vector(alloc)
    mu = arithmetic_means(cma)
    cov = cov_matrix(cma)
    inflation = cma["inflation_cpi"]

    rng = np.random.default_rng(args.seed)
    # shape: (paths, horizon, n_assets)
    draws = rng.multivariate_normal(mu, cov, size=(args.paths, args.horizon))
    port_annual = draws @ w  # (paths, horizon)
    growth = np.prod(1.0 + port_annual, axis=1)  # (paths,)
    terminal = args.portfolio * growth

    pct = lambda p: float(np.percentile(terminal, p))
    real_initial = args.portfolio * (1.0 + inflation) ** args.horizon
    double_target = args.portfolio * 2.0

    out = {
        "model": "monte_carlo",
        "paths": args.paths,
        "horizon_years": args.horizon,
        "allocation": alloc,
        "starting_value": args.portfolio,
        "terminal_value": {
            "p5": pct(5), "p10": pct(10), "p25": pct(25), "p50": pct(50),
            "p75": pct(75), "p90": pct(90), "p95": pct(95),
            "mean": float(np.mean(terminal)),
        },
        "annualized_return_p50": float((pct(50) / args.portfolio) ** (1.0 / args.horizon) - 1.0),
        "prob_preserve_real_value": float(np.mean(terminal >= real_initial)),
        "prob_double": float(np.mean(terminal >= double_target)),
        "real_value_threshold": float(real_initial),
        "inflation_assumption": inflation,
    }
    print(json.dumps(out, indent=2))


if __name__ == "__main__":
    main()
