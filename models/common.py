"""Shared loaders so every model draws from the SAME Step 4 assumptions.

This is the single source of capital-market assumptions (cma.json) and the
cached historical-style return series (historical_returns.csv). backtest.py,
monte_carlo.py and risk_metrics.py all import from here, which is what makes
the CLAUDE.md Step 6 check "backtest and Monte Carlo draw from the same
assumptions" structurally true rather than a matter of trust.
"""
import csv
import json
import os

import numpy as np

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "data")

# Canonical asset-class order used everywhere downstream.
ASSETS = ["us_equity", "bonds", "alternatives", "cash"]


def load_cma():
    with open(os.path.join(DATA, "cma.json")) as f:
        return json.load(f)


def expected_returns(cma, assets=ASSETS):
    """Geometric expected returns vector, in ASSETS order."""
    return np.array([cma["assets"][a]["expected_return"] for a in assets])


def volatilities(cma, assets=ASSETS):
    return np.array([cma["assets"][a]["volatility"] for a in assets])


def cov_matrix(cma, assets=ASSETS):
    vols = volatilities(cma, assets)
    corr = np.array([[cma["correlations"][a][b] for b in assets] for a in assets])
    return np.outer(vols, vols) * corr


def arithmetic_means(cma, assets=ASSETS):
    """Convert geometric expected returns to arithmetic means for simulation.

    mu_arith ~= mu_geo + 0.5 * sigma^2 (lognormal adjustment). Both the
    backtest series generator and the Monte Carlo use this same conversion so
    the two models remain on identical footing.
    """
    mu_geo = expected_returns(cma, assets)
    vols = volatilities(cma, assets)
    return mu_geo + 0.5 * vols ** 2


def load_returns():
    """Return (years list, NxK ndarray of annual returns) from the cached CSV."""
    path = os.path.join(DATA, "historical_returns.csv")
    years, rows = [], []
    with open(path) as f:
        for row in csv.DictReader(f):
            years.append(int(row["year"]))
            rows.append([float(row[a]) for a in ASSETS])
    return years, np.array(rows)


def weights_vector(allocation, assets=ASSETS):
    w = np.array([float(allocation.get(a, 0.0)) for a in assets])
    return w


def parse_allocation(s):
    """Parse an --allocation JSON string into a normalized weights dict."""
    alloc = json.loads(s)
    return {a: float(alloc.get(a, 0.0)) for a in ASSETS}
