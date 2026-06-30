"""Generate the cached historical-style annual return series.

IMPORTANT / DISCLOSURE: this does NOT scrape live prices. It produces a
deterministic (seeded) multivariate-normal series calibrated to the adopted
capital market assumptions in data/cma.json. The output is labeled in the
proposal methodology as "calibrated to published CMAs," not as raw historical
market data. This keeps the pipeline runnable offline and fully reproducible
while remaining honest about what the numbers are.

Run once to (re)create data/historical_returns.csv.
"""
import csv
import os

import numpy as np

from common import ASSETS, DATA, arithmetic_means, cov_matrix, load_cma

START_YEAR = 1990
END_YEAR = 2025  # inclusive
SEED = 19900101


def main():
    cma = load_cma()
    mu = arithmetic_means(cma)
    cov = cov_matrix(cma)
    years = list(range(START_YEAR, END_YEAR + 1))

    rng = np.random.default_rng(SEED)
    draws = rng.multivariate_normal(mu, cov, size=len(years))

    path = os.path.join(DATA, "historical_returns.csv")
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["year"] + ASSETS)
        for year, row in zip(years, draws):
            w.writerow([year] + [f"{x:.6f}" for x in row])

    print(f"wrote {len(years)} years ({START_YEAR}-{END_YEAR}) to {path}")


if __name__ == "__main__":
    main()
