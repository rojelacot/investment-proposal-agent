"""STEP 5.3 - Two-stage discounted cash flow for an individual equity.

Per CLAUDE.md this runs ONLY when a proposal includes individual equity
positions. An asset-class allocation (equity index sleeve, no single names)
does not trigger it. The script is provided and self-tests on a sample input
so the capability exists when a future proposal holds direct equities.

Usage:
    python dcf.py --fcf 5.0e9 --growth 0.08 --years 5 --terminal-growth 0.025 \
                  --discount 0.09 --shares 1.0e9 --net-debt 2.0e9
"""
import argparse
import json


def two_stage_dcf(fcf0, growth, years, terminal_growth, discount, shares, net_debt):
    pv_stage1 = 0.0
    fcf = fcf0
    for t in range(1, years + 1):
        fcf *= (1.0 + growth)
        pv_stage1 += fcf / (1.0 + discount) ** t
    terminal_fcf = fcf * (1.0 + terminal_growth)
    terminal_value = terminal_fcf / (discount - terminal_growth)
    pv_terminal = terminal_value / (1.0 + discount) ** years
    enterprise_value = pv_stage1 + pv_terminal
    equity_value = enterprise_value - net_debt
    return {
        "pv_stage1": pv_stage1,
        "pv_terminal": pv_terminal,
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "intrinsic_value_per_share": equity_value / shares,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--fcf", type=float, default=5.0e9)
    ap.add_argument("--growth", type=float, default=0.08)
    ap.add_argument("--years", type=int, default=5)
    ap.add_argument("--terminal-growth", type=float, default=0.025)
    ap.add_argument("--discount", type=float, default=0.09)
    ap.add_argument("--shares", type=float, default=1.0e9)
    ap.add_argument("--net-debt", type=float, default=2.0e9)
    args = ap.parse_args()

    result = two_stage_dcf(args.fcf, args.growth, args.years,
                           args.terminal_growth, args.discount,
                           args.shares, args.net_debt)
    print(json.dumps({"model": "dcf", "inputs": vars(args), "result": result}, indent=2))


if __name__ == "__main__":
    main()
