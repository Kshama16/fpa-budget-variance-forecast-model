"""Independent Python implementation of the workbook logic, used to verify the Excel results.

Run from tools/:  ../.venv/bin/python reference_model.py
"""
import numpy as np
import pandas as pd

from assumptions import (LINES, DEPARTMENTS, CURRENT_MONTH, VAR_PCT_THRESHOLD,
                         VAR_AMT_THRESHOLD, budget, actuals)

M = np.arange(1, 13)
CLOSED = M <= CURRENT_MONTH


def favorable(category, actual, budget_):
    # Revenue above budget is favorable; costs below budget are favorable.
    return actual - budget_ if category == "Revenue" else budget_ - actual


def line_table():
    b, a = budget(), actuals()
    rows = []
    for lid, name, dept, cat, method in LINES:
        bud, act = b[lid], a[lid]
        ytd_b, ytd_a = bud[CLOSED].sum(), act[CLOSED].sum()
        ratio = ytd_a / ytd_b if ytd_b else 1.0
        run_rate = act[(M > CURRENT_MONTH - 3) & CLOSED].sum() / 3
        if method == "Attainment":
            fcst_open = bud * ratio
        elif method == "Run-rate":
            fcst_open = np.full(12, run_rate)
        else:
            fcst_open = bud
        fcst = np.where(CLOSED, act, fcst_open)
        m_b, m_a = bud[CURRENT_MONTH - 1], act[CURRENT_MONTH - 1]
        ytd_fav = favorable(cat, ytd_a, ytd_b)
        ytd_pct = ytd_fav / abs(ytd_b) if ytd_b else 0.0
        flag = ""
        if abs(ytd_pct) >= VAR_PCT_THRESHOLD and abs(ytd_fav) >= VAR_AMT_THRESHOLD:
            flag = "Unfavorable" if ytd_fav < 0 else "Favorable"
        fy_b, fy_f = bud.sum(), fcst.sum()
        rows.append(dict(
            id=lid, line=name, dept=dept, cat=cat, method=method,
            month_budget=m_b, month_actual=m_a, month_fav=favorable(cat, m_a, m_b),
            ytd_budget=ytd_b, ytd_actual=ytd_a, ytd_fav=ytd_fav, ytd_pct=ytd_pct, flag=flag,
            fy_budget=fy_b, fy_forecast=fy_f, fy_fav=favorable(cat, fy_f, fy_b),
            **{f"fcst_{i}": v for i, v in enumerate(fcst, start=1)},
        ))
    return pd.DataFrame(rows)


def pnl(lines):
    """P&L summary: YTD budget/actual and FY budget/forecast with favorable variance."""
    def total(col, cat, dept=None):
        sel = lines["cat"] == cat
        if dept:
            sel &= lines["dept"] == dept
        return lines.loc[sel, col].sum()

    out = {}
    for col in ["ytd_budget", "ytd_actual", "fy_budget", "fy_forecast"]:
        rev, cogs = total(col, "Revenue"), total(col, "COGS")
        opex = {d: total(col, "Opex", d) for d in DEPARTMENTS}
        tot_opex = sum(opex.values())
        gp, ebitda = rev - cogs, rev - cogs - tot_opex
        out[col] = {"Revenue": rev, "COGS": cogs, "Gross Profit": gp, "Gross Margin %": gp / rev,
                    **{f"Opex - {d}": v for d, v in opex.items()},
                    "Total Opex": tot_opex, "EBITDA": ebitda, "EBITDA Margin %": ebitda / rev}
    return pd.DataFrame(out)


if __name__ == "__main__":
    lines = line_table()
    p = pnl(lines)
    pd.set_option("display.width", 200, "display.float_format", "{:,.2f}".format)
    print(p)
    print()
    print(lines.loc[lines.flag != "", ["id", "line", "dept", "ytd_budget", "ytd_actual", "ytd_fav", "ytd_pct", "flag"]]
          .sort_values("ytd_fav").to_string(index=False))
    print()
    print(lines[["id", "fy_budget", "fy_forecast", "fy_fav"]].to_string(index=False))
