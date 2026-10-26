"""Compares a recalculated workbook against the Python reference model.

Usage (from tools/):  ../.venv/bin/python verify_workbook.py <recalculated.xlsx>
"""
import sys
from openpyxl import load_workbook

from assumptions import LINES
from reference_model import line_table, pnl

FIRST = 4
TOL = 0.05  # $000s; inputs are rounded to $100, so $50 is well inside rounding


def main(path):
    wb = load_workbook(path, data_only=True)
    lines = line_table()
    p = pnl(lines)
    fails = 0

    def check(label, got, want, tol=TOL):
        nonlocal fails
        ok = got is not None and abs(got - want) <= tol
        fails += not ok
        if not ok:
            print(f"MISMATCH {label}: workbook {got} reference {want}")

    bva, fc, pl = wb["BvA"], wb["Forecast"], wb["P&L Summary"]
    for i, row in lines.iterrows():
        r = FIRST + i
        check(f"{row.id} month budget", bva.cell(r, 5).value, row.month_budget)
        check(f"{row.id} month actual", bva.cell(r, 6).value, row.month_actual)
        check(f"{row.id} YTD budget", bva.cell(r, 9).value, row.ytd_budget)
        check(f"{row.id} YTD actual", bva.cell(r, 10).value, row.ytd_actual)
        check(f"{row.id} YTD fav var", bva.cell(r, 11).value, row.ytd_fav)
        flag = bva.cell(r, 13).value or ""
        if flag != row.flag:
            fails += 1
            print(f"MISMATCH {row.id} flag: workbook '{flag}' reference '{row.flag}'")
        for m in range(1, 13):
            check(f"{row.id} forecast m{m}", fc.cell(r, 5 + m).value, row[f"fcst_{m}"])
        check(f"{row.id} FY forecast", fc.cell(r, 18).value, row.fy_forecast)

    col_of = {"ytd_budget": 3, "ytd_actual": 4, "fy_budget": 7, "fy_forecast": 8}
    for r in range(4, 4 + len(p.index)):
        label = pl.cell(r, 1).value
        for key, c in col_of.items():
            tol = 0.0005 if "%" in label else TOL
            check(f"P&L {label} {key}", pl.cell(r, c).value, p.loc[label, key], tol)

    ck = wb["Checks"]
    for r in range(4, ck.max_row + 1):
        if ck.cell(r, 2).value is False:
            fails += 1
            print(f"CHECK FAILED: {ck.cell(r, 1).value}")
    print(f"Lines: {len(LINES)}  P&L rows: {len(p.index)}  Checks sheet ALL PASS = {ck.cell(ck.max_row, 2).value}")
    print("ALL MATCH" if fails == 0 else f"{fails} MISMATCHES")
    return fails


if __name__ == "__main__":
    sys.exit(1 if main(sys.argv[1]) else 0)
