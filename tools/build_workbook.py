"""Builds model/FPA_Budget_Variance_Forecast_Model.xlsx with live Excel formulas.

Budget and Actuals are the only hard-coded data sheets (blue font). Every other figure
is a formula that recalculates when the inputs, budget, actuals or closed month change.
Run from tools/:  ../.venv/bin/python build_workbook.py
"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as col
from openpyxl.formatting.rule import CellIsRule

from assumptions import (COMPANY, FISCAL_YEAR, CURRENT_MONTH, VAR_PCT_THRESHOLD,
                         VAR_AMT_THRESHOLD, MONTHS, LINES, DEPARTMENTS, budget, actuals)

OUT = Path(__file__).resolve().parents[1] / "model" / "FPA_Budget_Variance_Forecast_Model.xlsx"

NAVY = "1F3864"
HDR = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor=NAVY)
INPUT = Font(color="0000FF")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14, color=NAVY)
NUM = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
TOP = Border(top=Side(style="thin"))

FIRST = 4                      # first data row on line-item sheets
LAST = FIRST + len(LINES) - 1  # last data row
MCOL = 5                       # month 1 column (E) on Budget/Actuals


def header(ws, row, labels, start=1):
    for i, text in enumerate(labels, start=start):
        c = ws.cell(row, i, text)
        c.font, c.fill = HDR, HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = TITLE
    ws["A2"] = sub
    ws["A2"].font = Font(italic=True, color="666666")


def month_range(sheet, r):
    return f"{sheet}!${col(MCOL)}{r}:${col(MCOL + 11)}{r}"


MONTH_HDR = f"Budget!${col(MCOL)}$3:${col(MCOL + 11)}$3"


def data_sheet(wb, name, values, note):
    ws = wb.create_sheet(name)
    title(ws, f"{COMPANY} - FY{FISCAL_YEAR} {name} ($000s)", note)
    header(ws, 3, ["ID", "Line item", "Department", "Category"] + list(range(1, 13)) + ["FY Total"])
    for i, m in enumerate(MONTHS):
        ws.cell(3, MCOL + i).number_format = f'"{m}"'
    for r, (lid, name_, dept, cat, _) in enumerate(LINES, start=FIRST):
        ws.cell(r, 1, lid)
        ws.cell(r, 2, name_)
        ws.cell(r, 3, dept)
        ws.cell(r, 4, cat)
        for i, v in enumerate(values[lid]):
            if v == v:  # skip NaN (open months)
                c = ws.cell(r, MCOL + i, float(v))
                c.font, c.number_format = INPUT, NUM
        t = ws.cell(r, MCOL + 12, f"=SUM({col(MCOL)}{r}:{col(MCOL + 11)}{r})")
        t.number_format, t.font = NUM, BOLD
    widths(ws, [11, 34, 12, 10] + [9] * 12 + [11])
    ws.freeze_panes = "E4"
    return ws


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[col(i)].width = w


def build():
    wb = Workbook()
    b, a = budget(), actuals()

    # ---------- Cover ----------
    ws = wb.active
    ws.title = "Cover"
    title(ws, f"{COMPANY} - FY{FISCAL_YEAR} Budget vs. Actual & Rolling Forecast",
          "Illustrative FP&A model. All figures in $ thousands. Data is fictional.")
    rows = [
        ("Purpose", "Monthly and YTD budget-to-actual reporting, variance flags, a rolling 9+3 forecast, and a management P&L."),
        ("How to use", "Close a month by entering actuals on the Actuals sheet and setting the closed month on Inputs. Everything else recalculates."),
        ("Conventions", "Blue = hard-coded input; black = formula. Variances are shown as favorable (+) / unfavorable (-) for both revenue and costs."),
        ("Inputs", "Closed month, variance flag thresholds (% and $)."),
        ("Budget / Actuals", "21 line items x 12 months by department and category (Revenue, COGS, Opex)."),
        ("BvA", "Current month and YTD budget vs. actual by line, with variance % and automatic flags."),
        ("Forecast", "Closed months = actuals; open months forecast by line method: Attainment (budget x YTD actual/budget), Run-rate (avg of last 3 closed months) or Budget."),
        ("P&L Summary", "Revenue to EBITDA: YTD budget vs. actual and FY budget vs. forecast, with margins."),
        ("Mgmt Report", "Key metrics and variance commentary for leadership."),
        ("Checks", "Integrity checks; all must show TRUE."),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        ws.cell(i, 1, k).font = BOLD
        ws.cell(i, 2, v).alignment = Alignment(wrap_text=True, vertical="top")
    widths(ws, [18, 110])

    # ---------- Inputs ----------
    ws = wb.create_sheet("Inputs")
    title(ws, "Inputs", "Change these to close a new month or tighten variance flags.")
    for r, (label, value, fmt, nm) in enumerate([
        ("Last closed month (1-12)", CURRENT_MONTH, "0", "CurMonth"),
        ("Variance flag threshold, %", VAR_PCT_THRESHOLD, "0.0%", "ThrPct"),
        ("Variance flag threshold, $000s", VAR_AMT_THRESHOLD, "#,##0", "ThrAmt"),
    ], start=4):
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, value)
        c.font, c.number_format = INPUT, fmt
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"Inputs!$B${r}")
    ws["A8"] = "Closed month name"
    ws["B8"] = '=INDEX({"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},CurMonth)'
    widths(ws, [34, 12])

    data_sheet(wb, "Budget", b, "Approved FY budget by month.")
    data_sheet(wb, "Actuals", a, "Closed actuals from the general ledger; open months left blank.")

    # ---------- BvA ----------
    ws = wb.create_sheet("BvA")
    title(ws, "Budget vs. Actual ($000s)", "Favorable variance = actual revenue above budget, or actual cost below budget.")
    ws["F2"] = '="Month: "&Inputs!B8'
    header(ws, 3, ["ID", "Line item", "Department", "Category",
                   "Month Budget", "Month Actual", "Month Var Fav/(Unfav)", "Month Var %",
                   "YTD Budget", "YTD Actual", "YTD Var Fav/(Unfav)", "YTD Var %", "YTD Flag"])
    for r in range(FIRST, LAST + 1):
        for c_ in range(1, 5):
            ws.cell(r, c_, f"=Budget!{col(c_)}{r}")
        ws.cell(r, 5, f"=INDEX({month_range('Budget', r)},CurMonth)")
        ws.cell(r, 6, f"=INDEX({month_range('Actuals', r)},CurMonth)")
        ws.cell(r, 7, f'=IF($D{r}="Revenue",F{r}-E{r},E{r}-F{r})')
        ws.cell(r, 8, f"=IF(E{r}=0,0,G{r}/ABS(E{r}))")
        ws.cell(r, 9, f"=SUMPRODUCT({month_range('Budget', r)}*({MONTH_HDR}<=CurMonth))")
        ws.cell(r, 10, f"=SUMPRODUCT({month_range('Actuals', r)}*({MONTH_HDR}<=CurMonth))")
        ws.cell(r, 11, f'=IF($D{r}="Revenue",J{r}-I{r},I{r}-J{r})')
        ws.cell(r, 12, f"=IF(I{r}=0,0,K{r}/ABS(I{r}))")
        ws.cell(r, 13, f'=IF(AND(ABS(L{r})>=ThrPct,ABS(K{r})>=ThrAmt),IF(K{r}<0,"Unfavorable","Favorable"),"")')
        for c_ in (5, 6, 7, 9, 10, 11):
            ws.cell(r, c_).number_format = NUM
        for c_ in (8, 12):
            ws.cell(r, c_).number_format = PCT
    rng = f"M{FIRST}:M{LAST}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Unfavorable"'],
                                                  fill=PatternFill("solid", fgColor="F8CBAD")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Favorable"'],
                                                  fill=PatternFill("solid", fgColor="C6EFCE")))
    widths(ws, [11, 34, 12, 10, 11, 11, 13, 10, 11, 11, 13, 10, 13])
    ws.freeze_panes = "E4"

    # ---------- Forecast ----------
    ws = wb.create_sheet("Forecast")
    title(ws, "Rolling Forecast ($000s)", "Closed months = actuals. Open months use the line's method (blue, editable).")
    labels = ["ID", "Line item", "Department", "Category", "Method"] + list(range(1, 13)) + \
             ["FY Forecast", "FY Budget", "FY Var Fav/(Unfav)", "FY Var %", "YTD Actual", "YTD Budget", "YTD Attainment", "3-mo Run-rate"]
    header(ws, 3, labels)
    F_M1 = 6  # month 1 column (F)
    for i, m in enumerate(MONTHS):
        ws.cell(3, F_M1 + i).number_format = f'"{m}"'
    c_fy, c_bud, c_var, c_pct, c_ya, c_yb, c_att, c_rr = range(F_M1 + 12, F_M1 + 20)
    for r, (_, _, _, _, method) in enumerate(LINES, start=FIRST):
        for c_ in range(1, 5):
            ws.cell(r, c_, f"=Budget!{col(c_)}{r}")
        mc = ws.cell(r, 5, method)
        mc.font = INPUT
        for i in range(12):
            bcell = f"Budget!{col(MCOL + i)}{r}"
            acell = f"Actuals!{col(MCOL + i)}{r}"
            hdr = f"{col(F_M1 + i)}$3"
            f = (f'=IF({hdr}<=CurMonth,{acell},IF($E{r}="Attainment",{bcell}*${col(c_att)}{r},'
                 f'IF($E{r}="Run-rate",${col(c_rr)}{r},{bcell})))')
            ws.cell(r, F_M1 + i, f).number_format = NUM
        ws.cell(r, c_fy, f"=SUM({col(F_M1)}{r}:{col(F_M1 + 11)}{r})")
        ws.cell(r, c_bud, f"=Budget!{col(MCOL + 12)}{r}")
        ws.cell(r, c_var, f'=IF($D{r}="Revenue",{col(c_fy)}{r}-{col(c_bud)}{r},{col(c_bud)}{r}-{col(c_fy)}{r})')
        ws.cell(r, c_pct, f"=IF({col(c_bud)}{r}=0,0,{col(c_var)}{r}/ABS({col(c_bud)}{r}))")
        ws.cell(r, c_ya, f"=SUMPRODUCT({month_range('Actuals', r)}*({MONTH_HDR}<=CurMonth))")
        ws.cell(r, c_yb, f"=SUMPRODUCT({month_range('Budget', r)}*({MONTH_HDR}<=CurMonth))")
        ws.cell(r, c_att, f"=IF({col(c_yb)}{r}=0,1,{col(c_ya)}{r}/{col(c_yb)}{r})")
        ws.cell(r, c_rr, f"=SUMPRODUCT({month_range('Actuals', r)}*({MONTH_HDR}>CurMonth-3)*({MONTH_HDR}<=CurMonth))/3")
        for c_ in (c_fy, c_bud, c_var, c_ya, c_yb, c_rr):
            ws.cell(r, c_).number_format = NUM
        ws.cell(r, c_pct).number_format = PCT
        ws.cell(r, c_att).number_format = "0.000"
    widths(ws, [11, 34, 12, 10, 11] + [9] * 12 + [11, 11, 12, 9, 11, 11, 11, 11])
    ws.freeze_panes = "F4"

    # ---------- P&L Summary ----------
    ws = wb.create_sheet("P&L Summary")
    title(ws, "P&L Summary ($000s)", '="YTD through "&Inputs!B8&" and full-year forecast vs. budget"')
    header(ws, 3, ["Line", "Sign", "YTD Budget", "YTD Actual", "YTD Var Fav/(Unfav)", "YTD Var %",
                   "FY Budget", "FY Forecast", "FY Var Fav/(Unfav)", "FY Var %"])
    bva = lambda c_: f"BvA!${c_}${FIRST}:${c_}${LAST}"
    fc = lambda c_: f"Forecast!${col(c_)}${FIRST}:${col(c_)}${LAST}"
    cat_b, dept_b = bva("D"), bva("C")
    src = {3: bva("I"), 4: bva("J"), 7: fc(c_bud), 8: fc(c_fy)}

    def sumifs(rng, cat, dept=None):
        f = f'SUMIFS({rng},{cat_b},"{cat}"'
        return f + (f',{dept_b},"{dept}")' if dept else ")")

    pl_rows = [("Revenue", 1, "Revenue", None), ("COGS", -1, "COGS", None), ("Gross Profit", 1, "GP", None),
               ("Gross Margin %", 0, "GM", None)]
    pl_rows += [(f"Opex - {d}", -1, "Opex", d) for d in DEPARTMENTS]
    pl_rows += [("Total Opex", -1, "TOT", None), ("EBITDA", 1, "EBITDA", None), ("EBITDA Margin %", 0, "EM", None)]
    rowof = {}
    for r, (label, sign, kind, dept) in enumerate(pl_rows, start=4):
        rowof[kind if kind not in ("Opex",) else f"Opex-{dept}"] = r
        ws.cell(r, 1, label)
        ws.cell(r, 2, sign)
        for c_ in (3, 4, 7, 8):
            L = col(c_)
            if kind in ("Revenue", "COGS"):
                f = "=" + sumifs(src[c_], kind)
            elif kind == "Opex":
                f = "=" + sumifs(src[c_], "Opex", dept)
            elif kind == "GP":
                f = f"={L}{rowof['Revenue']}-{L}{rowof['COGS']}"
            elif kind == "GM":
                f = f"=IF({L}{rowof['Revenue']}=0,0,{L}{rowof['GP']}/{L}{rowof['Revenue']})"
            elif kind == "TOT":
                f = f"=SUM({L}{rowof['Opex-' + DEPARTMENTS[0]]}:{L}{rowof['Opex-' + DEPARTMENTS[-1]]})"
            elif kind == "EBITDA":
                f = f"={L}{rowof['GP']}-{L}{rowof['TOT']}"
            else:
                f = f"=IF({L}{rowof['Revenue']}=0,0,{L}{rowof['EBITDA']}/{L}{rowof['Revenue']})"
            ws.cell(r, c_, f).number_format = PCT if sign == 0 else NUM
        if sign != 0:
            ws.cell(r, 5, f"=B{r}*(D{r}-C{r})").number_format = NUM
            ws.cell(r, 6, f"=IF(C{r}=0,0,E{r}/ABS(C{r}))").number_format = PCT
            ws.cell(r, 9, f"=B{r}*(H{r}-G{r})").number_format = NUM
            ws.cell(r, 10, f"=IF(G{r}=0,0,I{r}/ABS(G{r}))").number_format = PCT
        else:
            ws.cell(r, 5, f"=D{r}-C{r}").number_format = PCT
            ws.cell(r, 9, f"=H{r}-G{r}").number_format = PCT
        if kind in ("GP", "TOT", "EBITDA"):
            for c_ in range(1, 11):
                ws.cell(r, c_).font = BOLD
                ws.cell(r, c_).border = TOP
    ws.column_dimensions["B"].hidden = True
    widths(ws, [22, 5, 12, 12, 14, 10, 12, 12, 14, 10])
    pl = rowof

    # ---------- Mgmt Report ----------
    ws = wb.create_sheet("Mgmt Report")
    title(ws, f"Management Report - {COMPANY}", '="FY2026 performance through "&Inputs!B8&" close"')
    header(ws, 4, ["Key metric", "Budget", "Actual / Forecast", "Var Fav/(Unfav)", "Var %"])
    metrics = [("YTD Revenue", "Revenue", 3), ("YTD Gross Margin %", "GM", 3), ("YTD EBITDA", "EBITDA", 3),
               ("FY Revenue (forecast)", "Revenue", 7), ("FY EBITDA (forecast)", "EBITDA", 7),
               ("FY EBITDA Margin % (forecast)", "EM", 7)]
    for r, (label, kind, c0) in enumerate(metrics, start=5):
        pr = pl[kind]
        ws.cell(r, 1, label)
        is_pct = kind in ("GM", "EM")
        for j, c_ in enumerate((c0, c0 + 1, c0 + 2)):
            ws.cell(r, 2 + j, f"='P&L Summary'!{col(c_)}{pr}").number_format = PCT if is_pct else NUM
        if not is_pct:
            ws.cell(r, 5, f"='P&L Summary'!{col(c0 + 3)}{pr}").number_format = PCT
    ws["A12"] = "Flagged YTD variances (from BvA)"
    ws["A12"].font = BOLD
    ws["B12"] = '=COUNTIF(BvA!M:M,"Unfavorable")&" unfavorable, "&COUNTIF(BvA!M:M,"Favorable")&" favorable"'
    ws["A14"] = "Commentary (September close)"
    ws["A14"].font = TITLE
    commentary = [
        "1. YTD EBITDA is $1,102k vs. $1,482k budget ($380k / 26% unfavorable). Full-year EBITDA is forecast at $1,715k vs. $2,150k budget.",
        "2. Safety Equipment revenue is $695k (7%) below budget YTD after a distributor contract was lost in June. Its cost of goods shows $377k favorable, but only because volume fell; the net gross profit impact is about -$318k. Action: replacement distributor pipeline and pricing review.",
        "3. Industrial Supplies revenue is ahead of budget after the April price increase and is forecast $623k favorable for the year, partly offsetting Safety Equipment.",
        "4. Freight & Logistics is $139k (14%) unfavorable YTD following a carrier rate increase in March, rising faster than shipment volume. Action: re-bid carrier contracts before Q1.",
        "5. Trade Shows & Events is $185k over budget from an unbudgeted May trade show; Professional Fees are $118k over from one-off August legal and audit work. Both are treated as non-recurring in the forecast.",
        "6. Operations and IT salaries are $189k favorable from hiring lags through July. Positions are now filled, so this favorability should not be counted on for Q4.",
    ]
    for i, text in enumerate(commentary, start=15):
        c = ws.cell(i, 1, text)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=5)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 44
    widths(ws, [34, 14, 16, 16, 10])

    # ---------- Checks ----------
    ws = wb.create_sheet("Checks")
    title(ws, "Integrity checks", "Every check must be TRUE before the report is distributed.")
    header(ws, 3, ["Check", "Result"])
    F_first, F_last = col(F_M1), col(F_M1 + 11)
    A_first, A_last = col(MCOL), col(MCOL + 11)
    checks = [
        ("Closed months in the forecast equal actuals",
         f"=ABS(SUMPRODUCT((Forecast!{F_first}3:{F_last}3<=CurMonth)*(Forecast!{F_first}{FIRST}:{F_last}{LAST}-Actuals!{A_first}{FIRST}:{A_last}{LAST})))<0.001"),
        ("No missing actuals in closed months",
         f'=SUMPRODUCT((Actuals!{A_first}3:{A_last}3<=CurMonth)*(Actuals!{A_first}{FIRST}:{A_last}{LAST}=""))=0'),
        ("No actuals entered for open months",
         f'=SUMPRODUCT((Actuals!{A_first}3:{A_last}3>CurMonth)*(Actuals!{A_first}{FIRST}:{A_last}{LAST}<>""))=0'),
        ("P&L FY budget revenue ties to Budget sheet",
         f"=ABS('P&L Summary'!G{pl['Revenue']}-SUMIFS(Budget!Q{FIRST}:Q{LAST},Budget!D{FIRST}:D{LAST},\"Revenue\"))<0.001"),
        ("P&L YTD actual EBITDA ties to line-level BvA",
         f"=ABS('P&L Summary'!D{pl['EBITDA']}-(SUMIFS(BvA!J{FIRST}:J{LAST},BvA!D{FIRST}:D{LAST},\"Revenue\")-SUMIFS(BvA!J{FIRST}:J{LAST},BvA!D{FIRST}:D{LAST},\"<>Revenue\")))<0.001"),
        ("FY forecast = YTD actual + open-month forecast (all lines)",
         f"=ABS(SUM(Forecast!{col(c_fy)}{FIRST}:{col(c_fy)}{LAST})-SUM(Forecast!{col(c_ya)}{FIRST}:{col(c_ya)}{LAST})-SUMPRODUCT((Forecast!{F_first}3:{F_last}3>CurMonth)*Forecast!{F_first}{FIRST}:{F_last}{LAST}))<0.001"),
        ("Opex by department sums to total opex lines",
         f"=ABS('P&L Summary'!D{pl['TOT']}-SUMIFS(BvA!J{FIRST}:J{LAST},BvA!D{FIRST}:D{LAST},\"Opex\"))<0.001"),
    ]
    for r, (label, f) in enumerate(checks, start=4):
        ws.cell(r, 1, label)
        ws.cell(r, 2, f)
    ws.cell(len(checks) + 5, 1, "ALL CHECKS PASS").font = BOLD
    ws.cell(len(checks) + 5, 2, f"=AND(B4:B{len(checks) + 3})").font = BOLD
    widths(ws, [58, 10])

    wb.calculation.fullCalcOnLoad = True  # Excel recalculates every formula when the file opens
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    build()
