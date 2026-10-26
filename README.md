# FP&A Budget vs. Actual & Rolling Forecast Model

**Portfolio project | Financial Analyst / FP&A**

An Excel FP&A model for a fictional industrial distributor, Northwind Supply Co. It covers the monthly close cycle an FP&A analyst runs: an annual budget by department and line item, monthly and year-to-date **budget-to-actual reporting** with automatic variance flags, a **rolling 9+3 forecast**, a revenue-to-EBITDA **P&L summary**, and a **management report** with variance commentary.

Every number is a live Excel formula driven by the Budget, Actuals and Inputs sheets. To close a new month, enter its actuals and change one input cell. A separate Python implementation recalculates the model, and a script checks that the two agree on every line.

> **Illustrative data.** Northwind Supply Co. and all figures are fictional, generated from a fixed seed with realistic business events built in. Figures are in $ thousands.

📥 **[Download the workbook](model/FPA_Budget_Variance_Forecast_Model.xlsx)** (open in Excel; it recalculates on open)

---

## Results at September close

| $000s | YTD Budget | YTD Actual | Fav/(Unfav) | FY Budget | FY Forecast | Fav/(Unfav) |
|---|---|---|---|---|---|---|
| Revenue | 35,760 | 35,564 | (196) | 48,000 | 47,737 | (263) |
| Gross Profit | 13,358 | 13,079 | (279) | 17,930 | 17,556 | (374) |
| Gross Margin % | 37.4% | 36.8% | (0.6 pts) | 37.4% | 36.8% | (0.6 pts) |
| Total Opex | 11,876 | 11,977 | (101) | 15,780 | 15,840 | (60) |
| **EBITDA** | **1,482** | **1,102** | **(380)** | **2,150** | **1,715** | **(434)** |

### Variance commentary (from the Mgmt Report sheet)

1. **Safety Equipment revenue: −$695k (−7%) YTD** after a distributor contract was lost in June. Its cost of goods shows **+$377k "favorable"**, but only because volume fell, so the net gross profit impact is about −$318k. A line-by-line view alone would understate the problem.
2. **Industrial Supplies** is ahead of budget after the April price increase, forecast **+$623k** for the year.
3. **Freight: −$139k (−14%)** from a March carrier rate increase that is rising faster than volume.
4. **One-offs:** an unbudgeted May trade show (−$185k) and August legal and audit fees (−$118k). The forecast treats both as non-recurring (Budget method).
5. **Hiring lag:** Operations and IT salaries are **+$189k** favorable through July. The roles are now filled, so Q4 shouldn't count on this.

Seven of 21 lines are flagged automatically: 4 unfavorable and 3 favorable, using a variance of at least 5% *and* at least $25k.

## Workbook structure

| Sheet | What it does |
|---|---|
| **Cover** | Purpose, how to close a month, conventions, sheet guide |
| **Inputs** | Last closed month and the variance-flag thresholds (% and $), as named ranges |
| **Budget** | FY2026 budget: 21 line items × 12 months by department and category (Revenue, COGS, Opex) |
| **Actuals** | Closed actuals from the general ledger; open months left blank |
| **BvA** | Current-month and YTD budget vs. actual by line: favorable/(unfavorable) variance, variance %, automatic flag with conditional formatting |
| **Forecast** | Rolling forecast: closed months = actuals; open months by each line's method (editable) |
| **P&L Summary** | Revenue, COGS, gross profit and margin, opex by department, EBITDA and margin: YTD and FY |
| **Mgmt Report** | Key metrics linked live to the P&L, flagged-variance count, written commentary |
| **Checks** | 7 integrity checks rolled up to one ALL CHECKS PASS cell |

## Methodology

- **Variance sign:** shown as favorable (+) or unfavorable (−) for both revenue and costs, so a negative number always means "worse than budget".
- **Flags:** a line is flagged when its YTD variance meets *both* thresholds (default 5% and $25k), so small lines don't create noise.
- **Forecast methods** (set per line on the Forecast sheet):
  - **Attainment:** budget × YTD actual ÷ YTD budget. Used for revenue, cost of goods, freight and commissions, which scale with volume.
  - **Run-rate:** average of the last 3 closed months. Used for steady costs like salaries, rent and software.
  - **Budget:** used for lumpy or one-off lines (events, professional fees) so that one-time spikes aren't extrapolated.
- **Integrity checks:**
  - closed-month forecast equals actuals;
  - no missing actuals in closed months and none entered for open months;
  - P&L ties to the Budget sheet and to line-level BvA;
  - FY forecast = YTD actual + open-month forecast;
  - departments sum to total opex.

## Verification

`tools/reference_model.py` is an independent pandas implementation. `tools/recalc_and_verify.sh` recalculates the workbook headless in LibreOffice and compares every line's month and YTD figures, variance flags, all 12 forecast months, every P&L cell, and the Checks sheet:

```
Lines: 21  P&L rows: 12  Checks sheet ALL PASS = True
ALL MATCH
```

## Reproduce

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
(cd tools && ../.venv/bin/python build_workbook.py)     # writes the workbook
(cd tools && ../.venv/bin/python reference_model.py)    # Python results
tools/recalc_and_verify.sh                              # needs LibreOffice
```

The budget, actuals and thresholds live in `tools/assumptions.py` and feed both the workbook and the reference model.

---

Built by [Kshama Upadhyay](https://github.com/Kshama16).
