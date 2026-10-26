"""Budget and actuals for a sample company, shared by the workbook builder and the reference model.

Northwind Supply Co. is a fictional industrial distributor with three revenue lines.
All figures are in $ thousands ($000s). The FY2026 budget covers Jan-Dec; actuals are
closed through September (month 9). Every number is generated from a fixed seed, so the
data is illustrative, not from a real company.
"""
import numpy as np

COMPANY = "Northwind Supply Co."
FISCAL_YEAR = 2026
CURRENT_MONTH = 9          # last closed month
VAR_PCT_THRESHOLD = 0.05   # flag variances of at least 5% ...
VAR_AMT_THRESHOLD = 25.0   # ... and at least $25k
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

SEASON = np.array([0.075, 0.075, 0.085, 0.085, 0.09, 0.09, 0.08, 0.08, 0.085, 0.09, 0.085, 0.08])
SEASON = SEASON / SEASON.sum()
FLAT = np.full(12, 1 / 12)
TRADE_SHOWS = np.array([0.05, 0.05, 0.25, 0.05, 0.05, 0.05, 0.05, 0.05, 0.25, 0.05, 0.05, 0.05])

# id, line item, department, category, forecast method
LINES = [
    ("REV-A", "Industrial Supplies", "Sales", "Revenue", "Attainment"),
    ("REV-B", "Safety Equipment", "Sales", "Revenue", "Attainment"),
    ("REV-C", "Services & Installation", "Sales", "Revenue", "Attainment"),
    ("COGS-A", "Cost of Goods - Industrial Supplies", "Operations", "COGS", "Attainment"),
    ("COGS-B", "Cost of Goods - Safety Equipment", "Operations", "COGS", "Attainment"),
    ("COGS-C", "Cost of Services & Installation", "Operations", "COGS", "Attainment"),
    ("COGS-FRT", "Freight & Logistics", "Operations", "COGS", "Attainment"),
    ("SAL-SALES", "Salaries & Benefits", "Sales", "Opex", "Run-rate"),
    ("COMM-SALES", "Sales Commissions", "Sales", "Opex", "Attainment"),
    ("TE-SALES", "Travel & Entertainment", "Sales", "Opex", "Run-rate"),
    ("SAL-MKT", "Salaries & Benefits", "Marketing", "Opex", "Run-rate"),
    ("ADS-MKT", "Digital Advertising", "Marketing", "Opex", "Run-rate"),
    ("EVT-MKT", "Trade Shows & Events", "Marketing", "Opex", "Budget"),
    ("SAL-OPS", "Salaries & Benefits", "Operations", "Opex", "Run-rate"),
    ("RENT-OPS", "Warehouse Rent", "Operations", "Opex", "Run-rate"),
    ("MAINT-OPS", "Equipment Maintenance", "Operations", "Opex", "Run-rate"),
    ("SAL-GA", "Salaries & Benefits", "G&A", "Opex", "Run-rate"),
    ("PROF-GA", "Professional Fees", "G&A", "Opex", "Budget"),
    ("INS-GA", "Insurance", "G&A", "Opex", "Run-rate"),
    ("SAL-IT", "Salaries & Benefits", "IT", "Opex", "Run-rate"),
    ("SW-IT", "Software & Subscriptions", "IT", "Opex", "Run-rate"),
]
DEPARTMENTS = ["Sales", "Marketing", "Operations", "G&A", "IT"]


def budget():
    """FY2026 monthly budget by line, $000s."""
    b = {}
    b["REV-A"] = 24000 * SEASON
    b["REV-B"] = 14000 * SEASON
    b["REV-C"] = 10000 * SEASON
    b["COGS-A"] = 0.66 * b["REV-A"]
    b["COGS-B"] = 0.60 * b["REV-B"]
    b["COGS-C"] = 0.45 * b["REV-C"]
    b["COGS-FRT"] = 0.035 * (b["REV-A"] + b["REV-B"])
    b["SAL-SALES"] = 2400 * FLAT
    b["COMM-SALES"] = 0.03 * (b["REV-A"] + b["REV-B"] + b["REV-C"])
    b["TE-SALES"] = 360 * FLAT
    b["SAL-MKT"] = 960 * FLAT
    b["ADS-MKT"] = 720 * FLAT
    b["EVT-MKT"] = 480 * TRADE_SHOWS
    b["SAL-OPS"] = 3000 * FLAT
    b["RENT-OPS"] = 1440 * FLAT
    b["MAINT-OPS"] = 600 * FLAT
    b["SAL-GA"] = 1800 * FLAT
    b["PROF-GA"] = 480 * FLAT
    b["INS-GA"] = 360 * FLAT
    b["SAL-IT"] = 1200 * FLAT
    b["SW-IT"] = 540 * FLAT
    return {k: np.round(v, 1) for k, v in b.items()}


def actuals():
    """Closed actuals Jan-Sep, $000s. Months 10-12 are NaN (not yet closed).

    Built-in business events the analysis should find:
      - Price increase on Industrial Supplies from April (+4%).
      - Lost distributor contract on Safety Equipment from June (-14%).
      - Carrier rate increase on freight from March (+18% per unit shipped).
      - Extra trade show in May (+$180k).
      - Hiring lag in Operations and IT through July (-8% salaries).
      - One-off legal and audit fees in August (+$120k).
    """
    rng = np.random.default_rng(2026)
    b = budget()
    m = np.arange(1, 13)
    closed = m <= CURRENT_MONTH

    def noise(sd=0.015):
        return 1 + rng.normal(0, sd, 12)

    a = {}
    a["REV-A"] = b["REV-A"] * np.where(m >= 4, 1.04, 1.0) * noise()
    a["REV-B"] = b["REV-B"] * np.where(m >= 6, 0.86, 1.0) * noise()
    a["REV-C"] = b["REV-C"] * noise()
    a["COGS-A"] = 0.66 * a["REV-A"] * noise(0.008)
    a["COGS-B"] = 0.60 * a["REV-B"] * noise(0.008)
    a["COGS-C"] = 0.45 * a["REV-C"] * noise(0.008)
    a["COGS-FRT"] = 0.035 * (a["REV-A"] + a["REV-B"]) * np.where(m >= 3, 1.18, 1.0) * noise(0.01)
    lag = np.where(m <= 7, 0.92, 1.0)
    a["SAL-SALES"] = b["SAL-SALES"] * noise(0.005)
    a["COMM-SALES"] = 0.03 * (a["REV-A"] + a["REV-B"] + a["REV-C"]) * noise(0.01)
    a["TE-SALES"] = b["TE-SALES"] * noise(0.05)
    a["SAL-MKT"] = b["SAL-MKT"] * noise(0.005)
    a["ADS-MKT"] = b["ADS-MKT"] * noise(0.04)
    a["EVT-MKT"] = b["EVT-MKT"] * noise(0.03) + np.where(m == 5, 180.0, 0.0)
    a["SAL-OPS"] = b["SAL-OPS"] * lag * noise(0.005)
    a["RENT-OPS"] = b["RENT-OPS"].copy()
    a["MAINT-OPS"] = b["MAINT-OPS"] * noise(0.06)
    a["SAL-GA"] = b["SAL-GA"] * noise(0.005)
    a["PROF-GA"] = b["PROF-GA"] * noise(0.05) + np.where(m == 8, 120.0, 0.0)
    a["INS-GA"] = b["INS-GA"].copy()
    a["SAL-IT"] = b["SAL-IT"] * lag * noise(0.005)
    a["SW-IT"] = b["SW-IT"] * noise(0.02)
    return {k: np.where(closed, np.round(v, 1), np.nan) for k, v in a.items()}
