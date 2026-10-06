"""Build demo/Baseline-demo.xlsx: a small audit workbook that shows every Baseline Review feature.

Run:  python3 demo/make_demo.py
Sheets:
  Start here         - what to do, step by step
  FY25 Cash          - bank rec with a short total, a hard-coded FX rate, tickmarks with no legend,
                       no source and no reviewer line (5 failing checks)
  FY25 AR            - aging with a number typed over a formula, an inconsistent formula and a #DIV/0!
  FY25 Fixed Assets  - a clean sheet that passes every check
  PY Notes           - last year's review notes, ready to import
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.properties import CalcProperties

OUT = Path(__file__).with_name("Baseline-demo.xlsx")

BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
MUTED = Font(italic=True, color="666666")
HEAD_FILL = PatternFill("solid", fgColor="E8F1EE")
TOTAL_BORDER = Border(top=Side(style="thin"), bottom=Side(style="double"))
MONEY = "#,##0;(#,##0)"
PCT = "0.0%"


def header_row(ws, row, labels):
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font = BOLD
        c.fill = HEAD_FILL


def widths(ws, *w):
    for i, width in enumerate(w):
        ws.column_dimensions[chr(65 + i)].width = width


wb = Workbook()

# ─── Start here ──────────────────────────────────────────────────────────────
ws = wb.active
ws.title = "Start here"
lines = [
    ("Baseline Review: demo workbook", TITLE),
    ("", None),
    ("1. Open the add-in: Home → Review.", None),
    ('2. Click the "FY25 Cash" tab. The checklist shows 5 problems.', None),
    ("3. Expand a red item and click a cell link, e.g. B10, to jump to it.", None),
    ('4. Open "Last year\'s notes" in the add-in → Import from "PY Notes" sheet.', None),
    ('   Back on Review, the yellow box shows which of last year\'s notes are happening again.', None),
    ("5. Fix something live: change FY25 Cash B10 to =SUM(B6:B9) and watch it turn green.", None),
    ("   Also try B13 → =B12*B14 (links the FX rate) and add 'Reviewed by:' in A4.", None),
    ('6. Try "FY25 AR": a typed-over formula, an inconsistent formula and a #DIV/0!.', None),
    ('7. "FY25 Fixed Assets" is a clean sheet that passes every check.', None),
    ("", None),
    ("All figures are made up. Nothing in this file leaves Excel.", MUTED),
]
for r, (text, font) in enumerate(lines, start=1):
    c = ws.cell(row=r, column=1, value=text)
    if font:
        c.font = font
widths(ws, 95)

# ─── FY25 Cash (5 failing checks) ────────────────────────────────────────────
ws = wb.create_sheet("FY25 Cash")
ws["A1"] = "Cash: bank reconciliation, 31 Dec 2025"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: reconcile the operating account balance to the bank statement"
ws["A3"] = "Prepared by: JS 06/10/2025"
# Missing on purpose: "Reviewed by" line and a "Source" line.

header_row(ws, 5, ["Item", "Amount", "Tick"])
items = [
    ("Balance per bank statement", 482_150),
    ("Deposits in transit", 36_400),
    ("Outstanding cheques", -51_275),
    ("Bank error correction", 1_200),
]
for i, (label, amount) in enumerate(items):
    r = 6 + i
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=2, value=amount).number_format = MONEY
ws["C6"] = "✓"
ws["C7"] = "✓"
ws["C8"] = "✓"  # Tickmarks with no legend, on purpose.

ws["A10"] = "Adjusted bank balance"
ws["A10"].font = BOLD
ws["B10"] = "=SUM(B6:B8)"  # Stops short: leaves out B9, on purpose.
ws["B10"].number_format = MONEY
ws["B10"].font = BOLD
ws["B10"].border = TOTAL_BORDER

ws["A12"] = "USD sub-account (USD)"
ws["B12"] = 25_000
ws["B12"].number_format = MONEY
ws["A13"] = "USD sub-account (CAD)"
ws["B13"] = "=B12*1.37"  # Hard-coded FX rate, on purpose.
ws["B13"].number_format = MONEY
ws["A14"] = "FX rate USD→CAD"
ws["B14"] = 1.37
ws["C14"] = "per Bank of Canada 31 Dec close"
widths(ws, 34, 14, 34)

# ─── FY25 AR (3 failing checks) ──────────────────────────────────────────────
ws = wb.create_sheet("FY25 AR")
ws["A1"] = "Accounts receivable: aging, 31 Dec 2025"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: test the AR aging and allowance"
ws["A3"] = "Source: AR aging report from the client's system, 3 Jan 2026"
ws["A4"] = "Prepared by: JS 06/10/2025"
ws["C4"] = "Reviewed by:"

header_row(ws, 6, ["Customer", "0-30", "31-60", "61-90", "Total", "% of total"])
customers = [
    ("Northwind Traders", 18_200, 4_100, 900),
    ("Contoso Ltd", 22_750, 0, 3_300),
    ("Fabrikam Inc", 9_800, 2_250, 0),
    ("Adventure Works", 14_100, 6_400, 1_750),
    ("Tailspin Toys", 7_300, 0, 0),
]
for i, (name, a, b, c) in enumerate(customers):
    r = 7 + i
    ws.cell(row=r, column=1, value=name)
    for col, v in zip((2, 3, 4), (a, b, c)):
        ws.cell(row=r, column=col, value=v).number_format = MONEY
    ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})").number_format = MONEY
    ws.cell(row=r, column=6, value=f"=E{r}/$E$12").number_format = PCT

ws["E9"] = 14_050  # Typed over the row formula (true total is 12,050), on purpose.
ws["E9"].number_format = MONEY
ws["F10"] = "=E10/$E$13"  # Points at the wrong total row, on purpose.
ws["F10"].number_format = PCT

ws["A12"] = "Total"
ws["A12"].font = BOLD
for col in "BCDE":
    c = ws[f"{col}12"]
    c.value = f"=SUM({col}7:{col}11)"
    c.number_format = MONEY
    c.font = BOLD
    c.border = TOTAL_BORDER

ws["A15"] = "Allowance coverage"
ws["B15"] = "=E12/B16"  # B16 (allowance) left blank → #DIV/0!, on purpose.
ws["A16"] = "Allowance for doubtful accounts"
widths(ws, 30, 12, 12, 12, 14, 12)

# ─── FY25 Fixed Assets (clean) ───────────────────────────────────────────────
ws = wb.create_sheet("FY25 Fixed Assets")
ws["A1"] = "Fixed assets: additions and depreciation, FY2025"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: recalculate straight-line depreciation on FY25 additions"
ws["A3"] = "Source: fixed asset register export, 2 Jan 2026"
ws["A4"] = "Prepared by: JS 06/10/2025"
ws["C4"] = "Reviewed by:"

header_row(ws, 6, ["Asset", "Cost", "Useful life (yrs)", "Depreciation", "Net book value", "Tick"])
assets = [
    ("Delivery van", 48_000, 6),
    ("Laptops (12)", 21_600, 3),
    ("Warehouse racking", 35_500, 10),
    ("Forklift", 29_900, 8),
]
for i, (name, cost, life) in enumerate(assets):
    r = 7 + i
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=cost).number_format = MONEY
    ws.cell(row=r, column=3, value=life)
    ws.cell(row=r, column=4, value=f"=B{r}/C{r}").number_format = MONEY
    ws.cell(row=r, column=5, value=f"=B{r}-D{r}").number_format = MONEY
    ws.cell(row=r, column=6, value="✓")

ws["A11"] = "Total"
ws["A11"].font = BOLD
for col in "BDE":
    c = ws[f"{col}11"]
    c.value = f"=SUM({col}7:{col}10)"
    c.number_format = MONEY
    c.font = BOLD
    c.border = TOTAL_BORDER

ws["A13"] = "Tickmark legend"
ws["A13"].font = BOLD
ws["A14"] = "✓ = agreed cost to invoice and useful life to policy"
widths(ws, 30, 12, 16, 14, 16, 6)

# ─── PY Notes ────────────────────────────────────────────────────────────────
ws = wb.create_sheet("PY Notes")
header_row(ws, 1, ["Sheet", "Cell", "Review note"])
py_notes = [
    ("FY24 Cash", "B10", "Total doesn't foot: adjusted balance excludes the last reconciling item"),
    ("FY24 Cash", "B13", "Hard-coded FX rate, link to source"),
    ("FY24 Cash", "", "Please sign off and add the reviewer line"),
    ("FY24 AR", "E9", "Formula overwritten with a typed number, re-link the total"),
    ("FY24 AR", "", "Discuss allowance methodology with manager"),
    ("FY24 Fixed Assets", "", "Where is the tickmark legend?"),
    ("FY24 Fixed Assets", "D8", "Hard-coded useful life in depreciation formula"),
]
for r, row in enumerate(py_notes, start=2):
    for c, v in enumerate(row, start=1):
        ws.cell(row=r, column=c, value=v)
for row in ws.iter_rows(min_row=2, max_col=3):
    row[2].alignment = Alignment(wrap_text=True)
widths(ws, 20, 8, 70)

wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.active = 0
wb.save(OUT)
print(f"wrote {OUT}")
