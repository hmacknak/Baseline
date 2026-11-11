"""Build demo/Baseline-demo.xlsx: a small audit workbook that shows every Baseline Review feature.

Run:  python3 demo/make_demo.py
Sheets:
  Start here         - what to do, step by step
  FY25 Cash          - bank rec with a short total, a hard-coded FX rate, tickmarks with no legend,
                       no source and no reviewer line (5 failing checks)
  FY25 AR            - aging with a number typed over a formula, an inconsistent formula and a #DIV/0!
  FY25 Fixed Assets  - a clean sheet that passes every check
  AP-3 SURL          - search for unrecorded liabilities workpaper (the demo video's procedure)
  FY25 Revenue cut-off - sales recorded in the wrong year both ways, plus a post-year-end credit note
  FY25 AR confirmations - an unexplained difference, a non-response with no follow-up, an unsent confirm
  PY Notes           - last year's review notes, ready to import

The Cash sheet also carries an outstanding cheque list (stale, cleared-early and large uncleared items)
and a GL balance that the rec doesn't agree to.
"""

from datetime import date
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
DATE = "dd-mmm-yy"


def header_row(ws, row, labels):
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font = BOLD
        c.fill = HEAD_FILL


def table(ws, top, headers, rows, formats):
    """Write a header row and item rows; formats maps a 0-based column to a number format."""
    header_row(ws, top, headers)
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            if v is None:
                continue
            c = ws.cell(row=top + i, column=j + 1, value=v)
            if j in formats:
                c.number_format = formats[j]


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
    ("2. Settings tab: set year-end to 31 Dec 2025 (click 'Use it'). Save.", None),
    ('3. Click the "AP-3 SURL" tab. Baseline lists what to fix, most important first.', None),
    ("4. Expand an item to see why it matters and how to fix it. Click a cell reference to jump to it.", None),
    ("5. Use Step through: Issues to walk every problem, or Dates to check every date in order.", None),
    ("6. If a flag is fine, click Mark OK and give a reason. The reviewer can see it.", None),
    ('7. PY notes tab → Import from "PY Notes" sheet. The yellow box shows last year\'s issues happening again.', None),
    ("8. Fix something live: change AP-3 SURL F17 to =SUM(F8:F16) and watch it turn green.", None),
    ("9. Click Review all sheets to see every tab at once. FY25 Fixed Assets is ready for review.", None),
    ('10. Procedures: open "Cash confirmations" or "Expense vouching" (or "AP-3 SURL"). Baseline offers the', None),
    ("    matching reviewer. Set a threshold of 10000 in Settings to see every insight.", None),
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
ws["A3"] = "Prepared by: AB 06/10/2025"
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
ws["C14"] = "per central bank closing rate, 31 Dec"

ws["A11"] = "Balance per general ledger"
ws["B11"] = 468_475  # The rec is $1,200 off until B10 includes B9.
ws["B11"].number_format = MONEY

# Outstanding cheque detail (sums to the 51,275 above).
table(
    ws,
    17,
    ["Cheque #", "Payee", "Date", "Amount", "Cleared date", "Comment"],
    [
        ("1041", "Company A", date(2025, 5, 2), 2_400, None, "Not yet cleared"),  # stale
        ("1102", "Company B", date(2025, 12, 20), 31_500, date(2025, 12, 29), None),  # cleared before YE
        ("1110", "Company E", date(2025, 12, 30), 12_000, None, None),  # large, uncleared, no comment
        ("1111", "Company D", date(2025, 12, 31), 5_375, date(2026, 1, 6), None),
    ],
    {2: DATE, 3: MONEY, 4: DATE},
)
ws["A22"] = "Total"
ws["A22"].font = BOLD
ws["D22"] = "=SUM(D18:D21)"
ws["D22"].number_format = MONEY
ws["D22"].font = BOLD
ws["D22"].border = TOTAL_BORDER
widths(ws, 34, 18, 34, 12, 13, 16)

# ─── FY25 AR (3 failing checks) ──────────────────────────────────────────────
ws = wb.create_sheet("FY25 AR")
ws["A1"] = "Accounts receivable: aging, 31 Dec 2024"  # Not rolled forward, on purpose.
ws["A1"].font = TITLE
ws["A2"] = "Purpose: test the AR aging and allowance"
ws["A3"] = "Source: AR aging report from the client's system, 3 Jan 2026"
ws["A4"] = "Prepared by: AB 06/10/2025"
ws["C4"] = "Reviewed by:"

header_row(ws, 6, ["Customer", "0-30", "31-60", "61-90", "Total", "% of total"])
customers = [
    ("Company A", 18_200, 4_100, 900),
    ("Company B", 22_750, 0, 3_300),
    ("Company C", 9_800, 2_250, 0),
    ("Company D", 14_100, 6_400, 1_750),
    ("Company E", 7_300, 0, 0),
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
ws["A4"] = "Prepared by: AB 06/10/2025"
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
ws["A16"] = "Conclusion: additions and depreciation are fairly stated."
widths(ws, 30, 12, 16, 14, 16, 6)

# ─── AP-3 SURL (the featured procedure in the demo video) ─────────────────────
ws = wb.create_sheet("AP-3 SURL")
ws["A1"] = "Example Client Ltd.  |  AP-3  Search for unrecorded liabilities"
ws["A1"].font = TITLE
ws["A2"] = "Year ended 31 December 2025"
ws["A3"] = "Objective: identify liabilities at 31 Dec 2025 not recorded in accounts payable (completeness)."
ws["A4"] = "Source: Jan to Feb 2026 cash disbursements journal and vendor invoices. Selection: all payments over $5,000."
ws["A5"] = "Prepared by: AB 12/01/2026"
ws["E5"] = "Reviewed by:"
table(
    ws,
    7,
    ["Vendor", "Invoice #", "Inv date", "Service to", "Paid", "Amount", "In AP?", "Tick", "Conclusion"],
    [
        ("Vendor A", "INV-4471", date(2026, 1, 12), date(2025, 12, 18), date(2026, 1, 20), 18_400, "N", "✓", "Exception: Dec service, not in AP"),
        ("Vendor B", "BS-1102", date(2026, 1, 5), date(2026, 1, 31), date(2026, 1, 9), 6_000, "N", "✓", "Jan subscription, FY26"),
        ("Vendor C", "CP-88213", date(2025, 12, 28), date(2025, 12, 15), date(2026, 1, 15), 3_200, "N", "✓", "Investigate: Dec invoice"),
        ("Vendor D", "DF-5520", date(2026, 1, 8), None, date(2026, 1, 22), 12_750, "N", "✓", None),
        ("Vendor E", "EL-0931", date(2026, 1, 3), date(2025, 12, 10), date(2026, 1, 25), 9_000, "Y", "✓", "In AP listing at YE"),
        ("Vendor F", "FF-2290", date(2026, 1, 14), date(2026, 1, 10), date(2026, 1, 28), 2_150, "N", "✓", "FY26 expense"),
        ("Vendor G", "GR-7741", date(2026, 1, 20), date(2025, 12, 31), date(2026, 2, 3), 24_600, "Y", "✓", "Accrued at YE, see AP-2"),
        ("Vendor H", "HS-3310", date(2026, 2, 2), date(2026, 1, 31), date(2026, 2, 10), 7_800, "N", "✓", "FY26 expense"),
        ("Vendor I", "IC-1907", date(2026, 1, 9), date(2025, 12, 20), date(2026, 1, 30), 11_250, "N", "✓", None),
    ],
    {2: DATE, 3: DATE, 4: DATE, 5: MONEY},
)
ws["A17"] = "Total tested"
ws["A17"].font = BOLD
ws["F17"] = "=SUM(F8:F15)"  # Stops a row short: leaves out Vendor I (F16), on purpose.
ws["F17"].number_format = MONEY
ws["F17"].font = BOLD
ws["F17"].border = TOTAL_BORDER
ws["A19"] = "Population: payments over $5,000"
ws["F19"] = 1_284_600
ws["F19"].number_format = MONEY
ws["A20"] = "Coverage of population"
ws["F20"] = "=F17/1284600"  # Hard-coded population instead of a link to F19, on purpose.
ws["F20"].number_format = PCT
# Missing on purpose: tickmark legend (the ✓ marks in column H are unexplained).
ws["A22"] = "Conclusion: to be completed once exceptions are cleared."
ws["F11"].fill = PatternFill("solid", fgColor="FFFF00")  # Leftover highlight, on purpose.
ws["K12"] = 12_750  # Scratch number off to the side, on purpose.
ws["F13"] = "2,150"  # Amount stored as text, on purpose.
ws["E15"] = date(2016, 2, 10)  # Year typo in the paid date, on purpose.
ws["E15"].number_format = DATE
widths(ws, 20, 10, 10, 10, 10, 10, 7, 5, 30)

# ─── FY25 Revenue cut-off ────────────────────────────────────────────────────
ws = wb.create_sheet("FY25 Revenue cut-off")
ws["A1"] = "Revenue cut-off testing, 31 Dec 2025"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: test sales either side of year-end are recorded in the right period"
ws["A3"] = "Source: sales journal 20 Dec 2025 to 10 Jan 2026, shipping logs"
ws["A4"] = "Prepared by: AB 06/10/2025"
ws["D4"] = "Reviewed by:"
table(
    ws,
    6,
    ["Invoice #", "Customer", "Invoice date", "Ship date", "Amount", "Conclusion"],
    [
        ("INV-880", "Company D", date(2025, 12, 22), date(2025, 12, 21), 4_000, "OK"),
        ("INV-899", "Company E", date(2025, 12, 29), date(2025, 12, 29), 22_000, None),
        ("INV-901", "Company A", date(2025, 12, 31), date(2026, 1, 3), 48_000, "Shipped Jan"),
        ("INV-902", "Company B", date(2026, 1, 2), date(2025, 12, 30), 15_500, "Shipped Dec"),
        ("CN-14", "Company C", date(2026, 1, 6), date(2026, 1, 6), -9_800, "Credit note"),
        ("INV-910", "Company F", date(2026, 1, 8), date(2026, 1, 7), 7_250, "OK"),
    ],
    {2: DATE, 3: DATE, 4: MONEY},
)
widths(ws, 12, 20, 13, 13, 12, 18)

# ─── FY25 AR confirmations ───────────────────────────────────────────────────
ws = wb.create_sheet("FY25 AR confirmations")
ws["A1"] = "AR confirmations, 31 Dec 2025"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: confirm significant customer balances directly with customers"
ws["A3"] = "Source: AR aging at 31 Dec 2025, confirmation responses"
ws["A4"] = "Prepared by: AB 06/10/2025"
ws["D4"] = "Reviewed by:"
table(
    ws,
    6,
    ["Customer", "Balance per ledger", "Date sent", "Confirmed balance", "Explanation", "Alternative procedures"],
    [
        ("Company A", 23_200, date(2026, 1, 10), 23_200, None, None),
        ("Company B", 26_050, date(2026, 1, 10), 22_750, None, None),
        ("Company C", 14_050, date(2026, 1, 10), None, None, None),
        ("Company D", 22_250, date(2026, 1, 10), None, None, "Agreed to Jan receipts"),
        ("Company E", 17_300, None, None, None, None),
    ],
    {1: MONEY, 2: DATE, 3: MONEY},
)
widths(ws, 20, 17, 12, 17, 18, 22)

# ─── Cash confirmations ──────────────────────────────────────────────────────
ws = wb.create_sheet("Cash confirmations")
ws["A1"] = "Example Client Ltd.  |  C-2  Cash confirmations"
ws["A1"].font = TITLE
ws["A2"] = "Year ended 31 December 2025"
ws["A3"] = "Objective: confirm existence and completeness of all bank balances directly with each bank."
ws["A4"] = "Source: trial balance, prior-year bank replies, confirmation platform responses."
ws["A5"] = "Prepared by: AB 14/01/2026"
ws["F5"] = "Reviewed by:"
table(
    ws,
    7,
    ["Bank", "Account", "Currency", "Balance per books", "Balance per confirmation", "As at", "Date sent", "Date received", "Received via", "Explanation", "Conclusion"],
    [
        ("Bank A", "Operating", "CAD", 468_475, 468_475, date(2025, 12, 31), date(2026, 1, 5), date(2026, 1, 14), "Confirmation platform", None, "Agreed"),
        ("Bank A", "Payroll", "CAD", 36_400, 31_400, date(2025, 12, 31), date(2026, 1, 5), date(2026, 1, 14), "Confirmation platform", None, "Exception"),
        ("Bank B", "Current", "CAD", 84_260, 91_735, date(2025, 12, 31), date(2026, 1, 5), date(2026, 1, 15), "Confirmation platform", "Outstanding cheques 7,475 per bank rec C-1", "Agreed to bank rec"),
        ("Bank B", "Savings", "CAD", 125_000, None, date(2025, 12, 31), date(2026, 1, 5), None, None, None, "Chasing"),
        ("Bank C", "US dollar account", "USD", 25_000, 25_000, date(2025, 11, 30), date(2026, 1, 5), date(2026, 1, 20), "Forwarded by client controller (email)", None, "Agreed"),
        ("Bank D", "Dormant account", "CAD", 0, None, None, None, None, None, None, "Not used"),
        ("Bank A", "Credit line", "CAD", -15_000, -15_000, date(2025, 12, 31), date(2026, 1, 5), date(2026, 1, 14), "Confirmation platform", None, "Agreed"),
    ],
    {3: MONEY, 4: MONEY, 5: DATE, 6: DATE, 7: DATE},
)
ws["A15"] = "Conclusion: to be completed once differences and non-replies are cleared."
widths(ws, 10, 16, 8, 14, 14, 10, 10, 10, 18, 12, 12)

# ─── Expense vouching ────────────────────────────────────────────────────────
ws = wb.create_sheet("Expense vouching")
ws["A1"] = "Example Client Ltd.  |  E-4  Expense vouching"
ws["A1"].font = TITLE
ws["A2"] = "Year ended 31 December 2025"
ws["A3"] = "Objective: vouch a sample of operating expenses to support for occurrence, accuracy, cut-off and classification."
ws["A4"] = "Source: general ledger expense accounts; vendor invoices. Sample of 8 items."
ws["A5"] = "Prepared by: AB 20/01/2026"
ws["F5"] = "Reviewed by:"
table(
    ws,
    7,
    ["Sample #", "GL date", "Vendor", "Description", "Amount per GL", "Amount per invoice", "Invoice #", "Service date", "Approved?", "Correct account?", "Conclusion"],
    [
        ("1", date(2025, 3, 4), "Vendor A", "Office cleaning", 1_236.40, 1_236.40, "A-11", date(2025, 2, 28), "Y", "Y", "Agreed"),
        ("2", date(2025, 5, 9), "Vendor B", "Consulting", 8_500, 5_800, "B-20", date(2025, 5, 1), "Y", "Y", "Exception"),
        ("3", date(2025, 12, 30), "Vendor C", "Annual software support", 6_000, 6_000, "C-7", date(2026, 1, 15), "Y", "Y", "Agreed"),
        ("4", date(2025, 3, 13), "Vendor A", "Office cleaning", 1_236.40, 1_236.40, "A-11", date(2025, 2, 28), "Y", "Y", "Agreed"),
        ("5", date(2025, 8, 2), "Vendor D", "New forklift", 42_000, 42_000, "D-3", date(2025, 7, 30), "N", "N", "Exception"),
        ("6", date(2025, 9, 2), "Vendor E", "Courier", 300, 300, "E-9", date(2025, 9, 1), None, "Y", "Agreed"),
        ("7", date(2025, 10, 14), "Vendor F", "Utilities", 2_450, 2_450, "F-31", date(2025, 9, 30), "Y", "Y", "Agreed"),
        ("8", date(2025, 11, 20), "Vendor G", "Travel", 1_980, 1_980, "G-2", date(2025, 11, 12), "Y", "Y", None),
    ],
    {1: DATE, 4: MONEY, 5: MONEY, 7: DATE},
)
ws["A17"] = "Conclusion: to be completed once exceptions are evaluated and projected."
widths(ws, 9, 10, 10, 22, 12, 12, 9, 10, 9, 10, 12)

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
    ("AP-3 SURL", "", "Missing service dates on several items, so you can't conclude on the period"),
    ("AP-3 SURL", "", "January invoice for December services was not identified as an unrecorded liability"),
    ("AP-3 SURL", "H8", "Where is the tickmark legend?"),
    ("Cash confirmations", "", "Bank confirmation received via the client, not directly"),
    ("Expense vouching", "", "Duplicate payment not identified; vendor not queried. Follow up."),
    ("FY24 Revenue cut-off", "", "Dec 31 invoice shipped in January was not identified"),
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
