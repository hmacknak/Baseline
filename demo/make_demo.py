"""Build demo/Baseline-demo.xlsx: a small audit workbook that shows every Baseline Review feature.

Run:  python3 demo/make_demo.py
Sheets:
  Start here         - what to do, step by step
  FY25 Cash          - bank rec (fictional client) with a short total, a hard-coded FX rate, tickmarks with
                       no legend, no source and no reviewer line (5 failing checks)
  FY25 AR            - aging and ECL allowance with a number typed over a formula, an inconsistent
                       formula and a #DIV/0! (3 failing checks)
  FY25 Fixed Assets  - PP&E additions and depreciation: a clean sheet that passes every check
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
    ("3. Expand a red item and click a cell link, e.g. C14, to jump to it.", None),
    ('4. Open "Last year\'s notes" in the add-in → Import from "PY Notes" sheet.', None),
    ('   Back on Review, the yellow box shows which of last year\'s notes are happening again.', None),
    ("   Notes Baseline can't test automatically show as \"Check yourself\".", None),
    ("5. Fix something live: change FY25 Cash C14 to =SUM(C10:C13) and watch it turn green.", None),
    ("   Also try C21 → =C19*C20 (links the FX rate), then add a Reviewed by: line in A4,", None),
    ("   a Source: line in A5 and a Tickmark legend: line in A6.", None),
    ('6. Try "FY25 AR": a typed-over total (F10), an inconsistent formula (G9) and a #DIV/0!.', None),
    ('7. "FY25 Fixed Assets" is a clean sheet that passes every check.', None),
    ("", None),
    ("The client and all figures are made up. Nothing in this file leaves Excel.", MUTED),
]
for r, (text, font) in enumerate(lines, start=1):
    c = ws.cell(row=r, column=1, value=text)
    if font:
        c.font = font
widths(ws, 95)

CLIENT = "Prairie Ridge Mining Ltd."


def money(ws, addr, value, bold=False, border=False):
    c = ws[addr]
    c.value = value
    c.number_format = MONEY
    if bold:
        c.font = BOLD
    if border:
        c.border = TOTAL_BORDER


# ─── FY25 Cash (5 failing checks) ────────────────────────────────────────────
ws = wb.create_sheet("FY25 Cash")
ws["A1"] = f"{CLIENT} · WP 4.1 Cash: bank reconciliation, 31 Dec 2025"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: reconcile GL cash (acct 1010) to the bank and test the reconciling items"
ws["A3"] = "Prepared by: JS 14 Jan 2026"
# Missing on purpose: the reviewer line (A4), a source line (A5) and the tickmark legend (A6).

header_row(ws, 7, ["Item", "Ref", "CAD", "Tick", "Comment"])
cash_rows = [
    (8, "Balance per bank statement", "BS-1", 1_284_610, "Agreed to statement p.3"),
    (9, "Add: deposits in transit", "DIT-1", 86_420, "2 deposits; cleared 2 Jan 2026"),
    (10, "Less: chq 10442 · Pioneer Drilling", "", -41_300, "Outstanding at 31 Dec 2025"),
    (11, "Less: chq 10447 · Cobalt Freight", "", -18_950, "Outstanding at 31 Dec 2025"),
    (12, "Less: chq 10451 · Wascana Safety", "", -22_775, "Outstanding at 31 Dec 2025"),
    (13, "Less: chq 10455 · SaskPower", "", -9_420, "Outstanding at 31 Dec 2025"),
    (15, "Add: bank error (fee charged in error)", "BE-1", 3_200, "Reversed by bank 3 Jan 2026"),
    (17, "Balance per GL (acct 1010)", "TB", 1_281_785, "Agreed to trial balance"),
]
for r, label, ref, amount, comment in cash_rows:
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=2, value=ref)
    money(ws, f"C{r}", amount)
    ws.cell(row=r, column=4, value="✓").alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=5, value=comment)

ws["A14"] = "Total outstanding cheques"
ws["A14"].font = BOLD
money(ws, "C14", "=SUM(C10:C12)", bold=True, border=True)  # Stops short: leaves out C13, on purpose.
ws["A16"] = "Adjusted bank balance"
ws["A16"].font = BOLD
money(ws, "C16", "=C8+C9+C14+C15", bold=True, border=True)
ws["A18"] = "Difference"
ws["A18"].font = BOLD
money(ws, "C18", "=C16-C17", bold=True, border=True)
ws["E18"] = "Should be nil"
ws["E18"].font = MUTED

ws["A19"] = "USD sub-account (USD)"
money(ws, "C19", 58_300)
ws["D19"] = "✓"
ws["A20"] = "FX rate USD→CAD"
ws["C20"] = 1.37
ws["E20"] = "per Bank of Canada 31 Dec close"
ws["A21"] = "USD sub-account (CAD)"
money(ws, "C21", "=C19*1.37")  # Hard-coded FX rate, on purpose.
widths(ws, 38, 8, 14, 6, 34)

# ─── FY25 AR (3 failing checks) ──────────────────────────────────────────────
ws = wb.create_sheet("FY25 AR")
ws["A1"] = f"{CLIENT} · WP 5.2 Trade receivables: aging and ECL allowance, 31 Dec 2025"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: test the aging for accuracy and recalculate the expected credit loss (ECL) allowance"
ws["A3"] = "Source: AR aging report run 3 Jan 2026 (client system); control account GL 1200"
ws["A4"] = "Prepared by: JS 15 Jan 2026"
ws["E4"] = "Reviewed by:"

header_row(ws, 6, ["Customer", "Current", "31-60", "61-90", "90+", "Total", "% of total"])
customers = [
    ("Northern Reach Energy", 212_400, 38_200, 0, 0),
    ("Boreal Haulage Ltd.", 96_750, 41_300, 12_900, 0),
    ("Lakeshore Fabricators", 54_300, 22_100, 31_800, 48_600),
    ("Prairie Steel Works", 143_900, 27_450, 0, 0),
    ("Wascana Industrial Supply", 38_200, 9_800, 6_350, 2_150),
    ("Meridian Pipe & Fittings", 77_600, 15_000, 0, 11_200),
    ("Tallgrass Environmental", 28_950, 0, 0, 0),
    ("Kestrel Drilling Services", 64_100, 18_900, 7_700, 0),
]
for i, (name, *buckets) in enumerate(customers):
    r = 7 + i
    ws.cell(row=r, column=1, value=name)
    for col, v in zip((2, 3, 4, 5), buckets):
        ws.cell(row=r, column=col, value=v).number_format = MONEY
    money(ws, f"F{r}", f"=SUM(B{r}:E{r})")
    ws[f"F{r}"].font = BOLD
    ws[f"G{r}"] = f"=F{r}/F$15"
    ws[f"G{r}"].number_format = PCT

ws["F10"] = 168_850  # Typed over the row formula (true total is 171,350), on purpose.
ws["F10"].number_format = MONEY
ws["F10"].font = BOLD
ws["G9"] = "=F9/F16"  # Points at the wrong total row (F16 is blank) → #DIV/0!, on purpose.
ws["G9"].number_format = PCT

ws["A15"] = "Total"
ws["A15"].font = BOLD
for col in "BCDEF":
    money(ws, f"{col}15", f"=SUM({col}7:{col}14)", bold=True, border=True)
ws["G15"] = "=SUM(G7:G14)"
ws["G15"].number_format = PCT
ws["G15"].font = BOLD
ws["G15"].border = TOTAL_BORDER

ws["A16"] = "ECL rate (policy matrix)"
ws["A16"].font = MUTED
for col, rate in zip("BCDE", (0.005, 0.02, 0.08, 0.35)):
    ws[f"{col}16"] = rate
    ws[f"{col}16"].number_format = PCT
ws["A17"] = "Collective ECL by bucket"
for col in "BCDE":
    money(ws, f"{col}17", f"={col}15*{col}16")
money(ws, "F17", "=SUM(B17:E17)", bold=True)
ws["A18"] = "Specific provision · Lakeshore (disputed invoice)"
money(ws, "F18", 20_000)
ws["A19"] = "Required allowance"
ws["A19"].font = BOLD
money(ws, "F19", "=F17+F18", bold=True, border=True)
ws["A20"] = "Allowance booked per GL (1205)"
money(ws, "F20", 54_000)
ws["A21"] = "Difference"
ws["A21"].font = BOLD
money(ws, "F21", "=F19-F20", bold=True, border=True)
widths(ws, 36, 11, 11, 11, 11, 12, 11)

# ─── FY25 Fixed Assets (clean) ───────────────────────────────────────────────
ws = wb.create_sheet("FY25 Fixed Assets")
ws["A1"] = f"{CLIENT} · WP 7.1 PP&E: FY2025 additions and depreciation recalculation"
ws["A1"].font = TITLE
ws["A2"] = "Purpose: vouch FY2025 additions and recalculate straight-line depreciation"
ws["A3"] = "Source: fixed asset register export 2 Jan 2026; vendor invoices (sample of 5)"
ws["A4"] = "Prepared by: JS 16 Jan 2026"
ws["E4"] = "Reviewed by:"
ws["A5"] = "Tickmark legend: ✓ agreed to vendor invoice, approved capex and in-service date"
ws["A6"] = "Months in period (year ended 31 Dec 2025)"
ws["B6"] = 12

header_row(ws, 7, ["Asset (FY2025 addition)", "Cost", "Life (yrs)", "In service", "Months", "Depreciation", "Net book value", "Tick"])
assets = [
    ("Haul truck · CAT 777", 1_480_000, 10, "1 Mar 2025", 10),
    ("Crusher liner set", 312_500, 4, "1 Jun 2025", 7),
    ("Survey drone & GNSS", 48_900, 5, "1 Aug 2025", 5),
    ("Site office module", 215_000, 15, "1 Jul 2025", 6),
    ("Water treatment skid", 640_000, 12, "1 Oct 2025", 3),
]
for i, (name, cost, life, since, months) in enumerate(assets):
    r = 8 + i
    ws.cell(row=r, column=1, value=name)
    money(ws, f"B{r}", cost)
    ws.cell(row=r, column=3, value=life)
    ws.cell(row=r, column=4, value=since)
    ws.cell(row=r, column=5, value=months)
    money(ws, f"F{r}", f"=B{r}/C{r}*E{r}/$B$6")
    money(ws, f"G{r}", f"=B{r}-F{r}")
    ws.cell(row=r, column=8, value="✓").alignment = Alignment(horizontal="center")

ws["A13"] = "Total additions"
ws["A13"].font = BOLD
for col in "BFG":
    money(ws, f"{col}13", f"=SUM({col}8:{col}12)", bold=True, border=True)

header_row(ws, 14, ["Cost rollforward", "CAD"])
rollforward = [
    (15, "Opening cost per prior-year WP 7.1", 18_420_000),
    (16, "Additions (above)", "=B13"),
    (17, "Disposals (WP 7.3)", -225_000),
    (18, "Closing cost", "=SUM(B15:B17)"),
    (19, "Closing cost per GL (acct 1500)", 20_891_400),
    (20, "Difference", "=B18-B19"),
]
for r, label, v in rollforward:
    ws.cell(row=r, column=1, value=label)
    money(ws, f"B{r}", v)
for r in (18, 20):
    ws[f"A{r}"].font = BOLD
    money(ws, f"B{r}", ws[f"B{r}"].value, bold=True, border=True)
widths(ws, 34, 12, 10, 11, 9, 12, 15, 6)

# ─── PY Notes ────────────────────────────────────────────────────────────────
ws = wb.create_sheet("PY Notes")
header_row(ws, 1, ["Sheet", "Cell", "Review note"])
py_notes = [
    ("FY24 Cash", "C13", "Outstanding cheques total doesn't foot: the last cheque listed is excluded from the sum range"),
    ("FY24 Cash", "C20", "FX rate hard-coded in the formula; link to the Bank of Canada rate and cite it"),
    ("FY24 Cash", "A4", "Source not documented: cite the bank statement and its date"),
    ("FY24 Cash", "", "Please sign off and add the reviewer line"),
    ("FY24 Cash", "C10", "Cheques outstanding over 90 days: assess whether stale-dated and whether to reverse"),
    ("FY24 Cash", "C9", "Deposits in transit: agree to January bank activity and document clearing dates"),
    ("FY24 AR", "F9", "Customer total overwritten with a typed number; re-link the total"),
    ("FY24 AR", "G8", "% of total formula is inconsistent: one row points at the wrong total"),
    ("FY24 AR", "E16", "Support the 35% rate for the 90+ bucket against historic write-offs"),
    ("FY24 AR", "F18", "Lakeshore: obtain management or legal support for the specific provision"),
    ("FY24 Fixed Assets", "E8", "Useful life hard-coded in the depreciation formula; link to the life column"),
    ("FY24 Fixed Assets", "A13", "Where is the tickmark legend?"),
    ("FY24 Fixed Assets", "B8", "Confirm additions meet the capitalization threshold and note who approved the capex"),
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
