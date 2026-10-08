"""Export demo/Baseline-demo.xlsx to demo/video/workbook.json for the video studio."""

import json
from datetime import date, datetime
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
EPOCH = datetime(1899, 12, 30)


def plain(v):
    """Dates become Excel serial numbers, the way Office.js returns them."""
    if isinstance(v, datetime):
        return (v - EPOCH).days
    if isinstance(v, date):
        return (datetime(v.year, v.month, v.day) - EPOCH).days
    return v

wb = openpyxl.load_workbook(HERE.parent / "Baseline-demo.xlsx")
out = {"order": [], "sheets": {}}
for ws in wb.worksheets:
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            filled = bool(c.fill and c.fill.fgColor and c.fill.fgColor.rgb not in (None, "00000000"))
            cells[c.coordinate] = {"v": plain(c.value), "b": bool(c.font and c.font.b), "fmt": c.number_format, "fill": filled}
    out["order"].append(ws.title)
    out["sheets"][ws.title] = {
        "cells": cells,
        "widths": {k: v.width for k, v in ws.column_dimensions.items() if v.width},
    }
(HERE / "workbook.json").write_text(json.dumps(out))
print("wrote", HERE / "workbook.json")
