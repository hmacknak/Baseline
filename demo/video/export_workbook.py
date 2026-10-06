"""Export demo/Baseline-demo.xlsx to demo/video/workbook.json for the video studio."""

import json
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
wb = openpyxl.load_workbook(HERE.parent / "Baseline-demo.xlsx")
out = {"order": [], "sheets": {}}
for ws in wb.worksheets:
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            filled = bool(c.fill and c.fill.fgColor and c.fill.fgColor.rgb not in (None, "00000000"))
            cells[c.coordinate] = {"v": c.value, "b": bool(c.font and c.font.b), "fmt": c.number_format, "fill": filled}
    out["order"].append(ws.title)
    out["sheets"][ws.title] = {
        "cells": cells,
        "widths": {k: v.width for k, v in ws.column_dimensions.items() if v.width},
    }
(HERE / "workbook.json").write_text(json.dumps(out))
print("wrote", HERE / "workbook.json")
