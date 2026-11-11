"""xlsx -> list of SheetSnapshot-like dicts, mimicking what Office.js hands the add-in."""
import json, sys, warnings
from datetime import date, datetime, time
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter
warnings.filterwarnings("ignore")
EPOCH = datetime(1899, 12, 30)

def serial(v):
    if isinstance(v, datetime): return (v - EPOCH).total_seconds() / 86400
    if isinstance(v, date): return (datetime(v.year, v.month, v.day) - EPOCH).days
    if isinstance(v, time): return (v.hour * 3600 + v.minute * 60 + v.second) / 86400
    return v

def snap(path):
    out = []
    try:
        wf = openpyxl.load_workbook(path, data_only=False)
        wv = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return {"error": str(e)}
    for ws in wf.worksheets:
        if ws.sheet_state != "visible": continue
        wsv = wv[ws.title]
        pts = [(c.row, c.column) for row in ws.iter_rows() for c in row if c.value is not None]
        if not pts:
            out.append({"name": ws.title, "cells": [], "originRow": 0, "originCol": 0}); continue
        r0 = min(p[0] for p in pts); r1 = max(p[0] for p in pts)
        c0 = min(p[1] for p in pts); c1 = max(p[1] for p in pts)
        if (r1 - r0 + 1) * (c1 - c0 + 1) > 300000: r1 = r0 + 300000 // (c1 - c0 + 1)
        grid = []
        for r in range(r0, r1 + 1):
            line = []
            for c in range(c0, c1 + 1):
                cf = ws.cell(r, c); cv = wsv.cell(r, c)
                f = cf.value if isinstance(cf.value, str) and cf.value.startswith("=") else None
                if f is None and cf.data_type == "f": f = str(cf.value)
                v = cv.value if f else cf.value
                v = serial(v)
                if v is None: kind, v = "empty", ""
                elif isinstance(v, bool): kind = "boolean"
                elif isinstance(v, (int, float)): kind = "number"
                elif isinstance(v, str) and v.startswith("#") and v.rstrip("!?0/").upper() in ("#REF", "#DIV", "#N/A", "#VALUE", "#NAME", "#NUM", "#NULL", "#SPILL", "#CALC"): kind = "error"
                else: kind, v = "text", str(v)
                if f and v == "" : kind, v = "number", 0  # uncached formula result
                rgb = None
                try:
                    if cf.fill and cf.fill.fill_type == "solid" and cf.fill.fgColor is not None and cf.fill.fgColor.type == "rgb":
                        rgb = cf.fill.fgColor.rgb
                except Exception: pass
                fill = "#" + rgb[-6:].upper() if isinstance(rgb, str) and len(rgb) >= 6 and rgb[-6:].upper() != "FFFFFF" else None
                cell = {"row": r - 1, "col": c - 1, "address": f"{get_column_letter(c)}{r}", "kind": kind, "value": v, "formula": f, "numberFormat": cf.number_format}
                if fill: cell["fill"] = fill
                line.append(cell)
            grid.append(line)
        out.append({"name": ws.title, "cells": grid, "originRow": r0 - 1, "originCol": c0 - 1})
    return out

if __name__ == "__main__":
    res = {}
    for p in sys.argv[2:]:
        res[p] = snap(p)
    Path(sys.argv[1]).write_text(json.dumps(res))
    print("sheets:", sum(len(v) for v in res.values() if isinstance(v, list)), "files:", len(res))
