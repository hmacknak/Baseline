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


def rgb_of(colour):
    rgb = colour.rgb if colour is not None and colour.type == "rgb" else None
    return "#" + rgb[-6:].upper() if isinstance(rgb, str) and len(rgb) >= 6 else None


def look(c):
    """How the cell is formatted: fill, font colour, italics, size, borders, alignment."""
    out = {}
    if c.fill and c.fill.fill_type == "solid":
        fill = rgb_of(c.fill.fgColor)
        if fill:
            out["fill"] = fill
    if c.font:
        colour = rgb_of(c.font.color)
        if colour and colour != "#000000":
            out["color"] = colour
        if c.font.i:
            out["i"] = True
        if c.font.sz and c.font.sz != 11:
            out["sz"] = c.font.sz
    sides = {}
    for key, name in (("t", "top"), ("b", "bottom"), ("l", "left"), ("r", "right")):
        side = getattr(c.border, name, None) if c.border else None
        if side is not None and side.style:
            sides[key] = side.style
    if sides:
        out["bd"] = sides
    if c.alignment and c.alignment.horizontal in ("center", "right"):
        out["al"] = c.alignment.horizontal
    return out


wb = openpyxl.load_workbook(HERE.parent / "Baseline-demo.xlsx")
out = {"order": [], "sheets": {}}
for ws in wb.worksheets:
    cells = {}
    styles = {}  # Formatting on cells with nothing in them, e.g. the empty boxes of a testing table.
    for row in ws.iter_rows():
        for c in row:
            style = look(c)
            if c.value is None:
                if style:
                    styles[c.coordinate] = style
                continue
            cells[c.coordinate] = {"v": plain(c.value), "b": bool(c.font and c.font.b), "fmt": c.number_format, "fill": None, **style}
    out["order"].append(ws.title)
    out["sheets"][ws.title] = {
        "cells": cells,
        "styles": styles,
        "widths": {k: v.width for k, v in ws.column_dimensions.items() if v.width},
    }
(HERE / "workbook.json").write_text(json.dumps(out))
print("wrote", HERE / "workbook.json")
