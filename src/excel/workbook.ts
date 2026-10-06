/* global Excel, Office */

// Everything that talks to Excel lives here. The review engine never imports Office.js.

import { toA1 } from "../review/address";
import { PyNote } from "../review/notes";
import { Cell, CellKind, SheetSnapshot } from "../review/types";

/** Reading more than this many cells on every keystroke would make Excel sluggish. */
const MAX_CELLS = 100000;
const NOTES_SETTING = "baseline.pyNotes.v1";
const NOTES_SHEET = /^py\s*(review\s*)?notes?$/i;

function kindOf(type: string): CellKind {
  switch (type) {
    case "Empty":
      return "empty";
    case "Integer":
    case "Double":
      return "number";
    case "Boolean":
      return "boolean";
    case "Error":
      return "error";
    default:
      return "text";
  }
}

export async function readActiveSheet(): Promise<SheetSnapshot> {
  return Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const used = ws.getUsedRangeOrNullObject();
    used.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await ctx.sync();

    if (used.isNullObject) {
      return { name: ws.name, cells: [], originRow: 0, originCol: 0 };
    }

    let rows = used.rowCount;
    const cols = used.columnCount;
    let truncated = false;
    if (rows * cols > MAX_CELLS) {
      rows = Math.max(1, Math.floor(MAX_CELLS / cols));
      truncated = true;
    }
    const rng = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, rows, Math.min(cols, MAX_CELLS));
    rng.load(["values", "formulas", "valueTypes"]);
    await ctx.sync();

    const cells: Cell[][] = [];
    for (let r = 0; r < rng.values.length; r++) {
      const line: Cell[] = [];
      for (let c = 0; c < rng.values[r].length; c++) {
        const row = used.rowIndex + r;
        const col = used.columnIndex + c;
        const f = rng.formulas[r][c];
        const value = rng.values[r][c];
        line.push({
          row,
          col,
          address: toA1(row, col),
          kind: kindOf(String(rng.valueTypes[r][c])),
          value: value === undefined ? null : value,
          formula: typeof f === "string" && f.charAt(0) === "=" ? f : null,
        });
      }
      cells.push(line);
    }
    return { name: ws.name, cells, originRow: used.rowIndex, originCol: used.columnIndex, truncated };
  });
}

export async function selectCell(address: string): Promise<void> {
  await Excel.run(async ctx => {
    ctx.workbook.worksheets.getActiveWorksheet().getRange(address).select();
    await ctx.sync();
  });
}

/** Re-run `onChange` whenever any sheet is edited or a different sheet is opened. */
export async function watchWorkbook(onChange: () => void): Promise<void> {
  await Excel.run(async ctx => {
    ctx.workbook.worksheets.onChanged.add(async () => onChange());
    ctx.workbook.worksheets.onActivated.add(async () => onChange());
    await ctx.sync();
  });
}

/** Rows from a worksheet called "PY Notes" (or "PY Review Notes"), or null if there isn't one. */
export async function readNotesSheet(): Promise<unknown[][] | null> {
  return Excel.run(async ctx => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const match = sheets.items.filter(s => NOTES_SHEET.test(s.name.trim()))[0];
    if (!match) return null;
    const used = match.getUsedRangeOrNullObject(true);
    used.load(["isNullObject", "values"]);
    await ctx.sync();
    return used.isNullObject ? [] : (used.values as unknown[][]);
  });
}

// PY notes are stored inside the workbook, so they travel with the file when it is rolled forward.

export function loadNotes(): PyNote[] {
  const raw = Office.context.document.settings.get(NOTES_SETTING);
  return Array.isArray(raw) ? (raw as PyNote[]) : [];
}

export function saveNotes(notes: PyNote[]): Promise<void> {
  const settings = Office.context.document.settings;
  settings.set(NOTES_SETTING, notes);
  return new Promise((resolve, reject) => {
    settings.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}
