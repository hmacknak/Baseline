/* global Excel, Office */

// Everything that talks to Excel lives here. The review engine never imports Office.js.

import { toA1 } from "../review/address";
import { PyNote } from "../review/notes";
import { Cell, CellKind, ReviewContext, SheetSnapshot } from "../review/types";

/** Reading more than this many cells on every keystroke would make Excel sluggish. */
const MAX_CELLS = 100000;
const NOTES_SETTING = "baseline.pyNotes.v1";
const CONTEXT_SETTING = "baseline.context.v1";
const ASSIGNMENTS_SETTING = "baseline.assignments.v1";
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

/** Cells whose fill colours we read; above this, highlighting checks are skipped to keep Excel responsive. */
const MAX_FILL_CELLS = 20000;

function normaliseColour(c: unknown): string | undefined {
  if (typeof c !== "string") return undefined;
  const hex = c.trim().toUpperCase();
  // Excel reports no fill as white; treat white as unshaded.
  if (!/^#[0-9A-F]{6}$/.test(hex) || hex === "#FFFFFF") return undefined;
  return hex;
}

async function snapshotOf(ctx: Excel.RequestContext, ws: Excel.Worksheet): Promise<SheetSnapshot> {
  ws.load("name");
  // Values only: cells that are merely formatted (a border dragged down 10,000 rows) don't count as used.
  // Otherwise a small workpaper can look enormous, get cut off, and lose its highlighting check.
  const used = ws.getUsedRangeOrNullObject(true);
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
  rng.load(["values", "formulas", "valueTypes", "numberFormat"]);

  // Fill colours need ExcelApi 1.9; older Excel simply skips the highlighting check.
  let fills: OfficeExtension.ClientResult<Excel.CellProperties[][]> | null = null;
  if (rows * cols <= MAX_FILL_CELLS) {
    try {
      fills = rng.getCellProperties({ format: { fill: { color: true } } });
    } catch {
      fills = null;
    }
  }
  try {
    await ctx.sync();
  } catch (e) {
    if (!fills) throw e;
    // Retry without fill colours if this Excel can't provide them.
    fills = null;
    rng.load(["values", "formulas", "valueTypes", "numberFormat"]);
    await ctx.sync();
  }

  const cells: Cell[][] = [];
  for (let r = 0; r < rng.values.length; r++) {
    const line: Cell[] = [];
    for (let c = 0; c < rng.values[r].length; c++) {
      const row = used.rowIndex + r;
      const col = used.columnIndex + c;
      const f = rng.formulas[r][c];
      const value = rng.values[r][c];
      const fmt = rng.numberFormat && rng.numberFormat[r] ? rng.numberFormat[r][c] : undefined;
      const props = fills && fills.value && fills.value[r] ? fills.value[r][c] : undefined;
      line.push({
        row,
        col,
        address: toA1(row, col),
        kind: kindOf(String(rng.valueTypes[r][c])),
        value: value === undefined ? null : value,
        formula: typeof f === "string" && f.charAt(0) === "=" ? f : null,
        numberFormat: typeof fmt === "string" ? fmt : undefined,
        fill: normaliseColour(props && props.format && props.format.fill ? props.format.fill.color : undefined),
      });
    }
    cells.push(line);
  }
  return { name: ws.name, cells, originRow: used.rowIndex, originCol: used.columnIndex, truncated };
}

export async function readActiveSheet(): Promise<SheetSnapshot> {
  return Excel.run(async ctx => snapshotOf(ctx, ctx.workbook.worksheets.getActiveWorksheet()));
}

/** Every visible worksheet, in tab order, for the whole-workbook review. */
export async function readAllSheets(): Promise<SheetSnapshot[]> {
  return Excel.run(async ctx => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await ctx.sync();
    const out: SheetSnapshot[] = [];
    for (const ws of sheets.items) {
      if (ws.visibility !== Excel.SheetVisibility.visible) continue;
      out.push(await snapshotOf(ctx, ws));
    }
    return out;
  });
}

export async function activateSheet(name: string, address?: string): Promise<void> {
  await Excel.run(async ctx => {
    const ws = ctx.workbook.worksheets.getItem(name);
    ws.activate();
    if (address) ws.getRange(address).select();
    await ctx.sync();
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

// Settings are stored inside the workbook, so they travel with the file when it is rolled forward.

export function loadSetting<T>(key: string): T | undefined {
  const raw = Office.context.document.settings.get(key);
  return raw === null || raw === undefined ? undefined : (raw as T);
}

export function saveSetting(key: string, value: unknown): Promise<void> {
  const settings = Office.context.document.settings;
  settings.set(key, value);
  return new Promise((resolve, reject) => {
    settings.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(result.error);
    });
  });
}

export function loadNotes(): PyNote[] {
  const raw = loadSetting<PyNote[]>(NOTES_SETTING);
  return Array.isArray(raw) ? raw : [];
}

export function saveNotes(notes: PyNote[]): Promise<void> {
  return saveSetting(NOTES_SETTING, notes);
}

/** Year-end and threshold for this engagement. */
export function loadContext(): ReviewContext {
  return loadSetting<ReviewContext>(CONTEXT_SETTING) || {};
}

export function saveContext(ctx: ReviewContext): Promise<void> {
  return saveSetting(CONTEXT_SETTING, ctx);
}

/** Which procedure reviewer each sheet uses ("none" = the junior declined the suggestion). */
export function loadAssignments(): Record<string, string> {
  return loadSetting<Record<string, string>>(ASSIGNMENTS_SETTING) || {};
}

export function saveAssignments(map: Record<string, string>): Promise<void> {
  return saveSetting(ASSIGNMENTS_SETTING, map);
}
