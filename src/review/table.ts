// Reads the testing table on a procedure sheet: finds the header row, maps columns by their
// many possible names, and returns one record per tested item with cell addresses for jumping.

import { toAmount, toSerial } from "./dates";
import { Cell, SheetSnapshot } from "./types";

export type ColumnKind = "text" | "date" | "amount" | "flag";

export interface ColumnSpec {
  key: string;
  /** Shown to the junior, e.g. "invoice date". */
  label: string;
  kind: ColumnKind;
  /** Header texts that mean this column, lower case. Matched as whole header or as a phrase inside it. */
  names: string[];
}

export interface TableRow {
  /** 0-based worksheet row. */
  row: number;
  cells: Record<string, Cell | undefined>;
}

export interface Table {
  headerRow: number;
  columns: Record<string, number>;
  rows: TableRow[];
}

function norm(s: string): string {
  return s
    .toLowerCase()
    .replace(/[_\n\r]/g, " ")
    .replace(/[^a-z0-9#$?/ ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function headerMatches(header: string, names: string[]): number {
  const h = norm(header);
  if (!h) return 0;
  let best = 0;
  for (const n of names) {
    const name = norm(n);
    if (h === name) return 3;
    if ((" " + h + " ").indexOf(" " + name + " ") >= 0) best = Math.max(best, 2);
  }
  return best;
}

function rowOf(sheet: SheetSnapshot, r: number): Cell[] {
  return sheet.cells[r - sheet.originRow] || [];
}

export function isBlank(c: Cell | undefined): boolean {
  return !c || c.kind === "empty" || (typeof c.value === "string" && c.value.trim() === "");
}

/** Map columns on one candidate header row; each sheet column is used at most once, best match first. */
function mapRow(cells: Cell[], specs: ColumnSpec[]): Record<string, number> {
  const scored: Array<{ key: string; col: number; score: number }> = [];
  for (const c of cells) {
    if (!c || c.kind !== "text") continue;
    for (const spec of specs) {
      const score = headerMatches(String(c.value), spec.names);
      if (score) scored.push({ key: spec.key, col: c.col, score });
    }
  }
  scored.sort((a, b) => b.score - a.score || a.col - b.col);
  const columns: Record<string, number> = {};
  const usedCols = new Set<number>();
  for (const s of scored) {
    if (columns[s.key] !== undefined || usedCols.has(s.col)) continue;
    columns[s.key] = s.col;
    usedCols.add(s.col);
  }
  return columns;
}

const TOTAL_ROW = /^\s*(sub)?total\b|^\s*grand total/i;
const SCAN_ROWS = 40;

/**
 * Find the testing table. The header row is the row (within the first 40) that maps the most
 * columns; at least `minColumns` must match. Item rows run until two blank rows or a "Total" row.
 */
export function readTable(sheet: SheetSnapshot, specs: ColumnSpec[], minColumns = 2): Table | null {
  let best: { row: number; columns: Record<string, number> } | null = null;
  const last = Math.min(sheet.originRow + sheet.cells.length, sheet.originRow + SCAN_ROWS);
  for (let r = sheet.originRow; r < last; r++) {
    const columns = mapRow(rowOf(sheet, r), specs);
    const n = Object.keys(columns).length;
    if (n >= minColumns && (!best || n > Object.keys(best.columns).length)) best = { row: r, columns };
  }
  if (!best) return null;

  const keys = Object.keys(best.columns);
  const rows: TableRow[] = [];
  let blanks = 0;
  for (let r = best.row + 1; r < sheet.originRow + sheet.cells.length; r++) {
    const line = rowOf(sheet, r);
    const at = (col: number) => line[col - sheet.originCol];
    const cells: Record<string, Cell | undefined> = {};
    keys.forEach(k => (cells[k] = at(best!.columns[k])));
    const firstText = line.filter(c => c && c.kind === "text")[0];
    if (firstText && TOTAL_ROW.test(String(firstText.value))) break;
    if (keys.every(k => isBlank(cells[k]))) {
      if (++blanks >= 2) break;
      continue;
    }
    blanks = 0;
    rows.push({ row: r, cells });
  }
  return { headerRow: best.row, columns: best.columns, rows };
}

// ─── Typed accessors ─────────────────────────────────────────────────────────

export function dateOf(row: TableRow, key: string): number | null {
  const c = row.cells[key];
  return c && !isBlank(c) ? toSerial(c.value) : null;
}

export function amountOf(row: TableRow, key: string): number | null {
  const c = row.cells[key];
  return c && !isBlank(c) ? toAmount(c.value) : null;
}

export function textOf(row: TableRow, key: string): string {
  const c = row.cells[key];
  return c && !isBlank(c) ? String(c.value).trim() : "";
}

/** Yes/no answers such as "Y", "Yes", "Recorded", "✓", TRUE. Returns null when blank or unclear. */
export function flagOf(row: TableRow, key: string): boolean | null {
  const c = row.cells[key];
  if (!c || isBlank(c)) return null;
  if (typeof c.value === "boolean") return c.value;
  const s = String(c.value).trim().toLowerCase();
  if (/^(y|yes|true|✓|✔|x|recorded|in ap|accrued|received|done|ok)$/.test(s)) return true;
  if (/^(n|no|false|not recorded|none|not received|nr|n\/r)$/.test(s)) return false;
  return null;
}

function colRef(row: number, col: number): string {
  let n = col + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return `${s}${row + 1}`;
}

/** Address for a missing cell in a mapped column (blank cells may be absent from the snapshot). */
export function cellAddress(table: Table, row: TableRow, key: string): string {
  return row.cells[key]?.address ?? colRef(row.row, table.columns[key]);
}
