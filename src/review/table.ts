// Reads the testing table on a procedure sheet: finds the header row, maps columns by their
// many possible names, and returns one record per tested item with cell addresses for jumping.

import { numericDateOrder, toAmount, toSerial } from "./dates";
import { isNoteText } from "./layout";
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
  /** How numeric text dates on this sheet are read; see `dateOrderOf`. */
  dayFirst?: boolean;
}

export type DateOrder = "day-first" | "month-first" | "unknown";

export interface Table {
  headerRow: number;
  columns: Record<string, number>;
  rows: TableRow[];
  /** Whether dates typed as text on this sheet are day-first, month-first, or can't be told. */
  dateOrder: DateOrder;
}

/**
 * Dates typed as text ("05/01/2026") don't say whether the day or the month comes first. The sheet usually does:
 * another text date such as "25/01/2026" can only be day-first, and a date format such as dd/mm/yyyy on real dates
 * shows how the preparer writes them. With no evidence either way the order is unknown, and callers must not guess
 * silently.
 */
export function dateOrderOf(sheet: SheetSnapshot): DateOrder {
  let day = 0;
  let month = 0;
  for (const line of sheet.cells) {
    for (const c of line) {
      if (!c) continue;
      if (c.kind === "text") {
        const order = numericDateOrder(c.value);
        if (order === "day-first") day++;
        else if (order === "month-first") month++;
      } else if (c.kind === "number" && c.numberFormat) {
        // An explicit dd/mm format is a choice. "m/d/yyyy" is Excel's locale default code, so it proves nothing.
        const f = c.numberFormat.toLowerCase().replace(/\[[^\]]*\]/g, "");
        if (/^d{1,2}[/.\-]m{1,2}[/.\-]y{2,4}/.test(f)) day++;
      }
    }
  }
  if (day > month) return "day-first";
  if (month > day) return "month-first";
  return "unknown";
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
  const dateOrder = dateOrderOf(sheet);
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
    // A conclusion or note under the table ends it; it is not a tested item with every other column blank.
    const filled = keys.filter(k => !isBlank(cells[k]));
    if (filled.length === 1 && isNoteText(cells[filled[0]])) break;
    blanks = 0;
    rows.push({ row: r, cells, dayFirst: dateOrder === "day-first" });
  }
  return { headerRow: best.row, columns: best.columns, rows, dateOrder };
}

// ─── Typed accessors ─────────────────────────────────────────────────────────

export function dateOf(row: TableRow, key: string): number | null {
  const c = row.cells[key];
  return c && !isBlank(c) ? toSerial(c.value, row.dayFirst) : null;
}

export function amountOf(row: TableRow, key: string): number | null {
  const c = row.cells[key];
  return c && !isBlank(c) ? toAmount(c.value) : null;
}

export function textOf(row: TableRow, key: string): string {
  const c = row.cells[key];
  return c && !isBlank(c) ? String(c.value).trim() : "";
}

/** "N/A", "n.a.", "not applicable", or a dash: a deliberate non-answer, not something Baseline failed to read. */
export function isNotApplicable(text: string): boolean {
  return /^(n\s*[/.]?\s*a\.?|not applicable|nil|[-–—])$/i.test(text.trim());
}

/** Yes/no answers such as "Y", "Yes", "Recorded", "✓", TRUE. Returns null when blank or unclear. */
export function flagOf(row: TableRow, key: string): boolean | null {
  const c = row.cells[key];
  if (!c || isBlank(c)) return null;
  if (typeof c.value === "boolean") return c.value;
  const s = String(c.value).trim().toLowerCase();
  if (isNotApplicable(s)) return null;
  // The answer is the first word; anything after it is explanation ("No, accrued in January").
  if (/^(y|yes|true|ok|done)$/.test(s) || /^[✓✔]/.test(s) || /^(yes|recorded|in ap|accrued|received|approved|agreed)\b/.test(s)) return true;
  if (/^(n|no|false|nr|n\/r|none)$/.test(s) || /^[✗✘]/.test(s) || /^(no|not)\b/.test(s)) return false;
  // A bare "x" can mean ticked or crossed, so it is left unread rather than guessed.
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
