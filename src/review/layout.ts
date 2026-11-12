// Reads the shape of a workpaper: which cells are dates, where the tables are, and where
// the main working area ends. The basics checks build on these.

import { toSerial } from "./dates";
import { Cell, SheetSnapshot } from "./types";

export function cellAt(sheet: SheetSnapshot, row: number, col: number): Cell | undefined {
  const r = sheet.cells[row - sheet.originRow];
  return r ? r[col - sheet.originCol] : undefined;
}

export function allCells(sheet: SheetSnapshot): Cell[] {
  const out: Cell[] = [];
  for (const row of sheet.cells) for (const c of row) if (c) out.push(c);
  return out;
}

export function isBlank(c: Cell | undefined): boolean {
  return !c || c.kind === "empty" || (typeof c.value === "string" && c.value.trim() === "");
}

export function isText(c: Cell | undefined): c is Cell & { value: string } {
  return !!c && c.kind === "text" && typeof c.value === "string" && c.value.trim() !== "";
}

/** True for number formats that display a date, e.g. "dd-mmm-yy", "m/d/yyyy", "[$-409]d-mmm-yy". */
export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false;
  const bare = fmt
    .replace(/"[^"]*"/g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/\\./g, "")
    .toLowerCase();
  if (/^(general|@)$/.test(bare.trim())) return false;
  return /(^|[^a-z])(d{1,4}|m{3,5}|y{2,4})([^a-z]|$)/.test(bare) || /\b(m{1,2}[/\-.]d{1,2}|d{1,2}[/\-.]m{1,2})\b/.test(bare);
}

/** Serials below this are 1900 dates: in practice a zero or a time shown with a date format, not a date. */
const FIRST_REAL_DATE = 367;

/** A real Excel date: a number shown with a date format. */
export function isDateCell(c: Cell | undefined): boolean {
  return !!c && c.kind === "number" && typeof c.value === "number" && c.value >= FIRST_REAL_DATE && isDateFormat(c.numberFormat);
}

export function isNumeric(c: Cell | undefined): boolean {
  return !!c && c.kind === "number" && !isDateCell(c);
}

const DATE_LIKE_TEXT = /\d{1,4}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{1,2}[\s-](jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)|(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}/i;

/** Text that reads as a date, e.g. "05-Jan-26" typed so Excel stored it as text. */
export function textDate(c: Cell | undefined): number | null {
  if (!isText(c)) return null;
  const text = String(c.value).trim();
  return DATE_LIKE_TEXT.test(text) ? toSerial(text) : null;
}

// ─── Tables ──────────────────────────────────────────────────────────────────

export interface DetectedTable {
  headerRow: number;
  /** Column index → heading text. */
  headers: Map<number, string>;
  /** Item rows (0-based worksheet rows), excluding any total row. */
  rows: number[];
}

const TOTAL_ROW = /^\s*(sub[\s-]?)?total\b|^\s*grand total/i;
const NOTE_ROW = /^\s*(conclu(sion|ded)|notes?\b|tick\s*-?\s*marks?|legend|source|purpose|prepared\s*by|reviewed\s*by)/i;

function textCount(sheet: SheetSnapshot, r: number): Cell[] {
  const line = sheet.cells[r - sheet.originRow] || [];
  return line.filter(isText);
}

function isTotalRow(sheet: SheetSnapshot, r: number): boolean {
  const first = textCount(sheet, r)[0];
  return !!first && TOTAL_ROW.test(String(first.value));
}

/**
 * Tables are a header row of at least three headings followed by item rows that hold numbers or dates.
 * Items run until a blank row, a total row, or the end of the sheet.
 */
export function findTables(sheet: SheetSnapshot): DetectedTable[] {
  const tables: DetectedTable[] = [];
  const lastRow = sheet.originRow + sheet.cells.length - 1;
  let r = sheet.originRow;
  while (r <= lastRow) {
    const heads = textCount(sheet, r);
    if (heads.length < 3) {
      r++;
      continue;
    }
    const headers = new Map<number, string>();
    heads.forEach(h => headers.set(h.col, String(h.value).trim()));
    const rows: number[] = [];
    let next = r + 1;
    while (next <= lastRow) {
      if (isTotalRow(sheet, next)) break;
      const inRow = Array.from(headers.keys()).map(col => cellAt(sheet, next, col));
      if (inRow.every(isBlank)) break;
      // A lone sentence under the table (a conclusion or a note) ends it; it isn't an item with blanks.
      const filled = inRow.filter(c => !isBlank(c));
      if (filled.length === 1 && isText(filled[0]) && (filled[0].value.trim().length > 40 || NOTE_ROW.test(filled[0].value))) break;
      rows.push(next);
      next++;
    }
    const hasData = rows.some(row =>
      Array.from(headers.keys()).some(col => {
        const c = cellAt(sheet, row, col);
        return isNumeric(c) || isDateCell(c);
      }),
    );
    if (rows.length >= 2 && hasData) {
      tables.push({ headerRow: r, headers, rows });
      r = next;
    } else {
      r++;
    }
  }
  return tables;
}

/** First text in a row, used to name the row in messages, e.g. "row 11 (Vendor D)". */
export function rowLabel(sheet: SheetSnapshot, row: number): string {
  const t = textCount(sheet, row)[0];
  return t ? `row ${row + 1} (${String(t.value).trim()})` : `row ${row + 1}`;
}

/** Columns that hold the workpaper's main content: the tables, or failing that, wide rows. */
export function mainColumns(sheet: SheetSnapshot, tables: DetectedTable[]): { min: number; max: number } | null {
  let min = Infinity;
  let max = -Infinity;
  for (const t of tables) {
    for (const col of t.headers.keys()) {
      min = Math.min(min, col);
      max = Math.max(max, col);
    }
  }
  if (min === Infinity) {
    for (const line of sheet.cells) {
      const filled = line.filter(c => c && !isBlank(c));
      if (filled.length >= 3) {
        min = Math.min(min, filled[0].col);
        max = Math.max(max, filled[filled.length - 1].col);
      }
    }
  }
  return min === Infinity ? null : { min, max };
}
