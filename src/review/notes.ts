// Prior-year (PY) review notes: parse them, tag them to checklist items, and
// match them to this year's sheet so the review can say "this happened last year".

import { CheckResult } from "./checks";
import { CheckId } from "./types";

export interface PyNote {
  /** Sheet the note was raised on last year; empty means "any sheet". */
  sheet: string;
  /** Cell reference from last year, if given. Kept for context only; rows move year to year. */
  cell: string;
  text: string;
  tags: CheckId[];
}

export type PyStatus = "still-an-issue" | "looks-fixed" | "check-manually";

export interface PyNoteStatus {
  note: PyNote;
  status: PyStatus;
}

// ─── Tagging ─────────────────────────────────────────────────────────────────

const TAG_RULES: Array<{ id: CheckId; pattern: RegExp }> = [
  { id: "errors", pattern: /#ref|#div|#n\/a|#value|#name|formula error|\berrors?\b/i },
  { id: "hardcodes", pattern: /hard[\s-]?cod|\bplug(ged)?\b|typed (in )?number|link (it )?to (the )?source|not linked/i },
  { id: "overwritten", pattern: /over[\s-]?writ|typed over|pasted over|broken formula|formula (was )?replaced/i },
  { id: "totals", pattern: /\bfoot|cross[\s-]?foot|does(n'?t| not) (add|sum|tie)|total (is )?(wrong|off|excludes?|missing)|sum range/i },
  { id: "inconsistent", pattern: /inconsistent|formula (differs|doesn'?t match)|not consistent/i },
  { id: "external-links", pattern: /external link|links? to (another|other) (file|workbook)|broken link/i },
  { id: "signoff", pattern: /sign[\s-]?off|initial(s|led)?\b|prepared by|reviewed by|\bdate[ds]?\b.*\b(prep|sign)/i },
  { id: "header", pattern: /\bpurpose\b|\bobjective\b|\bsource\b|where (did|does) .* come from/i },
  { id: "tickmarks", pattern: /tick\s*marks?|legend/i },
];

/** "Link to source" is advice about a hard-code, not a note that the sheet lacks a documented source. */
const LINK_TO_SOURCE = /link(ed)? (it |this )?to (the )?source/gi;

export function tagNote(text: string): CheckId[] {
  return TAG_RULES.filter(r => {
    const subject = r.id === "header" ? text.replace(LINK_TO_SOURCE, "") : text;
    return r.pattern.test(subject);
  }).map(r => r.id);
}

// ─── Parsing ─────────────────────────────────────────────────────────────────

const CELL_REF = /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/;

function makeNote(sheet: string, cell: string, text: string): PyNote | null {
  const t = text.trim();
  if (!t) return null;
  return { sheet: sheet.trim(), cell: cell.trim(), text: t, tags: tagNote(t) };
}

/**
 * Parse pasted notes, one per line:
 *   Sheet | Cell | Note
 *   Sheet | Note
 *   Note
 * Tabs (pasted from Excel) work as separators too.
 */
export function parseNotesText(input: string): PyNote[] {
  const out: PyNote[] = [];
  for (const raw of input.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line) continue;
    const parts = line.split(/\s*[|\t]\s*/);
    let note: PyNote | null;
    if (parts.length >= 3) note = makeNote(parts[0], parts[1], parts.slice(2).join(" | "));
    else if (parts.length === 2) {
      // "B12 | note" is a cell, not a sheet.
      note = CELL_REF.test(parts[0]) ? makeNote("", parts[0], parts[1]) : makeNote(parts[0], "", parts[1]);
    } else note = makeNote("", "", parts[0]);
    if (note) out.push(note);
  }
  return out;
}

/**
 * Parse rows read from a "PY Notes" worksheet. If the first row looks like a header
 * (Sheet / Cell / Note), columns are matched by name; otherwise the layout is Sheet, Cell, Note.
 */
export function parseNotesRows(rows: unknown[][]): PyNote[] {
  if (!rows.length) return [];
  const header = rows[0].map(v => String(v ?? "").trim().toLowerCase());
  const find = (re: RegExp) => header.findIndex(h => re.test(h));
  let sheetCol = find(/^(sheet|tab|workpaper|wp|area)/);
  let cellCol = find(/^(cell|ref|reference|location)/);
  let noteCol = find(/(note|comment|issue|finding)/);
  let start = 1;
  if (noteCol < 0) {
    sheetCol = 0;
    cellCol = 1;
    noteCol = 2;
    start = 0;
  }
  const out: PyNote[] = [];
  for (let i = start; i < rows.length; i++) {
    const r = rows[i];
    const get = (k: number) => (k >= 0 && r[k] !== undefined && r[k] !== null ? String(r[k]) : "");
    const note = makeNote(get(sheetCol), get(cellCol), get(noteCol));
    if (note) out.push(note);
  }
  return out;
}

// ─── Matching to this year's sheet ───────────────────────────────────────────

/** "FY24 Cash", "Cash 2024" and "cash (PY)" all normalise to "cash". */
export function normaliseSheetName(name: string): string {
  return name
    .toLowerCase()
    .replace(/\b(fy|py|cy)\s*'?\d{0,4}\b/g, " ")
    .replace(/\b(19|20)\d{2}\b/g, " ")
    .replace(/\b\d{2}\b/g, " ")
    .replace(/[()[\]_\-.]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function notesForSheet(notes: PyNote[], sheetName: string): PyNote[] {
  const target = normaliseSheetName(sheetName);
  return notes.filter(n => !n.sheet || normaliseSheetName(n.sheet) === target);
}

/**
 * For each PY note on this sheet: is the same kind of issue still showing this year?
 * Notes we can't tie to an automatic check are flagged for a manual look.
 */
export function pyStatuses(notes: PyNote[], sheetName: string, results: CheckResult[]): PyNoteStatus[] {
  const failing = new Set<CheckId>();
  for (const r of results) if (r.findings.length) failing.add(r.check.id);
  return notesForSheet(notes, sheetName).map(note => {
    let status: PyStatus = "check-manually";
    if (note.tags.length) status = note.tags.some(t => failing.has(t)) ? "still-an-issue" : "looks-fixed";
    return { note, status };
  });
}

/** PY notes that relate to a given checklist item on this sheet. */
export function pyNotesForCheck(notes: PyNote[], sheetName: string, checkId: CheckId): PyNote[] {
  return notesForSheet(notes, sheetName).filter(n => n.tags.indexOf(checkId) >= 0);
}
