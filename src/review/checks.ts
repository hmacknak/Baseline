// The review checklist. Each check is a pure function of a SheetSnapshot so it can
// run on every edit and be tested without Excel.

import { lettersToCol, toA1 } from "./address";
import { findHardcodedNumbers, hasExternalLink, toR1C1 } from "./formula";
import { DetectedTable, findTables, isDateFormat } from "./layout";
import { completeness, dates, highlights, placeholders, strayNumbers, textNumbers } from "./quality";
import { Cell, Check, CheckId, Finding, ReviewContext, SheetSnapshot, Skip } from "./types";
import { looksLikeWorkpaper } from "./workpaper";

// ─── Grid helpers ────────────────────────────────────────────────────────────

function allCells(sheet: SheetSnapshot): Cell[] {
  const out: Cell[] = [];
  for (const row of sheet.cells) for (const c of row) if (c) out.push(c);
  return out;
}

function cellAt(sheet: SheetSnapshot, row: number, col: number): Cell | undefined {
  const r = sheet.cells[row - sheet.originRow];
  return r ? r[col - sheet.originCol] : undefined;
}

function isFormula(c: Cell | undefined): c is Cell & { formula: string } {
  return !!c && c.formula !== null;
}

function isNumberConstant(c: Cell | undefined): boolean {
  return !!c && c.formula === null && c.kind === "number";
}

function isEmpty(sheet: SheetSnapshot): boolean {
  return allCells(sheet).every(c => c.kind === "empty");
}

function textCells(sheet: SheetSnapshot): Cell[] {
  return allCells(sheet).filter(c => c.kind === "text" && typeof c.value === "string");
}

// Normalised relative formulas are computed once per snapshot.
const r1c1Cache = new WeakMap<Cell, string>();
function r1c1(c: Cell & { formula: string }): string {
  let v = r1c1Cache.get(c);
  if (v === undefined) {
    v = toR1C1(c.formula, c.row, c.col);
    r1c1Cache.set(c, v);
  }
  return v;
}

/** Neighbour pairs (before, after) along the row and along the column. */
function neighbourPairs(sheet: SheetSnapshot, c: Cell): Array<[Cell | undefined, Cell | undefined]> {
  return [
    [cellAt(sheet, c.row, c.col - 1), cellAt(sheet, c.row, c.col + 1)],
    [cellAt(sheet, c.row - 1, c.col), cellAt(sheet, c.row + 1, c.col)],
  ];
}

/** A flagged cell, with the message for it alone and for a run of cells with the same problem. */
interface Hit {
  cell: Cell;
  /** Cells with the same key have the same problem, e.g. the same formula copied down. */
  key: string;
  one: string;
  many(range: string, count: number): string;
}

/**
 * One finding per unbroken run of the same problem, down a column or along a row, instead of one per cell.
 * A formula with a hard-code copied down 200 rows is one thing to fix, not 200.
 */
function groupRuns(checkId: CheckId, hits: Hit[]): Finding[] {
  const out: Array<{ row: number; col: number; finding: Finding }> = [];
  const emit = (run: Hit[]) => {
    const first = run[0];
    const last = run[run.length - 1];
    out.push({
      row: first.cell.row,
      col: first.cell.col,
      finding: {
        checkId,
        cell: first.cell.address,
        message: run.length === 1 ? first.one : first.many(`${first.cell.address}:${last.cell.address}`, run.length),
      },
    });
  };
  const collect = (list: Hit[], vertical: boolean): Hit[] => {
    const along = (h: Hit) => (vertical ? h.cell.row : h.cell.col);
    const across = (h: Hit) => (vertical ? h.cell.col : h.cell.row);
    const sorted = list.slice().sort((a, b) => across(a) - across(b) || along(a) - along(b));
    const singles: Hit[] = [];
    let i = 0;
    while (i < sorted.length) {
      let j = i;
      while (
        j + 1 < sorted.length &&
        across(sorted[j + 1]) === across(sorted[i]) &&
        along(sorted[j + 1]) === along(sorted[j]) + 1 &&
        sorted[j + 1].key === sorted[i].key
      ) {
        j++;
      }
      if (j > i) emit(sorted.slice(i, j + 1));
      else singles.push(sorted[i]);
      i = j + 1;
    }
    return singles;
  };
  collect(collect(hits, true), false).forEach(h => emit([h]));
  return out.sort((a, b) => a.row - b.row || a.col - b.col).map(o => o.finding);
}

const tablesCache = new WeakMap<SheetSnapshot, DetectedTable[]>();
function tablesOf(sheet: SheetSnapshot): DetectedTable[] {
  let t = tablesCache.get(sheet);
  if (!t) {
    t = findTables(sheet);
    tablesCache.set(sheet, t);
  }
  return t;
}

// ─── Checks ──────────────────────────────────────────────────────────────────

const errors: Check = {
  id: "errors",
  title: "No formula errors",
  severity: "high",
  why: "An error like #REF! or #DIV/0! means a number on the page is wrong or missing. Reviewers stop reading the moment they see one.",
  fix: "Click into the cell, trace the inputs (Formulas → Trace Precedents) and fix the broken reference or the divide-by-zero.",
  run(sheet) {
    return groupRuns(
      "errors",
      allCells(sheet)
        .filter(c => c.kind === "error")
        .map(c => ({
          cell: c,
          key: String(c.value),
          one: `${c.value} error`,
          many: (range, n) => `${c.value} errors in ${range} (${n} cells)`,
        })),
    );
  },
};

const hardcodes: Check = {
  id: "hardcodes",
  title: "No hard-coded numbers inside formulas",
  severity: "high",
  why: "A number typed into a formula is invisible support. Nobody can tell where it came from, and it won't update when the source changes.",
  fix: "Put the number in its own labelled input cell with a source reference, then point the formula at that cell.",
  run(sheet) {
    const hits: Hit[] = [];
    for (const c of allCells(sheet)) {
      if (!isFormula(c)) continue;
      const nums = findHardcodedNumbers(c.formula);
      if (nums.length) {
        hits.push({
          cell: c,
          key: r1c1(c),
          one: `Hard-coded ${nums.join(", ")} in ${c.formula}`,
          many: (range, n) => `Hard-coded ${nums.join(", ")} in ${range} (${n} cells with the same formula), e.g. ${c.formula}`,
        });
      }
    }
    return groupRuns("hardcodes", hits);
  },
};

/**
 * Runs of numeric cells along a column (or row). A typed constant inside a run that is mostly the
 * same formula is almost always a formula that was pasted over.
 */
function constantsAmongFormulas(sheet: SheetSnapshot, vertical: boolean): Array<{ cell: Cell; example: Cell & { formula: string } }> {
  const out: Array<{ cell: Cell; example: Cell & { formula: string } }> = [];
  const lines = vertical ? sheet.cells[0]?.length ?? 0 : sheet.cells.length;
  const length = vertical ? sheet.cells.length : sheet.cells[0]?.length ?? 0;
  for (let i = 0; i < lines; i++) {
    let run: Cell[] = [];
    const flush = () => {
      const formulas = run.filter(isFormula) as Array<Cell & { formula: string }>;
      if (run.length >= 4 && formulas.length / run.length >= 0.6) {
        const counts = new Map<string, number>();
        formulas.forEach(f => counts.set(r1c1(f), (counts.get(r1c1(f)) || 0) + 1));
        let best = "";
        counts.forEach((n, k) => {
          if (!best || n > counts.get(best)!) best = k;
        });
        if (counts.get(best)! >= formulas.length * 0.5) {
          const example = formulas.filter(f => r1c1(f) === best)[0];
          // The first cell of a run is often an opening balance or input, so it's allowed.
          run.slice(1).forEach(c => {
            if (isNumberConstant(c)) out.push({ cell: c, example });
          });
        }
      }
      run = [];
    };
    for (let j = 0; j < length; j++) {
      const c = vertical ? sheet.cells[j]?.[i] : sheet.cells[i]?.[j];
      if (c && c.kind === "number" && (isFormula(c) || isNumberConstant(c))) run.push(c);
      else flush();
    }
    flush();
  }
  return out;
}

const overwritten: Check = {
  id: "overwritten",
  title: "No typed numbers among formulas",
  severity: "high",
  why: "A typed number in a row or column of formulas usually means someone pasted over a formula. The total looks right today and silently goes wrong when the inputs change.",
  fix: "Copy the formula back into the cell, or if the override is intentional, explain it in a comment.",
  run(sheet) {
    const out: Finding[] = [];
    const seen = new Set<string>();
    const add = (c: Cell, why: string) => {
      if (seen.has(c.address)) return;
      seen.add(c.address);
      out.push({ checkId: "overwritten", cell: c.address, message: `Typed value ${c.value} ${why}` });
    };
    for (const c of allCells(sheet)) {
      if (!isNumberConstant(c)) continue;
      for (const [a, b] of neighbourPairs(sheet, c)) {
        if (isFormula(a) && isFormula(b) && r1c1(a) === r1c1(b)) {
          add(c, `sits between matching formulas (${a.address}, ${b.address})`);
          break;
        }
      }
    }
    for (const vertical of [true, false]) {
      for (const { cell, example } of constantsAmongFormulas(sheet, vertical)) {
        add(cell, `in a ${vertical ? "column" : "row"} of formulas like ${example.address} (${example.formula})`);
      }
    }
    return out;
  },
};

const AGGREGATE = /^=\s*(SUM|SUBTOTAL|AGGREGATE|AVERAGE|COUNT|COUNTA|MIN|MAX|SUMPRODUCT)\(/i;

const inconsistent: Check = {
  id: "inconsistent",
  title: "Formulas consistent across rows and columns",
  severity: "medium",
  why: "When one formula differs from the identical ones beside it, it's usually a copy mistake, like a range that stops one row short.",
  fix: "Compare it with the neighbouring formula. If it should match, copy the neighbour across. If it's intentionally different, add a comment.",
  run(sheet) {
    const out: Finding[] = [];
    for (const c of allCells(sheet)) {
      if (!isFormula(c)) continue;
      for (const [a, b] of neighbourPairs(sheet, c)) {
        if (isFormula(a) && isFormula(b) && r1c1(a) === r1c1(b) && r1c1(c) !== r1c1(a)) {
          // A subtotal sitting between two item formulas is meant to be different.
          if (AGGREGATE.test(c.formula) && !AGGREGATE.test(a.formula)) continue;
          out.push({
            checkId: "inconsistent",
            cell: c.address,
            message: `${c.formula} doesn't match ${a.address} and ${b.address} (${a.formula})`,
          });
          break;
        }
      }
    }
    return out;
  },
};

// =SUM(A1:A9), and the filter-aware =SUBTOTAL(9,A1:A9) / =SUBTOTAL(109,A1:A9).
const SUM_RANGE = /^=\s*(?:SUM\(|SUBTOTAL\(\s*(?:10)?9\s*,)\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)\s*$/i;

const totals: Check = {
  id: "totals",
  title: "Totals include every line",
  severity: "high",
  why: "A SUM that stops short of the line items above it is the classic footing error. The total looks plausible but excludes amounts.",
  fix: "Extend the SUM range so it covers every line item between the first row and the total.",
  run(sheet) {
    const out: Finding[] = [];
    for (const c of allCells(sheet)) {
      if (!isFormula(c)) continue;
      const m = SUM_RANGE.exec(c.formula);
      if (!m) continue;
      const c1 = lettersToCol(m[1]);
      const r1 = parseInt(m[2], 10) - 1;
      const c2 = lettersToCol(m[3]);
      const r2 = parseInt(m[4], 10) - 1;

      // Vertical total: single column, same column as the total, ending above it.
      if (c1 === c2 && c1 === c.col && r2 < c.row - 1) {
        const skipped: string[] = [];
        for (let r = r2 + 1; r < c.row; r++) {
          const s = cellAt(sheet, r, c.col);
          if (s && s.kind === "number") skipped.push(s.address);
        }
        if (skipped.length) {
          out.push({
            checkId: "totals",
            cell: c.address,
            message: `${c.formula} leaves out ${skipped.join(", ")}`,
          });
        }
      }
      // A total that starts late: the range begins below the first items of the table it totals.
      if (c1 === c2 && c1 === c.col && r2 < c.row) {
        const table = tablesOf(sheet).filter(t => t.rows.indexOf(r1) >= 0)[0];
        if (table) {
          const before = table.rows
            .filter(r => r < r1)
            .map(r => cellAt(sheet, r, c.col))
            .filter((s): s is Cell => !!s && s.kind === "number" && !isDateFormat(s.numberFormat));
          if (before.length) {
            out.push({
              checkId: "totals",
              cell: c.address,
              message: `${c.formula} starts after ${before.map(s => s.address).join(", ")}`,
            });
          }
        }
      }
      // Horizontal total: single row, same row as the total, ending left of it.
      if (r1 === r2 && r1 === c.row && c2 < c.col - 1) {
        const skipped: string[] = [];
        for (let col = c2 + 1; col < c.col; col++) {
          const s = cellAt(sheet, c.row, col);
          if (s && s.kind === "number") skipped.push(s.address);
        }
        if (skipped.length) {
          out.push({
            checkId: "totals",
            cell: c.address,
            message: `${c.formula} leaves out ${skipped.join(", ")}`,
          });
        }
      }
    }
    return out;
  },
};

const externalLinks: Check = {
  id: "external-links",
  title: "No links to other files",
  severity: "medium",
  why: "Links to other workbooks break when files are moved, renamed or archived, and the reviewer can't see the source.",
  fix: "Paste the source figures into this file with a reference to where they came from, or document the link in the workpaper.",
  run(sheet) {
    return groupRuns(
      "external-links",
      allCells(sheet)
        .filter((c): c is Cell & { formula: string } => isFormula(c) && hasExternalLink(c.formula))
        .map(c => ({
          cell: c,
          key: r1c1(c),
          one: `Links to another file: ${c.formula}`,
          many: (range, n) => `Links to another file in ${range} (${n} cells), e.g. ${c.formula}`,
        })),
    );
  },
};

const DATE_IN_TEXT = /\d{1,2}[/\-.]\d{1,2}([/\-.]\d{2,4})?|\d{1,2}[\s-](jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)|(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2}/i;

/** Label cells such as "Prepared by:", whether something is filled in next to them, and whether that includes a date. */
function findLabel(sheet: SheetSnapshot, pattern: RegExp): { cell: Cell; filled: boolean; dated: boolean } | null {
  for (const c of textCells(sheet)) {
    const text = String(c.value);
    const m = pattern.exec(text);
    if (!m) continue;
    const parts: string[] = [];
    let dated = false;
    const rest = text.slice(m.index + m[0].length).replace(/^[\s:\-–]+/, "");
    if (rest) parts.push(rest);
    for (let k = 1; k <= 3; k++) {
      const right = cellAt(sheet, c.row, c.col + k);
      if (!right || right.kind === "empty" || String(right.value).trim() === "") continue;
      // Stop at the next label, e.g. "Reviewed by:" on the same row as "Prepared by:".
      if (right.kind === "text" && /:\s*$|\bby\b/i.test(String(right.value))) break;
      if (right.kind === "number" && isDateFormat(right.numberFormat)) dated = true;
      parts.push(String(right.value));
    }
    if (!parts.length) return { cell: c, filled: false, dated: false };
    dated = dated || parts.some(p => DATE_IN_TEXT.test(p));
    return { cell: c, filled: true, dated };
  }
  return null;
}

/**
 * Supporting data tabs (client listings, exports) aren't signed off and don't state a purpose; only workpapers do.
 * Shown as "not applied" so nobody reads it as "the sign-off is fine".
 */
function notAWorkpaper(sheet: SheetSnapshot): Skip | undefined {
  if (isEmpty(sheet) || looksLikeWorkpaper(sheet)) return undefined;
  return {
    reason:
      "Not applied: this sheet doesn't look like a workpaper (no workpaper reference, period heading, sign-off or conclusion line), so it was treated as supporting data.",
    blocking: false,
  };
}

// Sign-off labels start the cell ("Prepared by: AB"), so prose mentioning "the reviewer" doesn't count.
const PREPARED = /^\s*(prepared\s*by|preparer\b|prep'?d\s*by)/i;
const REVIEWED = /^\s*(reviewed\s*by|reviewer\b)/i;

const signoff: Check = {
  id: "signoff",
  title: "Preparer and reviewer sign-off",
  severity: "low",
  why: "Sign-offs show who did the work and when. A workpaper without them isn't evidence, and it's one of the most common review notes.",
  fix: 'Add "Prepared by: <initials>, <date>" and leave a "Reviewed by:" line for your senior.',
  skipped: notAWorkpaper,
  run(sheet) {
    if (isEmpty(sheet) || !looksLikeWorkpaper(sheet)) return [];
    const out: Finding[] = [];
    const prep = findLabel(sheet, PREPARED);
    if (!prep) out.push({ checkId: "signoff", message: 'No "Prepared by" sign-off on this sheet' });
    else if (!prep.filled)
      out.push({ checkId: "signoff", cell: prep.cell.address, message: '"Prepared by" is blank: add your initials and date' });
    else if (!prep.dated)
      out.push({ checkId: "signoff", cell: prep.cell.address, message: '"Prepared by" has no date' });
    const rev = findLabel(sheet, REVIEWED);
    if (!rev) out.push({ checkId: "signoff", message: 'No "Reviewed by" line for the reviewer to sign' });
    else if (rev.filled && !rev.dated)
      out.push({ checkId: "signoff", cell: rev.cell.address, message: '"Reviewed by" has no date' });
    return out;
  },
};

const header: Check = {
  id: "header",
  title: "Purpose, source and conclusion documented",
  severity: "low",
  why: "A reviewer needs to know what the sheet is for, where the numbers came from, and what you concluded before they can sign it off.",
  fix: 'Add "Purpose:" and "Source:" lines (document name, client contact or system report) at the top, and a "Conclusion:" at the bottom.',
  skipped: notAWorkpaper,
  run(sheet) {
    if (isEmpty(sheet) || !looksLikeWorkpaper(sheet)) return [];
    const texts = textCells(sheet).map(c => String(c.value));
    const out: Finding[] = [];
    if (!texts.some(t => /\b(purpose|objectives?)\b/i.test(t)))
      out.push({ checkId: "header", message: "No purpose or objective stated" });
    if (!texts.some(t => /\bsources?\b/i.test(t)))
      out.push({ checkId: "header", message: "No source documented for the numbers" });
    if (!texts.some(t => /^\s*conclu(sion|ded)\b/i.test(t)))
      out.push({ checkId: "header", message: "No conclusion documented" });
    return out;
  },
};

const TICKMARK = /^\s*[✓✔√✗✘^©®Ⓐ-Ⓩⓐ-ⓩ]\s*$/;

/** A column heading that says the cells under it are tickmarks. */
const TICK_HEADER = /^\s*(ticks?|tick\s*-?\s*marks?|t\/?m)\s*$/i;

const tickmarks: Check = {
  id: "tickmarks",
  title: "Tickmarks explained in a legend",
  severity: "low",
  why: "A tickmark only counts as evidence if the reader can see what it means, like agreed to invoice or footed.",
  fix: 'Add a "Tickmark legend" that explains every symbol used on the sheet, e.g. "✓ = agreed to invoice".',
  run(sheet) {
    const texts = textCells(sheet);
    // Tickmarks are the usual symbols anywhere, plus whatever short marks sit in a column headed "Tick".
    const marks = texts.filter(c => TICKMARK.test(String(c.value)));
    const seenMark = new Set(marks.map(c => c.address));
    for (const head of texts.filter(c => TICK_HEADER.test(String(c.value)))) {
      for (let r = head.row + 1; r < sheet.originRow + sheet.cells.length; r++) {
        const c = cellAt(sheet, r, head.col);
        if (!c || c.kind !== "text") continue;
        const v = String(c.value).trim();
        if (v.length >= 1 && v.length <= 2 && !seenMark.has(c.address)) {
          seenMark.add(c.address);
          marks.push(c);
        }
      }
    }
    if (!marks.length) return [];
    const isMark = (c: Cell) => seenMark.has(c.address);
    const legend = texts.filter(
      c =>
        !isMark(c) &&
        !TICK_HEADER.test(String(c.value)) &&
        /legend|tick\s*-?\s*marks?\b.*[=:]|[✓✔√✗✘^©®Ⓐ-Ⓩⓐ-ⓩ]\s*[=:\-–]|^\s*\S{1,2}\s*[=:]\s*\w/i.test(String(c.value)),
    );
    if (!legend.length) {
      return [
        {
          checkId: "tickmarks",
          cell: marks[0].address,
          message: `${marks.length} tickmark${marks.length > 1 ? "s" : ""} used but no legend found`,
        },
      ];
    }
    // Every symbol used must be explained somewhere in the legend.
    const legendText = legend.map(c => String(c.value)).join(" ");
    const out: Finding[] = [];
    const seen = new Set<string>();
    for (const m of marks) {
      const sym = String(m.value).trim();
      if (seen.has(sym)) continue;
      seen.add(sym);
      // A symbol is explained if it appears in the legend; a letter or number only if it is followed by = or :.
      const plain = /^[A-Za-z0-9]{1,2}$/.test(sym);
      const explained = plain
        ? new RegExp(`(^|[\\s(;,])${sym}\\s*[=:\\-–)]`, "i").test(legendText)
        : legendText.indexOf(sym) >= 0;
      if (!explained) {
        out.push({ checkId: "tickmarks", cell: m.address, message: `"${sym}" is used but not explained in the legend` });
      }
    }
    return out;
  },
};

export const CHECKS: Check[] = [
  errors,
  hardcodes,
  overwritten,
  totals,
  textNumbers,
  inconsistent,
  completeness,
  placeholders,
  dates,
  strayNumbers,
  highlights,
  externalLinks,
  signoff,
  header,
  tickmarks,
];

export function getCheck(id: CheckId): Check {
  return CHECKS.filter(c => c.id === id)[0];
}

export interface CheckResult {
  check: Check;
  findings: Finding[];
  /** Set when the check did not run; `findings` is then empty and must not be read as a pass. */
  skipped?: Skip;
}

/** Run one check, or record why it didn't run. */
export function runCheck(check: Check, sheet: SheetSnapshot, ctx?: ReviewContext): CheckResult {
  const skipped = check.skipped ? check.skipped(sheet, ctx) : undefined;
  return skipped ? { check, findings: [], skipped } : { check, findings: check.run(sheet, ctx) };
}

export function runChecklist(sheet: SheetSnapshot, ctx?: ReviewContext): CheckResult[] {
  return CHECKS.map(check => runCheck(check, sheet, ctx));
}

export { toA1 };
