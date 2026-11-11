// The review checklist. Each check is a pure function of a SheetSnapshot so it can
// run on every edit and be tested without Excel.

import { lettersToCol, toA1 } from "./address";
import { findHardcodedNumbers, hasExternalLink, toR1C1 } from "./formula";
import { isDateFormat } from "./layout";
import { completeness, dates, highlights, placeholders, strayNumbers, textNumbers } from "./quality";
import { Cell, Check, CheckId, Finding, ReviewContext, SheetSnapshot } from "./types";

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

// ─── Checks ──────────────────────────────────────────────────────────────────

const errors: Check = {
  id: "errors",
  title: "No formula errors",
  severity: "high",
  why: "An error like #REF! or #DIV/0! means a number on the page is wrong or missing. Reviewers stop reading the moment they see one.",
  fix: "Click into the cell, trace the inputs (Formulas → Trace Precedents) and fix the broken reference or the divide-by-zero.",
  run(sheet) {
    return allCells(sheet)
      .filter(c => c.kind === "error")
      .map(c => ({ checkId: "errors" as CheckId, cell: c.address, message: `${c.value} error` }));
  },
};

const hardcodes: Check = {
  id: "hardcodes",
  title: "No hard-coded numbers inside formulas",
  severity: "high",
  why: "A number typed into a formula is invisible support. Nobody can tell where it came from, and it won't update when the source changes.",
  fix: "Put the number in its own labelled input cell with a source reference, then point the formula at that cell.",
  run(sheet) {
    const out: Finding[] = [];
    for (const c of allCells(sheet)) {
      if (!isFormula(c)) continue;
      const nums = findHardcodedNumbers(c.formula);
      if (nums.length) {
        out.push({
          checkId: "hardcodes",
          cell: c.address,
          message: `Hard-coded ${nums.join(", ")} in ${c.formula}`,
        });
      }
    }
    return out;
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

const SUM_RANGE = /^=\s*SUM\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)\s*$/i;

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
    return allCells(sheet)
      .filter(c => isFormula(c) && hasExternalLink(c.formula!))
      .map(c => ({ checkId: "external-links" as CheckId, cell: c.address, message: `Links to another file: ${c.formula}` }));
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

// Sign-off labels start the cell ("Prepared by: AB"), so prose mentioning "the reviewer" doesn't count.
const PREPARED = /^\s*(prepared\s*by|preparer\b|prep'?d\s*by)/i;
const REVIEWED = /^\s*(reviewed\s*by|reviewer\b)/i;

const signoff: Check = {
  id: "signoff",
  title: "Preparer and reviewer sign-off",
  severity: "low",
  why: "Sign-offs show who did the work and when. A workpaper without them isn't evidence, and it's one of the most common review notes.",
  fix: 'Add "Prepared by: <initials>, <date>" and leave a "Reviewed by:" line for your senior.',
  run(sheet) {
    if (isEmpty(sheet)) return [];
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
  run(sheet) {
    if (isEmpty(sheet)) return [];
    const texts = textCells(sheet).map(c => String(c.value));
    const out: Finding[] = [];
    if (!texts.some(t => /purpose|objective/i.test(t)))
      out.push({ checkId: "header", message: "No purpose or objective stated" });
    if (!texts.some(t => /source/i.test(t)))
      out.push({ checkId: "header", message: "No source documented for the numbers" });
    if (!texts.some(t => /^\s*conclu(sion|ded)\b/i.test(t)))
      out.push({ checkId: "header", message: "No conclusion documented" });
    return out;
  },
};

const TICKMARK = /^\s*[✓✔√✗✘^©®Ⓐ-Ⓩⓐ-ⓩ]\s*$/;

const tickmarks: Check = {
  id: "tickmarks",
  title: "Tickmarks explained in a legend",
  severity: "low",
  why: "A tickmark only counts as evidence if the reader can see what it means, like agreed to invoice or footed.",
  fix: 'Add a "Tickmark legend" that explains every symbol used on the sheet, e.g. "✓ = agreed to invoice".',
  run(sheet) {
    const texts = textCells(sheet);
    const marks = texts.filter(c => TICKMARK.test(String(c.value)));
    if (!marks.length) return [];
    const legend = texts.filter(c => !TICKMARK.test(String(c.value)) && /legend|tick\s*mark|[✓✔√✗✘^©®Ⓐ-Ⓩⓐ-ⓩ]\s*[=:\-–]/i.test(String(c.value)));
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
      if (legendText.indexOf(sym) < 0) {
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
}

export function runChecklist(sheet: SheetSnapshot, ctx?: ReviewContext): CheckResult[] {
  return CHECKS.map(check => ({ check, findings: check.run(sheet, ctx) }));
}

export { toA1 };
