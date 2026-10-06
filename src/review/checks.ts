// The review checklist. Each check is a pure function of a SheetSnapshot so it can
// run on every edit and be tested without Excel.

import { lettersToCol, toA1 } from "./address";
import { findHardcodedNumbers, hasExternalLink, toR1C1 } from "./formula";
import { Cell, Check, CheckId, Finding, SheetSnapshot } from "./types";

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

const overwritten: Check = {
  id: "overwritten",
  title: "No numbers typed over formulas",
  severity: "high",
  why: "A typed number in the middle of a formula row usually means someone pasted over a formula. The total looks right today and silently goes wrong tomorrow.",
  fix: "Copy the formula from the neighbouring cell back into this one, or explain the override in a comment.",
  run(sheet) {
    const out: Finding[] = [];
    for (const c of allCells(sheet)) {
      if (!isNumberConstant(c)) continue;
      for (const [a, b] of neighbourPairs(sheet, c)) {
        if (isFormula(a) && isFormula(b) && r1c1(a) === r1c1(b)) {
          out.push({
            checkId: "overwritten",
            cell: c.address,
            message: `Typed value ${c.value} sits between matching formulas (${a.address}, ${b.address})`,
          });
          break;
        }
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

/** Label cells such as "Prepared by:" and whether something is filled in next to them. */
function findLabel(sheet: SheetSnapshot, pattern: RegExp): { cell: Cell; filled: boolean } | null {
  for (const c of textCells(sheet)) {
    const text = String(c.value);
    const m = pattern.exec(text);
    if (!m) continue;
    const rest = text.slice(m.index + m[0].length).replace(/^[\s:\-–]+/, "");
    if (rest.length > 0) return { cell: c, filled: true };
    for (let k = 1; k <= 3; k++) {
      const right = cellAt(sheet, c.row, c.col + k);
      if (right && right.kind !== "empty" && String(right.value).trim() !== "") return { cell: c, filled: true };
    }
    return { cell: c, filled: false };
  }
  return null;
}

const PREPARED = /prepared\s*by|preparer|prep'?d\s*by/i;
const REVIEWED = /reviewed\s*by|reviewer/i;

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
    if (!findLabel(sheet, REVIEWED))
      out.push({ checkId: "signoff", message: 'No "Reviewed by" line for the reviewer to sign' });
    return out;
  },
};

const header: Check = {
  id: "header",
  title: "Purpose and source documented",
  severity: "low",
  why: "A reviewer needs to know what the sheet is for and where the numbers came from before they can sign it off.",
  fix: 'Add a short "Purpose:" line and a "Source:" line (document name, client contact, or system report) at the top.',
  run(sheet) {
    if (isEmpty(sheet)) return [];
    const texts = textCells(sheet).map(c => String(c.value));
    const out: Finding[] = [];
    if (!texts.some(t => /purpose|objective/i.test(t)))
      out.push({ checkId: "header", message: "No purpose or objective stated" });
    if (!texts.some(t => /source/i.test(t)))
      out.push({ checkId: "header", message: "No source documented for the numbers" });
    return out;
  },
};

const TICKMARK = /^\s*[✓✔√✗✘^©®Ⓐ-Ⓩⓐ-ⓩ]\s*$/;

const tickmarks: Check = {
  id: "tickmarks",
  title: "Tickmarks explained in a legend",
  severity: "low",
  why: "A tickmark only counts as evidence if the reader can see what it means, like agreed to invoice or footed.",
  fix: 'Add a "Tickmark legend" box that explains every symbol used on the sheet.',
  run(sheet) {
    const texts = textCells(sheet);
    const marks = texts.filter(c => TICKMARK.test(String(c.value)));
    if (!marks.length) return [];
    const hasLegend = texts.some(c => /legend|tick\s*mark/i.test(String(c.value)));
    if (hasLegend) return [];
    return [
      {
        checkId: "tickmarks",
        cell: marks[0].address,
        message: `${marks.length} tickmark${marks.length > 1 ? "s" : ""} used but no legend found`,
      },
    ];
  },
};

export const CHECKS: Check[] = [
  errors,
  hardcodes,
  overwritten,
  totals,
  inconsistent,
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

export function runChecklist(sheet: SheetSnapshot): CheckResult[] {
  return CHECKS.map(check => ({ check, findings: check.run(sheet) }));
}

export { toA1 };
