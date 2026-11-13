// Workpaper-quality checks: completeness, placeholders, dates, numbers stored as text,
// stray numbers off to the side, and leftover highlighting.

import { findYearEndInText, formatSerial, toAmount, toSerial, ymdFromSerial } from "./dates";
import {
  DetectedTable,
  allCells,
  cellAt,
  findTables,
  isBlank,
  isDateCell,
  isNumeric,
  isText,
  mainColumns,
  rowLabel,
  textDate,
} from "./layout";
import { colToLetters } from "./address";
import { Cell, Check, Finding, ReviewContext, SheetSnapshot } from "./types";

const tablesCache = new WeakMap<SheetSnapshot, DetectedTable[]>();
function tablesOf(sheet: SheetSnapshot): DetectedTable[] {
  let t = tablesCache.get(sheet);
  if (!t) {
    t = findTables(sheet);
    tablesCache.set(sheet, t);
  }
  return t;
}

function list(cells: string[], max = 4): string {
  return cells.length <= max ? cells.join(", ") : `${cells.slice(0, max).join(", ")} and ${cells.length - max} more`;
}

// ─── Completeness ────────────────────────────────────────────────────────────

/** Columns where blanks are normal, e.g. comments or the tickmark column. */
const OPTIONAL_COLUMN = /comment|note|remark|explanation|tick|legend|^ref\b|w\/?p ref/i;

export const completeness: Check = {
  id: "table-complete",
  title: "Tables are complete",
  severity: "high",
  aliases: ["completeness"],
  why: "A blank date, amount or conclusion in a testing table means that item isn't evidenced. It's one of the most common reasons a workpaper comes back from review.",
  fix: "Fill in the blank cells. If something genuinely doesn't apply, write N/A with a short reason so the reviewer knows it was considered.",
  run(sheet) {
    const out: Finding[] = [];
    for (const t of tablesOf(sheet)) {
      for (const [col, heading] of t.headers) {
        if (OPTIONAL_COLUMN.test(heading)) continue;
        const blanks = t.rows.filter(r => isBlank(cellAt(sheet, r, col)));
        const filled = t.rows.length - blanks.length;
        // Only a column that is mostly filled in is expected to be complete.
        if (!blanks.length || filled < 2 || filled / t.rows.length < 0.6) continue;
        for (const r of blanks) {
          out.push({
            checkId: "table-complete",
            cell: `${colToLetters(col)}${r + 1}`,
            message: `${rowLabel(sheet, r)}: "${heading}" is blank`,
          });
        }
      }
    }
    return out;
  },
};

// ─── Placeholders ────────────────────────────────────────────────────────────

const PLACEHOLDER =
  /\b(tbd|tbc)\b|\bto be (completed|confirmed|updated|determined|added|done)\b|\bx{3,}\b|\?{2,}|\[(insert|add|enter|tbd|update)[^\]]*\]|\bpending\b|\bplaceholder\b/i;

/** Masked account or card numbers ("XXXX-1234", "****5678") are not placeholders. */
const MASKED_NUMBER = /[x*]{3,}[\s-]*\d+|\d+[\s-]*[x*]{3,}/gi;

export const placeholders: Check = {
  id: "placeholders",
  title: "No placeholders left (TBD, XXX, ???)",
  severity: "medium",
  aliases: ["completeness"],
  why: "Placeholder text tells the reviewer the work isn't finished. It's the first thing they'll send back.",
  fix: "Replace the placeholder with the real answer, or delete it if it no longer applies.",
  run(sheet) {
    return allCells(sheet)
      .filter(c => isText(c) && PLACEHOLDER.test(c.value.replace(MASKED_NUMBER, "")))
      .map(c => ({
        checkId: "placeholders",
        cell: c.address,
        message: `"${String(c.value).trim().slice(0, 60)}"`,
      }));
  },
};

// ─── Dates ───────────────────────────────────────────────────────────────────

/** All month-end dates mentioned in a text, e.g. both dates in "Year ended 31 Dec 2025 (PY: 31 Dec 2024)". */
function monthEndsIn(text: string): number[] {
  const out: number[] = [];
  const re =
    /\b(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?,?\s+(\d{4})\b|\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const found = findYearEndInText([m[0]]);
    if (found !== null) out.push(found);
  }
  return out;
}

const SIGNOFF_DATE = /(prepared|reviewed)\s*by[^0-9a-z]*[a-z .]*?(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}|\d{1,2}[\s-][a-z]{3}[a-z]*[\s-]\d{2,4})/i;

function signoffDate(sheet: SheetSnapshot, kind: "prepared" | "reviewed"): { cell: Cell; date: number } | null {
  for (const c of allCells(sheet)) {
    if (!isText(c) || !new RegExp(`^\\s*${kind}\\s*by`, "i").test(c.value)) continue;
    const inline = SIGNOFF_DATE.exec(c.value);
    if (inline) {
      const d = toSerial(inline[2]);
      if (d !== null) return { cell: c, date: d };
    }
    for (let k = 1; k <= 4; k++) {
      const right = cellAt(sheet, c.row, c.col + k);
      if (isDateCell(right)) return { cell: right!, date: Math.floor(right!.value as number) };
      const t = textDate(right);
      if (t !== null) return { cell: right!, date: t };
    }
  }
  return null;
}

const DATE_OUTLIER_DAYS = 400;

/** The order and style of a date format, ignoring padding: "dd-mmm-yy" and "d-mmm-yyyy" are both "dMy". */
function dateStyle(fmt: string): string {
  const bare = fmt
    .replace(/"[^"]*"/g, "")
    .replace(/\[[^\]]*\]/g, "")
    .split(";")[0]
    .toLowerCase();
  return (bare.match(/d+|m+|y+/g) || []).map(t => (t.charAt(0) === "m" ? (t.length >= 3 ? "M" : "m") : t.charAt(0))).join("");
}

export const dates: Check = {
  id: "dates",
  title: "Dates are consistent",
  severity: "medium",
  why: "Wrong dates are easy to miss and embarrassing to explain: a prior-year date left in the header after roll-forward, a typo in the year, or a date typed as text that won't sort or compare.",
  fix: "Check each flagged date against its source. Use Step through dates below to look at every date on the sheet in order.",
  run(sheet, ctx?: ReviewContext) {
    const out: Finding[] = [];

    // 1. Roll-forward: a prior-year year-end left in the heading rows.
    if (ctx && ctx.yearEnd) {
      const ye = ymdFromSerial(ctx.yearEnd);
      for (const c of allCells(sheet)) {
        if (!isText(c) || c.row > sheet.originRow + 8) continue;
        // A prior-year date beside this year's is a comparative column, not a heading left behind.
        const thisYear = new RegExp(`\\b${ye.year}\\b`);
        const rowMentionsThisYear = (sheet.cells[c.row - sheet.originRow] || []).some(
          o => !!o && ((isText(o) && thisYear.test(o.value)) || (isDateCell(o) && ymdFromSerial(o.value as number).year === ye.year)),
        );
        if (rowMentionsThisYear) continue;
        for (const d of monthEndsIn(c.value)) {
          const x = ymdFromSerial(d);
          if (x.year < ye.year && x.month === ye.month && !/prior|py\b|comparative|previous|last year/i.test(c.value)) {
            out.push({
              checkId: "dates",
              cell: c.address,
              message: `Says ${formatSerial(d)} but year-end is ${formatSerial(ctx.yearEnd)}. Rolled forward?`,
            });
          }
        }
      }
    }

    // 2. Per column: typos far from the other dates, dates typed as text, mixed formats.
    const byCol = new Map<number, Cell[]>();
    for (const c of allCells(sheet)) {
      if (isDateCell(c) || textDate(c) !== null) {
        if (!byCol.has(c.col)) byCol.set(c.col, []);
        byCol.get(c.col)!.push(c);
      }
    }
    for (const cells of byCol.values()) {
      const real = cells.filter(isDateCell);
      if (real.length >= 3) {
        const sorted = real.map(c => c.value as number).sort((a, b) => a - b);
        const median = sorted[Math.floor(sorted.length / 2)];
        const far = real.filter(c => Math.abs((c.value as number) - median) > DATE_OUTLIER_DAYS);
        const near = real.length - far.length;
        // A typo stands out only where the other dates sit close together (a testing table). Columns that
        // legitimately span years (acquisition dates, loan terms, lease starts) are left alone.
        if (far.length && near >= 3 && far.length <= Math.max(1, Math.floor(real.length * 0.2))) {
          for (const c of far) {
            out.push({
              checkId: "dates",
              cell: c.address,
              message: `${formatSerial(c.value as number)} is far from the other dates in this column. Typo in the year?`,
            });
          }
        }
        // "dd-mmm-yy" and "d-mmm-yyyy" read the same; only a different order or style is worth a flag.
        const styles = new Map<string, string>();
        real.forEach(c => {
          if (c.numberFormat && !styles.has(dateStyle(c.numberFormat))) styles.set(dateStyle(c.numberFormat), c.numberFormat);
        });
        if (styles.size > 1) {
          out.push({
            checkId: "dates",
            cell: real[0].address,
            message: `Column ${colToLetters(real[0].col)} mixes date formats (${Array.from(styles.values()).join(", ")})`,
          });
        }
      }
      if (real.length >= 2) {
        for (const c of cells) {
          if (isDateCell(c)) continue;
          out.push({
            checkId: "dates",
            cell: c.address,
            message: `"${String(c.value).trim()}" is typed as text, so Excel won't treat it as a date`,
          });
        }
      }
    }

    // 3. Sign-off order.
    const prep = signoffDate(sheet, "prepared");
    const rev = signoffDate(sheet, "reviewed");
    if (prep && rev && rev.date < prep.date) {
      out.push({
        checkId: "dates",
        cell: rev.cell.address,
        message: `Reviewed ${formatSerial(rev.date)}, before it was prepared (${formatSerial(prep.date)})`,
      });
    }
    return out;
  },
};

/** Every date on the sheet in reading order, for stepping through them one by one. */
export function dateCells(sheet: SheetSnapshot): Array<{ cell: string; date: number; asText: boolean }> {
  const out: Array<{ cell: string; date: number; asText: boolean }> = [];
  for (const line of sheet.cells) {
    for (const c of line) {
      if (isDateCell(c)) out.push({ cell: c.address, date: Math.floor(c.value as number), asText: false });
      else {
        const t = textDate(c);
        if (t !== null) out.push({ cell: c.address, date: t, asText: true });
      }
    }
  }
  return out;
}

// ─── Numbers stored as text ──────────────────────────────────────────────────

export const textNumbers: Check = {
  id: "text-numbers",
  title: "No numbers stored as text",
  severity: "high",
  why: "A number stored as text looks normal but is left out of SUM totals, so the total silently misses it.",
  fix: "Re-enter it as a number (or use Data → Text to Columns on the column). Excel's green triangle in the corner marks these cells.",
  run(sheet) {
    const out: Finding[] = [];
    for (const c of allCells(sheet)) {
      if (!isText(c) || textDate(c) !== null) continue;
      const raw = c.value.trim();
      const amount = toAmount(raw);
      if (amount === null || raw.length > 20) continue;
      // An ID like "1041" in a column of IDs is fine; flag it when its neighbours are real numbers.
      let numericNeighbours = 0;
      for (let k = -4; k <= 4; k++) if (k && isNumeric(cellAt(sheet, c.row + k, c.col))) numericNeighbours++;
      const looksLikeMoney = /[,$()]|\.\d{2}$/.test(raw);
      if (numericNeighbours >= 2 || (looksLikeMoney && numericNeighbours >= 1)) {
        out.push({ checkId: "text-numbers", cell: c.address, message: `"${raw}" is text, not a number` });
      }
    }
    return out;
  },
};

// ─── Stray numbers off to the side ──────────────────────────────────────────

/** Above this many, unlabelled numbers are a block of data (a second schedule), not scratch work. */
const MAX_STRAY = 8;

function hasNeighbour(sheet: SheetSnapshot, c: Cell): boolean {
  for (let dr = -1; dr <= 1; dr++)
    for (let dc = -1; dc <= 1; dc++) if ((dr || dc) && !isBlank(cellAt(sheet, c.row + dr, c.col + dc))) return true;
  return false;
}

export const strayNumbers: Check = {
  id: "stray",
  title: "No stray numbers off to the side",
  severity: "medium",
  why: "Unlabelled numbers outside the working area are usually leftover scratch work. A reviewer can't tell what they are or whether they matter.",
  fix: "Delete scratch calculations, or move them into the workpaper with a label explaining what they are.",
  run(sheet) {
    const tables = tablesOf(sheet);
    // "Off to the side" only means something once the sheet has a working area to be beside.
    if (!tables.length) return [];
    const area = mainColumns(sheet, tables);
    const width = sheet.cells[0] ? sheet.cells[0].length : 0;
    const inTable = new Set<number>();
    for (const t of tables) for (const r of t.rows) for (const col of t.headers.keys()) inTable.add(r * 16384 + col);
    const out: Finding[] = [];
    for (const line of sheet.cells) {
      // The first text in the row is its label; anything to its right is labelled.
      let firstText = Infinity;
      for (let i = 0; i < width; i++) {
        if (isText(line[i])) {
          firstText = line[i].col;
          break;
        }
      }
      for (const c of line) {
        if (!c || !(isNumeric(c) || isDateCell(c)) || inTable.has(c.row * 16384 + c.col)) continue;
        const labelled = firstText < c.col;
        // Outside the working area, only a label right beside the number explains it.
        const offSide =
          !!area && (c.col >= area.max + 2 || c.col <= area.min - 2) && !isText(cellAt(sheet, c.row, c.col - 1));
        const isolated = !labelled && !hasNeighbour(sheet, c);
        if (offSide || isolated) {
          out.push({
            checkId: "stray",
            cell: c.address,
            message: offSide ? `${c.address} sits outside the workpaper with no label` : `${c.address} has no label`,
          });
        }
      }
    }
    return out.length > MAX_STRAY ? [] : out;
  },
};

// ─── Leftover highlighting ───────────────────────────────────────────────────

function hsl(hex: string): { h: number; s: number; l: number } {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  h *= 60;
  return { h, s, l };
}

/** Bright highlighter colours, as opposed to the pale or dark fills used for headings. */
export function isHighlighter(hex: string): boolean {
  const { s, l } = hsl(hex);
  return s >= 0.6 && l >= 0.35 && l <= 0.85;
}

function colourName(hex: string): string {
  const { h } = hsl(hex);
  if (h < 15 || h >= 345) return "red";
  if (h < 40) return "orange";
  if (h < 70) return "yellow";
  if (h < 170) return "green";
  if (h < 200) return "cyan";
  if (h < 260) return "blue";
  if (h < 300) return "purple";
  return "pink";
}

const MAX_HIGHLIGHTS = 20;

const COLOUR_LEGEND = /(yellow|green|orange|pink|blue|red|purple|highlight\w*|shad\w+)\s*(=|:|-)\s*\w/i;

export const highlights: Check = {
  id: "highlights",
  title: "No leftover highlighting",
  severity: "low",
  why: "Bright highlights are usually notes-to-self (to check, to fix, ask client). Left in, they make the reviewer wonder what's still open.",
  fix: "Clear the fill once the point is resolved, or add a colour legend if the highlighting is meant to stay.",
  skipped(sheet) {
    if (sheet.fillsRead !== false) return undefined;
    return {
      reason: "Not checked: fill colours couldn't be read on this sheet (it is very large, or this version of Excel doesn't provide them).",
      blocking: true,
    };
  },
  run(sheet) {
    const cells = allCells(sheet);
    if (cells.some(c => isText(c) && COLOUR_LEGEND.test(c.value))) return [];
    const headerRows = new Set(tablesOf(sheet).map(t => t.headerRow));
    const byColour = new Map<string, Cell[]>();
    for (const c of cells) {
      if (!c.fill || !isHighlighter(c.fill) || headerRows.has(c.row)) continue;
      const name = colourName(c.fill);
      if (!byColour.has(name)) byColour.set(name, []);
      byColour.get(name)!.push(c);
    }
    const out: Finding[] = [];
    for (const [name, list_] of byColour) {
      // A whole banded row is styling, not a highlight.
      const rows = new Set(list_.map(c => c.row));
      const banner = Array.from(rows).every(r => list_.filter(c => c.row === r).length >= 4);
      if (banner) continue;
      // A colour used on many cells is a convention (input cells, a status column), not a note-to-self.
      if (list_.length > MAX_HIGHLIGHTS) continue;
      out.push({
        checkId: "highlights",
        cell: list_[0].address,
        message: `${list_.length} cell${list_.length > 1 ? "s" : ""} highlighted ${name}: ${list(list_.map(c => c.address))}`,
      });
    }
    return out;
  },
};
