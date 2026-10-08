// Building blocks for procedure reviewers: find the table once, check it's complete,
// and turn per-row rules into insights. A new procedure is mostly column names plus rules.

import { formatSerial } from "../dates";
import { ColumnSpec, Table, TableRow, cellAddress, isBlank, readTable, textOf } from "../table";
import { Cell, Check, Finding, Insight, InsightItem, ReviewContext, Severity, SheetSnapshot } from "../types";

export interface TableSpec {
  columns: ColumnSpec[];
  /** Columns that must be matched before we treat a header row as this table. */
  minColumns?: number;
}

const tableCache = new WeakMap<SheetSnapshot, Map<TableSpec, Table | null>>();

export function tableFor(sheet: SheetSnapshot, spec: TableSpec): Table | null {
  let bySpec = tableCache.get(sheet);
  if (!bySpec) {
    bySpec = new Map();
    tableCache.set(sheet, bySpec);
  }
  if (!bySpec.has(spec)) bySpec.set(spec, readTable(sheet, spec.columns, spec.minColumns ?? 3));
  return bySpec.get(spec)!;
}

function labelOf(spec: TableSpec, key: string): string {
  return spec.columns.filter(c => c.key === key)[0]?.label ?? key;
}

/** A short name for the row in messages, e.g. "row 14 (Acme Ltd)". */
export function rowName(row: TableRow, labelKeys: string[]): string {
  for (const k of labelKeys) {
    const t = textOf(row, k);
    if (t) return `row ${row.row + 1} (${t})`;
  }
  return `row ${row.row + 1}`;
}

/**
 * "Every tested item is filled in." Each entry in `required` is a column key, or a list of keys
 * where any one filled in is enough (e.g. cleared date OR comment).
 */
export function completenessCheck(opts: {
  id: string;
  spec: TableSpec;
  required: Array<string | string[]>;
  labelKeys: string[];
  tableHint: string;
}): Check {
  return {
    id: opts.id,
    title: "Every tested item is complete",
    severity: "high",
    aliases: ["completeness"],
    why: "A blank date, amount or conclusion means the item isn't evidenced. It's one of the most common reasons a workpaper comes back from review.",
    fix: "Fill in the missing cells. If something genuinely doesn't apply, write N/A with a short reason so the reviewer knows it was considered.",
    run(sheet) {
      const table = tableFor(sheet, opts.spec);
      if (!table) {
        return [{ checkId: opts.id, message: `Couldn't find the testing table. ${opts.tableHint}` }];
      }
      if (!table.rows.length) {
        return [{ checkId: opts.id, message: "The testing table has no items yet." }];
      }
      const out: Finding[] = [];
      for (const row of table.rows) {
        for (const req of opts.required) {
          const keys = typeof req === "string" ? [req] : req;
          const mapped = keys.filter(k => table.columns[k] !== undefined);
          if (!mapped.length) continue;
          if (mapped.every(k => isBlank(row.cells[k]))) {
            const what = mapped.map(k => labelOf(opts.spec, k)).join(" or ");
            out.push({
              checkId: opts.id,
              cell: cellAddress(table, row, mapped[0]),
              message: `${rowName(row, opts.labelKeys)}: ${what} missing`,
            });
          }
        }
      }
      return out;
    },
  };
}

/** "Your table is missing a column we need", reported once per missing column. */
export function columnsCheck(opts: { id: string; spec: TableSpec; needed: string[] }): Check {
  return {
    id: opts.id,
    title: "Testing table has the key columns",
    severity: "medium",
    why: "Without these columns the reviewer can't see the evidence for each item, and Baseline can't check the risky ones for you.",
    fix: "Add the missing columns to the testing table. Any common heading works, for example 'Invoice date' or 'Inv date'.",
    run(sheet) {
      const table = tableFor(sheet, opts.spec);
      if (!table) return [];
      return opts.needed
        .filter(k => table.columns[k] === undefined)
        .map(k => ({ checkId: opts.id, message: `No "${labelOf(opts.spec, k)}" column found` }));
    },
  };
}

/** An insight that tests every row of the table with one rule. */
export function rowInsight(opts: {
  id: string;
  title: string;
  severity: Severity;
  why: string;
  needs?: Array<keyof ReviewContext>;
  aliases?: string[];
  spec: TableSpec;
  /** Return a reason string to flag the row, or null to pass it. */
  test(row: TableRow, ctx: ReviewContext): { detail: string; cellKey: string; amount?: number | null } | null;
  labelKeys: string[];
}): Insight {
  return {
    id: opts.id,
    title: opts.title,
    severity: opts.severity,
    why: opts.why,
    needs: opts.needs,
    aliases: opts.aliases,
    run(sheet, ctx) {
      const table = tableFor(sheet, opts.spec);
      if (!table) return [];
      const items: InsightItem[] = [];
      for (const row of table.rows) {
        const hit = opts.test(row, ctx);
        if (!hit) continue;
        const label = opts.labelKeys.map(k => textOf(row, k)).filter(Boolean)[0] || `Row ${row.row + 1}`;
        items.push({
          cell: cellAddress(table, row, hit.cellKey),
          label,
          amount: hit.amount ?? undefined,
          detail: hit.detail,
        });
      }
      return items;
    },
  };
}

// ─── Detection ───────────────────────────────────────────────────────────────

export function sheetTexts(sheet: SheetSnapshot, limitRows = 40): string[] {
  const out: string[] = [];
  for (let r = 0; r < Math.min(sheet.cells.length, limitRows); r++) {
    for (const c of sheet.cells[r] as Cell[]) if (c && c.kind === "text") out.push(String(c.value));
  }
  return out;
}

/** Score 0..1 from the sheet name, words on the sheet, and how many table columns we recognise. */
export function detectScore(
  sheet: SheetSnapshot,
  opts: { name?: RegExp; text?: RegExp; spec?: TableSpec; columnsForFull?: number },
): number {
  let score = 0;
  if (opts.name && opts.name.test(sheet.name)) score += 0.5;
  if (opts.text && sheetTexts(sheet).some(t => opts.text!.test(t))) score += 0.3;
  if (opts.spec) {
    const table = tableFor(sheet, opts.spec);
    if (table) {
      const n = Object.keys(table.columns).length;
      score += 0.5 * Math.min(1, n / (opts.columnsForFull ?? 4));
    }
  }
  return Math.min(1, score);
}

export const fmtDate = formatSerial;

export function fmtAmount(n: number): string {
  const s = Math.round(Math.abs(n)).toLocaleString("en-US");
  return n < 0 ? `(${s})` : s;
}
