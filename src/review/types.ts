// Core data model for the review engine. Nothing here depends on Office.js,
// so every check can be unit-tested with plain objects.

export type CellKind = "empty" | "number" | "text" | "boolean" | "error";

export interface Cell {
  /** 0-based absolute row index on the worksheet. */
  row: number;
  /** 0-based absolute column index on the worksheet. */
  col: number;
  /** A1 address without sheet name, e.g. "B7". */
  address: string;
  kind: CellKind;
  value: string | number | boolean | null;
  /** The A1 formula including the leading "=", or null for constants and blanks. */
  formula: string | null;
}

export interface SheetSnapshot {
  name: string;
  /** Row-major grid covering the used range. Missing cells are treated as empty. */
  cells: Cell[][];
  /** Absolute row/column of cells[0][0]. */
  originRow: number;
  originCol: number;
  /** True when the used range was larger than we were willing to read. */
  truncated?: boolean;
}

export type Severity = "high" | "medium" | "low";

/** The always-on workpaper checks. Procedure reviewers add their own ids such as "surl.complete". */
export type BasicCheckId =
  | "errors"
  | "hardcodes"
  | "overwritten"
  | "inconsistent"
  | "totals"
  | "external-links"
  | "signoff"
  | "header"
  | "tickmarks";

export type CheckId = string;

export interface Finding {
  checkId: CheckId;
  /** A1 address the note is attached to; absent for sheet-level notes. */
  cell?: string;
  message: string;
}

export interface Check {
  id: CheckId;
  /** Checklist line, phrased as the thing that should be true. */
  title: string;
  severity: Severity;
  /** Plain-English "why a reviewer cares", shown when the junior asks why. */
  why: string;
  /** Plain-English "how to fix it". */
  fix: string;
  /** Extra PY-note tags this check answers to, e.g. "completeness". */
  aliases?: string[];
  run(sheet: SheetSnapshot, ctx?: ReviewContext): Finding[];
}

/** Engagement facts set once per workbook. Dates are Excel serial day numbers. */
export interface ReviewContext {
  yearEnd?: number;
  /** Amount above which an item is worth a closer look (e.g. clearly trivial / SAD threshold). */
  threshold?: number;
}

export interface InsightItem {
  /** Cell to jump to. */
  cell: string;
  /** Short label for the item, e.g. vendor or cheque number. */
  label: string;
  amount?: number;
  /** Why this item was flagged, in plain English. */
  detail: string;
}

export interface Insight {
  id: string;
  title: string;
  severity: Severity;
  why: string;
  /** Context the insight can't run without. */
  needs?: Array<keyof ReviewContext>;
  /** PY-note tags this insight answers to, e.g. "cutoff". */
  aliases?: string[];
  run(sheet: SheetSnapshot, ctx: ReviewContext): InsightItem[];
}

export interface Reviewer {
  id: string;
  name: string;
  /** Catalog row, e.g. "Cash", "Liabilities". */
  category: string;
  icon: string;
  summary: string;
  /** Always-on reviewers run on every sheet; procedure reviewers are chosen per sheet. */
  alwaysOn: boolean;
  /** Not built yet: shown in the catalog as "Coming soon". */
  comingSoon?: boolean;
  /** 0..1: how much this sheet looks like this procedure. */
  detect(sheet: SheetSnapshot): number;
  checks: Check[];
  insights: Insight[];
}
