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
  /** Excel number format, e.g. "dd-mmm-yy" or "#,##0". Absent when unknown. */
  numberFormat?: string;
  /** Fill colour as "#RRGGBB" when the cell is shaded; absent for no fill or when unknown. */
  fill?: string;
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
  /** Financial statement assertion at risk, e.g. "Completeness". */
  assertion?: string;
  /** Where the requirement comes from, e.g. "ISA 505.7; PCAOB AS 2310.14". */
  reference?: string;
  /** Wording and references to use instead when the workbook reports under US GAAP. */
  usGaap?: FrameworkText;
  run(sheet: SheetSnapshot, ctx?: ReviewContext): Finding[];
}

/** The accounting framework the client reports under. Decides which accounting references are shown. */
export type Framework = "ifrs" | "us-gaap";

/** Framework-specific replacement for an item's `why` and/or `reference`. */
export interface FrameworkText {
  why?: string;
  reference?: string;
}

/** Engagement facts set once per workbook. Dates are Excel serial day numbers. */
export interface ReviewContext {
  yearEnd?: number;
  /** Amount above which an item is worth a closer look (e.g. clearly trivial / SAD threshold). */
  threshold?: number;
  /** Defaults to IFRS when not set. */
  framework?: Framework;
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
  /** Financial statement assertion at risk, e.g. "Existence". */
  assertion?: string;
  /** Where the requirement comes from, e.g. "IAS 37.11". */
  reference?: string;
  /** Wording and references to use instead when the workbook reports under US GAAP. */
  usGaap?: FrameworkText;
  run(sheet: SheetSnapshot, ctx: ReviewContext): InsightItem[];
}

export interface Reviewer {
  id: string;
  name: string;
  /** Catalog row, e.g. "Cash", "Liabilities". */
  category: string;
  /** Two-to-four letter monogram shown on catalog cards, e.g. "BR". */
  code: string;
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
