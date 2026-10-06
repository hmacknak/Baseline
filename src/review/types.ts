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

export type CheckId =
  | "errors"
  | "hardcodes"
  | "overwritten"
  | "inconsistent"
  | "totals"
  | "external-links"
  | "signoff"
  | "header"
  | "tickmarks";

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
  run(sheet: SheetSnapshot): Finding[];
}
