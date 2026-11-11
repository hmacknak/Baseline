// "Mark OK": a junior or reviewer accepts a flag with a reason. Saved in the workbook so the
// reviewer can see what was accepted and why. If the flagged cell changes, the flag comes back.

import { CheckResult } from "./checks";
import { cellAt } from "./layout";
import { parseA1 } from "./address";
import { normaliseSheetName } from "./notes";
import { Finding, SheetSnapshot } from "./types";

export interface Dismissal {
  reason: string;
  /** What the cell held when it was accepted; a different value means the flag needs a fresh look. */
  fingerprint: string;
  /** ISO date it was accepted. */
  date: string;
}

export type Dismissals = Record<string, Dismissal>;

export interface DismissedFinding {
  finding: Finding;
  key: string;
  dismissal: Dismissal;
}

export interface ReviewedCheck extends CheckResult {
  dismissed: DismissedFinding[];
}

export function findingKey(sheetName: string, f: Finding): string {
  return `${normaliseSheetName(sheetName) || sheetName}|${f.checkId}|${f.cell ?? f.message}`;
}

export function fingerprint(sheet: SheetSnapshot, f: Finding): string {
  let content = "";
  const p = f.cell ? parseA1(f.cell) : null;
  if (p) {
    const c = cellAt(sheet, p.row, p.col);
    content = c ? String(c.formula ?? c.value ?? "") : "";
  }
  return `${content}|${f.message}`;
}

/** Split each check's findings into still-open and accepted. A check passes when nothing is open. */
export function applyDismissals(sheet: SheetSnapshot, results: CheckResult[], dismissals: Dismissals): ReviewedCheck[] {
  return results.map(r => {
    const open: Finding[] = [];
    const dismissed: DismissedFinding[] = [];
    for (const f of r.findings) {
      const key = findingKey(sheet.name, f);
      const d = dismissals[key];
      if (d && d.fingerprint === fingerprint(sheet, f)) dismissed.push({ finding: f, key, dismissal: d });
      else open.push(f);
    }
    return { check: r.check, findings: open, dismissed };
  });
}
