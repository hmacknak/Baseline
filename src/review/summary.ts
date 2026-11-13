// What the panel says at the top: how much is open on this sheet. Kept out of the panel code so it is tested.
//
// Every check ends in one of three states: it flagged something, it is clear, or it did not run. A check that
// did not run is never counted as clear.

export interface Summary {
  /** Checklist findings still open (not marked OK). */
  toFix: number;
  /** Items flagged by the procedure reviewer's insights. */
  highRisk: number;
  /** Checks and insights that could not run: a missing year-end, a missing column, unreadable fill colours. */
  notRun: number;
  /** Checks that applied to this sheet (those that don't apply are left out). */
  checks: number;
  clear: number;
  /** Only true when nothing is open and nothing was skipped. */
  allClear: boolean;
}

export function summarise(
  checks: Array<{ findings: unknown[]; skipped?: { blocking: boolean } }>,
  insights: Array<{ items: unknown[]; blocked?: string }> = [],
): Summary {
  const applied = checks.filter(r => !r.skipped);
  const toFix = applied.reduce((n, r) => n + r.findings.length, 0);
  const highRisk = insights.reduce((n, r) => n + r.items.length, 0);
  const notRun = insights.filter(r => r.blocked).length + checks.filter(r => r.skipped && r.skipped.blocking).length;
  return {
    toFix,
    highRisk,
    notRun,
    checks: applied.length,
    clear: applied.filter(r => r.findings.length === 0).length,
    allClear: toFix === 0 && highRisk === 0 && notRun === 0,
  };
}

/** One line for the top of the panel or a row of the whole-workbook view. */
export function summaryParts(s: Summary): string[] {
  const parts: string[] = [];
  if (s.toFix) parts.push(`${s.toFix} to fix`);
  if (s.highRisk) parts.push(`${s.highRisk} high-risk item${s.highRisk === 1 ? "" : "s"}`);
  // Nothing was found, but not everything could be checked: say both, in that order.
  if (!parts.length && s.notRun) parts.push("No issues found");
  if (s.notRun) parts.push(`${s.notRun} check${s.notRun === 1 ? "" : "s"} not run`);
  return parts;
}
