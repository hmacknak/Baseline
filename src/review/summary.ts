// What the panel says at the top: how much is open on this sheet. Kept out of the panel code so it is tested.

export interface Summary {
  /** Checklist findings still open (not marked OK). */
  toFix: number;
  /** Items flagged by the procedure reviewer's insights. */
  highRisk: number;
  /** Procedure insights that couldn't run, e.g. because the year-end isn't set. */
  notRun: number;
  checks: number;
  clear: number;
  /** Only true when nothing is open and nothing was skipped. */
  allClear: boolean;
}

export function summarise(
  checks: Array<{ findings: unknown[] }>,
  insights: Array<{ items: unknown[]; blocked?: string }> = [],
): Summary {
  const toFix = checks.reduce((n, r) => n + r.findings.length, 0);
  const highRisk = insights.reduce((n, r) => n + r.items.length, 0);
  const notRun = insights.filter(r => r.blocked).length;
  return {
    toFix,
    highRisk,
    notRun,
    checks: checks.length,
    clear: checks.filter(r => r.findings.length === 0).length,
    allClear: toFix === 0 && highRisk === 0 && notRun === 0,
  };
}

/** One line for the top of the panel or a row of the whole-workbook view. */
export function summaryParts(s: Summary): string[] {
  const parts: string[] = [];
  if (s.toFix) parts.push(`${s.toFix} to fix`);
  if (s.highRisk) parts.push(`${s.highRisk} high-risk item${s.highRisk === 1 ? "" : "s"}`);
  if (s.notRun) parts.push(`${s.notRun} procedure check${s.notRun === 1 ? "" : "s"} not run`);
  return parts;
}
