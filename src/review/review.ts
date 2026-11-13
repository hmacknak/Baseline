// Runs a full review of one sheet: the always-on basics, plus the procedure reviewer chosen for it.

import { CheckResult, runCheck } from "./checks";
import { basics, procedureReviewers } from "./reviewers/catalog";
import { Insight, InsightItem, ReviewContext, Reviewer, SheetSnapshot } from "./types";
import { markAsWorkpaper } from "./workpaper";

export interface InsightResult {
  insight: Insight;
  items: InsightItem[];
  /** Why the insight couldn't run, e.g. "Set the year-end to run this". */
  blocked?: string;
  /** True when the fix is in Settings (a missing year-end), as opposed to something on the sheet. */
  needsSettings?: boolean;
}

export interface ReviewResult {
  basics: CheckResult[];
  reviewer?: Reviewer;
  procedure: CheckResult[];
  insights: InsightResult[];
  /** Basics plus procedure checks, for PY-note matching and the score. */
  all: CheckResult[];
}

const NEEDS_LABEL: Record<keyof ReviewContext, string> = {
  yearEnd: "year-end",
  threshold: "threshold",
  framework: "accounting framework",
};

export function runReview(sheet: SheetSnapshot, ctx: ReviewContext, reviewer?: Reviewer): ReviewResult {
  if (reviewer) markAsWorkpaper(sheet);
  const basicResults = basics.checks.map(check => runCheck(check, sheet, ctx));
  const procedure = reviewer ? reviewer.checks.map(check => runCheck(check, sheet, ctx)) : [];
  const insights: InsightResult[] = reviewer
    ? reviewer.insights.map(insight => {
        const missing = (insight.needs || []).filter(k => ctx[k] === undefined || ctx[k] === null);
        if (missing.length) {
          return {
            insight,
            items: [],
            blocked: `Set the ${missing.map(k => NEEDS_LABEL[k]).join(" and ")} in Settings to run this`,
            needsSettings: true,
          };
        }
        const why = insight.unavailable ? insight.unavailable(sheet, ctx) : undefined;
        if (why) return { insight, items: [], blocked: why };
        return { insight, items: insight.run(sheet, ctx) };
      })
    : [];
  return { basics: basicResults, reviewer, procedure, insights, all: basicResults.concat(procedure) };
}

export const SUGGEST_AT = 0.6;

/** The procedure reviewer this sheet most looks like, if any looks likely enough. */
export function suggestReviewer(sheet: SheetSnapshot): { reviewer: Reviewer; score: number } | null {
  let best: { reviewer: Reviewer; score: number } | null = null;
  for (const reviewer of procedureReviewers()) {
    const score = reviewer.detect(sheet);
    if (score >= SUGGEST_AT && (!best || score > best.score)) best = { reviewer, score };
  }
  return best;
}
