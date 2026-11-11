// The reviewer catalog. To add a procedure: write src/review/reviewers/<name>.ts and list it here.
// See docs/ADDING_A_REVIEWER.md.

import { ENABLED_PROCEDURES } from "../../config";
import { CHECKS } from "../checks";
import { Reviewer } from "../types";
import { arConfirmations } from "./arConfirmations";
import { bankRec } from "./bankRec";
import { cashConfirmations } from "./cashConfirmations";
import { expenseVouching } from "./expenseVouching";
import { revenueCutoff } from "./revenueCutoff";
import { surl } from "./surl";

export const basics: Reviewer = {
  id: "basics",
  name: "Workpaper basics",
  category: "Always on",
  code: "WB",
  summary: "Errors, hard-codes, broken totals, sign-offs, purpose, source and tickmarks, on every sheet.",
  alwaysOn: true,
  detect: () => 1,
  checks: CHECKS,
  insights: [],
};

/** Every procedure reviewer that exists, switched on or not. */
export const ALL_PROCEDURES: Reviewer[] = [surl, cashConfirmations, expenseVouching, bankRec, revenueCutoff, arConfirmations];

/** What users see: the always-on basics plus the enabled procedures. */
export const CATALOG: Reviewer[] = [basics].concat(ALL_PROCEDURES.filter(r => ENABLED_PROCEDURES.indexOf(r.id) >= 0));

export const CATEGORY_ORDER = ["Always on", "Cash", "Liabilities", "Expenses", "Revenue", "Receivables"];

export function getReviewer(id: string): Reviewer | undefined {
  return CATALOG.filter(r => r.id === id)[0];
}

export function procedureReviewers(): Reviewer[] {
  return CATALOG.filter(r => !r.alwaysOn && !r.comingSoon);
}
