// The reviewer catalog. To add a procedure: write src/review/reviewers/<name>.ts and list it here.
// See docs/ADDING_A_REVIEWER.md.

import { CHECKS } from "../checks";
import { Reviewer } from "../types";
import { arConfirmations } from "./arConfirmations";
import { bankRec } from "./bankRec";
import { revenueCutoff } from "./revenueCutoff";
import { surl } from "./surl";

export const basics: Reviewer = {
  id: "basics",
  name: "Workpaper basics",
  category: "Always on",
  icon: "✓",
  summary: "Errors, hard-codes, broken totals, sign-offs, purpose, source and tickmarks, on every sheet.",
  alwaysOn: true,
  detect: () => 1,
  checks: CHECKS,
  insights: [],
};

function comingSoon(id: string, name: string, category: string, icon: string, summary: string): Reviewer {
  return { id, name, category, icon, summary, alwaysOn: false, comingSoon: true, detect: () => 0, checks: [], insights: [] };
}

export const CATALOG: Reviewer[] = [
  basics,
  bankRec,
  surl,
  revenueCutoff,
  arConfirmations,
  comingSoon("accruals", "Accrued liabilities", "Liabilities", "🧾", "Recalculates accruals and compares them to subsequent invoices."),
  comingSoon("debt", "Debt and covenants", "Liabilities", "📑", "Agrees loan balances to statements and recalculates covenant ratios."),
  comingSoon("revenue-vouching", "Revenue vouching", "Revenue", "🧮", "Vouches a sample of sales to contracts, shipping docs and cash."),
  comingSoon("ar-aging", "AR aging and allowance", "Receivables", "⏳", "Re-ages balances and challenges the allowance for old items."),
  comingSoon("fixed-assets", "Fixed asset additions", "Fixed assets", "🏗️", "Vouches additions and spots repairs that were capitalised."),
  comingSoon("depreciation", "Depreciation recalculation", "Fixed assets", "📉", "Recalculates depreciation and flags useful lives off policy."),
  comingSoon("inventory-count", "Inventory count", "Inventory", "📦", "Ties count sheets to the final listing and flags count differences."),
  comingSoon("payroll", "Payroll reasonableness", "Expenses", "👥", "Compares payroll to headcount and rates, and flags outliers."),
  comingSoon("prepaids", "Prepaid expenses", "Assets", "🗓️", "Recalculates prepaid amortisation from invoice periods."),
  comingSoon("journal-entries", "Journal entry testing", "Fraud risk", "🕵️", "Flags weekend, round-number and after-hours journal entries."),
];

export const CATEGORY_ORDER = ["Always on", "Cash", "Liabilities", "Revenue", "Receivables", "Fixed assets", "Inventory", "Assets", "Expenses", "Fraud risk"];

export function getReviewer(id: string): Reviewer | undefined {
  return CATALOG.filter(r => r.id === id)[0];
}

export function procedureReviewers(): Reviewer[] {
  return CATALOG.filter(r => !r.alwaysOn && !r.comingSoon);
}
