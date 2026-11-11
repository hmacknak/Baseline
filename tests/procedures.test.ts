import { describe, expect, it } from "vitest";
import { serialFromYMD as d } from "../src/review/dates";
import { parseNotesText, pyStatuses, tagNote } from "../src/review/notes";
import { runReview, suggestReviewer } from "../src/review/review";
import { explain } from "../src/review/framework";
import { CATALOG } from "../src/review/reviewers/catalog";
import { cashConfirmations } from "../src/review/reviewers/cashConfirmations";
import { expenseVouching } from "../src/review/reviewers/expenseVouching";
import { surl } from "../src/review/reviewers/surl";
import { Reviewer, SheetSnapshot } from "../src/review/types";
import { sheet } from "./helpers";

const YE = d(2025, 12, 31);
const ctx = { yearEnd: YE, threshold: 10000 };

function cells(r: Reviewer, s: SheetSnapshot, id: string, c = ctx): string[] {
  return r.insights.filter(i => i.id === id)[0].run(s, c).map(i => i.cell);
}

// ─── Catalog ─────────────────────────────────────────────────────────────────

describe("catalog", () => {
  it("offers exactly the three enabled procedures plus workpaper basics", () => {
    expect(CATALOG.map(r => r.id)).toEqual(["basics", "surl", "cash-confirmations", "expense-vouching"]);
  });

  it("gives every procedure check and insight a reason, an assertion and a reference", () => {
    for (const r of CATALOG.filter(x => !x.alwaysOn)) {
      for (const item of [...r.checks.filter(c => c.id.endsWith(".complete")), ...r.insights]) {
        expect(item.why.length, item.id).toBeGreaterThan(80);
        expect(item.assertion, item.id).toBeTruthy();
        expect(item.reference, item.id).toBeTruthy();
      }
    }
  });
});

// ─── SURL additions ─────────────────────────────────────────────────────────

describe("SURL: new rules", () => {
  const s = sheet(
    {
      A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Payment date", E1: "Amount", F1: "Recorded in AP?", G1: "Conclusion",
      A2: "Vendor A", B2: d(2026, 1, 4), C2: d(2026, 1, 2), D2: d(2026, 1, 9), E2: 5000, F2: "Y", G2: "ok",
      A3: "Vendor B", B3: d(2026, 1, 6), C3: d(2025, 12, 20), D3: d(2026, 1, 12), E3: 7000, F3: "Y", G3: "ok",
      A4: "Vendor C", B4: d(2026, 1, 8), C4: d(2026, 1, 5), D4: d(2026, 1, 15), E4: 3000, F4: "N", G4: "ok",
    },
    "SURL",
  );

  it("flags a year-end liability recorded for next year's service", () => {
    expect(cells(surl, s, "surl.recorded-next-year")).toEqual(["F2"]);
  });

  it("flags a search window that stops soon after year-end", () => {
    const f = surl.insights.filter(i => i.id === "surl.window")[0].run(s, ctx);
    expect(f).toHaveLength(1);
    expect(f[0].detail).toContain("15 days after year-end");
  });

  it("accepts a window that runs long enough", () => {
    const long = sheet({ ...Object.fromEntries(Object.entries({
      A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Payment date", E1: "Amount", F1: "Recorded in AP?",
      A2: "Vendor A", B2: d(2026, 1, 4), C2: d(2026, 1, 2), D2: d(2026, 1, 9), E2: 5000, F2: "N",
      A3: "Vendor B", B3: d(2026, 1, 6), C3: d(2026, 1, 3), D3: d(2026, 2, 12), E3: 7000, F3: "N",
      A4: "Vendor C", B4: d(2026, 1, 8), C4: d(2026, 1, 5), D4: d(2026, 1, 15), E4: 3000, F4: "N",
    })) }, "SURL");
    expect(surl.insights.filter(i => i.id === "surl.window")[0].run(long, ctx)).toEqual([]);
  });
});

// ─── Cash confirmations ─────────────────────────────────────────────────────

const cash = sheet(
  {
    A1: "Cash confirmations, 31 December 2025",
    A3: "Bank", B3: "Account", C3: "Currency", D3: "Balance per books", E3: "Balance per confirmation", F3: "As at",
    G3: "Date sent", H3: "Date received", I3: "Received via", J3: "Explanation", K3: "Conclusion",
    // Agrees, direct, year-end: clean.
    A4: "Bank A", B4: "Operating", C4: "CAD", D4: 482150, E4: 482150, F4: YE, G4: d(2026, 1, 5), H4: d(2026, 1, 14), I4: "Confirmation platform", K4: "Agreed",
    // Difference with no explanation.
    A5: "Bank A", B5: "Payroll", C5: "CAD", D5: 36400, E5: 31400, F5: YE, G5: d(2026, 1, 5), H5: d(2026, 1, 14), I5: "Confirmation platform", K5: "Exception",
    // No reply.
    A6: "Bank B", B6: "Savings", C6: "CAD", D6: 125000, F6: YE, G6: d(2026, 1, 5), K6: "Chasing",
    // Forwarded by the client, foreign currency, confirmed at the wrong date.
    A7: "Bank C", B7: "USD account", C7: "USD", D7: 25000, E7: 25000, F7: d(2025, 11, 30), G7: d(2026, 1, 5), H7: d(2026, 1, 20), I7: "Forwarded by client controller", K7: "Agreed",
    // Nil balance, never sent.
    A8: "Bank D", B8: "Old account", C8: "CAD", D8: 0, K8: "Dormant",
    // Overdraft.
    A9: "Bank A", B9: "Line", C9: "CAD", D9: -15000, E9: -15000, F9: YE, G9: d(2026, 1, 5), H9: d(2026, 1, 14), I9: "Confirmation platform", K9: "Agreed",
  },
  "Cash confirmations",
);

describe("cash confirmations", () => {
  it("is suggested for a bank confirmation sheet", () => {
    expect(suggestReviewer(cash)?.reviewer.id).toBe("cash-confirmations");
  });

  it("flags unexplained differences", () => {
    expect(cells(cashConfirmations, cash, "cash-confirmations.difference")).toEqual(["E5"]);
  });

  it("flags requests with no reply", () => {
    expect(cells(cashConfirmations, cash, "cash-confirmations.no-response")).toEqual(["H6"]);
  });

  it("flags replies that came through the client", () => {
    expect(cells(cashConfirmations, cash, "cash-confirmations.not-direct")).toEqual(["I7"]);
  });

  it("flags unconfirmed accounts, including nil balances", () => {
    const items = cashConfirmations.insights.filter(i => i.id === "cash-confirmations.not-sent")[0].run(cash, ctx);
    expect(items.map(i => i.cell)).toEqual(["G8"]);
    expect(items[0].detail).toContain("Nil-balance");
  });

  it("flags confirmations at another date, overdrafts and foreign currency", () => {
    expect(cells(cashConfirmations, cash, "cash-confirmations.date")).toEqual(["F7"]);
    expect(cells(cashConfirmations, cash, "cash-confirmations.overdraft")).toEqual(["D9"]);
    expect(cells(cashConfirmations, cash, "cash-confirmations.foreign")).toEqual(["C7"]);
  });

  it("is not offered for customer (AR) confirmations", () => {
    const ar = sheet({
      A1: "Customer", B1: "Balance per ledger", C1: "Date sent", D1: "Confirmed balance",
      A2: "Company A", B2: 100, C2: d(2026, 1, 5), D2: 100,
      A3: "Company B", B3: 200, C3: d(2026, 1, 5),
    }, "AR confirmations");
    expect(suggestReviewer(ar)).toBeNull();
  });
});

// ─── Expense vouching ───────────────────────────────────────────────────────

const exp = sheet(
  {
    A1: "Expense vouching, year ended 31 December 2025",
    A3: "Sample #", B3: "GL date", C3: "Vendor", D3: "Description", E3: "Amount per GL", F3: "Amount per invoice",
    G3: "Invoice #", H3: "Service date", I3: "Approved?", J3: "Correct account?", K3: "Conclusion",
    // Clean.
    A4: "1", B4: d(2025, 3, 4), C4: "Vendor A", D4: "Office cleaning", E4: 1200, F4: 1200, G4: "A-11", H4: d(2025, 2, 28), I4: "Y", J4: "Y", K4: "Agreed",
    // Amount doesn't agree.
    A5: "2", B5: d(2025, 5, 9), C5: "Vendor B", D5: "Consulting", E5: 8500, F5: 5800, G5: "B-20", H5: d(2025, 5, 1), I5: "Y", J5: "Y", K5: "Exception",
    // Service after year-end.
    A6: "3", B6: d(2025, 12, 30), C6: "Vendor C", D6: "Annual software support", E6: 6000, F6: 6000, G6: "C-7", H6: d(2026, 1, 15), I6: "Y", J6: "Y", K6: "Agreed",
    // Duplicate of sample 1 (same vendor, amount, invoice).
    A7: "4", B7: d(2025, 3, 18), C7: "Vendor A", D7: "Office cleaning", E7: 1200, F7: 1200, G7: "A-11", H7: d(2025, 2, 28), I7: "Y", J7: "Y", K7: "Agreed",
    // Not approved, wrong account, capital item.
    A8: "5", B8: d(2025, 8, 2), C8: "Vendor D", D8: "New forklift", E8: 42000, F8: 42000, G8: "D-3", H8: d(2025, 7, 30), I8: "N", J8: "N", K8: "Exception",
    // Approval blank.
    A9: "6", B9: d(2025, 9, 2), C9: "Vendor E", D9: "Courier", E9: 300, F9: 300, G9: "E-9", H9: d(2025, 9, 1), J9: "Y", K9: "Agreed",
  },
  "Expense vouching",
);

describe("expense vouching", () => {
  it("is suggested for an expense vouching sheet", () => {
    expect(suggestReviewer(exp)?.reviewer.id).toBe("expense-vouching");
  });

  it("flags amounts that don't agree to support", () => {
    expect(cells(expenseVouching, exp, "expense-vouching.amount")).toEqual(["F5"]);
  });

  it("flags services outside the year", () => {
    expect(cells(expenseVouching, exp, "expense-vouching.cutoff")).toEqual(["H6"]);
  });

  it("flags possible duplicates once, pointing at the repeat", () => {
    const items = expenseVouching.insights.filter(i => i.id === "expense-vouching.duplicates")[0].run(exp, ctx);
    expect(items.map(i => i.cell)).toEqual(["C7"]);
    expect(items[0].detail).toContain("row 4");
  });

  it("flags missing or refused approval", () => {
    expect(cells(expenseVouching, exp, "expense-vouching.approval")).toEqual(["I8", "I9"]);
  });

  it("flags likely capital items over threshold and wrong accounts", () => {
    expect(cells(expenseVouching, exp, "expense-vouching.capital")).toEqual(["D8"]);
    expect(cells(expenseVouching, exp, "expense-vouching.classification")).toEqual(["J8"]);
  });

  it("is not suggested for a SURL sheet with vendor and amount columns", () => {
    const s = sheet({
      A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Payment date", E1: "Amount", F1: "Recorded in AP?",
      A2: "Vendor A", B2: d(2026, 1, 4), C2: d(2026, 1, 2), D2: d(2026, 1, 9), E2: 5000, F2: "N",
    }, "Unrecorded liabilities");
    expect(suggestReviewer(s)?.reviewer.id).toBe("surl");
  });
});

// ─── Prior-year notes on the new procedures ─────────────────────────────────

describe("prior-year notes", () => {
  it("tags notes about the new issue types", () => {
    expect(tagNote("Bank confirmation received via the client, not directly")).toContain("confirm-control");
    expect(tagNote("Invoice 4471 paid twice")).toContain("duplicate");
    expect(tagNote("Sample 12 amount does not agree to the invoice")).toContain("vouch-amount");
  });

  it("shows a duplicate-payment note as happening again", () => {
    const r = runReview(exp, ctx, expenseVouching);
    const notes = parseNotesText("Expense vouching | Duplicate payment to a supplier not identified");
    expect(pyStatuses(notes, "Expense vouching", r.all, r.insights)[0].status).toBe("still-an-issue");
  });
});

// ─── Accounting framework ───────────────────────────────────────────────────

describe("accounting framework", () => {
  const items = CATALOG.filter(r => !r.alwaysOn).flatMap(r => [...r.checks, ...r.insights]);
  const IFRS_CITE = /\b(IAS|IFRS)\b/;

  it("defaults to IFRS references", () => {
    const capital = expenseVouching.insights.filter(i => i.id === "expense-vouching.capital")[0];
    expect(explain(capital, ctx).reference).toBe("IAS 16.7; IAS 16.12");
  });

  it("cites US GAAP instead when the workbook reports under US GAAP", () => {
    const us = { ...ctx, framework: "us-gaap" as const };
    const capital = expenseVouching.insights.filter(i => i.id === "expense-vouching.capital")[0];
    expect(explain(capital, us).reference).toContain("ASC 360");
    const overdraft = cashConfirmations.insights.filter(i => i.id === "cash-confirmations.overdraft")[0];
    expect(explain(overdraft, us).why).toContain("liability");
    expect(explain(overdraft, us).reference).toContain("ASC 305");
  });

  it("never shows an IFRS citation under US GAAP", () => {
    const us = { ...ctx, framework: "us-gaap" as const };
    for (const item of items) {
      const text = explain(item, us);
      expect(text.why, item.id).not.toMatch(IFRS_CITE);
      expect(text.reference || "", item.id).not.toMatch(IFRS_CITE);
    }
  });

  it("keeps audit references the same under either framework", () => {
    const us = { ...ctx, framework: "us-gaap" as const };
    const notDirect = cashConfirmations.insights.filter(i => i.id === "cash-confirmations.not-direct")[0];
    expect(explain(notDirect, us).reference).toBe(explain(notDirect, ctx).reference);
  });
});

describe("duplicate payments", () => {
  it("names occurrence and accuracy as the assertions at risk", () => {
    const dup = expenseVouching.insights.filter(i => i.id === "expense-vouching.duplicates")[0];
    expect(dup.assertion).toBe("Occurrence; Accuracy");
  });
});
