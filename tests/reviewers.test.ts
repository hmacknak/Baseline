import { describe, expect, it } from "vitest";
import { serialFromYMD as d } from "../src/review/dates";
import { pyStatuses, parseNotesText } from "../src/review/notes";
import { runReview, suggestReviewer } from "../src/review/review";
import { arConfirmations } from "../src/review/reviewers/arConfirmations";
import { bankRec } from "../src/review/reviewers/bankRec";
import { CATALOG, getReviewer } from "../src/review/reviewers/catalog";
import { revenueCutoff } from "../src/review/reviewers/revenueCutoff";
import { surl } from "../src/review/reviewers/surl";
import { Reviewer, SheetSnapshot } from "../src/review/types";
import { sheet } from "./helpers";

const YE = d(2025, 12, 31);
const ctx = { yearEnd: YE, threshold: 10000 };

function insightCells(r: Reviewer, s: SheetSnapshot, id: string, c = ctx): string[] {
  return r.insights.filter(i => i.id === id)[0].run(s, c).map(i => i.cell);
}

function completeness(r: Reviewer, s: SheetSnapshot): string[] {
  return r.checks[0].run(s, ctx).map(f => f.message);
}

// ─── SURL ────────────────────────────────────────────────────────────────────

const surlSheet = sheet(
  {
    A1: "Search for unrecorded liabilities, 31 Dec 2025",
    A3: "Vendor", B3: "Invoice date", C3: "Service date", D3: "Payment date", E3: "Amount", F3: "Recorded in AP?", G3: "Conclusion",
    // Jan invoice for December services, not recorded → unrecorded liability.
    A4: "Acme Plumbing", B4: d(2026, 1, 12), C4: d(2025, 12, 18), D4: d(2026, 1, 20), E4: 18400, F4: "N", G4: "Exception",
    // Jan invoice for January services → next year, fine.
    A5: "Bolt Software", B5: d(2026, 1, 5), C5: d(2026, 1, 1), D5: d(2026, 1, 9), E5: 6000, F5: "N", G5: "OK",
    // December invoice not in AP.
    A6: "City Power", B6: d(2025, 12, 28), C6: d(2025, 12, 15), D6: d(2026, 1, 15), E6: 3200, F6: "No", G6: "Exception",
    // Missing service date and conclusion.
    A7: "Delta Freight", B7: d(2026, 1, 8), D7: d(2026, 1, 22), E7: 12750, F7: "N",
    // Recorded properly.
    A8: "Echo Legal", B8: d(2026, 1, 3), C8: d(2025, 12, 10), D8: d(2026, 1, 25), E8: 9000, F8: "Y", G8: "Recorded",
  },
  "FY25 SURL",
);

describe("SURL reviewer", () => {
  it("flags post-year-end invoices for this year's services", () => {
    expect(insightCells(surl, surlSheet, "surl.post-ye-invoice")).toEqual(["C4"]);
  });

  it("flags pre-year-end invoices not in AP", () => {
    expect(insightCells(surl, surlSheet, "surl.pre-ye-unrecorded")).toEqual(["F6"]);
  });

  it("flags rows with no service date", () => {
    expect(insightCells(surl, surlSheet, "surl.no-service-date")).toEqual(["C7"]);
  });

  it("reports each missing cell", () => {
    expect(completeness(surl, surlSheet)).toEqual([
      "row 7 (Delta Freight): service date missing",
      "row 7 (Delta Freight): conclusion missing",
    ]);
  });

  it("treats a date exactly on year-end as this year", () => {
    const s = sheet({
      A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Amount", E1: "Recorded?",
      A2: "X", B2: YE + 1, C2: YE, D2: 5, E2: "N",
    });
    expect(insightCells(surl, s, "surl.post-ye-invoice")).toEqual(["C2"]);
  });

  it("explains when the table is missing", () => {
    expect(completeness(surl, sheet({ A1: "Notes only" }, "SURL"))[0]).toContain("Couldn't find the testing table");
  });
});

// ─── Bank rec ────────────────────────────────────────────────────────────────

const bankSheet = sheet(
  {
    A1: "Cash: bank reconciliation, 31 Dec 2025",
    A3: "Balance per bank statement", B3: 482150,
    A4: "Adjusted bank balance", B4: 467275,
    A5: "Balance per general ledger", B5: 468475,
    A7: "Cheque #", B7: "Payee", C7: "Date", D7: "Amount", E7: "Cleared date", F7: "Comment",
    A8: "1041", B8: "Northwind", C8: d(2025, 5, 2), D8: 2400, F8: "Not yet cleared",
    A9: "1102", B9: "Contoso", C9: d(2025, 12, 20), D9: 31500, E9: d(2025, 12, 29),
    A10: "1109", B10: "Fabrikam", C10: d(2026, 1, 4), D10: 7300, E10: d(2026, 1, 9),
    A11: "1110", B11: "Tailspin", C11: d(2025, 12, 30), D11: 12000,
    A12: "1111", B12: "Adventure", C12: d(2025, 12, 31), D12: 900, E12: d(2026, 1, 6),
  },
  "FY25 Cash",
);

describe("bank rec reviewer", () => {
  it("flags a rec that doesn't agree to the GL", () => {
    const items = bankRec.insights.filter(i => i.id === "bank-rec.difference")[0].run(bankSheet, ctx);
    expect(items).toHaveLength(1);
    expect(items[0].cell).toBe("B4");
    expect(items[0].amount).toBe(-1200);
  });

  it("flags stale, cleared-early, post-year-end and large uncleared items", () => {
    expect(insightCells(bankRec, bankSheet, "bank-rec.stale")).toEqual(["C8"]);
    expect(insightCells(bankRec, bankSheet, "bank-rec.cleared-before-ye")).toEqual(["E9"]);
    expect(insightCells(bankRec, bankSheet, "bank-rec.dated-after-ye")).toEqual(["C10"]);
    expect(insightCells(bankRec, bankSheet, "bank-rec.large-uncleared")).toEqual(["E11"]);
  });

  it("accepts a comment instead of a cleared date", () => {
    expect(completeness(bankRec, bankSheet)).toEqual(["row 11 (1110): cleared date or comment missing"]);
  });
});

// ─── Revenue cut-off ─────────────────────────────────────────────────────────

const cutoffSheet = sheet(
  {
    A1: "Revenue cut-off testing",
    A2: "Invoice #", B2: "Customer", C2: "Invoice date", D2: "Ship date", E2: "Amount", F2: "Conclusion",
    A3: "INV-901", B3: "Northwind", C3: d(2025, 12, 31), D3: d(2026, 1, 3), E3: 48000, F3: "Exception",
    A4: "INV-902", B4: "Contoso", C4: d(2026, 1, 2), D4: d(2025, 12, 30), E4: 15500, F4: "Exception",
    A5: "CN-14", B5: "Fabrikam", C5: d(2026, 1, 6), D5: d(2026, 1, 6), E5: -9800, F5: "Check",
    A6: "INV-899", B6: "Tailspin", C6: d(2025, 12, 29), D6: d(2025, 12, 29), E6: 22000,
    A7: "INV-880", B7: "Adventure", C7: d(2025, 12, 10), D7: d(2025, 12, 9), E7: 4000, F7: "OK",
  },
  "FY25 Revenue cut-off",
);

describe("revenue cut-off reviewer", () => {
  it("flags sales recorded in the wrong year both ways", () => {
    expect(insightCells(revenueCutoff, cutoffSheet, "revenue-cutoff.early")).toEqual(["D3"]);
    expect(insightCells(revenueCutoff, cutoffSheet, "revenue-cutoff.late")).toEqual(["C4"]);
  });

  it("flags credit notes after year-end and large near-year-end items with no conclusion", () => {
    expect(insightCells(revenueCutoff, cutoffSheet, "revenue-cutoff.credit-notes")).toEqual(["E5"]);
    expect(insightCells(revenueCutoff, cutoffSheet, "revenue-cutoff.large-near-ye")).toEqual(["F6"]);
  });

  it("reports the missing conclusion", () => {
    expect(completeness(revenueCutoff, cutoffSheet)).toEqual(["row 6 (INV-899): conclusion missing"]);
  });
});

// ─── AR confirmations ───────────────────────────────────────────────────────

const confSheet = sheet(
  {
    A1: "AR confirmations",
    A2: "Customer", B2: "Balance per ledger", C2: "Date sent", D2: "Confirmed balance", E2: "Explanation", F2: "Alternative procedures",
    A3: "Northwind", B3: 23200, C3: d(2026, 1, 10), D3: 23200,
    A4: "Contoso", B4: 26050, C4: d(2026, 1, 10), D4: 22750,
    A5: "Fabrikam", B5: 14050, C5: d(2026, 1, 10),
    A6: "Adventure", B6: 22250, C6: d(2026, 1, 10), F6: "Agreed to Jan receipts",
    A7: "Tailspin", B7: 17300,
  },
  "FY25 AR confirmations",
);

describe("AR confirmations reviewer", () => {
  it("flags unexplained differences, non-responses and unsent large balances", () => {
    expect(insightCells(arConfirmations, confSheet, "ar-confirmations.difference")).toEqual(["D4"]);
    expect(insightCells(arConfirmations, confSheet, "ar-confirmations.no-response")).toEqual(["F5"]);
    expect(insightCells(arConfirmations, confSheet, "ar-confirmations.not-sent")).toEqual(["C7"]);
  });

  it("reports the missing sent date", () => {
    expect(completeness(arConfirmations, confSheet)).toEqual(["row 7 (Tailspin): date sent missing"]);
  });
});

// ─── Catalog, suggestions, review ───────────────────────────────────────────

describe("suggestReviewer", () => {
  it.each([
    [surlSheet, "surl"],
    [bankSheet, "bank-rec"],
    [cutoffSheet, "revenue-cutoff"],
    [confSheet, "ar-confirmations"],
  ])("suggests the right reviewer for %#", (s, id) => {
    expect(suggestReviewer(s)?.reviewer.id).toBe(id);
  });

  it("suggests nothing for a generic sheet", () => {
    expect(suggestReviewer(sheet({ A1: "Fixed assets", A3: "Asset", B3: "Cost" }, "FY25 Fixed Assets"))).toBeNull();
  });
});

describe("runReview", () => {
  it("blocks date insights until year-end is set", () => {
    const r = runReview(surlSheet, {}, surl);
    const post = r.insights.filter(i => i.insight.id === "surl.post-ye-invoice")[0];
    expect(post.blocked).toContain("year-end");
    expect(post.items).toEqual([]);
  });

  it("combines basics and procedure checks for PY notes", () => {
    const r = runReview(surlSheet, ctx, surl);
    expect(r.all.length).toBe(r.basics.length + r.procedure.length);
    const notes = parseNotesText("FY24 SURL | Missing service dates on several items");
    expect(pyStatuses(notes, "FY25 SURL", r.all)[0].status).toBe("still-an-issue");
    // Without the procedure reviewer the note can't be judged automatically.
    expect(pyStatuses(notes, "FY25 SURL", r.basics)[0].status).toBe("check-manually");
  });

  it("has unique ids and resolvable reviewers", () => {
    const ids = CATALOG.map(r => r.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(getReviewer("surl")).toBe(surl);
  });
});

describe("PY notes and insights", () => {
  it("matches last year's cut-off note to this year's cut-off insight", () => {
    const r = runReview(cutoffSheet, ctx, revenueCutoff);
    const notes = parseNotesText("FY24 Revenue cut-off | Dec 31 invoice shipped in January was not identified");
    expect(pyStatuses(notes, "FY25 Revenue cut-off", r.all, r.insights)[0].status).toBe("still-an-issue");
  });

  it("calls it fixed when the insight finds nothing", () => {
    const clean = sheet({
      A1: "Invoice #", B1: "Invoice date", C1: "Ship date", D1: "Amount", E1: "Conclusion",
      A2: "INV-1", B2: d(2025, 12, 10), C2: d(2025, 12, 9), D2: 100, E2: "OK",
    }, "FY25 Cut-off");
    const r = runReview(clean, ctx, revenueCutoff);
    const notes = parseNotesText("Cut-off | Sale recorded in the wrong period");
    expect(pyStatuses(notes, "FY25 Cut-off", r.all, r.insights)[0].status).toBe("looks-fixed");
  });

  it("can't judge insight notes while the insight is blocked", () => {
    const r = runReview(cutoffSheet, {}, revenueCutoff);
    const notes = parseNotesText("FY24 Revenue cut-off | Cut-off error not identified");
    expect(pyStatuses(notes, "FY25 Revenue cut-off", r.all, r.insights)[0].status).toBe("check-manually");
  });
});
