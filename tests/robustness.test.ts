// Each flag has limits. These tests pin them down from both sides: the flag must still catch the real problem
// (the limit doesn't hold the tool back), and it must stay quiet on the legitimate look-alike (no noise).
// A check that can't evaluate something must say so; it must never read as a pass.

import { describe, expect, it } from "vitest";
import { getCheck, runChecklist } from "../src/review/checks";
import { serialFromYMD as d, toSerial } from "../src/review/dates";
import { parseNotesText, pyStatuses } from "../src/review/notes";
import { runReview } from "../src/review/review";
import { cashConfirmations } from "../src/review/reviewers/cashConfirmations";
import { expenseVouching } from "../src/review/reviewers/expenseVouching";
import { surl } from "../src/review/reviewers/surl";
import { summarise } from "../src/review/summary";
import { dateOrderOf, flagOf } from "../src/review/table";
import { sheet } from "./helpers";

const YE = d(2025, 12, 31);
const ctx = { yearEnd: YE, threshold: 10000 };
const msgs = (id: string, cells: Record<string, string | number | boolean>, name = "Sheet1", extra = {}) =>
  getCheck(id).run(sheet(cells, name, extra), ctx).map(f => `${f.cell ?? ""} ${f.message}`);

// ─── No silent passes ────────────────────────────────────────────────────────

describe("a check that can't run says so", () => {
  it("shows sign-off and purpose as not applied on supporting data, not as passed", () => {
    const data = sheet({ A1: "Customer", B1: "Invoice", C1: "Amount", A2: "Company A", B2: "1001", C2: 500 });
    const results = runChecklist(data, ctx);
    const signoff = results.filter(r => r.check.id === "signoff")[0];
    expect(signoff.skipped && signoff.skipped.blocking).toBe(false);
    expect(signoff.findings).toEqual([]);
    expect(results.filter(r => r.check.id === "header")[0].skipped).toBeTruthy();
  });

  it("applies sign-off once a procedure reviewer is on the sheet, whatever the sheet looks like", () => {
    const s = sheet({
      A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Amount", E1: "Recorded in AP?",
      A2: "Vendor A", B2: d(2026, 1, 4), C2: d(2025, 12, 20), D2: 5000, E2: "N",
    });
    const r = runReview(s, ctx, surl);
    const signoff = r.basics.filter(x => x.check.id === "signoff")[0];
    expect(signoff.skipped).toBeUndefined();
    expect(signoff.findings.length).toBeGreaterThan(0);
  });

  it("reports highlighting as not run when fill colours couldn't be read", () => {
    const s = sheet({ A1: "x", B2: 5 });
    s.fillsRead = false;
    const r = runChecklist(s, ctx).filter(x => x.check.id === "highlights")[0];
    expect(r.skipped && r.skipped.blocking).toBe(true);
    expect(summarise(runChecklist(s, ctx)).allClear).toBe(false);
  });

  it("does not let a skipped check mark a prior-year note as fixed", () => {
    const data = sheet({ A1: "Customer", B1: "Invoice", C1: "Amount", A2: "Company A", B2: "1001", C2: 500 }, "Listing");
    const notes = parseNotesText("Listing | Please sign off and add the reviewer line");
    expect(pyStatuses(notes, "Listing", runChecklist(data, ctx))[0].status).toBe("check-manually");
  });

  it("reports a procedure insight as not run when its column is missing", () => {
    const s = sheet({
      A1: "Bank", B1: "Account", C1: "Per books", D1: "Per bank", E1: "Date sent",
      A2: "Bank A", B2: "Operating", C2: 100, D2: 100, E2: d(2026, 1, 5),
    }, "Cash confirmations");
    const r = runReview(s, ctx, cashConfirmations);
    const notDirect = r.insights.filter(i => i.insight.id === "cash-confirmations.not-direct")[0];
    expect(notDirect.blocked).toContain("sent / received via");
    expect(notDirect.needsSettings).toBeFalsy();
    expect(summarise(r.all, r.insights).allClear).toBe(false);
    // Rules whose columns are there still run.
    expect(r.insights.filter(i => i.insight.id === "cash-confirmations.difference")[0].blocked).toBeUndefined();
  });

  it("reports every insight as not run when there is no testing table", () => {
    const r = runReview(sheet({ A1: "Notes on expenses", A2: "Nothing here yet" }, "Expense vouching"), ctx, expenseVouching);
    expect(r.insights.every(i => !!i.blocked)).toBe(true);
  });
});

// ─── Answers Baseline can't read ─────────────────────────────────────────────

describe("unreadable answers are reported, not skipped", () => {
  const s = sheet({
    A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Amount", E1: "Recorded in AP?", F1: "Conclusion",
    A2: "Vendor A", B2: d(2026, 1, 4), C2: d(2025, 12, 20), D2: 5000, E2: "No, accrued in January", F2: "x",
    A3: "Vendor B", B3: d(2026, 1, 6), C3: "December", D3: "see invoice", E3: "Partly", F3: "x",
    A4: "Vendor C", B4: d(2026, 1, 8), C4: "N/A", D4: 3000, E4: "N/A", F4: "x",
  }, "SURL");

  it("reads an answer that starts with yes or no", () => {
    const r = runReview(s, ctx, surl);
    const flagged = r.insights.filter(i => i.insight.id === "surl.post-ye-invoice")[0].items.map(i => i.cell);
    expect(flagged).toEqual(["C2"]);
  });

  it("lists the cells it couldn't read", () => {
    const found = surl.checks.filter(c => c.id === "surl.readable")[0].run(s, ctx).map(f => f.cell);
    expect(found).toEqual(["C3", "D3", "E3"]);
  });

  it("treats N/A as an answer, and a bare x as unclear", () => {
    const row = (v: string) => ({ row: 1, cells: { k: { row: 1, col: 0, address: "A2", kind: "text" as const, value: v, formula: null } } });
    expect(flagOf(row("N/A"), "k")).toBeNull();
    expect(flagOf(row("x"), "k")).toBeNull();
    expect(flagOf(row("Yes - see AP-2"), "k")).toBe(true);
    expect(flagOf(row("Not in AP"), "k")).toBe(false);
  });
});

describe("dates typed as text", () => {
  it("reads day-first when the sheet shows that is how dates are written", () => {
    const s = sheet({ A1: "25/01/2026", A2: "05/01/2026" });
    expect(dateOrderOf(s)).toBe("day-first");
    expect(toSerial("05/01/2026", true)).toBe(d(2026, 1, 5));
    expect(toSerial("05/01/2026")).toBe(d(2026, 5, 1));
  });

  it("asks for a real date when the order can't be told", () => {
    const s = sheet({
      A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Amount", E1: "Recorded in AP?",
      A2: "Vendor A", B2: "05/01/2026", C2: d(2025, 12, 20), D2: 5000, E2: "N",
    }, "SURL");
    expect(dateOrderOf(s)).toBe("unknown");
    const found = surl.checks.filter(c => c.id === "surl.readable")[0].run(s, ctx);
    expect(found.map(f => f.cell)).toEqual(["B2"]);
    expect(found[0].message).toContain("day/month or month/day");
  });

  it("uses the sheet's order in the procedure rules", () => {
    const s = sheet({
      A1: "Vendor", B1: "Invoice date", C1: "Service date", D1: "Amount", E1: "Recorded in AP?",
      A2: "Vendor A", B2: "05/01/2026", C2: "20/12/2025", D2: 5000, E2: "N",
    }, "SURL");
    const r = runReview(s, ctx, surl);
    expect(r.insights.filter(i => i.insight.id === "surl.post-ye-invoice")[0].items[0].detail).toContain("5 Jan 2026");
  });
});

// ─── Noise ───────────────────────────────────────────────────────────────────

describe("one problem is one flag", () => {
  it("groups the same hard-code copied down a column", () => {
    const cells: Record<string, string | number> = {};
    for (let r = 1; r <= 30; r++) {
      cells[`A${r}`] = r;
      cells[`B${r}`] = `=A${r}*1.13`;
    }
    const found = msgs("hardcodes", cells);
    expect(found).toHaveLength(1);
    expect(found[0]).toContain("B1:B30 (30 cells");
  });

  it("groups the same hard-code copied along a row", () => {
    expect(msgs("hardcodes", { A1: 1, B1: 2, C1: 3, A2: "=A1*1.13", B2: "=B1*1.13", C2: "=C1*1.13" })).toHaveLength(1);
  });

  it("keeps different hard-codes apart", () => {
    expect(msgs("hardcodes", { A1: 1, A2: 2, B1: "=A1*1.13", B2: "=A2*1.05" })).toHaveLength(2);
  });

  it("groups a column of the same error", () => {
    const found = msgs("errors", { A1: "#N/A", A2: "#N/A", A3: "#N/A", C1: "#REF!" });
    expect(found).toEqual(["A1 #N/A errors in A1:A3 (3 cells)", "C1 #REF! error"]);
  });
});

describe("look-alikes that are not problems", () => {
  it("does not call a subtotal between item formulas inconsistent", () => {
    const cells = { A1: 2, B1: 3, C1: "=A1*B1", A2: 4, B2: 5, C2: "=A2*B2", C3: "=SUM(C1:C2)", A4: 6, B4: 7, C4: "=A4*B4" };
    expect(msgs("inconsistent", cells)).toEqual([]);
  });

  it("still calls a different item formula inconsistent", () => {
    const cells = { A1: 2, B1: 3, C1: "=A1*B1", A2: 4, B2: 5, C2: "=A2+B2", A3: 6, B3: 7, C3: "=A3*B3" };
    expect(msgs("inconsistent", cells)).toHaveLength(1);
  });

  it("does not call a masked account number a placeholder", () => {
    expect(msgs("placeholders", { A1: "Account XXXX-4471", A2: "Card ****5678" })).toEqual([]);
    expect(msgs("placeholders", { A1: "Vendor XXX" })).toHaveLength(1);
  });

  it("does not read a sign-off form as a table with blanks", () => {
    const s = {
      A1: "Prepared by:", B1: "Date:", C1: "Reviewed by:",
      A2: "AB", B2: d(2026, 1, 12),
      A3: "CD",
    };
    expect(msgs("table-complete", s, "T", { formats: { B2: "dd-mmm-yy" } })).toEqual([]);
  });
});

// ─── Reach ───────────────────────────────────────────────────────────────────

describe("totals", () => {
  const items = { A1: "Item", B1: "Amount", C1: "Ref", A2: "Rent", B2: 100, C2: "a", A3: "Power", B3: 200, C3: "b", A4: "Water", B4: 300, C4: "c" };

  it("catches a total that starts after the first item", () => {
    expect(msgs("totals", { ...items, A5: "Total", B5: "=SUM(B3:B4)" })).toEqual(["B5 =SUM(B3:B4) starts after B2"]);
  });

  it("catches a SUBTOTAL that stops short", () => {
    expect(msgs("totals", { ...items, A5: "Total", B5: "=SUBTOTAL(9,B2:B3)" })).toEqual(["B5 =SUBTOTAL(9,B2:B3) leaves out B4"]);
  });

  it("accepts a total that covers every item, under a numeric year heading", () => {
    const cells = { A1: "Item", B1: 2025, C1: 2024, A2: "Rent", B2: 100, C2: 90, A3: "Power", B3: 200, C3: 180, A4: "Total", B4: "=SUM(B2:B3)", C4: "=SUM(C2:C3)" };
    expect(msgs("totals", cells)).toEqual([]);
  });
});

describe("tickmarks", () => {
  it("treats letters in a Tick column as tickmarks that need a legend", () => {
    const cells = { A1: "Item", B1: "Amount", C1: "Tick", A2: "Rent", B2: 100, C2: "A", A3: "Power", B3: 200, C3: "B" };
    expect(msgs("tickmarks", cells)).toHaveLength(1);
    const withLegend = { ...cells, A6: "Tickmark legend", A7: "A = agreed to invoice", A8: "B = agreed to bank statement" };
    expect(msgs("tickmarks", withLegend)).toEqual([]);
    const half = { ...cells, A6: "Tickmark legend", A7: "A = agreed to invoice" };
    expect(msgs("tickmarks", half)).toEqual(['C3 "B" is used but not explained in the legend']);
  });

  it("does not treat a Y/N answer column as tickmarks", () => {
    expect(msgs("tickmarks", { A1: "Item", B1: "Approved?", A2: "Rent", B2: "Y", A3: "Power", B3: "N" })).toEqual([]);
  });
});

describe("a conclusion under a testing table", () => {
  it("is not read as another tested item", () => {
    const s = sheet({
      A1: "Bank", B1: "Account", C1: "Per books", D1: "Per bank", E1: "Date sent",
      A2: "Bank A", B2: "Operating", C2: 100, D2: 100, E2: d(2026, 1, 5),
      A3: "Bank B", B3: "Savings", C3: 200, D3: 200, E3: d(2026, 1, 5),
      A5: "Conclusion: to be completed once differences and non-replies are cleared.",
    }, "Cash confirmations");
    const r = runReview(s, ctx, cashConfirmations);
    expect(r.insights.filter(i => i.insight.id === "cash-confirmations.not-sent")[0].items).toEqual([]);
    expect(r.procedure.filter(c => c.check.id === "cash-confirmations.complete")[0].findings).toEqual([]);
  });
});

describe("bank reply status written in words", () => {
  const base = {
    A1: "Bank", B1: "Account", C1: "Per books", D1: "Per bank", E1: "Date sent", F1: "Date received",
    A2: "Bank A", B2: "Operating", C2: 100, E2: d(2026, 1, 5),
  };
  const noReply = (received: string) =>
    runReview(sheet({ ...base, F2: received }, "Cash confirmations"), ctx, cashConfirmations)
      .insights.filter(i => i.insight.id === "cash-confirmations.no-response")[0].items.length;

  it("still flags no reply when the cell says so in words", () => {
    expect(noReply("Not yet")).toBe(1);
    expect(noReply("Outstanding")).toBe(1);
  });

  it("accepts words that say a reply came in", () => {
    expect(noReply("Received by email 14 Jan")).toBe(0);
  });
});
