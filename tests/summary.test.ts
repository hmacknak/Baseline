import { describe, expect, it } from "vitest";
import { summarise, summaryParts } from "../src/review/summary";
import { looksLikeWorkpaper } from "../src/review/workpaper";
import { sheet } from "./helpers";

describe("panel summary", () => {
  it("is all clear only when nothing is open and nothing was skipped", () => {
    expect(summarise([{ findings: [] }, { findings: [] }]).allClear).toBe(true);
    expect(summarise([{ findings: [1] }]).allClear).toBe(false);
  });

  it("never says all clear while a procedure insight has flagged an item", () => {
    const s = summarise([{ findings: [] }], [{ items: [1, 2] }]);
    expect(s.allClear).toBe(false);
    expect(summaryParts(s)).toEqual(["2 high-risk items"]);
  });

  it("never says all clear while a procedure check could not run", () => {
    const s = summarise([{ findings: [] }], [{ items: [], blocked: "Set the year-end" }]);
    expect(s.allClear).toBe(false);
    expect(summaryParts(s)).toEqual(["1 procedure check not run"]);
  });

  it("lists checklist findings first", () => {
    expect(summaryParts(summarise([{ findings: [1, 2, 3] }], [{ items: [1] }]))).toEqual(["3 to fix", "1 high-risk item"]);
  });
});

describe("workpaper or supporting data", () => {
  it("recognises a workpaper by its reference, period heading or sign-off line", () => {
    expect(looksLikeWorkpaper(sheet({ A1: "x", B2: 1 }, "C-2 Cash"))).toBe(true);
    expect(looksLikeWorkpaper(sheet({ A1: "x", B2: 1 }, "FY25 AR"))).toBe(true);
    expect(looksLikeWorkpaper(sheet({ A1: "Payroll testing", A2: "Year ended 30 June 2026" }))).toBe(true);
    expect(looksLikeWorkpaper(sheet({ A1: "Listing", A9: "Prepared by:" }))).toBe(true);
  });

  it("treats a plain listing or export as supporting data", () => {
    expect(looksLikeWorkpaper(sheet({ A1: "Customer", B1: "Amount", A2: "Company A", B2: 5 }, "Sheet1"))).toBe(false);
    expect(looksLikeWorkpaper(sheet({ A1: "id", B1: "value", A2: "1", B2: 5 }, "dec2"))).toBe(false);
    expect(looksLikeWorkpaper(sheet({ A1: "Human resources", B1: 5 }, "Data"))).toBe(false);
    expect(looksLikeWorkpaper(sheet({ A1: "Sheet", B1: "Note", A2: "FY24 Cash", B2: "Sign off" }, "PY Notes"))).toBe(false);
  });
});
