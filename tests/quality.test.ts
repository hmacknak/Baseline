import { describe, expect, it } from "vitest";
import { getCheck } from "../src/review/checks";
import { serialFromYMD as d } from "../src/review/dates";
import { isDateFormat } from "../src/review/layout";
import { dateCells, isHighlighter } from "../src/review/quality";
import { sheet } from "./helpers";

const DATE = "dd-mmm-yy";
const ctx = { yearEnd: d(2025, 12, 31) };

describe("typed numbers among formulas", () => {
  it("flags a constant in a column of matching formulas, even with no matching neighbour on both sides", () => {
    const s = sheet({
      A1: 1, A2: 2, A3: 3, A4: 4, A5: 5,
      B1: "=A1*2", B2: "=A2*2", B3: "=A3*2", B4: 99, B5: 10,
    });
    expect(getCheck("overwritten").run(s).map(f => f.cell)).toEqual(["B4", "B5"]);
  });

  it("allows an opening balance at the top of a formula column", () => {
    const s = sheet({ A1: 100, A2: "=A1+5", A3: "=A2+5", A4: "=A3+5", A5: "=A4+5" });
    expect(getCheck("overwritten").run(s)).toEqual([]);
  });

  it("leaves columns of inputs alone", () => {
    expect(getCheck("overwritten").run(sheet({ A1: 1, A2: 2, A3: 3, A4: 4 }))).toEqual([]);
  });
});

describe("tables are complete", () => {
  const s = sheet({
    A1: "Vendor", B1: "Invoice date", C1: "Amount", D1: "Comment",
    A2: "Vendor A", B2: d(2026, 1, 5), C2: 100,
    A3: "Bolt", B3: d(2026, 1, 6), C3: 200, D3: "ok",
    A4: "City", C4: 300,
    A5: "Delta", B5: d(2026, 1, 9), C5: 400,
    A6: "Total", C6: "=SUM(C2:C5)",
  }, "T", { formats: { B2: DATE, B3: DATE, B5: DATE } });

  it("flags blanks in mostly-filled columns, but not optional comment columns", () => {
    expect(getCheck("table-complete").run(s).map(f => f.message)).toEqual(['row 4 (City): "Invoice date" is blank']);
  });
});

describe("placeholders", () => {
  it.each(["TBD", "Conclusion: to be completed", "XXX", "Amount ???", "[insert client name]", "pending client response"])(
    "flags %s",
    text => {
      expect(getCheck("placeholders").run(sheet({ A1: text }))).toHaveLength(1);
    },
  );

  it("ignores normal words", () => {
    expect(getCheck("placeholders").run(sheet({ A1: "Expenses tested", A2: "Box 10" }))).toEqual([]);
  });
});

describe("dates", () => {
  it("recognises date number formats", () => {
    expect(["dd-mmm-yy", "m/d/yyyy", "[$-409]d-mmm-yy;@", "yyyy-mm-dd"].every(isDateFormat)).toBe(true);
    expect(["General", "#,##0", "0.0%", "#,##0;(#,##0)", "@"].some(isDateFormat)).toBe(false);
  });

  it("flags a prior-year year-end left in the heading after roll-forward", () => {
    const f = getCheck("dates").run(sheet({ A1: "Accounts payable, 31 December 2024" }), ctx);
    expect(f[0].message).toContain("Rolled forward");
  });

  it("does not flag a labelled prior-year comparative", () => {
    expect(getCheck("dates").run(sheet({ A1: "Prior year balance at 31 December 2024" }), ctx)).toEqual([]);
  });

  it("flags a year typo, a date typed as text and mixed formats", () => {
    const s = sheet(
      { A1: d(2026, 1, 5), A2: d(2026, 1, 9), A3: d(2016, 1, 12), A4: "15-Jan-26", A5: d(2026, 1, 20) },
      "T",
      { formats: { A1: DATE, A2: DATE, A3: DATE, A5: "m/d/yyyy" } },
    );
    const msgs = getCheck("dates").run(s, ctx).map(f => `${f.cell} ${f.message}`);
    expect(msgs.some(m => m.startsWith("A3") && m.includes("Typo"))).toBe(true);
    expect(msgs.some(m => m.startsWith("A4") && m.includes("typed as text"))).toBe(true);
    expect(msgs.some(m => m.includes("mixes date formats"))).toBe(true);
  });

  it("flags a review dated before preparation", () => {
    const s = sheet({ A1: "Prepared by: AB 20-Jan-26", A2: "Reviewed by: CD 15-Jan-26" });
    expect(getCheck("dates").run(s, ctx)[0].message).toContain("before it was prepared");
  });

  it("lists every date in reading order for stepping through", () => {
    const s = sheet({ B1: d(2026, 1, 5), A2: "9-Jan-26", A1: "Label" }, "T", { formats: { B1: DATE } });
    expect(dateCells(s).map(x => x.cell)).toEqual(["B1", "A2"]);
  });
});

describe("numbers stored as text", () => {
  it("flags a text number among real numbers", () => {
    const s = sheet({ A1: 100, A2: 200, A3: "1,250", A4: 300 });
    expect(getCheck("text-numbers").run(s).map(f => f.cell)).toEqual(["A3"]);
  });

  it("leaves ID columns alone", () => {
    expect(getCheck("text-numbers").run(sheet({ A1: "1041", A2: "1042", A3: "1043" }))).toEqual([]);
  });
});

describe("stray numbers", () => {
  const base = {
    A1: "Item", B1: "Amount", C1: "Tick",
    A2: "Cash", B2: 100, C2: "✓",
    A3: "AR", B3: 200, C3: "✓",
    A5: "Total", B5: "=SUM(B2:B3)",
  };

  it("flags an unlabelled number off to the side", () => {
    expect(getCheck("stray").run(sheet({ ...base, G3: 1250 })).map(f => f.cell)).toEqual(["G3"]);
  });

  it("flags an isolated unlabelled number below the work", () => {
    expect(getCheck("stray").run(sheet({ ...base, B9: 77 })).map(f => f.cell)).toEqual(["B9"]);
  });

  it("accepts labelled numbers", () => {
    expect(getCheck("stray").run(sheet({ ...base, A8: "Population", B8: 5000 }))).toEqual([]);
  });
});

describe("leftover highlighting", () => {
  it("knows highlighter colours from heading fills", () => {
    expect(isHighlighter("#FFFF00")).toBe(true);
    expect(isHighlighter("#FFFF99")).toBe(true);
    expect(isHighlighter("#E8F1EE")).toBe(false);
    expect(isHighlighter("#D9D9D9")).toBe(false);
    expect(isHighlighter("#1F4E78")).toBe(false);
  });

  it("flags highlighted cells, grouped by colour", () => {
    const s = sheet({ A1: "x", B2: 5, B3: 6 }, "T", { fills: { B2: "#FFFF00", B3: "#FFFF00" } });
    expect(getCheck("highlights").run(s)[0].message).toBe("2 cells highlighted yellow: B2, B3");
  });

  it("accepts highlighting explained by a colour legend", () => {
    const s = sheet({ A1: "Yellow = agreed to bank", B2: 5 }, "T", { fills: { B2: "#FFFF00" } });
    expect(getCheck("highlights").run(s)).toEqual([]);
  });
});

describe("sign-off dates and tickmark legend", () => {
  it("wants a date on the preparer sign-off", () => {
    const f = getCheck("signoff").run(sheet({ A1: "Prepared by: AB", A2: "Reviewed by:" }));
    expect(f.map(x => x.message)).toEqual(['"Prepared by" has no date']);
  });

  it("flags a tickmark symbol missing from the legend", () => {
    const s = sheet({ A1: 5, B1: "✓", B2: "^", A5: "Tickmark legend", A6: "✓ = agreed to invoice" });
    expect(getCheck("tickmarks").run(s).map(f => f.message)).toEqual(['"^" is used but not explained in the legend']);
  });
});

describe("sign-off labels", () => {
  it("ignores prose that mentions the reviewer", () => {
    const f = getCheck("signoff").run(sheet({ A1: "Prepared by: AB 05/01/2026", A2: "Reviewed by:", A4: "Send to the reviewer when done" }));
    expect(f).toEqual([]);
  });
});
