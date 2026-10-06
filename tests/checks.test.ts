import { describe, expect, it } from "vitest";
import { getCheck, runChecklist } from "../src/review/checks";
import { CheckId } from "../src/review/types";
import { sheet } from "./helpers";

const run = (id: CheckId, cells: Record<string, string | number | boolean>) => getCheck(id).run(sheet(cells));

describe("errors", () => {
  it("flags error values", () => {
    const f = run("errors", { A1: "#REF!", A2: 3 });
    expect(f).toHaveLength(1);
    expect(f[0].cell).toBe("A1");
  });
});

describe("hardcodes", () => {
  it("flags numbers in formula maths", () => {
    const f = run("hardcodes", { A1: 100, B1: "=A1*1.05", C1: "=ROUND(B1,2)" });
    expect(f.map(x => x.cell)).toEqual(["B1"]);
  });
});

describe("overwritten", () => {
  it("flags a typed number between matching formulas", () => {
    const f = run("overwritten", {
      A1: 1, A2: 2, A3: 3,
      B1: "=A1*2", B2: 99, B3: "=A3*2",
    });
    expect(f.map(x => x.cell)).toEqual(["B2"]);
  });

  it("leaves ordinary input columns alone", () => {
    expect(run("overwritten", { A1: 1, A2: 2, A3: 3 })).toEqual([]);
  });
});

describe("inconsistent", () => {
  it("flags the odd formula out in a row", () => {
    const f = run("inconsistent", {
      B1: 1, C1: 1, D1: 1, B2: 1, C2: 1, D2: 1, B3: 1, C3: 1, D3: 1,
      B4: "=SUM(B1:B3)", C4: "=SUM(C1:C2)", D4: "=SUM(D1:D3)",
    });
    expect(f.map(x => x.cell)).toEqual(["C4"]);
  });

  it("is quiet when formulas match", () => {
    expect(
      run("inconsistent", { B1: 1, C1: 1, D1: 1, B2: "=B1*2", C2: "=C1*2", D2: "=D1*2" }),
    ).toEqual([]);
  });
});

describe("totals", () => {
  it("flags a vertical SUM that stops short", () => {
    const f = run("totals", { B1: 10, B2: 20, B3: 30, B4: "=SUM(B1:B2)" });
    expect(f).toHaveLength(1);
    expect(f[0].message).toContain("B3");
  });

  it("flags a horizontal SUM that stops short", () => {
    const f = run("totals", { A1: 1, B1: 2, C1: 3, D1: "=SUM(A1:B1)" });
    expect(f[0].message).toContain("C1");
  });

  it("is quiet for complete totals and blank spacer rows", () => {
    expect(run("totals", { B1: 10, B2: 20, B4: "=SUM(B1:B2)" })).toEqual([]);
    expect(run("totals", { B1: 10, B2: 20, B3: "=SUM(B1:B2)" })).toEqual([]);
  });
});

describe("external-links", () => {
  it("flags links to other files", () => {
    expect(run("external-links", { A1: "='[Budget.xlsx]Sheet1'!A1" })).toHaveLength(1);
    expect(run("external-links", { A1: "=Sheet2!A1" })).toEqual([]);
  });
});

describe("signoff", () => {
  it("flags missing sign-offs", () => {
    const f = run("signoff", { A1: "Cash lead", B2: 5 });
    expect(f).toHaveLength(2);
  });

  it("flags a blank Prepared by", () => {
    const f = run("signoff", { A1: "Prepared by:", A2: "Reviewed by:" });
    expect(f).toHaveLength(1);
    expect(f[0].cell).toBe("A1");
  });

  it("accepts filled sign-offs inline or in the next cell", () => {
    expect(run("signoff", { A1: "Prepared by: JS 01/10", A2: "Reviewed by:", B2: "MK" })).toEqual([]);
  });

  it("ignores empty sheets", () => {
    expect(run("signoff", {})).toEqual([]);
  });
});

describe("header", () => {
  it("wants purpose and source", () => {
    expect(run("header", { A1: "Cash" })).toHaveLength(2);
    expect(run("header", { A1: "Purpose: test cash", A2: "Source: bank statement" })).toEqual([]);
  });
});

describe("tickmarks", () => {
  it("flags tickmarks with no legend", () => {
    expect(run("tickmarks", { A1: 5, B1: "✓", B2: "✓" })).toHaveLength(1);
    expect(run("tickmarks", { A1: 5, B1: "✓", A5: "Tickmark legend" })).toEqual([]);
  });
});

describe("runChecklist", () => {
  it("returns one result per checklist item", () => {
    const results = runChecklist(sheet({ A1: "x" }));
    expect(results.map(r => r.check.id)).toEqual([
      "errors", "hardcodes", "overwritten", "totals", "inconsistent",
      "external-links", "signoff", "header", "tickmarks",
    ]);
  });
});
