import { describe, expect, it } from "vitest";
import { findHardcodedNumbers, hasExternalLink, lex, toR1C1 } from "../src/review/formula";

describe("lex", () => {
  it("separates refs, functions, numbers and strings", () => {
    const types = lex('=SUM(A1:B2)*1.05&"x"').map(t => t.type);
    expect(types).toEqual(["func", "op", "ref", "op", "ref", "op", "op", "number", "op", "string"]);
  });

  it("keeps quoted sheet names together", () => {
    const t = lex("='My Sheet'!A1+1");
    expect(t[0]).toEqual({ type: "sheet", text: "'My Sheet'!" });
  });
});

describe("findHardcodedNumbers", () => {
  it.each([
    ["=A1*1.05", ["1.05"]],
    ["=SUM(A1:A5)+500", ["500"]],
    ["=A1/12", ["12"]],
    ["=1000+2000", ["1000", "2000"]],
    ["=500", ["500"]],
    ["=A1*5%", ["5%"]],
    ["=A1-250", ["250"]],
  ])("flags %s", (formula, expected) => {
    expect(findHardcodedNumbers(formula)).toEqual(expected);
  });

  it.each([
    "=SUM(A1:A5)",
    "=ROUND(A1,2)",
    "=ROUND(A1,-2)",
    "=VLOOKUP(A1,Table,3,FALSE)",
    "=A1*1",
    "=A1+0",
    "=SUM(1:1)",
    "=SUM(Sheet2!A1:A3)",
    '=IF(A1="2024",B1,C1)',
    "=INDEX({1,2,3},2)",
    "=LOG10(A1)",
    "=1",
  ])("does not flag %s", formula => {
    expect(findHardcodedNumbers(formula)).toEqual([]);
  });
});

describe("hasExternalLink", () => {
  it("detects links to other workbooks", () => {
    expect(hasExternalLink("='[Budget.xlsx]Sheet1'!A1")).toBe(true);
    expect(hasExternalLink("=[1]Sheet1!A1")).toBe(true);
    expect(hasExternalLink("=[Budget.xlsx]Sheet1!A1*2")).toBe(true);
  });

  it("ignores same-workbook refs and structured references", () => {
    expect(hasExternalLink("=Sheet2!A1")).toBe(false);
    expect(hasExternalLink("=SUM(Table1[Amount])")).toBe(false);
    expect(hasExternalLink("=Table1[@Amount]*2")).toBe(false);
  });
});

describe("toR1C1", () => {
  it("makes copied formulas compare equal", () => {
    // B5 =SUM(B2:B4) copied to C5 =SUM(C2:C4)
    expect(toR1C1("=SUM(B2:B4)", 4, 1)).toBe(toR1C1("=SUM(C2:C4)", 4, 2));
  });

  it("distinguishes a range that stops short", () => {
    expect(toR1C1("=SUM(B2:B4)", 4, 1)).not.toBe(toR1C1("=SUM(C2:C3)", 4, 2));
  });

  it("respects absolute references", () => {
    expect(toR1C1("=A1*$F$1", 0, 1)).toBe("=RC[-1]*R1C6");
    expect(toR1C1("=A2*$F$1", 1, 1)).toBe("=RC[-1]*R1C6");
  });
});
