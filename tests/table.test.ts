import { describe, expect, it } from "vitest";
import { findYearEndInText, formatSerial, serialFromYMD, toAmount, toSerial } from "../src/review/dates";
import { flagOf, readTable } from "../src/review/table";
import { sheet } from "./helpers";

const YE = serialFromYMD(2025, 12, 31);

describe("dates", () => {
  it("matches Excel serials", () => {
    expect(YE).toBe(46022);
    expect(formatSerial(46022)).toBe("31 Dec 2025");
  });

  it.each([
    [46022, YE],
    ["2025-12-31", YE],
    ["31-Dec-2025", YE],
    ["31 Dec 2025", YE],
    ["Dec 31, 2025", YE],
    ["12/31/2025", YE],
    ["31/12/2025", YE],
    ["01/02/2026", serialFromYMD(2026, 1, 2)],
  ])("reads %s", (input, expected) => {
    expect(toSerial(input)).toBe(expected);
  });

  it("ignores things that aren't dates", () => {
    expect(toSerial(1200)).toBeNull();
    expect(toSerial("Acme Ltd")).toBeNull();
    expect(toSerial("")).toBeNull();
  });

  it.each([
    [1234.5, 1234.5],
    ["$1,234.50", 1234.5],
    ["(1,234)", -1234],
    ["-500", -500],
    ["n/a", null],
  ])("amount %s", (input, expected) => {
    expect(toAmount(input)).toBe(expected);
  });

  it("finds a year-end in a title", () => {
    expect(findYearEndInText(["Cash: bank reconciliation, 31 Dec 2025"])).toBe(YE);
    expect(findYearEndInText(["Year ended December 31, 2025"])).toBe(YE);
    expect(findYearEndInText(["Meeting on 15 Dec 2025"])).toBeNull();
  });
});

describe("readTable", () => {
  const specs = [
    { key: "vendor", label: "vendor", kind: "text" as const, names: ["vendor", "supplier"] },
    { key: "invoiceDate", label: "invoice date", kind: "date" as const, names: ["invoice date", "inv date"] },
    { key: "amount", label: "amount", kind: "amount" as const, names: ["amount", "invoice amount"] },
  ];

  it("finds the header row below a title and stops at Total", () => {
    const s = sheet({
      A1: "SURL testing",
      A3: "Supplier", B3: "Inv date", C3: "Invoice amount",
      A4: "Acme", B4: YE + 5, C4: 100,
      A5: "Bolt Co", C5: 200,
      A6: "Total", C6: "=SUM(C4:C5)",
    });
    const t = readTable(s, specs)!;
    expect(t.headerRow).toBe(2);
    expect(t.columns).toEqual({ vendor: 0, invoiceDate: 1, amount: 2 });
    expect(t.rows.map(r => r.row)).toEqual([3, 4]);
  });

  it("returns null when the headers aren't there", () => {
    expect(readTable(sheet({ A1: "Item", B1: "Amount" }), specs, 3)).toBeNull();
  });

  it("reads yes/no flags", () => {
    const s = sheet({ A1: "Vendor", B1: "Invoice date", C1: "Amount", A2: "Y", A3: "No", A4: "maybe" });
    const t = readTable(s, specs)!;
    expect(t.rows.map(r => flagOf(r, "vendor"))).toEqual([true, false, null]);
  });
});
