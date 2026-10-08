import { describe, expect, it } from "vitest";
import { getCheck, runChecklist } from "../src/review/checks";
import {
  normaliseSheetName,
  notesForSheet,
  parseNotesRows,
  parseNotesText,
  pyNotesForCheck,
  pyStatuses,
  tagNote,
} from "../src/review/notes";
import { sheet } from "./helpers";

describe("tagNote", () => {
  it.each([
    ["Hard-coded growth rate, link to source", "hardcodes"],
    ["Total doesn't foot", "totals"],
    ["Please sign off and date", "signoff"],
    ["Where is the tickmark legend?", "tickmarks"],
    ["#REF! in column D", "errors"],
    ["Formula overwritten with a typed number", "overwritten"],
  ])("%s → %s", (text, tag) => {
    expect(tagNote(text)).toContain(tag);
  });

  it('treats "link to source" as a hard-code note only', () => {
    expect(tagNote("Hard-coded rate, link to source")).toEqual(["hardcodes"]);
    expect(tagNote("No source documented")).toEqual(["header"]);
  });

  it("leaves unrelated notes untagged", () => {
    expect(tagNote("Discuss with manager re: accrual judgement")).toEqual([]);
  });
});

describe("parseNotesText", () => {
  it("supports sheet|cell|note, sheet|note, cell|note and bare notes", () => {
    const notes = parseNotesText(
      ["Cash | B12 | Hard-coded FX rate", "Cash | Missing sign-off", "C4 | Total doesn't foot", "Purpose unclear", ""].join("\n"),
    );
    expect(notes.map(n => [n.sheet, n.cell, n.text])).toEqual([
      ["Cash", "B12", "Hard-coded FX rate"],
      ["Cash", "", "Missing sign-off"],
      ["", "C4", "Total doesn't foot"],
      ["", "", "Purpose unclear"],
    ]);
  });

  it("accepts tab-separated rows pasted from Excel", () => {
    expect(parseNotesText("AR\tD7\tHard code")[0]).toMatchObject({ sheet: "AR", cell: "D7", text: "Hard code" });
  });
});

describe("parseNotesRows", () => {
  it("matches columns by header name", () => {
    const notes = parseNotesRows([
      ["Review note", "Workpaper", "Cell"],
      ["Hard-coded rate", "Cash", "B2"],
      ["", "Cash", ""],
    ]);
    expect(notes).toEqual([{ sheet: "Cash", cell: "B2", text: "Hard-coded rate", tags: ["hardcodes"] }]);
  });

  it("falls back to Sheet, Cell, Note when there is no header", () => {
    expect(parseNotesRows([["Cash", "B2", "Sign off"]])[0].sheet).toBe("Cash");
  });
});

describe("sheet matching", () => {
  it("ignores years and FY labels", () => {
    expect(normaliseSheetName("FY24 Cash")).toBe("cash");
    expect(normaliseSheetName("Cash 2025")).toBe("cash");
    expect(normaliseSheetName("Cash (PY)")).toBe("cash");
  });

  it("includes notes with no sheet", () => {
    const notes = parseNotesText("Cash | a\nAR | b\nc");
    expect(notesForSheet(notes, "FY25 Cash").map(n => n.text)).toEqual(["a", "c"]);
  });
});

describe("pyStatuses", () => {
  const notes = parseNotesText(
    ["Cash | Hard-coded FX rate", "Cash | Please sign off", "Cash | Discuss accrual with manager"].join("\n"),
  );

  it("marks repeat issues, fixed issues and manual ones", () => {
    const s = sheet(
      { A1: "Purpose: cash", A2: "Source: bank", A3: "Prepared by: JS", A4: "Reviewed by: MK", B5: 2, C5: "=B5*1.3" },
      "Cash",
    );
    const statuses = pyStatuses(notes, "Cash", runChecklist(s)).map(x => x.status);
    expect(statuses).toEqual(["still-an-issue", "looks-fixed", "check-manually"]);
  });

  it("finds PY notes for a specific check", () => {
    expect(pyNotesForCheck(notes, "FY25 Cash", getCheck("hardcodes")).map(n => n.text)).toEqual(["Hard-coded FX rate"]);
  });
});

describe("tagNote precision", () => {
  it("doesn't treat every use of 'error' as a formula error", () => {
    expect(tagNote("Cut-off error not identified")).not.toContain("errors");
    expect(tagNote("Formula error in total")).toContain("errors");
  });
});
