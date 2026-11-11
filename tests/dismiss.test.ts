import { describe, expect, it } from "vitest";
import { runChecklist } from "../src/review/checks";
import { applyDismissals, findingKey, fingerprint } from "../src/review/dismiss";
import { sheet } from "./helpers";

describe("Mark OK", () => {
  const s = sheet({ A1: 100, B1: "=A1*1.05" }, "FY25 Cash");
  const results = runChecklist(s);
  const hard = results.filter(r => r.check.id === "hardcodes")[0].findings[0];

  it("moves an accepted flag out of the open list", () => {
    const d = { [findingKey(s.name, hard)]: { reason: "Contractual 5% uplift", fingerprint: fingerprint(s, hard), date: "2026-01-10" } };
    const r = applyDismissals(s, results, d).filter(x => x.check.id === "hardcodes")[0];
    expect(r.findings).toEqual([]);
    expect(r.dismissed[0].dismissal.reason).toBe("Contractual 5% uplift");
  });

  it("brings the flag back when the cell changes", () => {
    const d = { [findingKey(s.name, hard)]: { reason: "ok", fingerprint: fingerprint(s, hard), date: "2026-01-10" } };
    const changed = sheet({ A1: 100, B1: "=A1*1.07" }, "FY25 Cash");
    const r = applyDismissals(changed, runChecklist(changed), d).filter(x => x.check.id === "hardcodes")[0];
    expect(r.findings).toHaveLength(1);
  });
});
