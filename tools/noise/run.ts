// Runs every Baseline check over snapshots made by xlsx2snap.py and prints what fired, per check.
// See tools/noise/README.md.
import * as fs from "fs";
import { runChecklist } from "../../src/review/checks";
import { suggestReviewer } from "../../src/review/review";
const corpus = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const ctx = { yearEnd: 46022, threshold: 10000 };
const per: Record<string, { sheets: number; findings: number; samples: string[] }> = {};
let sheets = 0, crashes = 0, maxMs = 0, sugg: string[] = [];
for (const [file, list] of Object.entries<any>(corpus)) {
  if (!Array.isArray(list)) continue;
  for (const s of list) {
    if (!s.cells.length) continue;
    sheets++;
    const t = Date.now();
    try {
      for (const r of runChecklist(s, ctx)) {
        if (!r.findings.length) continue;
        const p = (per[r.check.id] ||= { sheets: 0, findings: 0, samples: [] });
        p.sheets++; p.findings += r.findings.length;
        if (p.samples.length < 12) p.samples.push(`${file.split("/").slice(-1)[0]}!${s.name} ${r.findings[0].cell || ""} ${r.findings[0].message}`.slice(0, 200));
      }
      const sg = suggestReviewer(s);
      if (sg) sugg.push(`${file.split("/").slice(-1)[0]}!${s.name} -> ${sg.reviewer.id} ${sg.score.toFixed(2)}`);
    } catch (e) { crashes++; console.log("CRASH", file, s.name, String(e).slice(0, 200)); }
    maxMs = Math.max(maxMs, Date.now() - t);
  }
}
console.log(`sheets=${sheets} crashes=${crashes} slowest=${maxMs}ms`);
for (const [id, p] of Object.entries(per).sort((a, b) => b[1].sheets - a[1].sheets)) {
  console.log(`\n## ${id}: ${p.sheets} sheets, ${p.findings} findings`);
  p.samples.forEach(x => console.log("   " + x));
}
console.log("\n## suggestions:", sugg.length); sugg.forEach(x => console.log("   " + x));
