// Records the Baseline demo video: simulated Excel on the left, the real add-in on the right.
//
//   npm run build && python3 demo/video/export_workbook.py
//   npm i --no-save playwright-core && node demo/video/record.mjs
//   ffmpeg -i demo/video/out/raw.webm -c:v libx264 -crf 18 -pix_fmt yuv420p -movflags +faststart demo/Baseline-demo-video.mp4
//
// Set CHROMIUM_PATH if Chromium isn't at the default location.
import { chromium } from "playwright-core";
import fs from "fs";
import path from "path";

const DIR = path.dirname(new URL(import.meta.url).pathname);
const DIST = path.resolve(DIR, "../../dist");
const OUT = path.join(DIR, "out");
fs.mkdirSync(OUT, { recursive: true });

const files = {
  "/studio.html": [path.join(DIR, "studio.html"), "text/html"],
  "/office.js": [path.join(DIR, "mock-office.js"), "application/javascript"],
};
const workbook = JSON.parse(fs.readFileSync(path.join(DIR, "workbook.json"), "utf8"));

const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || "/opt/pw-browsers/chromium" });
const context = await browser.newContext({
  viewport: { width: 1280, height: 720 },
  recordVideo: { dir: OUT, size: { width: 1280, height: 720 } },
});
await context.route("**/*", route => {
  const url = new URL(route.request().url());
  if (url.hostname === "appsforoffice.microsoft.com") {
    return route.fulfill({ contentType: "application/javascript", body: fs.readFileSync(files["/office.js"][0]) });
  }
  if (url.hostname !== "demo.local") return route.abort();
  if (files[url.pathname]) return route.fulfill({ path: files[url.pathname][0], contentType: files[url.pathname][1] });
  if (url.pathname.startsWith("/addin/")) {
    const p = path.join(DIST, url.pathname.slice("/addin/".length));
    if (fs.existsSync(p)) return route.fulfill({ path: p });
  }
  return route.fulfill({ status: 404, body: "not found" });
});

const page = await context.newPage();
const errors = [];
page.on("pageerror", e => errors.push(String(e)));
await page.goto("http://demo.local/studio.html");
const addin = page.frameLocator("#addin");

const wait = ms => page.waitForTimeout(ms);
const caption = async (html, ms = 0) => { await page.evaluate(h => window.__caption(h), html); if (ms) await wait(ms); };
const card = async (html, ms = 0) => { await page.evaluate(h => window.__card(h), html); if (ms) await wait(ms); };

async function moveTo(locator, dx = 0.5, dy = 0.5) {
  const b = await locator.boundingBox();
  const x = b.x + b.width * dx, y = b.y + b.height * dy;
  await page.evaluate(([x, y]) => window.__cursor(x, y), [x, y]);
  await wait(800);
  return { x, y };
}
async function click(locator, opts = {}) {
  const { x, y } = await moveTo(locator, opts.dx ?? 0.5, opts.dy ?? 0.5);
  await page.evaluate(([x, y]) => window.__ripple(x, y), [x, y]);
  await locator.click();
  await wait(opts.after ?? 500);
}
async function editCell(a, text, opts = {}) {
  const td = page.locator(`#cell-${a}`);
  const { x, y } = await moveTo(td, 0.5, 0.5);
  await page.evaluate(([x, y]) => window.__ripple(x, y), [x, y]);
  await page.evaluate(a => window.__select(a), a);
  await wait(300);
  for (let i = 1; i <= text.length; i++) {
    await page.evaluate(t => window.__typing(t), text.slice(0, i));
    await wait(55);
  }
  await wait(350);
  await page.evaluate(([a, t, value, fmt]) => {
    const v = value !== undefined ? value : /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t;
    window.__setCell(a, v, fmt);
  }, [a, text, opts.value, opts.fmt]);
  await wait(1000); // debounce (600 ms) + render
}
async function activate(name) {
  const tab = page.locator(`.tabs .t[data-sheet="${name}"]`);
  const { x, y } = await moveTo(tab);
  await page.evaluate(([x, y]) => window.__ripple(x, y), [x, y]);
  await page.evaluate(n => window.__activate(n), name);
  await wait(1100);
}
const panelTop = async () => {
  await addin.locator("body").evaluate(() => window.scrollTo({ top: 0, behavior: "smooth" }));
  await wait(600);
};
const item = title => addin.locator(".item", { hasText: title }).locator(".item-head");
const insight = title => addin.locator(".insight", { hasText: title });
async function show(locator, block = "start") {
  await locator.evaluate((el, b) => el.scrollIntoView({ behavior: "smooth", block: b }), block);
  await wait(700);
}
async function useSuggested() {
  await panelTop();
  await moveTo(addin.locator("#suggest"), 0.3, 0.4);
  await caption("Baseline recognises the workpaper and <b>suggests the right reviewer</b>.", 3200);
  await click(addin.locator("#suggest button.primary"), { after: 1200 });
}

// ─── Title card ───────────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>A first-pass review, live in Excel: workpaper basics and audit procedures</p>
  <div class="note">Demo · simulated spreadsheet, real add-in · all names fictional</div></div>`);
workbook.active = "AP-3 SURL";
await page.evaluate(wb => window.__init(wb), workbook);
await addin.locator(".item").first().waitFor();
await wait(3500);
await card(null, 900);

// ─── 1. The workpaper ───────────────────────────────────────────────────────
await caption("An associate is preparing <b>AP-3, the search for unrecorded liabilities</b>.<small>Baseline reviews it the way a senior would on a first pass, live, as they work.</small>", 4800);

// ─── 2. Year-end, once ──────────────────────────────────────────────────────
await caption("Set the year-end and threshold once. They're saved in the workbook.", 600);
await click(addin.locator('.tab[data-tab="settings"]'), { after: 500 });
await click(addin.locator("#yearEndHint button"), { after: 400 });
await click(addin.locator("#threshold"), { after: 200 });
await addin.locator("#threshold").type("10000", { delay: 60 });
await click(addin.locator("#saveSettings"), { after: 600 });
await click(addin.locator('.tab[data-tab="review"]'), { after: 600 });
await panelTop();

// ─── 3. What to fix, most important first ────────────────────────────────────
await moveTo(addin.locator("#score"), 0.2, 0.5);
await caption("Everything a reviewer would send back, <b>most important first</b>.", 3800);
await moveTo(addin.locator("#checklist"), 0.4, 0.3);
await caption("Hard-coded numbers, a short total, a number stored as text, blanks in the testing table…", 4400);
await click(item("No hard-coded numbers inside formulas"), { after: 600 });
await moveTo(addin.locator("#checklist .item-body").first(), 0.4, 0.4);
await caption("Each flag explains <b>why a reviewer cares</b> and <b>how to fix it</b>.", 4600);
await click(addin.locator("#checklist .item-body .cell-link").first(), { after: 600 });
await caption("One click jumps to the cell.", 2600);
await click(item("No hard-coded numbers inside formulas"), { after: 300 });

// ─── 4. Dates ────────────────────────────────────────────────────────────────
await caption("Dates get their own pass.", 400);
await click(addin.locator('.step-mode[data-mode="dates"]'), { after: 400 });
for (let i = 0; i < 3; i++) await click(addin.locator("#stepNext"), { after: 500 });
await caption("<b>Step through every date</b> on the sheet, one at a time.", 3400);
await click(item("Dates are consistent"), { after: 600 });
await caption("And Baseline spots the odd one out: <b>a 2016 date in a column of 2026 dates</b>.", 4400);
await click(item("Dates are consistent"), { after: 300 });

// ─── 5. Tidiness and Mark OK ─────────────────────────────────────────────────
await caption("It also catches loose ends: <b>leftover highlighting</b> and <b>stray numbers off to the side</b>.", 4200);
await click(item("No leftover highlighting"), { after: 500 });
await click(addin.locator("#checklist .item", { hasText: "No leftover highlighting" }).locator(".ok-link").first(), { after: 300 });
await addin.locator(".ok-reason").type("Agreed to client follow-up, 14 Jan", { delay: 35 });
await caption("If a flag is fine, <b>Mark OK with a reason</b>. The reviewer can see it.", 1800);
await click(addin.locator(".ok-form .primary"), { after: 1200 });

// ─── 6. Last year's notes ────────────────────────────────────────────────────
await caption("Load last year's review notes…", 400);
await click(addin.locator('.tab[data-tab="notes"]'), { after: 500 });
await click(addin.locator("#importNotes"), { after: 1100 });
await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
await panelTop();
await moveTo(addin.locator("#pyPanel"), 0.3, 0.3);
await caption("…and Baseline shows which of <b>last year's issues are happening again</b>.", 4400);

// ─── 7. Clear the points ─────────────────────────────────────────────────────
await caption("The associate clears the points before the senior ever opens the file…", 400);
await editCell("D11", "05-Jan-26", { value: 46027, fmt: "dd-mmm-yy" });
await editCell("I11", "Jan freight, FY26");
await editCell("I16", "Exception: Dec service, not in AP");
await editCell("F13", "2150", { fmt: "#,##0;(#,##0)" });
await editCell("E15", "10-Feb-26", { value: 46063, fmt: "dd-mmm-yy" });
await editCell("F17", "=SUM(F8:F16)");
await editCell("F20", "=F17/F19");
await editCell("K12", "");
await editCell("A22", "Conclusion: two December liabilities not recorded; see exceptions.");
await editCell("A23", "Tickmark legend: ✓ = agreed to vendor invoice and cleared cheque");
await panelTop();
await moveTo(addin.locator("#score"), 0.2, 0.5);
await caption("<b>Ready for review.</b> Last year's notes now show <b>Looks fixed</b>.", 4600);

// ─── 8. The procedure: SURL ─────────────────────────────────────────────────
await caption("The basics are clean. Now <b>the procedure itself</b>.", 2600);
await useSuggested();
await show(insight("Invoices after year-end for this year's services"));
await moveTo(insight("Invoices after year-end for this year's services").locator(".insight-item").first(), 0.4, 0.5);
await caption("It reads the testing table: <b>January invoices for December services</b> that were never recorded.", 4600);
await moveTo(insight("Invoices after year-end for this year's services").locator(".why-block"), 0.4, 0.4);
await caption("Every finding says <b>why it matters</b>, the <b>assertion</b> at risk and the <b>standard</b> behind it.", 5200);

// ─── 9. Cash confirmations ───────────────────────────────────────────────────
await caption(null, 200);
await activate("Cash confirmations");
await useSuggested();
await show(insight("Not sent or received directly"));
await moveTo(insight("Not sent or received directly").locator(".insight-item").first(), 0.4, 0.5);
await caption("Cash confirmations: a bank reply <b>forwarded by the client</b> isn't direct evidence.", 4400);
await moveTo(insight("Not sent or received directly").locator(".why-meta"), 0.4, 0.5);
await caption("ISA 505 and PCAOB AS 2310 require the auditor to <b>control the confirmation</b>.", 4600);
await show(insight("Overdrawn accounts"), "center");
await caption("It also catches unexplained differences, missing replies, unconfirmed accounts and overdrafts.", 4200);

// ─── 10. Expense vouching ────────────────────────────────────────────────────
await caption(null, 200);
await activate("Expense vouching");
await useSuggested();
await show(insight("Possible duplicate payments"));
await moveTo(insight("Possible duplicate payments").locator(".insight-item").first(), 0.4, 0.5);
await caption("Expense vouching: <b>the same invoice paid twice</b>, sitting in the sample.", 4200);
await moveTo(insight("Possible duplicate payments").locator(".py-callout, .why-block").first(), 0.4, 0.4);
await caption("And it remembers <b>last year's review note</b> about exactly this.", 4400);
await show(insight("Looks like a capital item"), "center");
await caption("Plus amounts that don't agree, cut-off, missing approval and <b>a forklift booked as an expense</b>.", 4600);

// ─── 11. Whole workbook ──────────────────────────────────────────────────────
await caption(null, 200);
await click(addin.locator("#reviewAll"), { after: 1200 });
await caption("One click reviews <b>every sheet</b> before it goes to the senior.", 4600);
await caption(null, 400);

// ─── End card ────────────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>Complete workpapers. Fewer review notes. Juniors who learn as they go.</p></div>`, 4500);

const video = page.video();
await context.close();
await browser.close();
const src = await video.path();
fs.copyFileSync(src, path.join(OUT, "raw.webm"));
console.log("errors:", errors);
console.log("video:", path.join(OUT, "raw.webm"));
