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

const browser = await chromium.launch({ executablePath: process.env.CHROMIUM_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome" });
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
async function editCell(a, text, ms = 55) {
  const td = page.locator(`#cell-${a}`);
  const { x, y } = await moveTo(td, 0.5, 0.5);
  await page.evaluate(([x, y]) => window.__ripple(x, y), [x, y]);
  await page.evaluate(a => window.__select(a), a);
  await wait(300);
  for (let i = 1; i <= text.length; i++) {
    await page.evaluate(t => window.__typing(t), text.slice(0, i));
    await wait(ms);
  }
  await wait(350);
  await page.evaluate(([a, t]) => {
    const v = /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t;
    window.__setCell(a, v);
  }, [a, text]);
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

// ─── Title card ───────────────────────────────────────────────────────────────
await card(`<div><div class="logo">✓</div><h1>Baseline Review</h1>
  <p>Live review feedback for junior staff, inside Excel</p>
  <div class="note">Demo · simulated spreadsheet, real add-in</div></div>`);
workbook.active = "FY25 Cash";
await page.evaluate(wb => window.__init(wb), workbook);
await addin.locator(".item").first().waitFor();
await wait(3500);
await card(null, 900);

// ─── 1. Live review ──────────────────────────────────────────────────────────
await caption("A junior is preparing the year-end bank reconciliation.<small>Baseline reviews the sheet the way a senior would, live, as they work.</small>", 4200);
await moveTo(addin.locator("#score"), 0.15, 0.5);
await caption("Before anyone opens it for review, Baseline has already found <b>5 issues</b>.", 3600);

await click(item("Totals include every line"), { after: 700 });
await caption("Each issue says <b>where</b> it is, <b>why it matters</b>, and <b>how to fix it</b>.", 4200);
await click(addin.locator(".item-body .cell-link").first(), { after: 600 });
await caption("Click the cell reference to jump straight to the problem.", 3000);

// ─── 2. Last year's notes ────────────────────────────────────────────────────
await caption("Now load <b>last year's review notes</b>, pasted in or imported from a sheet.", 1200);
await click(addin.locator('.tab[data-tab="notes"]'), { after: 700 });
await click(addin.locator("#importNotes"), { after: 1200 });
await addin.locator(".notes-table").scrollIntoViewIfNeeded();
await wait(1500);
await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
await panelTop();
await caption("Baseline matches them to this year's sheet:<small>4 of last year's 6 points are happening again. The other 2 need your judgement.</small>", 4800);
await moveTo(addin.locator(".py-callout"), 0.3, 0.5);
await caption("Heads up, <b>this happened last year</b>.", 3000);

// ─── 3. Fix it live ──────────────────────────────────────────────────────────
await caption("The junior fixes the total…", 600);
await editCell("C14", "=SUM(C10:C13)");
await panelTop();
await caption("…and the checklist updates instantly. Last year's footing note shows <b>Looks fixed</b>.", 3800);

await caption("Link the FX rate instead of hard-coding it…", 400);
await editCell("C21", "=C19*C20");
await caption("Add the reviewer line, a source and a tickmark legend…", 400);
await editCell("A4", "Reviewed by:");
await editCell("A5", "Source: RBC operating account statement at 31 Dec 2025 (p.3)", 38);
await editCell("A6", "Tickmark legend: ✓ agreed to bank statement or trial balance", 38);
await panelTop();
await moveTo(addin.locator("#score"), 0.15, 0.5);
await caption("<b>9 of 9 checks clear.</b> Ready for review.<small>The two Check yourself points stay listed for the junior to work through.</small>", 4800);

// ─── 4. Other sheets ─────────────────────────────────────────────────────────
await caption("Now the receivables aging and ECL allowance…", 300);
await activate("FY25 AR");
await panelTop();
await click(item("No numbers typed over formulas"), { after: 600 });
await caption("It catches a number <b>typed over a formula</b>, a formula pointing at the <b>wrong row</b>, and the <b>#DIV/0!</b> that follows.", 4600);
await click(addin.locator(".item", { hasText: "No numbers typed over formulas" }).locator(".cell-link").first(), { after: 1200 });
await click(item("Formulas consistent across rows and columns"), { after: 600 });
await click(addin.locator(".item", { hasText: "Formulas consistent across rows and columns" }).locator(".cell-link").first(), { after: 1500 });
await caption("The junior re-links the total and fixes the reference.", 500);
await editCell("F10", "=SUM(B10:E10)");
await editCell("G9", "=F9/F$15");
await panelTop();
await moveTo(addin.locator("#score"), 0.15, 0.5);
await caption("<b>9 of 9 again.</b> The 90+ loss rate and the Lakeshore provision stay listed as <b>Check yourself</b>.", 5200);
await caption(null);
await activate("FY25 Fixed Assets");
await panelTop();
await caption("A clean workpaper passes every check, and last year's notes show as fixed.", 5000);
await caption(null, 400);

// ─── End card ────────────────────────────────────────────────────────────────
await card(`<div><div class="logo">✓</div><h1>Baseline Review</h1>
  <p>Immediate feedback, so juniors build skills while the work gets done.</p>
  <p style="margin-top:22px;font-size:20px">hmacknak.github.io/Baseline</p></div>`, 4500);

const video = page.video();
await context.close();
await browser.close();
const src = await video.path();
fs.copyFileSync(src, path.join(OUT, "raw.webm"));
console.log("errors:", errors);
console.log("video:", path.join(OUT, "raw.webm"));
