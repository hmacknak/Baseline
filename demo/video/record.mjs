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

// ─── Title card ───────────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>AP-3 Search for unrecorded liabilities</p>
  <div class="note">Demo · simulated spreadsheet, real add-in</div></div>`);
workbook.active = "AP-3 SURL";
await page.evaluate(wb => window.__init(wb), workbook);
await addin.locator(".item").first().waitFor();
await wait(3500);
await card(null, 900);

// ─── 1. The workpaper and the reviewer ───────────────────────────────────────
await caption("A first-year associate is preparing <b>AP-3, the search for unrecorded liabilities</b>.<small>Nine January and February payments selected for testing.</small>", 4600);
await moveTo(addin.locator("#suggest"), 0.3, 0.3);
await caption("Baseline recognises the procedure and offers the <b>SURL reviewer</b>.", 3400);
await click(addin.locator("#suggest button.primary"), { after: 900 });

// ─── 2. Engagement settings ──────────────────────────────────────────────────
await caption("Year-end and threshold are set once for the engagement.", 600);
await click(addin.locator('.tab[data-tab="settings"]'), { after: 500 });
await click(addin.locator("#yearEndHint button"), { after: 400 });
await click(addin.locator("#threshold"), { after: 200 });
for (const ch of "10000") { await addin.locator("#threshold").press(ch); await wait(90); }
await click(addin.locator("#saveSettings"), { after: 600 });
await click(addin.locator('.tab[data-tab="review"]'), { after: 600 });
await panelTop();

// ─── 3. Insights: the exceptions ─────────────────────────────────────────────
await moveTo(addin.locator("#insights"), 0.3, 0.08);
await caption("Baseline tests every selection against year-end.<small>Two January invoices are for <b>December services</b> and were never recorded.</small>", 5200);
await moveTo(addin.locator("#insights .over").first(), 0.4, 0.5);
await caption("Together <b>$29,650</b>, above the $10,000 threshold.", 3600);
await click(addin.locator("#insights .cell-link").first(), { after: 500 });
await caption("Click any item to jump straight to the selection.", 3000);

// ─── 4. Procedure completeness ───────────────────────────────────────────────
await addin.locator("#procList").scrollIntoViewIfNeeded();
await wait(400);
await moveTo(addin.locator("#procList"), 0.3, 0.2);
await caption("It also checks the testing is complete:<small>Delta Freight has no service date, and two selections have no conclusion.</small>", 4600);

// ─── 5. Workpaper basics ─────────────────────────────────────────────────────
await addin.locator("#checklist").scrollIntoViewIfNeeded();
await wait(500);
await moveTo(addin.locator("#checklist"), 0.4, 0.35);
await caption("<b>Workpaper basics</b> run on every sheet, like a senior's first pass:<small>a total that stops a row short, a hard-coded number, tickmarks with no legend.</small>", 5200);
await click(item("No hard-coded numbers inside formulas"), { after: 600 });
await addin.locator(".item.open, #checklist .item-body").first().scrollIntoViewIfNeeded().catch(() => {});
await wait(400);
await moveTo(addin.locator("#checklist .item-body").first(), 0.4, 0.4);
await caption("Each flag explains <b>why a reviewer cares</b> and <b>how to fix it</b>.<small>Here the coverage calculation hard-codes the population instead of linking to it.</small>", 5600);
await click(addin.locator("#checklist .item-body .cell-link").first(), { after: 600 });
await caption("One click takes the associate to the cell.", 2800);

// ─── 6. Last year's review notes ─────────────────────────────────────────────
await caption("Load last year's review notes for this workpaper…", 400);
await click(addin.locator('.tab[data-tab="notes"]'), { after: 500 });
await click(addin.locator("#importNotes"), { after: 1100 });
await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
await panelTop();
await moveTo(addin.locator("#pyPanel"), 0.3, 0.3);
await caption("…last year's reviewer raised <b>the same issues</b>. Baseline flags that they are happening again.", 4600);

// ─── 7. The associate clears the review points ──────────────────────────────
await caption("The associate clears the points before the senior ever opens the file…", 400);
await editCell("D11", "05-Jan-26", { value: 46027, fmt: "dd-mmm-yy" });
await editCell("I11", "Jan freight, FY26");
await editCell("I16", "Exception: Dec service, not in AP");
await editCell("F17", "=SUM(F8:F16)");
await editCell("F20", "=F17/F19");
await editCell("A23", "Tickmark legend: ✓ = agreed to vendor invoice and cleared cheque");
await panelTop();
await moveTo(addin.locator("#score"), 0.15, 0.5);
await caption("<b>All checks clear.</b> Last year's notes on service dates and the tickmark legend now show <b>Looks fixed</b>.", 4800);

// ─── 8. What's left is the real finding ─────────────────────────────────────
await moveTo(addin.locator("#insights"), 0.3, 0.08);
await caption("What's left is the real finding: <b>unrecorded liabilities</b>, documented and ready for the summary of unadjusted differences.", 5600);
await caption(null, 400);

// ─── End card ────────────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>Complete workpapers. Sharper testing. Juniors who learn as they go.</p>
  <p style="margin-top:22px;font-size:20px">hmacknak.github.io/Baseline</p></div>`, 4500);

const video = page.video();
await context.close();
await browser.close();
const src = await video.path();
fs.copyFileSync(src, path.join(OUT, "raw.webm"));
console.log("errors:", errors);
console.log("video:", path.join(OUT, "raw.webm"));
