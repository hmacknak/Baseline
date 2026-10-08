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
async function editCell(a, text) {
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
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>Live review feedback for junior staff, inside Excel</p>
  <div class="note">Demo · simulated spreadsheet, real add-in</div></div>`);
workbook.active = "FY25 Cash";
await page.evaluate(wb => window.__init(wb), workbook);
await addin.locator(".item").first().waitFor();
await wait(3500);
await card(null, 900);

// ─── 1. Suggests the right reviewer ──────────────────────────────────────────
await caption("A junior is preparing the bank reconciliation.<small>Baseline reviews it the way a senior would, live, as they work.</small>", 4000);
await moveTo(addin.locator("#suggest"), 0.3, 0.3);
await caption("It recognises the procedure and suggests the right <b>reviewer</b> from its catalog.", 3600);
await click(addin.locator("#suggest button.primary"), { after: 900 });

// ─── 2. Settings once per workbook ───────────────────────────────────────────
await caption("Set the year-end and threshold once. They're saved in the workbook.", 600);
await click(addin.locator('.tab[data-tab="settings"]'), { after: 500 });
await click(addin.locator("#yearEndHint button"), { after: 400 });
await click(addin.locator("#threshold"), { after: 200 });
for (const ch of "10000") { await addin.locator("#threshold").press(ch); await wait(90); }
await click(addin.locator("#saveSettings"), { after: 600 });
await click(addin.locator('.tab[data-tab="review"]'), { after: 600 });
await panelTop();

// ─── 3. Insights ─────────────────────────────────────────────────────────────
await moveTo(addin.locator("#insights"), 0.3, 0.06);
await caption("<b>Insights</b>: the high-risk items a senior would ask about.<small>The rec is $1,200 off the GL, a cheque cleared before year-end, a stale cheque, a large uncleared item.</small>", 5200);
await click(addin.locator("#insights .cell-link").first(), { after: 500 });
await caption("Click any item to jump straight to it.", 2600);

// ─── 4. Last year's notes ────────────────────────────────────────────────────
await caption("Load last year's review notes…", 400);
await click(addin.locator('.tab[data-tab="notes"]'), { after: 500 });
await click(addin.locator("#importNotes"), { after: 1100 });
await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
await panelTop();
await moveTo(addin.locator("#pyPanel"), 0.3, 0.2);
await caption("…and Baseline shows which of last year's issues are <b>happening again</b>.", 4000);

// ─── 5. Fix it live ──────────────────────────────────────────────────────────
await caption("The junior fixes the total…", 500);
await editCell("B10", "=SUM(B6:B9)");
await panelTop();
await caption("…the rec now agrees to the GL, and last year's footing note shows <b>Looks fixed</b>.", 4400);

// ─── 6. The catalog ──────────────────────────────────────────────────────────
await caption(null);
await click(addin.locator('.tab[data-tab="catalog"]'), { after: 700 });
await caption("A <b>catalog of reviewers</b>, one for each procedure. Workpaper basics always run.", 3800);
await addin.locator(".cat-row", { hasText: "Liabilities" }).scrollIntoViewIfNeeded();
await click(addin.locator(".card", { hasText: "Search for unrecorded liabilities" }), { after: 700 });
await panelTop();
await caption("Each one knows its test: what must be filled in, and what's risky.", 3800);

// ─── 7. SURL ─────────────────────────────────────────────────────────────────
await caption(null);
await activate("FY25 SURL");
await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
await click(addin.locator("#suggest button.primary"), { after: 1000 });
await panelTop();
await moveTo(addin.locator("#insights"), 0.3, 0.1);
await caption("On the search for unrecorded liabilities: a <b>January invoice for December work</b> that never made it into AP…", 4600);
await addin.locator("#procList").scrollIntoViewIfNeeded();
await wait(500);
await moveTo(addin.locator("#procList"), 0.3, 0.2);
await caption("…and a reminder that row 10 is <b>missing its service date and conclusion</b>.", 4200);

// ─── 8. Revenue cut-off ──────────────────────────────────────────────────────
await caption(null);
await activate("FY25 Revenue cut-off");
await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
await click(addin.locator("#suggest button.primary"), { after: 1000 });
await panelTop();
await moveTo(addin.locator("#insights"), 0.3, 0.1);
await caption("Revenue cut-off: sales booked in the <b>wrong year</b>, both ways, plus a credit note after year-end.", 5000);
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
