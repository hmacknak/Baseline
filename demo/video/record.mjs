// Records the Baseline demo video: simulated Excel on the left, the real add-in on the right.
//
//   npm run build && python3 demo/video/export_workbook.py
//   npm i --no-save playwright-core && node demo/video/record.mjs full    (about 2.5 minutes)
//                                       node demo/video/record.mjs short   (under 1 minute)
//   ffmpeg -i demo/video/out/raw-full.webm -map_metadata -1 -c:v libx264 -crf 18 -pix_fmt yuv420p -movflags +faststart demo/Baseline-demo-video.mp4
//   ffmpeg -i demo/video/out/raw-short.webm -map_metadata -1 -c:v libx264 -crf 18 -pix_fmt yuv420p -movflags +faststart demo/Baseline-demo-short.mp4
//
// Storyboard: what Baseline is, then three parts (workpaper basics; procedure review and the why; last
// year's review notes). Year-end and threshold are pre-set, not clicked through.
//
// Set CHROMIUM_PATH if Chromium isn't at the default location.
import { chromium } from "playwright-core";
import fs from "fs";
import path from "path";

const MODE = process.argv[2] === "short" ? "short" : "full";
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

const opens = title => addin.locator(".item", { hasText: title });
const linkIn = (title, cell) => opens(title).locator(".cell-link", { hasText: new RegExp(`^${cell}$`) }).first();

// Year-end 31 Dec 2025 and a 12,500 threshold, as if set earlier in Settings.
await page.evaluate(() => {
  window.__seedSettings = { "baseline.context.v1": { yearEnd: 46022, threshold: 12500 } };
});

// The video says what Baseline is, then shows three things: the basics, the procedure review with its
// "why", and last year's review notes. Each part opens with a short title card.
const part = async (n, title, sub) => {
  await caption(null, 150);
  await card(`<div><div class="partno">${n}</div><h1 class="parttitle">${title}</h1><p>${sub}</p></div>`, 3000);
  await card(null, 800);
};
const SURL_FLAG = "Invoices after year-end for this year's services";

// ─── Cover ───────────────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>Live feedback on your workpaper, as you work.</p>
  <ul class="three"><li>Checks the basics</li><li>Reviews the procedure and explains why</li><li>Remembers last year's review notes</li></ul>
  <div class="note">Demo · simulated spreadsheet, real add-in · all names fictional</div></div>`);
workbook.active = "AP-3 SURL";
await page.evaluate(wb => window.__init(wb), workbook);
await addin.locator(".item").first().waitFor();
if (MODE === "short") {
  // The short cut starts with last year's notes already loaded.
  await addin.locator('.tab[data-tab="notes"]').click();
  await addin.locator("#importNotes").click();
  await addin.locator('.tab[data-tab="review"]').click();
}
await wait(MODE === "short" ? 5000 : 6500);
await card(null, 800);

async function basicsFinding(title, cell, text, ms) {
  await click(item(title), { after: 500 });
  await click(linkIn(title, cell), { after: 500 });
  await caption(text, ms);
  await click(item(title), { after: 250 });
}

async function theWhy() {
  await show(insight(SURL_FLAG));
  await moveTo(insight(SURL_FLAG).locator(".insight-item").first(), 0.4, 0.5);
  await caption("A <b>January invoice for a December service</b>, never recorded.", 4200);
  await moveTo(insight(SURL_FLAG).locator(".why-block"), 0.4, 0.3);
  await caption("<b>Why it matters:</b> the company already owed this at year-end, so liabilities are understated.", 6000);
  await moveTo(insight(SURL_FLAG).locator(".why-meta"), 0.4, 0.5);
  await caption("Every flag names the <b>assertion at risk</b> and the <b>standard behind it</b>.", 5200);
}

if (MODE === "full") {
  // ─── The sheet ──────────────────────────────────────────────────────────────
  await caption("A workpaper in Excel. Baseline reads it and <b>gives feedback live</b>, in the panel on the right.", 5200);

  // ─── Part 1: workpaper basics ───────────────────────────────────────────────
  await part(1, "Workpaper basics", "Is it complete, consistent, and does it add up?");
  await moveTo(addin.locator("#checklist"), 0.4, 0.4);
  await caption("Checked on every sheet, <b>most important first</b>.", 3600);
  await basicsFinding("Totals include every line", "F17", "A total that <b>skips a line</b>. One click jumps to the cell.", 4200);
  await basicsFinding("Tables are complete", "D11", "A <b>blank</b> in the testing table: this item has no service date.", 4200);
  await basicsFinding("No placeholders left", "A22", "A conclusion still marked <b>to be completed</b>.", 3800);
  await basicsFinding("Dates are consistent", "E15", "A <b>2016</b> date in a column of 2026 dates.", 3800);
  await caption("Fix it, and the feedback updates <b>as you type</b>.", 300);
  await editCell("F17", "=SUM(F8:F16)");
  await panelTop();
  await moveTo(addin.locator("#score"), 0.2, 0.5);
  await wait(2600);

  // ─── Part 2: procedure review and the why ───────────────────────────────────
  await part(2, "Procedure review", "It knows what the procedure is for, and explains why each flag matters.");
  await panelTop();
  await moveTo(addin.locator("#suggest"), 0.3, 0.4);
  await caption("Baseline recognises a <b>search for unrecorded liabilities</b> and offers the matching review.", 4200);
  await click(addin.locator("#suggest button.primary"), { after: 1200 });
  await theWhy();
  await caption(null, 200);
  await activate("Cash confirmations");
  await panelTop();
  await click(addin.locator("#suggest button.primary"), { after: 1100 });
  await show(insight("Not sent or received directly"));
  await moveTo(insight("Not sent or received directly").locator(".insight-item").first(), 0.4, 0.5);
  await caption("On bank confirmations: a reply <b>forwarded by the client</b>.", 4000);
  await moveTo(insight("Not sent or received directly").locator(".why-block"), 0.4, 0.4);
  await caption("<b>Why it matters:</b> the auditor must control the confirmation, or it could be altered.", 6200);

  // ─── Part 3: last year's review notes ───────────────────────────────────────
  await part(3, "Last year's review notes", "Considered while you work, not after the review.");
  await activate("AP-3 SURL");
  await caption("Load last year's review notes…", 400);
  await click(addin.locator('.tab[data-tab="notes"]'), { after: 500 });
  await click(addin.locator("#importNotes"), { after: 1100 });
  await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
  await panelTop();
  await moveTo(addin.locator("#pyPanel"), 0.3, 0.3);
  await caption("…and Baseline shows which of them are <b>happening again this year</b>.", 5400);
  await show(insight(SURL_FLAG).locator(".py-callout"), "center");
  await moveTo(insight(SURL_FLAG).locator(".py-callout"), 0.4, 0.5);
  await caption("The reminder sits <b>on the flag itself</b>: this happened last year.", 5200);
  await caption(null, 400);
} else {
  // ─── Short cut: one of each ─────────────────────────────────────────────────
  await moveTo(addin.locator("#checklist"), 0.4, 0.4);
  await caption("<b>The basics:</b> complete, consistent, and it adds up.", 2600);
  await basicsFinding("Totals include every line", "F17", "A total that <b>skips a line</b>. One click jumps to the cell.", 3800);
  await panelTop();
  await caption("<b>The procedure:</b> Baseline recognises it and reviews the testing.", 600);
  await click(addin.locator("#suggest button.primary"), { after: 1100 });
  await show(insight(SURL_FLAG));
  await moveTo(insight(SURL_FLAG).locator(".insight-item").first(), 0.4, 0.5);
  await caption("A <b>January invoice for a December service</b>, never recorded.", 3600);
  await moveTo(insight(SURL_FLAG).locator(".why-meta"), 0.4, 0.5);
  await caption("<b>Why it matters</b>, the assertion at risk and the standard, on every flag.", 5600);
  await panelTop();
  await moveTo(addin.locator("#pyPanel"), 0.3, 0.3);
  await caption("<b>Last year's review notes:</b> which issues are happening again.", 5200);
  await caption(null, 300);
}

// ─── End card ────────────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>Feedback while you work. The why behind every flag.</p></div>`, 5000);

const video = page.video();
await context.close();
await browser.close();
const src = await video.path();
fs.copyFileSync(src, path.join(OUT, `raw-${MODE}.webm`));
console.log("errors:", errors);
console.log("video:", path.join(OUT, `raw-${MODE}.webm`));
