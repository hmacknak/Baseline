// Records the Baseline demo video: simulated Excel on the left, the real add-in on the right.
//
//   npm run build && python3 demo/video/export_workbook.py
//   npm i --no-save playwright-core && node demo/video/record.mjs full    (about 2.5 minutes)
//                                       node demo/video/record.mjs short   (under 1 minute)
//   ffmpeg -i demo/video/out/raw-full.webm -map_metadata -1 -c:v libx264 -crf 18 -pix_fmt yuv420p -movflags +faststart demo/Baseline-demo-video.mp4
//   ffmpeg -i demo/video/out/raw-short.webm -map_metadata -1 -c:v libx264 -crf 18 -pix_fmt yuv420p -movflags +faststart demo/Baseline-demo-short.mp4
//
// Storyboard: one finding per sheet, captions carry the story, year-end and threshold are pre-set
// (shown in a caption, not clicked through).
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

// ─── 1. Title card ───────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>Juniors get review notes for the same small mistakes.<br>Baseline flags them while they work.</p>
  <div class="note">Demo · simulated spreadsheet, real add-in · all names fictional</div></div>`);
workbook.active = "AP-3 SURL";
await page.evaluate(wb => window.__init(wb), workbook);
await addin.locator(".item").first().waitFor();
if (MODE === "short") {
  // Load last year's notes off screen.
  await addin.locator('.tab[data-tab="notes"]').click();
  await addin.locator("#importNotes").click();
  await addin.locator('.tab[data-tab="review"]').click();
}
await wait(5200);
await card(null, 800);

async function totalSkipsALine() {
  await click(item("Totals include every line"), { after: 600 });
  await click(linkIn("Totals include every line", "F17"), { after: 700 });
  await caption("A total that skips a line. Baseline <b>jumps to the cell</b> and explains why.", 1200);
  await moveTo(opens("Totals include every line").locator(".why-block, .why").first(), 0.4, 0.4);
  await wait(4600);
  await click(item("Totals include every line"), { after: 300 });
}

async function lastYearsNotes(load) {
  if (load) {
    await caption("Load last year's review notes…", 400);
    await click(addin.locator('.tab[data-tab="notes"]'), { after: 500 });
    await click(addin.locator("#importNotes"), { after: 1000 });
    await click(addin.locator('.tab[data-tab="review"]'), { after: 300 });
  }
  await panelTop();
  await moveTo(addin.locator("#pyPanel"), 0.3, 0.3);
  await caption("Last year's review notes are loaded. <b>Issues that repeat are flagged.</b>", 5400);
}

async function forwardedBankReply() {
  await caption(null, 200);
  await activate("Cash confirmations");
  await panelTop();
  await click(addin.locator("#suggest button.primary"), { after: 1100 });
  await show(insight("Not sent or received directly"));
  await moveTo(insight("Not sent or received directly").locator(".insight-item").first(), 0.4, 0.5);
  await caption("Bank reply <b>forwarded by the client</b>, not received directly.", 4200);
  await moveTo(insight("Not sent or received directly").locator(".why-meta"), 0.4, 0.5);
  await caption("Not reliable evidence until the auditor controls it. <b>ISA 505 / AS 2310.</b>", 5200);
}

if (MODE === "full") {
  // ─── 2. Ranked issues ───────────────────────────────────────────────────────
  await caption("A first-pass review of <b>AP-3, the search for unrecorded liabilities</b>.", 400);
  await moveTo(addin.locator("#score"), 0.2, 0.5);
  await wait(2600);
  await moveTo(addin.locator("#checklist"), 0.4, 0.4);
  await caption("Issues ranked, <b>most important first</b>. Year-end and threshold are already set.", 4600);

  // ─── 3. Total skips a line ──────────────────────────────────────────────────
  await totalSkipsALine();

  // ─── 4. The 2016 date ───────────────────────────────────────────────────────
  await click(item("Dates are consistent"), { after: 500 });
  await click(linkIn("Dates are consistent", "E15"), { after: 600 });
  await caption("Dates are checked together, so <b>a 2016 among 2026</b> stands out.", 4600);
  await click(item("Dates are consistent"), { after: 300 });

  // ─── 5. Mark OK with a reason ───────────────────────────────────────────────
  await click(item("No leftover highlighting"), { after: 500 });
  await click(opens("No leftover highlighting").locator(".ok-link").first(), { after: 300 });
  await addin.locator(".ok-reason").type("Highlight marks the item queried with the client", { delay: 30 });
  await caption("<b>Mark OK needs a reason.</b> If the cell changes, the flag comes back.", 1600);
  await click(addin.locator(".ok-form .primary"), { after: 2600 });

  // ─── 6. Last year's notes ───────────────────────────────────────────────────
  await lastYearsNotes(true);

  // ─── 7. Two fixes ───────────────────────────────────────────────────────────
  await caption("Fix the cell and the flag clears…", 300);
  await editCell("F17", "=SUM(F8:F16)");
  await editCell("A23", "Tickmark legend: ✓ = agreed to vendor invoice and cleared cheque");
  await panelTop();
  await moveTo(addin.locator("#pyPanel"), 0.3, 0.5);
  await caption("…and last year's tickmark note now shows <b>Looks fixed</b>.", 4400);

  // ─── 8–9. SURL reviewer ─────────────────────────────────────────────────────
  await caption(null, 200);
  await moveTo(addin.locator("#suggest"), 0.3, 0.4);
  await caption("Baseline recognises the workpaper type and <b>suggests the matching reviewer</b>.", 3600);
  await click(addin.locator("#suggest button.primary"), { after: 1200 });
  await show(insight("Invoices after year-end for this year's services"));
  await moveTo(insight("Invoices after year-end for this year's services").locator(".insight-item").first(), 0.4, 0.5);
  await caption("A <b>January invoice for December services</b>, never recorded.", 4200);
  await moveTo(insight("Invoices after year-end for this year's services").locator(".why-block"), 0.4, 0.3);
  await caption("Each finding says why it matters, the <b>assertion at risk</b> and the <b>standard behind it</b>.", 4000);
  await moveTo(insight("Invoices after year-end for this year's services").locator(".why-meta"), 0.4, 0.5);
  await wait(4000);

  // ─── 10. Cash confirmations ─────────────────────────────────────────────────
  await forwardedBankReply();

  // ─── 11. Expense vouching ───────────────────────────────────────────────────
  await caption(null, 200);
  await activate("Expense vouching");
  await panelTop();
  await click(addin.locator("#suggest button.primary"), { after: 1100 });
  await show(insight("Possible duplicate payments"));
  await moveTo(insight("Possible duplicate payments").locator(".insight-item").first(), 0.4, 0.5);
  await caption("<b>Same invoice paid twice</b>, nine days apart.", 4000);
  await moveTo(insight("Possible duplicate payments").locator(".py-callout"), 0.4, 0.5);
  await caption("Last year's review raised <b>the same issue</b>.", 4600);

  // ─── 12. Whole workbook ─────────────────────────────────────────────────────
  await caption(null, 200);
  await click(addin.locator("#reviewAll"), { after: 1200 });
  await caption("One click reviews <b>every sheet</b>. Settings are saved in the workbook.", 5400);
  await caption(null, 400);
} else {
  // ─── Short cut: one basic, the repeat callout, the bank reply ───────────────
  await moveTo(addin.locator("#checklist"), 0.4, 0.4);
  await caption("A first-pass review, <b>live in Excel</b>, most important first.", 3000);
  await totalSkipsALine();
  await lastYearsNotes(false);
  await forwardedBankReply();
  await caption(null, 300);
}

// ─── 13. End card ────────────────────────────────────────────────────────────
await card(`<div><div class="logo">B</div><h1>Baseline Review</h1>
  <p>The review note, delivered before the review.</p></div>`, 5000);

const video = page.video();
await context.close();
await browser.close();
const src = await video.path();
fs.copyFileSync(src, path.join(OUT, `raw-${MODE}.webm`));
console.log("errors:", errors);
console.log("video:", path.join(OUT, `raw-${MODE}.webm`));
