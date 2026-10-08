/* global Office */

import { CheckResult } from "../review/checks";
import { findYearEndInText, formatSerial, isoFromSerial, serialFromYMD } from "../review/dates";
import {
  PyNote,
  PyStatus,
  normaliseSheetName,
  parseNotesRows,
  parseNotesText,
  pyNotesForCheck,
  pyStatuses,
  tagNote,
} from "../review/notes";
import { InsightResult, ReviewResult, runReview, suggestReviewer } from "../review/review";
import { fmtAmount, sheetTexts } from "../review/reviewers/procedure";
import { CATALOG, CATEGORY_ORDER, getReviewer } from "../review/reviewers/catalog";
import { ReviewContext, Reviewer, SheetSnapshot } from "../review/types";
import { CURRENT_LABEL, CURRENT_VERSION, applyUpdate, autoUpdateOnOpen, fetchLatest, isUpdate } from "../update";
import {
  loadAssignments,
  loadContext,
  loadNotes,
  readActiveSheet,
  readNotesSheet,
  saveAssignments,
  saveContext,
  saveNotes,
  selectCell,
  watchWorkbook,
} from "../excel/workbook";

const DEBOUNCE_MS = 600;
const MAX_SHOWN = 20;
const DECLINED = "none";

let notes: PyNote[] = [];
let ctx: ReviewContext = {};
let assignments: Record<string, string> = {};
let sheet: SheetSnapshot | null = null;
let result: ReviewResult | null = null;
let lastReview = 0;
let catalogSelection: string | null = null;
const expanded = new Set<string>();
/** High-risk insights start open; this remembers the ones the junior closed. */
const collapsed = new Set<string>();

// ─── Small DOM helpers (text only, never innerHTML with workbook content) ────

function $(id: string): HTMLElement {
  return document.getElementById(id) as HTMLElement;
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function toast(msg: string) {
  const t = $("toast");
  t.textContent = msg;
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 1800);
}

function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

function cellLink(address: string): HTMLButtonElement {
  const link = el("button", "cell-link", address);
  link.title = `Go to ${address}`;
  link.addEventListener("click", () => selectCell(address).catch(() => toast("Couldn't select that cell")));
  return link;
}

// ─── Which reviewer is on this sheet ─────────────────────────────────────────

function sheetKey(): string {
  return sheet ? normaliseSheetName(sheet.name) || sheet.name : "";
}

function assignedReviewer(): Reviewer | undefined {
  const id = assignments[sheetKey()];
  return id && id !== DECLINED ? getReviewer(id) : undefined;
}

async function assign(id: string, message?: string) {
  const next = { ...assignments, [sheetKey()]: id };
  try {
    await saveAssignments(next);
    assignments = next;
    if (message) toast(message);
    await review();
  } catch (e) {
    console.error("[Baseline] saving reviewer failed", e);
    toast("Couldn't save to this workbook");
  }
}

// ─── Review loop ─────────────────────────────────────────────────────────────

let timer: ReturnType<typeof setTimeout> | undefined;
function scheduleReview() {
  if (timer) clearTimeout(timer);
  timer = setTimeout(review, DEBOUNCE_MS);
}

async function review() {
  try {
    sheet = await readActiveSheet();
    result = runReview(sheet, ctx, assignedReviewer());
    lastReview = Date.now();
    renderReview();
    renderCatalog();
    renderSettingsHint();
  } catch (e) {
    console.error("[Baseline] review failed", e);
    $("status").textContent = "Couldn't read this sheet. Try clicking into it again.";
  }
}

function renderStatus() {
  const status = $("status");
  status.textContent = "";
  if (!sheet) return;
  status.appendChild(el("span", "live"));
  const secs = Math.round((Date.now() - lastReview) / 1000);
  const when = secs < 5 ? "just now" : secs < 60 ? `${secs}s ago` : `${Math.round(secs / 60)}m ago`;
  status.appendChild(document.createTextNode(`Live · reviewing "${sheet.name}" · updated ${when}`));
}

// ─── Review tab ──────────────────────────────────────────────────────────────

const STATUS_LABEL: Record<PyStatus, string> = {
  "still-an-issue": "Again this year",
  "looks-fixed": "Looks fixed",
  "check-manually": "Check yourself",
};

function renderSuggestion() {
  const box = $("suggest");
  box.textContent = "";
  box.hidden = true;
  if (!sheet || !sheet.cells.length || assignments[sheetKey()]) return;
  const s = suggestReviewer(sheet);
  if (!s) return;
  box.hidden = false;
  const text = el("p");
  text.appendChild(document.createTextNode("This looks like a "));
  text.appendChild(el("strong", undefined, s.reviewer.name));
  text.appendChild(document.createTextNode(". Use that reviewer?"));
  box.appendChild(text);
  const buttons = el("div", "buttons");
  const use = el("button", "primary", "Use it");
  use.addEventListener("click", () => assign(s.reviewer.id, `${s.reviewer.name} reviewer on`));
  const other = el("button", undefined, "Choose another");
  other.addEventListener("click", () => {
    catalogSelection = null;
    showTab("catalog");
  });
  const no = el("button", "quiet", "Not now");
  no.addEventListener("click", () => assign(DECLINED));
  buttons.appendChild(use);
  buttons.appendChild(other);
  buttons.appendChild(no);
  box.appendChild(buttons);
}

function renderPyPanel() {
  const panel = $("pyPanel");
  panel.textContent = "";
  if (!sheet || !result) return;
  const statuses = pyStatuses(notes, sheet.name, result.all, result.insights);
  panel.hidden = statuses.length === 0;
  if (!statuses.length) return;

  const repeats = statuses.filter(s => s.status === "still-an-issue").length;
  panel.appendChild(
    el(
      "h2",
      undefined,
      repeats
        ? `Last year's reviewer flagged this. ${repeats} of ${statuses.length} issue${statuses.length > 1 ? "s" : ""} showing again`
        : `Last year's reviewer left ${plural(statuses.length, "note")} on this sheet`,
    ),
  );
  const order: PyStatus[] = ["still-an-issue", "check-manually", "looks-fixed"];
  const list = el("ul");
  statuses
    .slice()
    .sort((a, b) => order.indexOf(a.status) - order.indexOf(b.status))
    .forEach(s => {
      const li = el("li");
      li.appendChild(el("span", `tag ${s.status}`, STATUS_LABEL[s.status]));
      li.appendChild(el("span", undefined, s.note.cell ? `${s.note.text} (was ${s.note.cell})` : s.note.text));
      list.appendChild(li);
    });
  panel.appendChild(list);
}

function renderChecklist(list: HTMLElement, results: CheckResult[]) {
  list.textContent = "";
  if (!sheet) return;

  for (const { check, findings } of results) {
    const pass = findings.length === 0;
    const pyRelated = pyNotesForCheck(notes, sheet.name, check);
    const item = el("li", `item ${pass ? "pass" : "fail"} ${check.severity}`);

    const head = el("button", "item-head");
    head.setAttribute("aria-expanded", String(expanded.has(check.id)));
    head.appendChild(el("span", "mark", pass ? "✓" : "✗"));
    head.appendChild(el("span", "item-title", check.title));
    if (!pass && pyRelated.length) head.appendChild(el("span", "py-flag", "PY"));
    if (!pass) head.appendChild(el("span", "count", String(findings.length)));
    head.addEventListener("click", () => {
      if (expanded.has(check.id)) expanded.delete(check.id);
      else expanded.add(check.id);
      renderReview();
    });
    item.appendChild(head);

    if (expanded.has(check.id)) {
      const body = el("div", "item-body");
      if (pass) body.appendChild(el("p", "why", "Nothing to fix here."));

      findings.slice(0, MAX_SHOWN).forEach(f => {
        const row = el("div", "finding");
        if (f.cell) row.appendChild(cellLink(f.cell));
        row.appendChild(el("span", "finding-msg", f.message));
        body.appendChild(row);
      });
      if (findings.length > MAX_SHOWN) body.appendChild(el("p", "more", `…and ${findings.length - MAX_SHOWN} more`));

      if (!pass && pyRelated.length) {
        const callout = el("div", "py-callout");
        callout.appendChild(el("strong", undefined, "Heads up, this happened last year: "));
        callout.appendChild(document.createTextNode(pyRelated.map(n => `"${n.text}"`).join("; ")));
        body.appendChild(callout);
      }

      if (!pass) {
        const why = el("p", "why");
        why.appendChild(el("strong", undefined, "Why it matters: "));
        why.appendChild(document.createTextNode(check.why));
        body.appendChild(why);
        const fix = el("p", "why");
        fix.appendChild(el("strong", undefined, "How to fix: "));
        fix.appendChild(document.createTextNode(check.fix));
        body.appendChild(fix);
      }
      item.appendChild(body);
    }
    list.appendChild(item);
  }
}

function renderInsight(r: InsightResult): HTMLElement {
  const key = `insight:${r.insight.id}`;
  const hits = r.items.length;
  const startsOpen = hits > 0 && r.insight.severity === "high";
  const open = startsOpen ? !collapsed.has(key) : expanded.has(key);
  const card = el("div", `insight ${r.blocked ? "blocked" : hits ? `hit ${r.insight.severity}` : "clear"}`);

  const head = el("button", "insight-head");
  head.setAttribute("aria-expanded", String(open));
  head.appendChild(el("span", "mark", r.blocked ? "–" : hits ? "!" : "✓"));
  head.appendChild(el("span", "item-title", r.insight.title));
  if (hits && sheet && pyNotesForCheck(notes, sheet.name, r.insight).length) {
    head.appendChild(el("span", "py-flag", "PY"));
  }
  if (hits) {
    const total = r.items.reduce((s, i) => s + (i.amount ?? 0), 0);
    const meta = el("span", "insight-meta", total ? `${hits} · $${fmtAmount(total)}` : String(hits));
    head.appendChild(meta);
  }
  head.addEventListener("click", () => {
    if (startsOpen) {
      if (open) collapsed.add(key);
      else collapsed.delete(key);
    } else if (open) expanded.delete(key);
    else expanded.add(key);
    renderReview();
  });
  card.appendChild(head);

  if (open) {
    const body = el("div", "insight-body");
    if (r.blocked) {
      body.appendChild(el("p", "why", r.blocked + "."));
      const go = el("button", "link", "Open settings");
      go.addEventListener("click", () => showTab("settings"));
      body.appendChild(go);
    } else if (!hits) {
      body.appendChild(el("p", "why", "Nothing flagged."));
    } else {
      r.items.slice(0, MAX_SHOWN).forEach(i => {
        const row = el("div", "insight-item");
        row.appendChild(cellLink(i.cell));
        const text = el("div", "insight-text");
        text.appendChild(el("strong", undefined, i.label));
        text.appendChild(el("span", undefined, i.detail));
        row.appendChild(text);
        if (i.amount !== undefined) row.appendChild(el("span", "insight-amount", fmtAmount(i.amount)));
        body.appendChild(row);
      });
      if (hits > MAX_SHOWN) body.appendChild(el("p", "more", `…and ${hits - MAX_SHOWN} more`));
      const pyRelated = sheet ? pyNotesForCheck(notes, sheet.name, r.insight) : [];
      if (pyRelated.length) {
        const callout = el("div", "py-callout");
        callout.appendChild(el("strong", undefined, "Heads up, this happened last year: "));
        callout.appendChild(document.createTextNode(pyRelated.map(n => `"${n.text}"`).join("; ")));
        body.appendChild(callout);
      }
      const total = r.items.reduce((s, i) => s + Math.abs(i.amount ?? 0), 0);
      if (ctx.threshold && hits > 1 && total >= ctx.threshold) {
        body.appendChild(
          el("p", "over", `Together $${fmtAmount(total)}: above your $${fmtAmount(ctx.threshold)} threshold.`),
        );
      }
      const why = el("p", "why");
      why.appendChild(el("strong", undefined, "Why it's risky: "));
      why.appendChild(document.createTextNode(r.insight.why));
      body.appendChild(why);
    }
    card.appendChild(body);
  }
  return card;
}

function renderProcedure() {
  const section = $("procSection");
  const add = $("addReviewer");
  const reviewer = result?.reviewer;
  section.hidden = !reviewer;
  add.hidden = !!reviewer || !sheet || !sheet.cells.length || !$("suggest").hidden;
  if (!reviewer || !result) return;

  $("procIcon").textContent = reviewer.code;
  $("procName").textContent = reviewer.name;
  const box = $("insights");
  box.textContent = "";
  const flagged = result.insights.reduce((n, r) => n + r.items.length, 0);
  const head = el("p", "insights-head");
  head.appendChild(el("strong", undefined, "Insights"));
  head.appendChild(
    document.createTextNode(flagged ? ` · ${plural(flagged, "high-risk item")} to look at` : " · nothing flagged"),
  );
  box.appendChild(head);
  result.insights.forEach(r => box.appendChild(renderInsight(r)));
  renderChecklist($("procList"), result.procedure);
}

function renderScore() {
  const score = $("score");
  score.textContent = "";
  if (!sheet || !result) return;
  if (!sheet.cells.length) {
    score.textContent = "This sheet is empty. Start working and the checklist will update as you go.";
    return;
  }
  const passed = result.all.filter(r => r.findings.length === 0).length;
  score.appendChild(el("strong", undefined, `${passed}/${result.all.length}`));
  score.appendChild(
    document.createTextNode(passed === result.all.length ? " checks clear. Ready for review." : " checks clear"),
  );
}

function renderReview() {
  renderStatus();
  renderSuggestion();
  renderScore();
  renderPyPanel();
  renderProcedure();
  if (result) renderChecklist($("checklist"), result.basics);
  $("truncated").hidden = !(sheet && sheet.truncated);
}

// ─── Catalog tab ─────────────────────────────────────────────────────────────

function reviewerMeta(r: Reviewer): string {
  if (r.comingSoon) return "Coming soon";
  if (r.alwaysOn) return `Always on · ${plural(r.checks.length, "check")}`;
  return `${plural(r.checks.length, "check")} · ${plural(r.insights.length, "insight")}`;
}

function renderCatalogDetail() {
  const box = $("catalogDetail");
  box.textContent = "";
  const r = catalogSelection ? getReviewer(catalogSelection) : undefined;
  box.hidden = !r;
  if (!r) return;
  const current = assignedReviewer();

  const top = el("div", "detail-top");
  top.appendChild(el("span", "monogram large", r.code));
  const titles = el("div");
  titles.appendChild(el("h2", undefined, r.name));
  titles.appendChild(el("p", "help", r.summary));
  top.appendChild(titles);
  box.appendChild(top);

  if (r.comingSoon) {
    box.appendChild(el("p", "help", "Coming soon. Tell us what this reviewer should catch."));
  } else {
    if (r.checks.length) {
      box.appendChild(el("h3", undefined, "Checklist"));
      const ul = el("ul");
      r.checks.forEach(c => ul.appendChild(el("li", undefined, c.title)));
      box.appendChild(ul);
    }
    if (r.insights.length) {
      box.appendChild(el("h3", undefined, "Finds high-risk items"));
      const ul = el("ul");
      r.insights.forEach(i => ul.appendChild(el("li", undefined, i.title)));
      box.appendChild(ul);
    }
  }

  const buttons = el("div", "buttons");
  if (!r.alwaysOn && !r.comingSoon && sheet) {
    if (current && current.id === r.id) {
      const remove = el("button", undefined, "Remove from this sheet");
      remove.addEventListener("click", () => assign(DECLINED, "Reviewer removed"));
      buttons.appendChild(remove);
    } else {
      const use = el("button", "primary", `Use on "${sheet.name}"`);
      use.addEventListener("click", async () => {
        await assign(r.id, `${r.name} reviewer on`);
        showTab("review");
      });
      buttons.appendChild(use);
    }
  }
  const close = el("button", "quiet", "Close");
  close.addEventListener("click", () => {
    catalogSelection = null;
    renderCatalog();
  });
  buttons.appendChild(close);
  box.appendChild(buttons);
}

function renderCatalog() {
  renderCatalogDetail();
  const wrap = $("catalogRows");
  wrap.textContent = "";
  const current = assignedReviewer();
  const suggested = sheet ? suggestReviewer(sheet) : null;
  const categories = CATEGORY_ORDER.filter(c => CATALOG.some(r => r.category === c));
  for (const cat of categories) {
    const row = el("section", "cat-row");
    row.appendChild(el("h3", undefined, cat));
    const strip = el("div", "cat-strip");
    CATALOG.filter(r => r.category === cat).forEach(r => {
      const on = r.alwaysOn || (current && current.id === r.id);
      const card = el("button", `card${r.comingSoon ? " soon" : ""}${on ? " on" : ""}${catalogSelection === r.id ? " selected" : ""}`);
      card.appendChild(el("span", "monogram", r.code));
      card.appendChild(el("span", "card-name", r.name));
      card.appendChild(el("span", "card-meta", reviewerMeta(r)));
      if (on) card.appendChild(el("span", "badge", r.alwaysOn ? "Always on" : "On this sheet"));
      else if (suggested && suggested.reviewer.id === r.id) card.appendChild(el("span", "badge suggested", "Suggested"));
      card.addEventListener("click", () => {
        catalogSelection = r.id;
        renderCatalog();
        $("tab-catalog").scrollTop = 0;
        window.scrollTo({ top: 0, behavior: "smooth" });
      });
      strip.appendChild(card);
    });
    row.appendChild(strip);
    wrap.appendChild(row);
  }
}

// ─── Settings tab ────────────────────────────────────────────────────────────

function renderSettings() {
  ($("yearEnd") as HTMLInputElement).value = ctx.yearEnd ? isoFromSerial(ctx.yearEnd) : "";
  ($("threshold") as HTMLInputElement).value = ctx.threshold !== undefined ? String(ctx.threshold) : "";
  renderSettingsHint();
}

/** Offer a year-end spotted in the sheet title, e.g. "Bank rec, 31 Dec 2025". */
function renderSettingsHint() {
  const hint = $("yearEndHint");
  hint.textContent = "";
  hint.hidden = true;
  if (ctx.yearEnd || !sheet) return;
  const found = findYearEndInText(sheetTexts(sheet, 6));
  if (found === null) return;
  hint.hidden = false;
  hint.appendChild(document.createTextNode(`This sheet mentions ${formatSerial(found)}. `));
  const use = el("button", "link", "Use it");
  use.addEventListener("click", () => {
    ($("yearEnd") as HTMLInputElement).value = isoFromSerial(found);
  });
  hint.appendChild(use);
}

function wireSettings() {
  $("saveSettings").addEventListener("click", async () => {
    const ye = ($("yearEnd") as HTMLInputElement).value;
    const th = ($("threshold") as HTMLInputElement).value;
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(ye);
    const next: ReviewContext = {};
    if (m) next.yearEnd = serialFromYMD(+m[1], +m[2], +m[3]);
    if (th.trim() !== "" && isFinite(Number(th))) next.threshold = Number(th);
    try {
      await saveContext(next);
      ctx = next;
      toast("Settings saved");
      await review();
    } catch (e) {
      console.error("[Baseline] saving settings failed", e);
      toast("Couldn't save settings to this workbook");
    }
  });
}

// ─── Updates ─────────────────────────────────────────────────────────────────

const UPDATE_CHECK_MS = 30 * 60 * 1000;
let pendingVersion: string | null = null;

/** Quietly check in the background; offer the update rather than reloading mid-task. */
async function checkInBackground() {
  const latest = await fetchLatest();
  if (latest && isUpdate(CURRENT_VERSION, latest.version)) {
    pendingVersion = latest.version;
    $("updateBanner").hidden = false;
  }
}

function wireUpdates() {
  $("versionLabel").textContent = CURRENT_LABEL;
  $("updateNow").addEventListener("click", () => {
    if (pendingVersion) applyUpdate(pendingVersion);
  });
  $("checkUpdates").addEventListener("click", async () => {
    const status = $("updateStatus");
    status.hidden = false;
    status.textContent = "Checking…";
    const latest = await fetchLatest();
    if (!latest) {
      status.textContent = "Couldn't reach the update server. Check your connection and try again.";
    } else if (CURRENT_VERSION === "dev") {
      status.textContent = `Development build. The published version is ${latest.label}.`;
    } else if (isUpdate(CURRENT_VERSION, latest.version)) {
      status.textContent = `Updating to ${latest.label}…`;
      applyUpdate(latest.version);
    } else {
      status.textContent = "You're on the latest version.";
    }
  });
}

// ─── Notes tab ───────────────────────────────────────────────────────────────

function renderNotesTab() {
  const count = $("notesCount");
  count.hidden = notes.length === 0;
  count.textContent = String(notes.length);

  const wrap = $("notesList");
  wrap.textContent = "";
  if (!notes.length) return;
  wrap.appendChild(el("p", "help", `${plural(notes.length, "note")} saved in this workbook:`));
  const table = el("table", "notes-table");
  const head = el("tr");
  ["Sheet", "Cell", "Note"].forEach(h => head.appendChild(el("th", undefined, h)));
  table.appendChild(head);
  notes.forEach(n => {
    const tr = el("tr");
    tr.appendChild(el("td", undefined, n.sheet || "Any"));
    tr.appendChild(el("td", undefined, n.cell));
    tr.appendChild(el("td", undefined, n.text));
    table.appendChild(tr);
  });
  wrap.appendChild(table);
}

async function storeNotes(next: PyNote[], message: string) {
  try {
    await saveNotes(next);
    notes = next;
    renderNotesTab();
    renderReview();
    toast(message);
  } catch (e) {
    console.error("[Baseline] saving notes failed", e);
    toast("Couldn't save notes to this workbook");
  }
}

function wireNotesTab() {
  $("saveNotes").addEventListener("click", () => {
    const input = $("notesInput") as HTMLTextAreaElement;
    const parsed = parseNotesText(input.value);
    if (!parsed.length) {
      toast("Paste at least one note first");
      return;
    }
    input.value = "";
    storeNotes(notes.concat(parsed), `Saved ${plural(parsed.length, "note")}`);
  });

  $("importNotes").addEventListener("click", async () => {
    const rows = await readNotesSheet().catch(() => null);
    if (rows === null) {
      toast('No sheet called "PY Notes" in this workbook');
      return;
    }
    const parsed = parseNotesRows(rows);
    if (!parsed.length) {
      toast('The "PY Notes" sheet has no notes');
      return;
    }
    storeNotes(parsed, `Imported ${plural(parsed.length, "note")}`);
  });

  $("clearNotes").addEventListener("click", () => {
    if (!notes.length) return;
    storeNotes([], "Notes cleared");
  });
}

// ─── Tabs and start-up ───────────────────────────────────────────────────────

function showTab(name: string) {
  const tabs = Array.prototype.slice.call(document.querySelectorAll(".tab")) as HTMLButtonElement[];
  tabs.forEach(t => {
    const on = t.dataset.tab === name;
    t.classList.toggle("active", on);
    t.setAttribute("aria-selected", String(on));
    $(`tab-${t.dataset.tab}`).hidden = !on;
  });
  if (name === "settings") renderSettings();
  if (name === "catalog") renderCatalog();
}

function wireTabs() {
  const tabs = Array.prototype.slice.call(document.querySelectorAll(".tab")) as HTMLButtonElement[];
  tabs.forEach(tab => tab.addEventListener("click", () => showTab(tab.dataset.tab || "review")));
  $("procChange").addEventListener("click", () => {
    catalogSelection = assignedReviewer()?.id ?? null;
    showTab("catalog");
  });
  $("addReviewer").addEventListener("click", () => {
    catalogSelection = null;
    showTab("catalog");
  });
}

Office.onReady(async () => {
  // Switch to a newer published build before doing any work; everything is stored in the workbook.
  if (await autoUpdateOnOpen()) return;
  wireTabs();
  wireUpdates();
  wireNotesTab();
  wireSettings();
  // Re-tag on load so notes saved by an older version pick up improved keyword rules.
  notes = loadNotes().map(n => ({ ...n, tags: tagNote(n.text) }));
  ctx = loadContext();
  assignments = loadAssignments();
  renderNotesTab();
  await review();
  try {
    await watchWorkbook(scheduleReview);
  } catch (e) {
    console.error("[Baseline] live updates unavailable", e);
    toast("Live updates unavailable in this Excel version");
  }
  setInterval(renderStatus, 15000);
  setInterval(checkInBackground, UPDATE_CHECK_MS);
});

export {};
