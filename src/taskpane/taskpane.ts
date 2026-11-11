/* global Office */

import { PROCEDURES_ENABLED } from "../config";
import { runChecklist } from "../review/checks";
import { Dismissals, ReviewedCheck, applyDismissals, findingKey, fingerprint } from "../review/dismiss";
import { dateCells } from "../review/quality";
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
import { explain, frameworkOf } from "../review/framework";
import { InsightResult, ReviewResult, runReview, suggestReviewer } from "../review/review";
import { fmtAmount, sheetTexts } from "../review/reviewers/procedure";
import { CATALOG, CATEGORY_ORDER, getReviewer } from "../review/reviewers/catalog";
import { ReviewContext, Reviewer, SheetSnapshot } from "../review/types";
import { CURRENT_LABEL, CURRENT_VERSION, applyUpdate, autoUpdateOnOpen, fetchLatest, isUpdate } from "../update";
import {
  activateSheet,
  loadSetting,
  readAllSheets,
  saveSetting,
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
const DISMISSED_SETTING = "baseline.dismissed.v1";
const SEVERITY_ORDER = ["high", "medium", "low"];

let notes: PyNote[] = [];
let ctx: ReviewContext = {};
let assignments: Record<string, string> = {};
let sheet: SheetSnapshot | null = null;
let result: ReviewResult | null = null;
let lastReview = 0;
let catalogSelection: string | null = null;
let dismissals: Dismissals = {};
/** Workpaper basics after "Mark OK" is applied. */
let basicsView: ReviewedCheck[] = [];
let procedureView: ReviewedCheck[] = [];
/** The finding whose "Mark OK" reason box is open. */
let marking: string | null = null;
let showClear = false;
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

/** "Why it matters" block: the reasoning, then the assertion at risk and the standard it comes from. */
function whyBlock(why: string, assertion?: string, reference?: string): HTMLElement {
  const box = el("div", "why-block");
  const p = el("p", "why");
  p.appendChild(el("strong", undefined, "Why it matters: "));
  p.appendChild(document.createTextNode(why));
  box.appendChild(p);
  if (assertion || reference) {
    const meta = el("p", "why-meta");
    if (assertion) {
      meta.appendChild(el("span", "why-label", "Assertion"));
      meta.appendChild(document.createTextNode(assertion));
    }
    if (reference) {
      meta.appendChild(el("span", "why-label", "Reference"));
      meta.appendChild(document.createTextNode(reference));
    }
    box.appendChild(meta);
  }
  return box;
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
    result = runReview(sheet, ctx, PROCEDURES_ENABLED ? assignedReviewer() : undefined);
    basicsView = applyDismissals(sheet, result.basics, dismissals);
    procedureView = applyDismissals(sheet, result.procedure, dismissals);
    lastReview = Date.now();
    renderReview();
    if (PROCEDURES_ENABLED) renderCatalog();
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
  text.appendChild(document.createTextNode("This sheet looks like "));
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
  const statuses = pyStatuses(notes, sheet.name, basicsView.concat(procedureView), result.insights);
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

async function saveDismissals(next: Dismissals, message: string) {
  try {
    await saveSetting(DISMISSED_SETTING, next);
    dismissals = next;
    marking = null;
    toast(message);
    await review();
  } catch (e) {
    console.error("[Baseline] saving accepted flags failed", e);
    toast("Couldn't save to this workbook");
  }
}

/** A flagged item, with "Mark OK" and, while open, the reason box. */
function renderFinding(f: ReviewedCheck["findings"][number]): HTMLElement {
  const wrap = el("div", "finding-wrap");
  const row = el("div", "finding");
  if (f.cell) row.appendChild(cellLink(f.cell));
  row.appendChild(el("span", "finding-msg", f.message));
  const key = sheet ? findingKey(sheet.name, f) : "";
  const ok = el("button", "ok-link", "Mark OK");
  ok.title = "Accept this flag with a reason";
  ok.addEventListener("click", () => {
    marking = marking === key ? null : key;
    renderReview();
  });
  row.appendChild(ok);
  wrap.appendChild(row);

  if (marking === key && sheet) {
    const form = el("div", "ok-form");
    const input = el("input", "ok-reason") as HTMLInputElement;
    input.placeholder = "Why is this OK? e.g. agreed rate per contract";
    const save = el("button", "primary", "Save");
    const cancel = el("button", "quiet", "Cancel");
    const snap = sheet;
    const submit = () => {
      const reason = input.value.trim();
      if (!reason) {
        toast("Add a short reason first");
        return;
      }
      const next = {
        ...dismissals,
        [key]: { reason, fingerprint: fingerprint(snap, f), date: new Date().toISOString().slice(0, 10) },
      };
      saveDismissals(next, "Marked OK");
    };
    save.addEventListener("click", submit);
    input.addEventListener("keydown", e => {
      if (e.key === "Enter") submit();
      if (e.key === "Escape") {
        marking = null;
        renderReview();
      }
    });
    cancel.addEventListener("click", () => {
      marking = null;
      renderReview();
    });
    form.appendChild(input);
    form.appendChild(save);
    form.appendChild(cancel);
    wrap.appendChild(form);
    setTimeout(() => input.focus(), 0);
  }
  return wrap;
}

function renderCheckItem(r: ReviewedCheck): HTMLElement {
  const { check, findings, dismissed } = r;
  const pass = findings.length === 0;
  const pyRelated = sheet ? pyNotesForCheck(notes, sheet.name, check) : [];
  const item = el("li", `item ${pass ? "pass" : "fail"} ${check.severity}`);

  const head = el("button", "item-head");
  head.setAttribute("aria-expanded", String(expanded.has(check.id)));
  head.appendChild(el("span", "mark", pass ? "✓" : "✗"));
  head.appendChild(el("span", "item-title", check.title));
  if (!pass && pyRelated.length) head.appendChild(el("span", "py-flag", "PY"));
  if (!pass) head.appendChild(el("span", "count", String(findings.length)));
  else if (dismissed.length) head.appendChild(el("span", "count ok-count", `${dismissed.length} OK`));
  head.addEventListener("click", () => {
    if (expanded.has(check.id)) expanded.delete(check.id);
    else expanded.add(check.id);
    renderReview();
  });
  item.appendChild(head);

  if (expanded.has(check.id)) {
    const body = el("div", "item-body");
    if (pass && !dismissed.length) body.appendChild(el("p", "why", "Nothing to fix here."));

    findings.slice(0, MAX_SHOWN).forEach(f => body.appendChild(renderFinding(f)));
    if (findings.length > MAX_SHOWN) body.appendChild(el("p", "more", `…and ${findings.length - MAX_SHOWN} more`));

    if (!pass && pyRelated.length) {
      const callout = el("div", "py-callout");
      callout.appendChild(el("strong", undefined, "Heads up, this happened last year: "));
      callout.appendChild(document.createTextNode(pyRelated.map(n => `"${n.text}"`).join("; ")));
      body.appendChild(callout);
    }

    if (!pass) {
      const text = explain(check, ctx);
      body.appendChild(whyBlock(text.why, check.assertion, text.reference));
      const fix = el("p", "why");
      fix.appendChild(el("strong", undefined, "How to fix: "));
      fix.appendChild(document.createTextNode(check.fix));
      body.appendChild(fix);
    }

    if (dismissed.length) {
      const box = el("div", "accepted");
      box.appendChild(el("p", "accepted-head", `Marked OK (${dismissed.length})`));
      dismissed.forEach(d => {
        const row = el("div", "finding");
        if (d.finding.cell) row.appendChild(cellLink(d.finding.cell));
        const text = el("span", "finding-msg");
        text.appendChild(el("span", "accepted-reason", d.dismissal.reason));
        text.appendChild(el("span", "accepted-meta", ` · ${d.finding.message} · ${d.dismissal.date}`));
        row.appendChild(text);
        const undo = el("button", "ok-link", "Undo");
        undo.addEventListener("click", () => {
          const next = { ...dismissals };
          delete next[d.key];
          saveDismissals(next, "Flag restored");
        });
        row.appendChild(undo);
        box.appendChild(row);
      });
      body.appendChild(box);
    }
    item.appendChild(body);
  }
  return item;
}

/** Open items first, most important first; checks that are clear fold into one line. */
function renderChecklist(list: HTMLElement, results: ReviewedCheck[]) {
  list.textContent = "";
  if (!sheet) return;
  const open = results
    .filter(r => r.findings.length)
    .sort((a, b) => SEVERITY_ORDER.indexOf(a.check.severity) - SEVERITY_ORDER.indexOf(b.check.severity));
  const clear = results.filter(r => !r.findings.length);
  open.forEach(r => list.appendChild(renderCheckItem(r)));
  if (!clear.length) return;
  const toggle = el("li", "item clear-toggle");
  const btn = el("button", "item-head");
  btn.appendChild(el("span", "mark", "✓"));
  btn.appendChild(el("span", "item-title", `${plural(clear.length, "check")} clear`));
  btn.appendChild(el("span", "count", showClear ? "Hide" : "Show"));
  btn.addEventListener("click", () => {
    showClear = !showClear;
    renderReview();
  });
  toggle.appendChild(btn);
  list.appendChild(toggle);
  if (showClear) clear.forEach(r => list.appendChild(renderCheckItem(r)));
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
      const text = explain(r.insight, ctx);
      body.appendChild(whyBlock(text.why, r.insight.assertion, text.reference));
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
  renderChecklist($("procList"), procedureView);
}

function renderScore() {
  const score = $("score");
  score.textContent = "";
  if (!sheet || !result) return;
  if (!sheet.cells.length) {
    score.textContent = "This sheet is empty. Start working and the checklist will update as you go.";
    return;
  }
  const all = basicsView.concat(procedureView);
  const toFix = all.reduce((n, r) => n + r.findings.length, 0);
  const passed = all.filter(r => r.findings.length === 0).length;
  if (!toFix) {
    score.appendChild(el("strong", "all-clear", "No open issues"));
    score.appendChild(document.createTextNode(` · all ${all.length} checks clear`));
    return;
  }
  score.appendChild(el("strong", undefined, `${toFix} to fix`));
  score.appendChild(document.createTextNode(` · ${passed} of ${all.length} checks clear`));
}

// ─── Step through issues or dates ───────────────────────────────────────────

type StepMode = "issues" | "dates";
let stepMode: StepMode = "issues";
let stepIndex = -1;

interface Step {
  cell: string;
  text: string;
}

function steps(): Step[] {
  if (!sheet) return [];
  if (stepMode === "dates") {
    return dateCells(sheet).map(d => ({
      cell: d.cell,
      text: `${formatSerial(d.date)}${d.asText ? " (typed as text)" : ""}`,
    }));
  }
  const out: Step[] = [];
  for (const r of basicsView.concat(procedureView)) {
    for (const f of r.findings) if (f.cell) out.push({ cell: f.cell, text: f.message });
  }
  return out;
}

function renderStepper() {
  const box = $("stepper");
  box.hidden = !sheet || !sheet.cells.length;
  const list = steps();
  Array.prototype.slice.call(document.querySelectorAll(".step-mode")).forEach((b: HTMLButtonElement) => {
    b.classList.toggle("active", b.dataset.mode === stepMode);
  });
  if (stepIndex >= list.length) stepIndex = list.length - 1;
  ($("stepPrev") as HTMLButtonElement).disabled = !list.length;
  ($("stepNext") as HTMLButtonElement).disabled = !list.length;
  const what = stepMode === "dates" ? "date" : "issue";
  $("stepPos").textContent = list.length
    ? stepIndex >= 0
      ? `${stepIndex + 1} of ${list.length}`
      : plural(list.length, what)
    : `No ${what}s`;
  const detail = $("stepDetail");
  detail.textContent = "";
  if (stepIndex >= 0 && list[stepIndex]) {
    detail.appendChild(el("span", "step-cell", list[stepIndex].cell));
    detail.appendChild(document.createTextNode(list[stepIndex].text));
  }
}

function step(delta: number) {
  const list = steps();
  if (!list.length) return;
  stepIndex = stepIndex < 0 ? (delta > 0 ? 0 : list.length - 1) : (stepIndex + delta + list.length) % list.length;
  selectCell(list[stepIndex].cell).catch(() => toast("Couldn't select that cell"));
  renderStepper();
}

function wireStepper() {
  Array.prototype.slice.call(document.querySelectorAll(".step-mode")).forEach((b: HTMLButtonElement) => {
    b.addEventListener("click", () => {
      stepMode = (b.dataset.mode as StepMode) || "issues";
      stepIndex = -1;
      renderStepper();
    });
  });
  $("stepPrev").addEventListener("click", () => step(-1));
  $("stepNext").addEventListener("click", () => step(1));
}

// ─── Whole-workbook review ───────────────────────────────────────────────────

async function reviewWorkbook() {
  const panel = $("workbookPanel");
  panel.hidden = false;
  panel.textContent = "Reviewing every sheet…";
  try {
    const sheets = await readAllSheets();
    panel.textContent = "";
    const head = el("div", "wb-head");
    head.appendChild(el("strong", undefined, "Whole workbook"));
    const close = el("button", "link", "Close");
    close.addEventListener("click", () => {
      panel.hidden = true;
    });
    head.appendChild(close);
    panel.appendChild(head);
    let needing = 0;
    const list = el("ul", "wb-list");
    for (const snap of sheets) {
      if (!snap.cells.length) continue;
      const checks = applyDismissals(snap, runChecklist(snap, ctx), dismissals);
      const open = checks.filter(c => c.findings.length);
      const count = open.reduce((n, c) => n + c.findings.length, 0);
      if (count) needing++;
      const li = el("li");
      const btn = el("button", `wb-row ${count ? "fail" : "pass"}`);
      btn.appendChild(el("span", "mark", count ? "✗" : "✓"));
      const text = el("span", "wb-text");
      text.appendChild(el("strong", undefined, snap.name));
      text.appendChild(
        el(
          "span",
          "wb-sub",
          count
            ? open
                .sort((a, b) => SEVERITY_ORDER.indexOf(a.check.severity) - SEVERITY_ORDER.indexOf(b.check.severity))
                .slice(0, 2)
                .map(c => c.check.title)
                .join(" · ")
            : "No open issues",
        ),
      );
      btn.appendChild(text);
      if (count) btn.appendChild(el("span", "count", String(count)));
      btn.addEventListener("click", () => activateSheet(snap.name).catch(() => toast("Couldn't open that sheet")));
      li.appendChild(btn);
      list.appendChild(li);
    }
    panel.appendChild(
      el("p", "help small", needing ? `${plural(needing, "sheet")} need attention. Tap one to go to it.` : "No open issues on any sheet. This is a first pass, not a review sign-off."),
    );
    panel.appendChild(list);
  } catch (e) {
    console.error("[Baseline] workbook review failed", e);
    panel.textContent = "Couldn't read every sheet. Try again.";
  }
}

function renderReview() {
  renderStatus();
  if (PROCEDURES_ENABLED) renderSuggestion();
  renderScore();
  renderStepper();
  renderPyPanel();
  if (PROCEDURES_ENABLED) renderProcedure();
  renderChecklist($("checklist"), basicsView);
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
  ($("framework") as HTMLSelectElement).value = frameworkOf(ctx);
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
    if (($("framework") as HTMLSelectElement).value === "us-gaap") next.framework = "us-gaap";
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
  wireStepper();
  $("reviewAll").addEventListener("click", reviewWorkbook);
  if (!PROCEDURES_ENABLED) {
    Array.prototype.slice.call(document.querySelectorAll(".procedures-only")).forEach((n: HTMLElement) => {
      n.hidden = true;
    });
  }
  wireNotesTab();
  wireSettings();
  // Re-tag on load so notes saved by an older version pick up improved keyword rules.
  notes = loadNotes().map(n => ({ ...n, tags: tagNote(n.text) }));
  ctx = loadContext();
  dismissals = loadSetting<Dismissals>(DISMISSED_SETTING) || {};
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
