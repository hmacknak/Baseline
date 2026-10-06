/* global Office */

import { CheckResult, runChecklist } from "../review/checks";
import { PyNote, PyStatus, parseNotesRows, parseNotesText, pyNotesForCheck, pyStatuses, tagNote } from "../review/notes";
import { CheckId, SheetSnapshot } from "../review/types";
import { loadNotes, readActiveSheet, readNotesSheet, saveNotes, selectCell, watchWorkbook } from "../excel/workbook";

const DEBOUNCE_MS = 600;
const MAX_FINDINGS_SHOWN = 20;

let notes: PyNote[] = [];
let sheet: SheetSnapshot | null = null;
let results: CheckResult[] = [];
let lastReview = 0;
const expanded = new Set<CheckId>();

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

// ─── Review loop ─────────────────────────────────────────────────────────────

let timer: ReturnType<typeof setTimeout> | undefined;
function scheduleReview() {
  if (timer) clearTimeout(timer);
  timer = setTimeout(review, DEBOUNCE_MS);
}

async function review() {
  try {
    sheet = await readActiveSheet();
    results = runChecklist(sheet);
    lastReview = Date.now();
    renderReview();
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

function renderPyPanel() {
  const panel = $("pyPanel");
  panel.textContent = "";
  if (!sheet) return;
  const statuses = pyStatuses(notes, sheet.name, results);
  panel.hidden = statuses.length === 0;
  if (!statuses.length) return;

  const repeats = statuses.filter(s => s.status === "still-an-issue").length;
  panel.appendChild(
    el(
      "h2",
      undefined,
      repeats
        ? `Last year's reviewer flagged this. ${repeats} of ${statuses.length} issue${statuses.length > 1 ? "s" : ""} showing again`
        : `Last year's reviewer left ${statuses.length} note${statuses.length > 1 ? "s" : ""} on this sheet`,
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

function renderChecklist() {
  const list = $("checklist");
  list.textContent = "";
  if (!sheet) return;

  for (const { check, findings } of results) {
    const pass = findings.length === 0;
    const pyRelated = pyNotesForCheck(notes, sheet.name, check.id);
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
      renderChecklist();
    });
    item.appendChild(head);

    if (expanded.has(check.id)) {
      const body = el("div", "item-body");
      if (pass) body.appendChild(el("p", "why", "Nothing to fix here."));

      findings.slice(0, MAX_FINDINGS_SHOWN).forEach(f => {
        const row = el("div", "finding");
        if (f.cell) {
          const link = el("button", "cell-link", f.cell);
          link.title = `Go to ${f.cell}`;
          link.addEventListener("click", () => selectCell(f.cell!).catch(() => toast("Couldn't select that cell")));
          row.appendChild(link);
        }
        row.appendChild(el("span", "finding-msg", f.message));
        body.appendChild(row);
      });
      if (findings.length > MAX_FINDINGS_SHOWN) {
        body.appendChild(el("p", "more", `…and ${findings.length - MAX_FINDINGS_SHOWN} more`));
      }

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

function renderScore() {
  const score = $("score");
  score.textContent = "";
  if (!sheet) return;
  if (!sheet.cells.length) {
    score.textContent = "This sheet is empty. Start working and the checklist will update as you go.";
    return;
  }
  const passed = results.filter(r => r.findings.length === 0).length;
  score.appendChild(el("strong", undefined, `${passed}/${results.length}`));
  score.appendChild(
    document.createTextNode(passed === results.length ? " checks clear. Ready for review." : " checks clear"),
  );
}

function renderReview() {
  renderStatus();
  renderScore();
  renderPyPanel();
  renderChecklist();
  $("truncated").hidden = !(sheet && sheet.truncated);
}

// ─── Notes tab ───────────────────────────────────────────────────────────────

function renderNotesTab() {
  const count = $("notesCount");
  count.hidden = notes.length === 0;
  count.textContent = String(notes.length);

  const wrap = $("notesList");
  wrap.textContent = "";
  if (!notes.length) return;
  wrap.appendChild(el("p", "help", `${notes.length} note${notes.length > 1 ? "s" : ""} saved in this workbook:`));
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
    storeNotes(notes.concat(parsed), `Saved ${parsed.length} note${parsed.length > 1 ? "s" : ""}`);
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
    storeNotes(parsed, `Imported ${parsed.length} note${parsed.length > 1 ? "s" : ""}`);
  });

  $("clearNotes").addEventListener("click", () => {
    if (!notes.length) return;
    storeNotes([], "Notes cleared");
  });
}

// ─── Tabs and start-up ───────────────────────────────────────────────────────

function wireTabs() {
  const tabs = Array.prototype.slice.call(document.querySelectorAll(".tab")) as HTMLButtonElement[];
  tabs.forEach(tab => {
    tab.addEventListener("click", () => {
      tabs.forEach(t => {
        const on = t === tab;
        t.classList.toggle("active", on);
        t.setAttribute("aria-selected", String(on));
        $(`tab-${t.dataset.tab}`).hidden = !on;
      });
    });
  });
}

Office.onReady(async () => {
  wireTabs();
  wireNotesTab();
  // Re-tag on load so notes saved by an older version pick up improved keyword rules.
  notes = loadNotes().map(n => ({ ...n, tags: tagNote(n.text) }));
  renderNotesTab();
  await review();
  try {
    await watchWorkbook(scheduleReview);
  } catch (e) {
    console.error("[Baseline] live updates unavailable", e);
    toast("Live updates unavailable in this Excel version");
  }
  setInterval(renderStatus, 15000);
});

export {};
