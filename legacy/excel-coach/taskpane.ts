/* global Office, Excel */

// ═══════════════════════════════════════════════════════════════════════════════
// EXCEL COACH - Keyboard Efficiency Trainer for Audit Professionals
// ═══════════════════════════════════════════════════════════════════════════════
// Automatic detection via Excel change events. Tracks shortcut usage to build
// muscle memory and reduce time spent on repetitive Excel tasks.
// ═══════════════════════════════════════════════════════════════════════════════

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

type Skill = {
  id: string;
  name: string;
  shortcut: string;
  benefit: string;
  icon: string;
  target: number;
};

type CoachState = {
  version: number;
  progress: Record<string, number>;
  expanded: Record<string, boolean>;
};

// ─────────────────────────────────────────────────────────────────────────────
// Skills Configuration - Essential Shortcuts for Audit Work
// ─────────────────────────────────────────────────────────────────────────────

const SKILLS: Skill[] = [
  {
    id: "autosum",
    name: "AutoSum",
    shortcut: "Alt + =",
    benefit: "Instantly sum columns for tick marks and footing",
    icon: "Σ",
    target: 15,
  },
  {
    id: "edit_cell",
    name: "Edit in Cell",
    shortcut: "F2",
    benefit: "Jump into edit mode without double-clicking",
    icon: "✎",
    target: 20,
  },
  {
    id: "lock_ref",
    name: "Absolute Reference",
    shortcut: "F4",
    benefit: "Lock cell references when copying formulas across workpapers",
    icon: "$",
    target: 12,
  },
  {
    id: "copy_paste",
    name: "Copy & Paste",
    shortcut: "Ctrl+C / Ctrl+V",
    benefit: "Replicate data and formulas efficiently",
    icon: "⧉",
    target: 25,
  },
  {
    id: "paste_values",
    name: "Paste Values",
    shortcut: "Alt, H, V, V",
    benefit: "Remove formulas when finalizing workpapers",
    icon: "V",
    target: 10,
  },
  {
    id: "flash_fill",
    name: "Flash Fill",
    shortcut: "Ctrl+E",
    benefit: "Extract or reformat data patterns instantly",
    icon: "⚡",
    target: 8,
  },
  {
    id: "select_column",
    name: "Select Column",
    shortcut: "Ctrl+Space",
    benefit: "Highlight entire columns for formatting or review",
    icon: "▐",
    target: 10,
  },
  {
    id: "select_row",
    name: "Select Row",
    shortcut: "Shift+Space",
    benefit: "Select entire rows for deletion or insertion",
    icon: "▬",
    target: 10,
  },
  {
    id: "insert_row",
    name: "Insert Row/Column",
    shortcut: "Ctrl+Shift++",
    benefit: "Add rows for new line items without using the ribbon",
    icon: "+",
    target: 8,
  },
  {
    id: "delete_row",
    name: "Delete Row/Column",
    shortcut: "Ctrl+−",
    benefit: "Remove unnecessary rows quickly during cleanup",
    icon: "−",
    target: 8,
  },
  {
    id: "goto",
    name: "Go To / Find",
    shortcut: "Ctrl+G / F5",
    benefit: "Navigate large workbooks or find specific cells",
    icon: "→",
    target: 10,
  },
];

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

const STORAGE_KEY = "excel_coach_v3";
const VERSION = 3;
const DEDUPE_MS = 3000;

// ─────────────────────────────────────────────────────────────────────────────
// Runtime State
// ─────────────────────────────────────────────────────────────────────────────

let state: CoachState;
const lastLog = new Map<string, number>();
const prevFormulas = new Map<string, string>();

// ─────────────────────────────────────────────────────────────────────────────
// State Management
// ─────────────────────────────────────────────────────────────────────────────

function freshState(): CoachState {
  const p: Record<string, number> = {};
  const e: Record<string, boolean> = {};
  SKILLS.forEach(s => { p[s.id] = 0; e[s.id] = false; });
  return { version: VERSION, progress: p, expanded: e };
}

function loadState(): CoachState {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return freshState();
    const s = JSON.parse(raw) as CoachState;
    if (s.version !== VERSION) return freshState();
    SKILLS.forEach(sk => {
      if (!(sk.id in s.progress)) s.progress[sk.id] = 0;
      if (!(sk.id in s.expanded)) s.expanded[sk.id] = false;
    });
    return s;
  } catch {
    return freshState();
  }
}

function saveState() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch {}
}

function resetState() {
  localStorage.removeItem(STORAGE_KEY);
  state = freshState();
  saveState();
  render();
  flash("Progress reset");
}

// ─────────────────────────────────────────────────────────────────────────────
// UI: Flash Notification
// ─────────────────────────────────────────────────────────────────────────────

function flash(msg: string) {
  const el = document.getElementById("flash");
  if (!el) return;
  el.textContent = msg;
  el.classList.remove("show");
  void el.offsetWidth;
  el.classList.add("show");
  setTimeout(() => el.classList.remove("show"), 1500);
}

// ─────────────────────────────────────────────────────────────────────────────
// UI: Render Skill Cards
// ─────────────────────────────────────────────────────────────────────────────

function render() {
  const container = document.getElementById("skillsContainer");
  if (!container) return;
  container.innerHTML = "";

  SKILLS.forEach(s => {
    const count = state.progress[s.id] ?? 0;
    const open = state.expanded[s.id] ?? false;
    const done = count >= s.target;
    const pct = Math.min(100, Math.round((count / s.target) * 100));

    const card = document.createElement("div");
    card.className = `skill${open ? " open" : ""}${done ? " complete" : ""}`;
    card.innerHTML = `
      <div class="skill-header">
        <span class="skill-toggle">▶</span>
        <span class="skill-icon">${s.icon}</span>
        <span class="skill-name">${s.name}</span>
        <span class="skill-count">${count}/${s.target}</span>
      </div>
      <div class="skill-body">
        <div class="skill-shortcut">${s.shortcut}</div>
        <p class="skill-benefit">${s.benefit}</p>
        <div class="skill-bar">
          <div class="skill-fill" style="width:${pct}%"></div>
        </div>
      </div>`;
    card.querySelector(".skill-header")?.addEventListener("click", () => {
      state.expanded[s.id] = !open;
      saveState();
      render();
    });
    container.appendChild(card);
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Detection: Log Skill Usage
// ─────────────────────────────────────────────────────────────────────────────

function log(skillId: string, label: string) {
  const now = Date.now();
  if ((lastLog.get(skillId) ?? 0) + DEDUPE_MS > now) return;
  lastLog.set(skillId, now);

  state.progress[skillId] = (state.progress[skillId] ?? 0) + 1;
  saveState();
  render();
  flash(label);
  console.info(`[Coach] ${label}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Detection: Excel Change Handler
// ─────────────────────────────────────────────────────────────────────────────

async function onCellChange(event: Excel.WorksheetChangedEventArgs) {
  if (!event?.address) return;

  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const rng = ws.getRange(event.address);
      rng.load(["formulas", "values", "rowCount", "columnCount", "address"]);
      await ctx.sync();

      const formulas = rng.formulas as string[][];
      const values = rng.values as any[][];
      const rows = rng.rowCount;
      const cols = rng.columnCount;
      const cellCount = rows * cols;
      const addr = rng.address; // e.g. "Sheet1!A1:C3"

      // Parse base address
      const baseMatch = addr.match(/!?([A-Z]+)(\d+)/i);
      const baseCol = baseMatch ? colNum(baseMatch[1]) : 1;
      const baseRow = baseMatch ? parseInt(baseMatch[2], 10) : 1;

      // Flags for what we detected this change
      let foundSum = false;
      let foundEdit = false;
      let foundLock = false;
      let foundPasteVal = false;
      let foundFlashFill = false;

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const f = formulas[r]?.[c] ?? "";
          const v = values[r]?.[c];
          const isFormula = typeof f === "string" && f.startsWith("=");
          const cellKey = `${colLetter(baseCol + c)}${baseRow + r}`;
          const fUpper = f.toUpperCase();

          // Track formula cells for paste-values detection
          const wasFormula = prevFormulas.has(cellKey);
          if (isFormula) {
            prevFormulas.set(cellKey, f);
          } else if (wasFormula) {
            // Was a formula, now it's a plain value → paste values!
            prevFormulas.delete(cellKey);
            foundPasteVal = true;
          }

          // Detect SUM formula
          if (isFormula && fUpper.includes("SUM(")) {
            foundSum = true;
          }

          // Detect locked reference ($)
          if (isFormula && f.includes("$")) {
            foundLock = true;
          }

          // Detect any formula edit (excluding SUM and $ which have their own quests)
          if (isFormula && !foundSum && !foundLock) {
            foundEdit = true;
          }

          // Detect single-cell value edit (F2 on a non-formula cell)
          if (!isFormula && cellCount === 1 && v !== "" && v !== null) {
            foundEdit = true;
          }
        }
      }

      // Detect Flash Fill: many cells filled at once with non-formula values
      // Flash Fill typically fills 3+ cells with plain values
      if (cellCount >= 3 && !foundSum && !foundLock) {
        let allValues = true;
        for (let r = 0; r < rows && allValues; r++) {
          for (let c = 0; c < cols && allValues; c++) {
            if (formulas[r]?.[c]?.startsWith("=")) allValues = false;
          }
        }
        if (allValues) foundFlashFill = true;
      }

      // Log detected shortcuts (only log each type once per change event)
      if (foundSum) log("autosum", "AutoSum ✓");
      if (foundLock) log("lock_ref", "Lock Ref ✓");
      if (foundEdit && !foundSum && !foundLock) log("edit_cell", "Edit Cell ✓");
      if (foundPasteVal) log("paste_values", "Paste Values ✓");
      if (foundFlashFill && !foundPasteVal) log("flash_fill", "Flash Fill ✓");

      // Multi-cell paste (copy/paste)
      if (cellCount >= 2 && !foundFlashFill && !foundPasteVal) {
        log("copy_paste", "Copy/Paste ✓");
      }
    });
  } catch (e) {
    console.warn("[Coach] Detection error:", e);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Detection: Selection Change Handler (for row/column selection)
// ─────────────────────────────────────────────────────────────────────────────

async function onSelectionChange(event: Excel.WorksheetSelectionChangedEventArgs) {
  if (!event?.address) return;

  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const sel = ws.getRange(event.address);
      sel.load(["rowCount", "columnCount", "address"]);
      await ctx.sync();

      const rows = sel.rowCount;
      const cols = sel.columnCount;

      // Detect full column selection (e.g. "A:A" or selection spans many rows)
      // Full column = 1 column, 1048576 rows (Excel's max)
      if (cols === 1 && rows >= 10000) {
        log("select_column", "Select Column ✓");
      }

      // Detect full row selection (e.g. "1:1" or selection spans many columns)
      if (rows === 1 && cols >= 1000) {
        log("select_row", "Select Row ✓");
      }
    });
  } catch (e) {
    console.warn("[Coach] Selection error:", e);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Detection: Row/Column Insert/Delete
// ─────────────────────────────────────────────────────────────────────────────

let lastRowCount = 0;
let lastColCount = 0;

async function checkRowColChanges() {
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const used = ws.getUsedRangeOrNullObject();
      used.load(["rowCount", "columnCount"]);
      await ctx.sync();

      if (used.isNullObject) {
        lastRowCount = 0;
        lastColCount = 0;
        return;
      }

      const rows = used.rowCount;
      const cols = used.columnCount;

      if (lastRowCount > 0 || lastColCount > 0) {
        if (rows > lastRowCount || cols > lastColCount) {
          log("insert_row", "Insert Row/Col ✓");
        } else if (rows < lastRowCount || cols < lastColCount) {
          log("delete_row", "Delete Row/Col ✓");
        }
      }

      lastRowCount = rows;
      lastColCount = cols;
    });
  } catch {}
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function colNum(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

function colLetter(num: number): string {
  let s = "";
  while (num > 0) {
    num--;
    s = String.fromCharCode((num % 26) + 65) + s;
    num = Math.floor(num / 26);
  }
  return s;
}

// ─────────────────────────────────────────────────────────────────────────────
// Snapshot existing formulas on load (for paste-values detection)
// ─────────────────────────────────────────────────────────────────────────────

async function snapshotFormulas() {
  try {
    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const used = ws.getUsedRangeOrNullObject();
      used.load(["formulas", "address", "rowCount", "columnCount"]);
      await ctx.sync();

      if (used.isNullObject) return;

      const formulas = used.formulas as string[][];
      const addr = used.address;
      const baseMatch = addr.match(/!?([A-Z]+)(\d+)/i);
      const baseCol = baseMatch ? colNum(baseMatch[1]) : 1;
      const baseRow = baseMatch ? parseInt(baseMatch[2], 10) : 1;

      let count = 0;
      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < (formulas[r]?.length ?? 0); c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) {
            const cellKey = `${colLetter(baseCol + c)}${baseRow + r}`;
            prevFormulas.set(cellKey, f);
            count++;
          }
        }
      }
      console.info(`[Coach] Snapshot: ${count} formulas tracked`);
    });
  } catch (e) {
    console.warn("[Coach] Snapshot error:", e);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Initialization
// ─────────────────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  state = loadState();
  saveState();
  render();

  // Reset button
  document.getElementById("resetDemo")?.addEventListener("click", resetState);

  // Excel event listeners
  if (typeof Excel !== "undefined") {
    await snapshotFormulas();

    await Excel.run(async ctx => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      
      // Listen for cell changes
      ws.onChanged.add(async (e) => {
        await onCellChange(e);
        await checkRowColChanges();
      });

      // Listen for selection changes
      ws.onSelectionChanged.add(onSelectionChange);

      await ctx.sync();
      console.info("[Coach] Excel listeners attached");
    }).catch(e => console.error("[Coach] Setup error:", e));

    // Initial row/col count
    await checkRowColChanges();
  }
});

export {};
