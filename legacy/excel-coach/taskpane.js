/* global Office */

const TIPS = [
  { id: "autosum_alt_equals", title: "AutoSum instantly", body: "Instead of typing =SUM(...), press <b>Alt + =</b>. Goal: use it in real work this week.", weeklyGoal: 15 },
  { id: "filters_ctrl_shift_l", title: "Toggle filters fast", body: "Press <b>Ctrl + Shift + L</b> to add/remove filters.", weeklyGoal: 12 },
  { id: "freeze_panes", title: "Freeze panes", body: "Use View → Freeze Panes to keep headers visible while scrolling.", weeklyGoal: 8 },
  { id: "format_painter", title: "Format Painter", body: "Use the Format Painter to copy formatting quickly.", weeklyGoal: 10 },
  { id: "quick_fill", title: "Flash Fill", body: "Use Flash Fill (Ctrl+E) to auto-complete patterns.", weeklyGoal: 7 },
  { id: "named_ranges", title: "Named ranges", body: "Give ranges names for easier formulas and navigation.", weeklyGoal: 6 },
  { id: "pivot_basics", title: "Pivot basics", body: "Create a PivotTable to summarize data quickly.", weeklyGoal: 5 },
  { id: "conditional_format", title: "Conditional formatting", body: "Highlight important values with conditional formatting.", weeklyGoal: 9 },
  { id: "keyboard_nav", title: "Keyboard navigation", body: "Use Ctrl+Arrow keys and Ctrl+Home/End to move around sheets faster.", weeklyGoal: 12 },
  { id: "text_to_columns", title: "Text to Columns", body: "Split text into columns using Data → Text to Columns.", weeklyGoal: 4 },
];

const STORAGE_KEY = "excel_coach_state_v1";
const STATE_VERSION = 1;

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function mondayOfThisWeekISO() {
  const d = new Date();
  const day = d.getDay(); // Sun=0..Sat=6
  const diffToMonday = (day === 0 ? -6 : 1) - day;
  d.setDate(d.getDate() + diffToMonday);

  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${da}`;
}

function tipIndexFromWeekStart(weekStartISO) {
  var epoch = new Date("2020-01-06T00:00:00Z");
  var wk = new Date(weekStartISO + "T00:00:00Z");
  var diffMs = wk.getTime() - epoch.getTime();
  var weeks = Math.floor(diffMs / (7 * 24 * 60 * 60 * 1000));
  var idx = ((weeks % TIPS.length) + TIPS.length) % TIPS.length;
  return idx;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (parsed.version !== STATE_VERSION) return null;

    if (typeof parsed.repsThisWeek === "number") {
      const repsByTip = {};
      const completedByTip = {};
      repsByTip[parsed.currentTipId] = parsed.repsThisWeek || 0;
      completedByTip[parsed.currentTipId] = !!parsed.completedThisWeek;
      parsed.repsByTip = repsByTip;
      parsed.completedByTip = completedByTip;
      delete parsed.repsThisWeek;
      delete parsed.completedThisWeek;
    }

    return parsed;
  } catch (e) {
    return null;
  }
}

function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function defaultState() {
  return {
    version: STATE_VERSION,
    currentTipId: TIPS[0].id,
    weekStartISO: mondayOfThisWeekISO(),
    repsByTip: {},
    completedByTip: {},
    lastFlushISO: todayISO(),
  };
}

function dailyFlush(state) {
  const t = todayISO();
  if (state.lastFlushISO === t) return state;
  return Object.assign({}, state, { lastFlushISO: t });
}

function weeklyRollover(state) {
  const currentWeekStart = mondayOfThisWeekISO();
  if (state.weekStartISO === currentWeekStart) return state;
  var idx = tipIndexFromWeekStart(currentWeekStart);
  var tip = TIPS[idx];

  var next = Object.assign({}, state);
  next.currentTipId = tip.id;
  next.weekStartISO = currentWeekStart;
  next.repsByTip = {};
  next.completedByTip = {};
  return next;
}

function getTip(state) {
  for (var i = 0; i < TIPS.length; i++) {
    if (TIPS[i].id === state.currentTipId) return TIPS[i];
  // Compiled JS removed — build from TypeScript (keep source-only here).
  return TIPS[0];
