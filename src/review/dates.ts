// Dates as Excel serial day numbers (1 = 1 Jan 1900, 45657 = 31 Dec 2024).
// Working in serials keeps comparisons trivial and matches what Excel hands us.

const EPOCH_MS = Date.UTC(1899, 11, 30);
const DAY_MS = 86400000;

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

export function serialFromYMD(year: number, month: number, day: number): number {
  return Math.round((Date.UTC(year, month - 1, day) - EPOCH_MS) / DAY_MS);
}

export function ymdFromSerial(serial: number): { year: number; month: number; day: number } {
  const d = new Date(EPOCH_MS + Math.floor(serial) * DAY_MS);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

/** "31 Dec 2025" */
export function formatSerial(serial: number): string {
  const { year, month, day } = ymdFromSerial(serial);
  return `${day} ${MONTH_NAMES[month - 1]} ${year}`;
}

/** "2025-12-31", for <input type="date">. */
export function isoFromSerial(serial: number): string {
  const { year, month, day } = ymdFromSerial(serial);
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function valid(y: number, m: number, d: number): boolean {
  return y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31;
}

function monthIndex(word: string): number {
  return MONTHS.indexOf(word.slice(0, 3).toLowerCase()) + 1;
}

/**
 * Turn a cell value into a serial date, or null.
 * Numbers in a plausible range (1990–2100) are serials, which is how Excel stores dates. Text accepts
 * 2025-12-31, 31-Dec-2025, 31 Dec 2025, Dec 31, 2025 and 12/31/2025 (numeric dates are read
 * month-first, like en-US Excel, unless the first number is over 12).
 */
export function toSerial(value: unknown): number | null {
  if (typeof value === "number") {
    return value >= 32874 && value <= 73051 ? Math.floor(value) : null;
  }
  if (typeof value !== "string") return null;
  const s = value.trim();
  if (!s) return null;
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m && valid(+m[1], +m[2], +m[3])) return serialFromYMD(+m[1], +m[2], +m[3]);
  m = /^(\d{1,2})[\s\-/.]+([A-Za-z]{3,9})[\s\-/.,]+(\d{2,4})$/.exec(s);
  if (m && monthIndex(m[2])) {
    const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    if (valid(y, monthIndex(m[2]), +m[1])) return serialFromYMD(y, monthIndex(m[2]), +m[1]);
  }
  m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/.exec(s);
  if (m && monthIndex(m[1]) && valid(+m[3], monthIndex(m[1]), +m[2])) {
    return serialFromYMD(+m[3], monthIndex(m[1]), +m[2]);
  }
  m = /^(\d{1,2})[/.](\d{1,2})[/.](\d{2,4})$/.exec(s);
  if (m) {
    const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    let month = +m[1];
    let day = +m[2];
    if (month > 12) [month, day] = [day, month];
    if (valid(y, month, day)) return serialFromYMD(y, month, day);
  }
  return null;
}

/** Amounts arrive as numbers, or as text like "$1,234.50" or "(1,234)". */
export function toAmount(value: unknown): number | null {
  if (typeof value === "number") return isFinite(value) ? value : null;
  if (typeof value !== "string") return null;
  let s = value.trim();
  if (!s) return null;
  let negative = false;
  if (/^\(.*\)$/.test(s)) {
    negative = true;
    s = s.slice(1, -1);
  }
  s = s.replace(/[$€£,\s]|CAD|USD/gi, "");
  if (s.charAt(0) === "-") {
    negative = !negative;
    s = s.slice(1);
  }
  if (!/^\d+(\.\d+)?$/.test(s)) return null;
  const n = Number(s);
  return negative ? -n : n;
}

/** Find a year-end mentioned in sheet titles: "31 Dec 2025", "December 31, 2025", "year ended 2025-12-31". */
export function findYearEndInText(texts: string[]): number | null {
  for (const t of texts) {
    const m =
      /\b(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?,?\s+(\d{4})\b/i.exec(t) ||
      /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b/i.exec(t);
    if (!m) continue;
    const dayFirst = /^\d/.test(m[1]);
    const day = +(dayFirst ? m[1] : m[2]);
    const month = monthIndex(dayFirst ? m[2] : m[1]);
    const year = +m[3];
    if (!valid(year, month, day)) continue;
    // Only month-ends look like year-ends.
    const serial = serialFromYMD(year, month, day);
    if (ymdFromSerial(serial + 1).day === 1) return serial;
  }
  return null;
}
