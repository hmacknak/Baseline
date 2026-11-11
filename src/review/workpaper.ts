// Is this sheet a workpaper, or supporting data (a client listing, a TB export, a calculation tab)?
// Documentation checks (sign-off, purpose, source, conclusion) only make sense on a workpaper; on a
// pasted data tab they fire every time and drown out real findings.

import { allCells, isText } from "./layout";
import { SheetSnapshot } from "./types";

/** Text that only appears on a workpaper: sign-off lines, purpose, source, conclusion, tickmark legend. */
const MARKER =
  /^\s*(prepared\s*by|preparer\b|prep'?d\s*by|reviewed\s*by|reviewer\b|purpose\b|objectives?\b|sources?\b|conclu(sion|ded)\b|scope\b|procedures?( performed)?\s*:|tick\s*-?\s*marks?|legend\b|w\/?p\s*ref|workpaper|working paper)/i;

const MONTH = "(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\\.?";

/** A reporting-period heading, e.g. "Year ended 31 December 2025", "As at 30 June 2026", "FY25". */
const PERIOD = new RegExp(
  [
    "\\b(year|period|quarter|month|half[- ]year|twelve months)s?\\s+end(ed|ing)\\b",
    "\\bas\\s+(at|of)\\s+\\d",
    "\\bas\\s+(at|of)\\s+" + MONTH,
    "\\bfor the (year|period|quarter)\\b",
    "\\bFY\\s?'?\\d{2,4}\\b",
    "\\b\\d{1,2}\\s+" + MONTH + ",?\\s+\\d{4}\\b",
    "\\b" + MONTH + "\\s+\\d{1,2},\\s+\\d{4}\\b",
  ].join("|"),
  "i",
);

/** Sheet names that carry a workpaper reference, e.g. "C-2 Cash", "AP-3 SURL", "E4", "FY25 Cash". */
const WP_NAME = /^\s*([A-Z]{1,3}\s?[-.]?\s?\d{1,3}(\.\d+)?\b|[Ff][Yy]\s?'?\d{2,4}\b)/;

/** Last year's review notes pasted in for Baseline to read: reference material, not a workpaper. */
const NOTES_SHEET = /^\s*py\s*(review\s*)?notes?\s*$/i;

const HEADING_ROWS = 6;

const cache = new WeakMap<SheetSnapshot, boolean>();

export function looksLikeWorkpaper(sheet: SheetSnapshot): boolean {
  const hit = cache.get(sheet);
  if (hit !== undefined) return hit;
  const notes = NOTES_SHEET.test(sheet.name);
  let yes = !notes && WP_NAME.test(sheet.name);
  if (!yes && !notes) {
    for (const c of allCells(sheet)) {
      if (!isText(c)) continue;
      const text = String(c.value);
      if (MARKER.test(text) || (c.row < sheet.originRow + HEADING_ROWS && PERIOD.test(text))) {
        yes = true;
        break;
      }
    }
  }
  cache.set(sheet, yes);
  return yes;
}
