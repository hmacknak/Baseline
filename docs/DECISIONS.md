# Decisions & Open Questions

## Provenance (2026-10-06)
- This project began from an earlier prototype (an Excel keyboard-shortcut coach). That code has since been removed.

## 2026-10-06: v0.1 Review Bot
- **First product = live review checklist + prior-year notes.** Chosen as the simplest high-value version.
  The original shortcut-coach prototype was later removed.
- **Local only, no AI.** All checks run inside Excel, so no client data leaves the workbook and the
  client-data approval question doesn't block v0.1. AI "talk" is a saved idea.
- **Checks are deterministic rules** (`src/review/checks.ts`), tuned to avoid noise: hard-code detection only
  flags numbers used in arithmetic (not function arguments like `ROUND(x,2)`), and ignores 0 and 1.
- **PY notes are tagged by keyword** (`src/review/notes.ts`) and stored in workbook settings so they roll
  forward with the file. Sheet names are matched ignoring years/FY labels.
- **Common-issue analysis deferred** (owner: "could come later, just save the thought"). Recorded in VISION.md.
- **Hosting: GitHub Pages** (owner approved publishing). `.github/workflows/pages.yml` runs tests, builds and
  publishes `dist/` to the configured hosting address on every push to `main`. The production manifest
  points there; `npm start` still uses localhost for development. Only built add-in code is published, never
  workbook data.
- `npm run validate` uses Microsoft's online validator; it couldn't run in the build sandbox (network blocked).

## 2026-10-08: v0.2 Reviewer catalog
- **Two kinds of reviewer.** *Workpaper basics* (the v0.1 checks) run on every sheet. *Procedure reviewers* are
  chosen per sheet from a catalog. Many procedures are planned, so each reviewer is a small data-driven file
  (`src/review/reviewers/`) built on a shared table reader (`src/review/table.ts`).
- **Suggest, then confirm** (owner's choice). Baseline scores each reviewer against the sheet name, the words on the sheet
  and the recognised column headings; a score of 0.6 or more shows "This looks like a … Use it?". The choice (or "Not
  now") is saved per sheet in workbook settings, keyed by the sheet name without year labels so it rolls forward.
- **Insights are panel-only** (owner's choice). High-risk items are listed in the add-in with click-to-jump.
  Nothing is written to the workbook.
- **Year-end and threshold** are set once per workbook (Settings tab, saved in the workbook). Insights that need
  them say so instead of guessing. Baseline offers a year-end it spots in a sheet title, but never applies it silently.
- **Dates** are handled as Excel serial numbers. Text dates are read month-first (en-US Excel) unless the first
  number is over 12.
- **PY notes can match insights**, not just checklist items (e.g. a cut-off note becomes *Again this year* when the
  cut-off insight finds items). Bare "error" no longer tags a note as a formula error ("cut-off error" was misread).
- **"Coming soon" cards** show planned reviewers in the catalog so the direction is visible.

## 2026-10-08: Professional presentation, no emoji
- **No emoji anywhere**: add-in, website, docs, demo, video. The audience is professional audit users.
  Reviewers are identified by a short monogram (`code`, e.g. "BR" Bank reconciliation, "UL" Unrecorded
  liabilities) rendered as a plain badge. The add-in icon is a "B" monogram.
- Allowed symbols: ✓ / ✗ as pass/fail marks in the checklist, and audit tickmarks (✓ √ ^ …) inside workpapers,
  which are standard audit notation.

## 2026-10-08: Self-updating add-in
- Every publish writes `version.json` (the commit id) next to the add-in and bakes the same id into the code.
- **On opening**, the panel checks `version.json`; if a newer build is live it reloads into it once (guarded per
  version, so a lagging cache can't cause a loop). Safe because settings, reviewer choices and PY notes live in
  the workbook.
- **While open**, it checks every 30 minutes and shows an "Update now" banner instead of reloading mid-task.
- **Settings → Check for updates** shows the current version and applies an update on demand.
- Bundles are content-hashed (`taskpane.<hash>.js`), so Excel can't keep serving a stale cached copy.
- The manifest URL never changes, so nobody needs to re-upload the add-in file to get updates.

## 2026-11-11: Focus on Workpaper basics
- **Procedure reviewers switched off** (`src/config.ts`, `PROCEDURES_ENABLED = false`). The code and tests stay;
  the Catalog tab, suggestions, insights and threshold setting are hidden.
- **Basics grew from 9 to 15 checks**: typed numbers anywhere in a row/column of formulas (not just between two),
  numbers stored as text, table completeness (blank cells in mostly-filled columns; comment/tick columns optional),
  placeholders (TBD, XXX, ???, "to be completed"), dates (prior-year year-end left after roll-forward, year typos,
  text dates, mixed formats, review dated before preparation), stray unlabelled numbers off to the side, leftover
  highlighting (bright fills; pale/dark heading fills and colour legends are fine), sign-offs must carry a date,
  a conclusion must be documented, and every tickmark symbol must be in the legend.
- **Mark OK with a reason**: any flag can be accepted with a reason, saved in the workbook. If the cell changes, the
  flag comes back. Accepted flags are listed under each check with Undo.
- **Step through** Issues or Dates one at a time (selects each cell), and **Review all sheets** for a one-pass summary.
- Reading fill colours needs ExcelApi 1.9; on older Excel the highlighting check simply finds nothing.

## 2026-11-12: Procedures back on: SURL, cash confirmations, expense vouching
- `PROCEDURES_ENABLED = true`; only `ENABLED_PROCEDURES` (SURL, cash confirmations, expense vouching) appear in
  the catalog. Bank reconciliation, revenue cut-off and AR confirmations stay in the code with their tests.
  "Coming soon" cards were removed.
- Each procedure check and insight now carries the **assertion** at risk and the **reference** (ISA, PCAOB AS or
  IFRS paragraph) behind it, shown under "Why it matters". Research and sources: `docs/PROCEDURES.md`.
- New: cash confirmations (difference, no reply, not direct, not sent incl. nil balances, wrong date, overdraft,
  foreign currency) and expense vouching (amount vs support, cut-off, duplicates, approval, capital items, wrong
  account). SURL adds over-accrual for next year's services and a short search window.
- Cash confirmations are only suggested when the sheet has a bank column, so customer (AR) confirmations aren't
  mistaken for bank ones. Expense vouching is only suggested when the sheet's name or wording says so.
- The demo video is not re-recorded for every change; only on request.

## 2026-11-11: Demo video storyboard v2
- Full demo cut to about 2.5 minutes, plus a 54-second cut for email and chat (`demo/Baseline-demo-short.mp4`).
  One finding per sheet; year-end and threshold are pre-set in the recording and mentioned in a caption.
- Wording follows an external review of the captions: "first-pass review", never implying it replaces review; no
  "Ready for review" state shown in the video.
- Demo workbook made more realistic: an explained bank difference beside the unexplained one, a duplicate paid nine
  days apart at a non-round amount, last year's notes in a reviewer's voice.

## 2026-11-11: "No open issues" instead of "Ready for review"
- A clean sheet now says "No open issues" (and the workbook view adds that this is a first pass, not a review
  sign-off). "Ready for review" read as the tool approving the work, which it can't do.
## 2026-11-11: US GAAP references
- New Settings choice: accounting framework, IFRS (default) or US GAAP, saved in the workbook. Under US GAAP the
  procedure findings cite the ASC (and FASB Concepts Statement 8 for the liability definition) instead of IFRS, and
  the overdraft explanation changes because US GAAP presents overdrafts as liabilities.
- Audit references (ISA, PCAOB AS) don't change with the framework.
- Unresolved: the FASB Codification itself couldn't be opened; paragraph-level US GAAP references were checked
  against Deloitte, PwC and KPMG guides. Only ASC 830-20-35-2 is cited at paragraph level.

## 2026-11-11: Less noise on ordinary workbooks
Measured on 180 sheets from 130 public sample workbooks (`tools/noise`), then fixed what fired wrongly:
- **Sign-off and purpose/source/conclusion** now apply only to sheets that look like workpapers (a workpaper
  reference or FY in the tab name, a period heading, or any sign-off/purpose/conclusion/legend line). Client
  listings and exports are no longer flagged. Before: 180 of 180 sheets flagged. After: 3.
  Trade-off: a brand-new tab with no heading at all gets no sign-off reminder until it has one.
- **Dates:** a "typo" is only called where the other dates sit close together, so fixed-asset, loan and lease
  dates aren't flagged. Comparative columns ("31 December 2024" beside this year's) aren't "rolled forward".
  Padding differences (dd-mmm-yy vs d-mmm-yyyy) aren't "mixed formats". Zero shown as a date is ignored.
- **Hard-codes:** unit conversions (/12, *100, /1000, /365, ^2, /2) are not flagged. Rates, prices and added
  amounts still are.
- **Stray numbers:** only checked once the sheet has a working table; a block of more than 8 unlabelled numbers
  is treated as data. Also about 12x faster on wide sheets.
- **Highlighting:** a colour on more than 20 cells is treated as a convention.
- **Panel:** "No open issues" is never shown while a procedure insight has flagged something or could not run.
  Step-through and "Review all sheets" include procedure findings. Sheets are sized by their contents, not their
  formatting, so a small workpaper with formatting dragged down is no longer cut off.
- Not done yet: marking a procedure finding OK; expiring last year's "Mark OK" decisions on roll-forward.

## 2026-11-12: Demo workbook looks like a workpaper
- The three procedure sheets now use a standard workpaper layout: client, subject and period top left; a boxed
  W/P reference and sign-offs top right; objective and source; a boxed testing table with a dark heading row;
  red tickmarks; a blue cross-reference to the lead schedule. Generic conventions, not any firm's template.
- The planted findings are unchanged. The simulated spreadsheet used for the video now shows borders, font
  colours and fills. Both videos were re-recorded.
- A lone sentence directly under a table (a conclusion or note) is no longer read as a table row with blanks.

## 2026-11-12: Demo video in three parts
- The video now says what Baseline is ("Live feedback on your workpaper, as you work") and shows three things, each
  behind a short title card: workpaper basics; procedure review and the why; last year's review notes. The "why"
  gets the most screen time because it is the part that teaches.
- Dropped from the video: the "fewer review notes" framing, workpaper-reference jargon in captions, Mark OK and
  Review all sheets. Captions sit over the sheet so they never cover the panel.
- The download page and README use the same description.

## 2026-11-13: Download page redesigned
- The page now leads with what Baseline is ("Live feedback on your workpaper, as you work"), then the three things
  it does, an example flag with its why, what it checks, and how to get started. Same single static file, no
  scripts, no external fonts or trackers, light and dark.
- States the limits plainly: nothing leaves Excel; a first pass, not a replacement for review; needs Microsoft 365
  or Excel on the web.
- The footer credits the two people who built it, at the owner's request. This is a deliberate exception to the
  "nothing personal in the repo" rule; no other personal details are included.

## 2026-11-13: Every flag says when it couldn't check
- A check now ends flagged, clear, or **not run with a reason**; a check that couldn't run is never shown as a
  pass and "No open issues" is never shown while one is outstanding. Reasons: setting missing, column missing, no
  testing table, fill colours unreadable, not a workpaper.
- New procedure check "Baseline can read every answer" reports values the rules can't read instead of skipping
  them. Yes/no answers are read by their first word; a bare "x" is reported as unclear (it used to count as yes).
- Text dates are read day-first or month-first from the evidence on the sheet; with none, the cell is reported.
- Repeated findings are grouped into one flag with a range (errors, hard-codes, external links).
- Fewer false alarms: subtotals between item formulas, masked account numbers, sign-off forms read as tables, a
  conclusion under a testing table read as a tested item, "Not yet" in a received-date cell read as a reply.
- More reach: totals that start late, SUBTOTAL totals, letter tickmarks in a "Tick" column, sign-off applied to
  any sheet with a procedure reviewer on it.
- The full register of what each flag can and can't see is `docs/LIMITATIONS.md`.
