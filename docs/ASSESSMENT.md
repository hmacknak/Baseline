# Baseline: code and privacy assessment

Snapshot: `main` at `19c49e2`, 24 commits across all branches, 137 automated tests passing.
Everything below was checked in the code. Where something can only be confirmed in real Excel or in GitHub
settings, it says **Not verifiable from code**.

---

## Top 10 findings (highest risk first)

| # | Finding | Evidence | Severity | Effort | Fix |
|---|---|---|---|---|---|
| 1 | **Personal and identifying information is in the repo and its history, and the repo is public.** | `CLAUDE.md:3` (health), `CLAUDE.md:6` (firm types), personal handle in 7 files (see privacy audit), commit author emails and timezone (history section) | High | M | Stage 1: fresh repo from cleaned files. |
| 2 | **Whoever controls the personal GitHub account controls code that runs inside every user's Excel.** The add-in loads from GitHub Pages every time it opens and auto-switches to the newest build. It has permission to read and write the whole workbook. | `webpack.config.js:27`, `manifest.xml:17,30`, `src/update.ts:45,55-61`, `src/taskpane/taskpane.ts:924` | High | M | Firm-controlled hosting, protected `main`, tagged releases with a manual approval, optional version pin. |
| 3 | **Every push to `main` goes live immediately.** No tag, no approval step, no lint. `workflow_dispatch` has no branch filter, so a manual run may deploy any branch (whether GitHub's Pages environment blocks that is **not verifiable from code**). | `.github/workflows/pages.yml:3-7,27-30,35-44`; lint exists but isn't run: `package.json:16` | High | S | Deploy on version tags only, with environment approval; add lint to CI. |
| 4 | **"Nothing leaves Excel" is true of today's code but isn't enforced.** No content-security policy, so a future change could send data anywhere. GitHub also sees every user's IP address, on every panel open and every 30 minutes. | No CSP in `src/taskpane/taskpane.html`; calls at `src/update.ts:22-23`, `src/taskpane/taskpane.ts:790,950`; claim at `site/index.html:89`, `README.md:74` | Medium | S | Add a CSP allowing only Microsoft's Office script and the add-in's own host; reword the claim. |
| 5 | **Text users type is saved inside the workbook and travels with it.** That covers prior-year notes and "Mark OK" reasons. A reason like "agreed with [client CFO]" goes wherever the file goes. | `src/excel/workbook.ts:170-180`, keys at `:12-14`; reasons saved at `src/taskpane/taskpane.ts:226` | Medium | S | Say so in the UI, and add a "clear Baseline data" button. |
| 6 | **Table reading is brittle.** Procedure reviewers look only at the first 40 rows and pick one table. The basics checks stop a table at its first blank row. Merged cells and hidden rows aren't handled at all (no code mentions them). | `src/review/table.ts:81,89-93,108`; `src/review/layout.ts:99` | Medium | M | Scan the whole sheet, support several tables, and add merged/hidden-row tests in real Excel. |
| 7 | **Dates and amounts can be misread.** Text dates are read month-first, so 05/01/2026 is May 1. Two-digit years always mean 20xx. Any number between 32874 and 73051 counts as a date in a date column. Amounts like "1.234,56", "500-" and "CR"/"DR" aren't understood. | `src/review/dates.ts:43,47,68,75-92` | Medium | S | Make the date order a setting; extend amount parsing; add tests. |
| 8 | **"Looks fixed" on a prior-year note can be falsely reassuring.** It means no issue of the same *type* is on the sheet, not that the specific note was dealt with. | `src/review/notes.ts:25-50,152-180`, cell ignored at `:10` | Medium | S | Relabel to "No similar issue found"; match on the cell where one is given. |
| 9 | **Large sheets may be slow.** Up to 100,000 cells are re-read 0.6 seconds after every edit, and fill colours for up to 20,000 cells. "Review all sheets" reads every sheet one after another. Real performance is **not verifiable from code**. | `src/excel/workbook.ts:10,33,56-58,65,113-123`; `src/taskpane/taskpane.ts:40` | Medium | M | Re-read only changed areas; test on a large real workbook. |
| 10 | **Developer tools have 58 known vulnerabilities (8 critical).** None are in the code users receive. One dev tool (`office-addin-usage-data`) sends usage data to Microsoft. | `npm audit`: 58 total; `npm audit --omit=dev`: 0 | Low | S | Upgrade dev dependencies; opt out of Office CLI telemetry. |

---

## 1. Architecture map

```
Excel workbook
   │  read by src/excel/workbook.ts  (snapshotOf: values, formulas, number formats, fill colours)
   ▼
SheetSnapshot  (src/review/types.ts:6-28)
   │  checked by src/review/  (no Office code here, so it is fully testable)
   │     checks.ts      → 15 "Workpaper basics" checks       (list at checks.ts:371-387)
   │     quality.ts     → completeness, dates, stray numbers, highlighting, etc.
   │     layout.ts      → finds tables, dates, the main working area
   │     notes.ts       → prior-year note matching
   │     dismiss.ts     → "Mark OK" filtering
   │     reviewers/     → procedure reviewers (switched off: src/config.ts)
   │     review.ts      → runs basics + the chosen reviewer   (review.ts:28-41)
   ▼
src/taskpane/taskpane.ts  → draws the panel; re-reviews on every edit (taskpane.ts:133-136, 944)
```

- **Where findings are created:** inside each check's `run()` (e.g. `checks.ts`, `quality.ts`), and inside reviewer insights (`reviewers/procedure.ts`, `rowInsight`).
- **What a finding carries** (`types.ts:50-55`): `checkId`, an optional `cell` (e.g. "F17"), and a plain-English `message`. Severity, "why it matters" and "how to fix" live on the check, not on the finding (`types.ts:57-69`). Insight items carry `cell`, `label`, an optional `amount` and `detail` (`types.ts:78-85`).

## 2. Duplication

- **The four procedure reviewers don't repeat logic.** They already share helpers: `completenessCheck`, `columnsCheck`, `rowInsight` and `detectScore` in `reviewers/procedure.ts`. What repeats is column-name lists ("invoice date", "amount", "conclusion"): `surl.ts:13,21,28`, `revenueCutoff.ts:16,24,29`. A shared list of common column names is worth it, but only that.
- **Real duplication is elsewhere:**
  - Two separate table readers: `table.ts:87` (reviewers) and `layout.ts:82` (basics). They already disagree: "Total" detection differs between `table.ts:80` and `layout.ts:66`.
  - Copied helpers: `cellAt`/`allCells` at `checks.ts:12,18` and `layout.ts:7,12`; `isBlank` at `table.ts:55` and `layout.ts:18`.
  - Three different "date inside text" patterns: `checks.ts:259`, `layout.ts:47`, `quality.ts:107`.
  - Fix: one table reader and one helpers file. Low risk once tests cover both uses.

## 3. Fragility

- **How reviewers find headers:** each reviewer lists synonyms per column. A heading counts if it exactly equals a synonym or contains one as a whole phrase (`table.ts:39-49`). The row matching the most columns within the first 40 rows wins, with at least 3 matches needed (`table.ts:81,87-93`; `procedure.ts:22`). Sheet suggestions score the sheet name, words on the sheet and matched columns (`procedure.ts:155-168`), and need 0.6 (`review.ts:43`).
- **What breaks:**
  - **Renamed headers:** headings not in the synonym lists are ignored, e.g. "Svc dt" (`surl.ts:13-28`).
  - **Merged cells:** not handled. **Not verifiable from code** how Excel reports them; probably only the first cell holds the text.
  - **Hidden rows or columns:** read and checked like visible ones; nothing tells the user. Hidden *sheets* are skipped by "Review all sheets" (`workbook.ts:120`).
  - **Several tables:** reviewers see only one (`table.ts:93`). Basics handle several (`layout.ts:82-114`).
  - **Blank rows:** reviewers stop after two (`table.ts:108`); basics stop after one, so rows below a gap aren't checked for completeness (`layout.ts:99`).
  - **Very large sheets:** cut off after 100,000 cells, with a warning (`workbook.ts:56-58`). Highlight checks are skipped above 20,000 cells (`workbook.ts:33,65`).
- **Parsing:**
  - **Dates** are Excel serial numbers. Text dates are read month-first unless the first number is over 12 (`dates.ts:43-71`).
  - **Currency** symbols removed: `$ € £`, "CAD", "USD" (`dates.ts:85`).
  - **Negatives:** a leading "-" or (brackets) (`dates.ts:80-92`). Not supported: trailing "-", "CR"/"DR", or comma-as-decimal formats.

## 4. Prior-year note matching

- **How it works:** each note is tagged by keyword rules (`notes.ts:25-50`), e.g. "hard-cod" means hard-code. Notes match a sheet by name with years and "FY" removed (`notes.ts:125-138`). A note shows **Again this year** if any check with that tag is failing on the sheet, **Looks fixed** if those checks pass, and **Check yourself** if no check covers the tag (`notes.ts:152-180`).
- **False positives:**
  - Any other hard-code on the sheet shows the note as "Again this year".
  - Notes with no sheet apply to every sheet (`notes.ts:138`).
  - Sheet names differing only by a two-digit number collide: "Note 12" equals "Note 13" (`notes.ts:130`).
  - Broad words: "source" tags the note as a header issue (`notes.ts:33`).
- **False negatives:** wording the rules don't know becomes "Check yourself", and "Looks fixed" can appear while the specific issue remains (finding 8).
- **Tested:** tagging examples, parsing, year-stripping, the three statuses and insight matching (`tests/notes.test.ts`, 17 tests; `tests/reviewers.test.ts`). **Not tested:** the name collision, notes without a sheet, and real reviewer wording at scale.

## 5. Network and data flow

| What | Where | Sends workbook data? |
|---|---|---|
| Office.js from Microsoft's CDN | `src/taskpane/taskpane.html:9` | No (standard for all add-ins) |
| Add-in code and icons from GitHub Pages | `manifest.xml` (built by `webpack.config.js:27,132`) | No, but GitHub logs IP and browser |
| Version check: GET `version.json` on open and every 30 min | `src/update.ts:22-23`, `taskpane.ts:790,950,924` | No: no body, no workbook values |
| Session storage (update loop guard) | `src/update.ts:59-60` | Stays in the browser |
| Workbook settings: notes, year-end, "Mark OK" reasons | `src/excel/workbook.ts:12-14,170-180` | Saved **inside the file**, so it travels with it |
| Fonts, analytics, cookies, other telemetry | None found in `src/` or `site/` | n/a |

**Verdict:** today's code sends no workbook content anywhere. The claim isn't *enforced* (finding 4), and data typed into Baseline does travel inside the workbook (finding 5).

## 6. Auto-update

- **How:** each build stamps a version, and `version.json` publishes the latest (`webpack.config.js` version stamping; `src/update.ts`). On opening, the panel reloads into a newer build once (`update.ts:55-61`); while open it checks every 30 minutes and shows an "Update now" banner (`taskpane.ts:790,794-800`).
- **Code is loaded from** `https://hmacknak.github.io/Baseline/` (`webpack.config.js:27`).
- **Who can change it:** anyone with write access to the GitHub repo, or anyone who takes over the personal account. Branch protection is **not verifiable from code**; all work so far was pushed straight to `main`.
- **Pinning a version:** not possible. Only development builds skip updates (`update.ts:55`). Rollback means reverting on `main` and re-publishing.

## 7. Release process

- **Today:**
  - push to `main` → tests, type-check, build → deploy (`pages.yml:3-44`)
  - pull requests run the checks but don't deploy (`pages.yml:36`)
  - no tags, approvals or release notes
- **Controlled release:**
  1. Work happens on a branch, with a pull request and one required review.
  2. CI must pass: tests, type-check, lint, build.
  3. Merge to protected `main` (no deploy).
  4. Create a version tag (e.g. `v0.3.0`).
  5. The tag triggers a build into a "production" environment that needs manual approval.
  6. Deploy.

  Keep the last good build so rollback is "re-deploy the previous tag".

## 8. Tests

- **Covered:** 137 tests in 8 files: formula parsing (26), basics checks (17 + 28), reviewers (25), table and date parsing (19), prior-year notes (17), Mark OK (2), update decision (3). They run in CI (`pages.yml:28-29`).
- **Not covered:**
  - everything that touches Excel: `src/excel/workbook.ts`, the panel `src/taskpane/taskpane.ts`, and the live update reload (`src/update.ts:20-46`)
  - lint (not run in CI)
- **Messy real workpapers not covered:**
  - merged header cells and two-row headers
  - hidden or filtered rows
  - Excel Tables with structured references beyond the basic case
  - subtotals partway down a table
  - side-by-side tables
  - day-first text dates
  - "CR/DR" or trailing-minus amounts
  - very large sheets
  - cross-sheet references and pivot tables
  - comment threads
  - protected sheets

## 9. Security

- **npm audit:** 58 issues (8 critical, 29 high), all in developer tooling (vitest, webpack-dev-server, office-addin-* CLIs). Code shipped to users: 0. Every issue has an available fix.
- **XSS (scripts injected via workbook text):** low risk. The panel only inserts text with `textContent`, never HTML (`taskpane.ts:64-80`; no `innerHTML` in `src/`). The demo-video tool does use `innerHTML` and `Function(...)` (`demo/video/studio.html:115,168-176,197`); it is never shipped, but must not be reused in the product.
- **Manifest permission:** `ReadWriteDocument` (`manifest.xml:30`). The add-in writes only its own settings, so read-only plus settings may be possible; whether Office allows saving settings with `ReadDocument` is **not verifiable from code**.
  - The manifest ID was inherited from the original prototype (`manifest.xml:7`). A firm deployment needs its own.
- **Licences:** all packages are MIT, ISC, Apache-2.0, BSD or similar; `node-forge` is BSD-3 *or* GPL-2 (dev only, take BSD). Shipped runtime deps: `core-js`, `regenerator-runtime` (MIT).
  - The project is marked `UNLICENSED` (`package.json:6`). Parts of the build setup come from Microsoft's MIT Office add-in template (the original `package.json` named it, removed in `1410d62`), so MIT notice obligations still apply.
  - Code ownership of the original prototype by account `7h7f7btq9p-bit` is **not verifiable from code**.

## 10. Migration blockers (moving to a firm-owned setup)

- Personal URLs: `webpack.config.js:27`, `manifest.xml:15,17`, `site/index.html:90`, `README.md:47,79`, `docs/DECISIONS.md:43`, `demo/video/record.mjs:177`, plus the video's end card.
- Hosting is GitHub Pages under a personal account (`pages.yml`). A firm would use its own approved host and Microsoft 365 centralised deployment.
- Personal-account references: `docs/DECISIONS.md:4,6,7,30` (other accounts); commit authors in history.
- Secrets: none found in files or history (searched for common token and key patterns). Nothing in the code needs a secret.
- Firm-specific content: no firm name appears anywhere (0 matches in files and history). The audit terms used (SURL, PY, tickmark, AP-3 style references) are generic.

---

## Privacy and de-identification audit

### Current files

| File:line | What | Neutral replacement |
|---|---|---|
| `CLAUDE.md:3` | Personal health information and working preferences | Remove from repo; keep in personal assistant settings |
| `CLAUDE.md:6` | Firm-type and regulator references ("Big 4 / PCAOB") | "Professional audit users" |
| `webpack.config.js:27` | Personal GitHub handle in hosting URL | Build-time setting, e.g. `https://addins.example.com/baseline/` |
| `manifest.xml:15` | Personal repo URL (support link) | `https://example.com/support` |
| `manifest.xml:17` | Personal domain | Firm host domain |
| `site/index.html:90` | Personal repo link | Remove, or the firm's internal link |
| `README.md:47,79` | Personal site URL | Relative or placeholder link |
| `docs/DECISIONS.md:43` | Personal site URL | "the hosted site" |
| `docs/DECISIONS.md:4,6,7,30` | Other GitHub accounts (`7h7f7btq9p-bit`, `going-concern`) | "an earlier prototype" |
| `docs/DECISIONS.md:66` | "audit committees" (audience) | "professional audiences" |
| `demo/video/record.mjs:177` and the end card in `demo/Baseline-demo-video.mp4` | Personal site URL on screen | Placeholder URL, then re-record |
| `demo/make_demo.py:239` and inside `demo/Baseline-demo.xlsx` and the video | "Cedar Valley Foods Ltd." (fictional, but could match a real company; **not verifiable**) | "Example Client Ltd" |
| `demo/make_demo.py:126` | Real institution ("Bank of Canada") | "central bank closing rate" |
| `demo/make_demo.py:138-141,164-168,297-302,321-325` and `tests/reviewers.test.ts:86-89,121-124,152-156,169` | Microsoft sample company names (Northwind, Contoso, Fabrikam, Tailspin, Adventure Works, Wide World): fictional but trademarked | "Customer A" to "Customer F" |
| `demo/make_demo.py:251-259`, `tests/reviewers.test.ts:31-39,59-60`, `tests/quality.test.ts:33`, `tests/table.test.ts:29,61`, `src/review/layout.ts:119`, `src/review/reviewers/procedure.ts:30` | Invented vendor names (Acme Plumbing, Delta Freight, …) | "Vendor A" to "Vendor I" (low risk; optional) |
| `demo/make_demo.py:93,159,203,244,290,314`, `tests/notes.test.ts:89` | Sign-off initials "JS", "MK" | "AB", "CD" |
| Amounts and dates in demo/tests | Invented round figures; none look taken from a real file (**not verifiable** that they never were) | Keep, or regenerate |
| `legacy/excel-coach/*` | No personal data; "Audit Professionals" wording only (`taskpane.ts:4,30`) | Delete the folder (unused) |
| Metadata: `demo/Baseline-demo.xlsx` | Author "openpyxl"; no person or company | None needed |
| Metadata: `.mp4`, `docs/img/*.png`, `assets/*.png` | Only encoder tags; no EXIF, GPS or names | None needed |
| Secrets, tokens, machine paths | None found | n/a |

### Git history (not removed by editing the latest files)

| Item | Where in history |
|---|---|
| Personal Apple relay email `7h7f7btq9p@privaterelay.appleid.com` | Author of 4 original commits (`3fe7588`, `92764eb`, `1b48da0`, `460f066`) |
| Personal handle and email `hmacknak <158846598+hmacknak@users.noreply.github.com>` | Author of 18 commits |
| Location hint (timezone −06:00) | 5 commits by `7h7f7btq9p-bit` and on branch `update-demo-video` |
| Assistant session links (`claude.ai/code/session_…`) | 19 commit messages |
| Health and firm-type text (CLAUDE.md) | Commit `aea551a` onward |
| Personal URLs | From `9e253c7` onward |
| Demo names (Cedar Valley, Bank of Canada, Microsoft sample names) | From `17a29fe` / `98a0db5` onward |
| "Prairie Ridge Mining Ltd." and other fictional clients | Remote branch `origin/update-demo-video` (`f2bf6e2`) |
| Six old versions each of the demo video and workbook (on-screen names and URL) | Binary history of `demo/` |
| Original template placeholders (contoso.com) | `1b48da0` |
| Outside git, on GitHub itself | Pull request #1 and its comment, Actions run logs (show commit messages), a public Pages site, and possibly clones by anyone (repo is public; 1 star) |

### Recommendation: (b) start a fresh repo

Rewriting history scrubs old commits but leaves copies GitHub keeps (pull request refs, Actions logs, existing clones) and breaks every link to old commits. A fresh repo built from the cleaned files is simpler and certain, at the cost of losing the commit history (the design record is kept in `docs/DECISIONS.md`). Agree with your lean: **(b)**. Then make the old repo private and archive it, and delete its Pages site.

---

## Staged plan (one branch and one pull request per stage; tests must pass before merge; nothing to `main` directly)

| Stage | Goal | What changes | How it's tested | Roll back |
|---|---|---|---|---|
| **1. Privacy cleanup** | No personal, client-like or account data anywhere | **In the repo:** apply every replacement in the privacy audit; delete `CLAUDE.md` (move its preferences to your personal assistant settings) and `legacy/`; make the hosting URL a build setting; regenerate the demo workbook and video. **On GitHub:** create a fresh repo from the cleaned files with a neutral author; make the old repo private, archive it and take down its Pages site. | Search for every term in this audit (0 matches); full test suite; inspect file metadata; watch the new video | Keep the old repo private (not deleted) until the new one is confirmed |
| 2. Release controls | No code reaches users without approval | Protected `main`, required review, tag-triggered deploy with manual approval, lint in CI, actions pinned | A test tag deploys only after approval; a push to `main` doesn't deploy | Re-enable the old workflow file |
| 3. Hosting and trust | Firm-controlled code source; data can't leave | Configurable host; content-security policy; optional version pin; clear wording on what's stored in the workbook | Browser test confirms blocked requests; update test with a pinned version | Revert the pull request; previous tag redeploys |
| 4. Dependency hygiene | Clean audit | Upgrade dev tools; opt out of CLI telemetry | `npm audit` shows 0 high/critical; full tests and build | Restore previous lockfile |
| 5. Robustness | Fewer misses and false alarms | One shared table reader; scan all rows; multiple tables; day-first date setting; extra amount formats; prior-year wording "No similar issue found" | New tests for every messy case in section 8; a set of real-world (de-identified) sample workpapers | Revert the pull request |
| 6. Real-Excel validation | Proof outside the simulator | Manual test script in Excel desktop and web: merged cells, hidden rows, large sheet, performance timing | Signed-off checklist with timings | n/a (test only) |
