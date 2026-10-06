# Decisions & Open Questions

## Provenance (2026-10-06)
- This repository continues `7h7f7btq9p-bit/Baseline` (public), copied with full git history.
- Related copies, not yet reviewed (private, no access at time of copy):
  - `7h7f7btq9p-bit/baseline-excel` — appears to have the `src/`, `assets/`, `dist/` layout and a GitHub Pages workflow.
  - `going-concern/Baseline--Prototype-for-testing` — a fork of Baseline, 3 commits ahead, with a Vercel (Node 20) setup.
- Branch `copilot/generate-readme-file` is an old branch already merged into `main` (PR #1); kept for reference only.

## Current state
- `taskpane.ts` is the active app ("Excel Coach"): 11 shortcut skills, automatic detection from
  Excel change and selection events, per-skill progress bars, progress kept in `localStorage`.
- `taskpane.js` is an older, different app: 10 tips rotating weekly (the "one upgrade per week" idea
  described in the README).
- `temp_replace.html` is empty.

## Known problems
1. **Does not build.** `webpack.config.js` expects `src/taskpane/taskpane.ts`, `src/taskpane/taskpane.html`
   and `assets/*`; none exist in this repo.
2. **Detection is heuristic.** A paste of 3 or more plain values counts as Flash Fill, any formula with `$` counts
   as F4, any `SUM(` counts as Alt+=, and so on. The Office.js API cannot see keystrokes, so the add-in
   can only infer shortcut use from cell changes.
3. **Dev-only manifest.** URLs point at `https://localhost:3000`, and the support URL is the Contoso placeholder.
4. Event listeners attach only to the worksheet active at load time.
5. README promise ("one weekly upgrade") and the current UI (all 11 skills at once) disagree.

## Open questions (to discuss before feature work)
- Product direction: weekly single-habit coaching vs. a full shortcut tracker vs. something else?
- Audience: audit and accounting staff (the current copy) or general Excel users?
- Distribution: GitHub Pages / Vercel hosting plus a sideloaded manifest, or AppSource?
- Should the `baseline-excel` or `going-concern` copies be merged in?
