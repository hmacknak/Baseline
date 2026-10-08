# Baseline: Vision

## The problem
Juniors aren't getting work done because their skills aren't being built. Their skills aren't being built
because feedback comes too late (at review, days later), too vaguely, or not at all. Seniors repeat the
same review notes over and over, and nothing turns those notes into lasting skill.

## The idea
**Immediate feedback inside real work.** Baseline notices issues as juniors work, explains them like a patient
senior, and builds the skill so the same note doesn't come back. It uses technology to find efficiencies for
both the junior (fewer redo cycles) and the senior (less repeat reviewing).

The loop:
1. **Notice:** spot an issue in real work.
2. **Nudge:** say it right away, in one line, where the work is happening.
3. **Build:** track repeat issues and focus on one habit at a time.
4. **Show progress:** fewer repeat notes, faster turnaround.

## First version: Review Bot (Excel add-in)
1. The junior clicks **Review this sheet**.
2. The bot leaves cell-linked **review notes**, like a senior would. Clicking a note jumps to the cell.
3. **Talk:** the junior can ask "why does this matter?" or "how do I fix it?" on any note.
4. Repeat notes become this week's skill focus (this reuses the existing Excel Coach code).

How it reviews:
- **Rule checks** (always right, no AI): hard-coded numbers in formulas, `#REF!` and other errors, totals that don't
  add up, formulas that don't match their neighbours.
- **AI review** (Claude): judgement calls such as missing sign-offs, unclear explanations, and missing tickmark legends.
- A small server holds the API key; the add-in never contains it.

## Status
- **v0.1 (built):** live 9-item checklist in Excel plus prior-year notes ("this happened last year"). Local only,
  no AI.

- **v0.2 (built):** reviewer catalog. Workpaper basics always run; procedure reviewers (bank rec, SURL, revenue
  cut-off, AR confirmations) are suggested per sheet and add completeness checks plus insights, which surface
  high-risk items from the testing data. Year-end and threshold are set once per workbook.

## Reviewer catalog (direction)
Think "Netflix of procedures": rows by area (Cash, Liabilities, Revenue, Receivables, Fixed assets…), each card a
reviewer for one test. The aim is twofold: **workpapers come back complete** ("you missed a date here"), and
**the testing itself gets sharper**, because the reviewer runs the analysis a senior would (dates around
year-end, items over threshold, unexplained differences) and points the junior straight at the risky items.
Ten more reviewers are listed as "Coming soon"; adding one is documented in `docs/ADDING_A_REVIEWER.md`.

## Saved ideas (later)
- **Common-issue analysis:** analyse a whole set of review notes (across sheets, juniors, engagements) to show
  which issues come up most, then focus the checklist and coaching on those. The note tagging in
  `src/review/notes.ts` is the starting point.
- **Talk:** ask "why?" or "how do I fix it?" on any note and get a senior-style explanation from Claude. This needs a
  decision on client data and a small server to hold the API key.
- **Skill building:** turn repeat issues into this week's habit (reuse the legacy Excel Coach).

## Open questions
- Client data: practice/dummy workpapers only, or real ones with the firm's approval?
- The first set of real (anonymised) senior review notes, which become the first checks and set the bot's voice.
- Who pays: the firm or the junior?
- Beyond Excel later: Word, Outlook, Teams, audit software?
