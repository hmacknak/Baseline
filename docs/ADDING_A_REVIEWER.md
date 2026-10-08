# Adding a procedure reviewer

A procedure reviewer knows one audit test. It has three parts:

1. **Columns.** What the testing table's headings can be called.
2. **Checklist.** Usually "every tested item is complete" and "the table has the key columns".
3. **Insights.** Rules that flag high-risk rows, such as "invoice dated after year-end for a December service".

Most new reviewers are about 100 lines, mostly column names and plain-English text.

## Steps

1. **Copy an existing reviewer** in `src/review/reviewers/`. `surl.ts` is a good template.
2. **List the columns.** Give every heading variant you've seen in real workpapers; matching ignores case and
   punctuation, and "Inv date" matches a column headed "Inv date (per invoice)".
   ```ts
   { key: "invoiceDate", label: "invoice date", kind: "date", names: ["invoice date", "inv date", "date of invoice"] }
   ```
   `kind` is `text`, `date`, `amount` or `flag` (Y/N, Yes/No, ✓, Recorded…).
3. **Detection.** In `detect`, give `detectScore` a regex for the sheet name and a regex for words on the
   sheet. A score of 0.6 or more makes Baseline suggest the reviewer.
4. **Completeness.** `completenessCheck({ required: [...] })` lists the columns every item must have. Use
   `["clearedDate", "comment"]` when either one is enough.
5. **Insights.** Each `rowInsight` gets a `test(row, ctx)` that returns `null` (fine) or
   `{ detail, cellKey, amount }` (flag it). Use `dateOf`, `amountOf`, `flagOf`, `textOf` to read cells;
   `ctx.yearEnd` and `ctx.threshold` come from Settings. Declare `needs: ["yearEnd"]` if the rule needs them.
   Write `why` for a junior: what could be wrong and why a reviewer cares.
6. **Last year's notes.** Give an insight `aliases: ["your-tag"]`, then add a matching rule to `TAG_RULES` in
   `src/review/notes.ts`, so a PY note like "stale cheque not followed up" shows as *Again this year*.
7. **Register it** in `CATALOG` in `src/review/reviewers/catalog.ts`. Replace the "Coming soon" entry if
   there is one, and add any new category to `CATEGORY_ORDER`.
8. **Test it** in `tests/reviewers.test.ts`: build a small sheet with `sheet({...})`, plant one example of each
   risk and one clean row, and assert exactly the planted rows are flagged.
9. **Demo it.** Add a sheet to `demo/make_demo.py` with planted issues, then run `python3 demo/make_demo.py`.

Run `npm test` and `npm run build`. Pushing to `main` republishes the add-in.

## Good insight rules

- **Flag what a senior would ask about**, not everything unusual. Each false alarm teaches juniors to ignore the panel.
- **Year-end is part of the year.** Use `<= ctx.yearEnd` for "this year" and `> ctx.yearEnd` for "next year".
- **Blank isn't false.** `flagOf` returns `null` for blank. Decide deliberately whether a blank answer should
  be flagged, and usually leave it to the completeness check.
- **Point at the cell to fix** (`cellKey`). A missing service date points at the service date column, not the vendor.
