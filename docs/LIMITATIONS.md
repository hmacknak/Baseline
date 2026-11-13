# What each flag can and can't see

Baseline reads cell values, formulas, number formats and fill colours. It does not understand the client, the
engagement or the source documents. This page says, for every flag, what it catches, what it deliberately leaves
alone, and what it cannot see. The behaviour described here is covered by tests in `tests/robustness.test.ts`,
`tests/quality.test.ts` and `tests/procedures.test.ts`; when a limit changes, change the test and this page together.

## The rule every flag follows

A flag ends in one of three states, and the panel shows which:

| State | Shown as | Meaning |
|---|---|---|
| Flagged | cross or exclamation mark | Something on the sheet matches the rule. |
| Clear | tick | The rule ran on this sheet and found nothing. |
| Not run / not applied | dash, with the reason | The rule could not be evaluated, or does not apply. Never shown as a tick. |

"No open issues" appears only when nothing is flagged **and** nothing was left not run. If nothing is flagged but
something could not run, the panel says "No issues found · N checks not run".

Reasons a flag is not run:

- **Year-end or threshold not set.** Fixed in Settings.
- **The column it reads isn't in the testing table.** For example, no "received via" column on a bank
  confirmation sheet. The panel names the column.
- **No testing table found.** Said once for the whole procedure.
- **Fill colours couldn't be read.** Very large sheets (over 20,000 cells) or an older Excel.
- **Not a workpaper.** Sign-off and purpose/source/conclusion are not applied to supporting data tabs. Once a
  procedure reviewer is switched on for a sheet, it is always treated as a workpaper.

A value the procedure rules can't read (a date typed as "December", "Partly" in a yes/no column, text in an amount
column) is reported by "Baseline can read every answer" instead of being skipped silently.

## Workpaper basics

| Flag | Catches | Leaves alone on purpose | Cannot see |
|---|---|---|---|
| Formula errors | Every error cell. A run of the same error is one flag with its range. | | Errors hidden by IFERROR. |
| Hard-coded numbers in formulas | Numbers added, subtracted or used as a rate or price. The same formula copied down or across is one flag. | 0 and 1; unit conversions (/12, *100, /1000, /365, ^2, /2); function arguments such as the 2 in ROUND(x,2). | Thresholds inside IF tests (IF(A1>5000,...)); numbers hidden in defined names. |
| Typed numbers among formulas | A constant between matching formulas, or in a row or column that is mostly one formula. | The first cell of a run (an opening balance). | A single override where neighbours don't match each other. |
| Totals include every line | SUM or SUBTOTAL(9/109) that stops short of numbers below or beside it, or starts after the first items of its table. | Year headings such as 2025 above the items. | Totals written as A1+A2+A3, SUM with several ranges, totals on another sheet. |
| Numbers stored as text | Text that parses as a number beside real numbers. | Columns of IDs. | |
| Formulas consistent | A formula that differs from identical formulas on both sides. | Subtotals and other aggregates between item formulas. | A wrong formula at the end of a run (one neighbour only). |
| Tables are complete | Blanks in a mostly filled column of a table (a heading row of three or more headings with data under it). | Comment, note and tickmark columns; a conclusion or note under the table; rows of "Label:" form fields. | Tables with fewer than three headings; merged cells (Excel reports them as blank). |
| Placeholders | TBD, TBC, XXX, ???, "to be completed", "[insert ...]", "pending". | Masked account numbers such as XXXX-1234. | Placeholders in other languages. |
| Dates | A prior-year year-end left in the heading; a date far from the others in a tightly dated column; dates typed as text; mixed date styles; review dated before preparation. | Long-dated columns (asset, loan and lease dates); comparative columns; zero shown as a date; padding differences in formats. | Whether a plausible date is the right date. |
| Stray numbers | Unlabelled numbers outside the working table. | Sheets with no table; a block of more than 8 numbers (treated as data). | |
| Leftover highlighting | Bright fills on up to 20 cells of one colour. | Heading rows; banded rows; colours explained in a legend; a colour used on more than 20 cells. | Conditional formatting; font colours. Not run above 20,000 cells. |
| Links to other files | Formulas pointing at another workbook. Repeats are grouped. | | Links inside defined names or charts. |
| Sign-off | Missing or undated preparer; missing reviewer line; undated reviewer. | Supporting data tabs (not applied). | Whether the initials are the right person. |
| Purpose, source, conclusion | None of the three words found. | Supporting data tabs (not applied). | Whether what is written is adequate. |
| Tickmarks | Symbols, or short marks in a column headed "Tick", that the legend doesn't explain. | Y/N answer columns. | Tickmarks drawn with a symbol font such as Wingdings outside a "Tick" column. |

## Procedure reviews

All three read one testing table per sheet: the row within the first 40 rows that matches the most known column
headings, and the rows under it until a total, a conclusion or note, or two blank rows.

| Limit | What Baseline does about it |
|---|---|
| A heading it doesn't recognise | The rules that need that column show "Not run" and name the column. Add the heading or rename it. |
| A second table on the same sheet | Not read. Put each procedure on its own sheet. |
| Yes/no answers with explanations | The first word decides ("No, accrued in January" is no). "N/A" is accepted as an answer. A bare "x" is reported as unclear. |
| Dates typed as text, such as 05/01/2026 | Read day-first or month-first according to the rest of the sheet. If the sheet gives no evidence, the cell is reported and no guess is made. |
| Status words in a date column ("Not yet", "Outstanding") | Treated as no date. A bank reply marked that way is still flagged as no reply. |
| Amounts in thousands | Not detected. Set the threshold in the same units as the sheet. |
| Financial years that aren't 12 months | Expense cut-off assumes a 12-month year ending on the year-end date. |

### Search for unrecorded liabilities

Catches: post-year-end invoices for services on or before year-end that aren't marked recorded; pre-year-end
invoices marked not recorded; items recorded at year-end for next year's services; items with no service date; a
search that ends within 30 days of year-end.

Cannot see: whether the selection covers the population; invoices that were never selected; whether "recorded"
is true in the ledger.

### Cash confirmations

Catches: book and bank balances that differ with no explanation; requests with no reply; replies that came through
the client; accounts with no request, including nil balances; confirmations at a date other than year-end;
overdrafts; balances in another currency.

Cannot see: accounts missing from the list; whether an explanation is valid; the content of the reply (loans,
guarantees, restrictions). "Through the client" is judged from words in the "received via" column.

### Expense vouching

Catches: ledger and support amounts that differ; service or invoice dates outside the year; the same vendor,
amount and invoice twice; approval refused or blank; descriptions that read like capital items over the
threshold; items marked as in the wrong account.

Cannot see: whether the support is genuine; duplicates outside the sample; capital items described in words it
doesn't know.

## Whole-sheet limits

- Sheets are read up to 100,000 cells. Beyond that the panel says the sheet was cut off.
- Hidden sheets are not reviewed by "Review all sheets". Hidden rows and columns on a visible sheet are.
- Nothing is read from comments, notes, shapes, charts or pivot tables.

## How noise is measured

`tools/noise` runs every check over a folder of real workbooks and prints what fired. On 180 sheets from 130
public sample workbooks: no crashes; error flags fell from 1,242 to 154 and hard-code flags from 65 to 30 once
repeats were grouped; sign-off and purpose flags fell from 180 sheets to 3 once supporting data was told apart
from workpapers.
