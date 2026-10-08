// Bank reconciliation: the rec agrees to the GL, and outstanding items are real, current and clearing.

import { toAmount } from "../dates";
import { amountOf, dateOf } from "../table";
import { Cell, Reviewer, SheetSnapshot } from "../types";
import { TableSpec, columnsCheck, completenessCheck, detectScore, fmtAmount, fmtDate, rowInsight } from "./procedure";

const spec: TableSpec = {
  minColumns: 3,
  columns: [
    {
      key: "number",
      label: "cheque #",
      kind: "text",
      names: ["cheque #", "cheque no", "cheque number", "check #", "check no", "check number", "chq #", "chq no", "reference", "ref"],
    },
    { key: "payee", label: "payee", kind: "text", names: ["payee", "vendor", "description", "item", "name"] },
    { key: "issueDate", label: "date", kind: "date", names: ["date", "cheque date", "check date", "issue date", "date issued", "date written", "deposit date"] },
    { key: "amount", label: "amount", kind: "amount", names: ["amount", "$"] },
    {
      key: "clearedDate",
      label: "cleared date",
      kind: "date",
      names: ["cleared", "cleared date", "date cleared", "clearing date", "cleared bank", "cleared on", "date cleared bank", "cleared per bank"],
    },
    { key: "comment", label: "comment", kind: "text", names: ["comment", "comments", "notes", "explanation", "status"] },
  ],
};

const LABEL = ["number", "payee"];
const STALE_DAYS = 182;

/** Number to the right of a label cell such as "Balance per general ledger". */
function valueRightOf(sheet: SheetSnapshot, pattern: RegExp): { cell: Cell; value: number } | null {
  for (const row of sheet.cells) {
    for (let i = 0; i < row.length; i++) {
      const c = row[i];
      if (!c || c.kind !== "text" || !pattern.test(String(c.value))) continue;
      for (let j = i + 1; j < Math.min(row.length, i + 4); j++) {
        const v = toAmount(row[j]?.value);
        if (row[j] && row[j].kind === "number" && v !== null) return { cell: row[j], value: v };
      }
    }
  }
  return null;
}

const ADJUSTED_BANK = /adjusted (bank )?balance|balance per bank,? adjusted|reconciled (bank )?balance/i;
const PER_BOOKS = /balance per (books|gl|g\/l|general ledger|ledger)|per general ledger|gl balance/i;

export const bankRec: Reviewer = {
  id: "bank-rec",
  name: "Bank reconciliation",
  category: "Cash",
  icon: "🏦",
  summary: "Ties the rec to the GL and flags stale, misdated or uncleared outstanding items.",
  alwaysOn: false,
  detect: sheet =>
    detectScore(sheet, {
      name: /bank|cash|outstanding/i,
      text: /outstanding (cheque|check)|deposits? in transit|balance per bank|bank reconciliation/i,
      spec,
      columnsForFull: 4,
    }),
  checks: [
    completenessCheck({
      id: "bank-rec.complete",
      spec,
      labelKeys: LABEL,
      required: ["payee", "issueDate", "amount", ["clearedDate", "comment"]],
      tableHint: "List outstanding items with columns like Cheque #, Payee, Date, Amount, Cleared date.",
    }),
    columnsCheck({ id: "bank-rec.columns", spec, needed: ["issueDate", "amount", "clearedDate"] }),
  ],
  insights: [
    {
      id: "bank-rec.difference",
      aliases: ["rec-difference"],
      title: "Rec doesn't agree to the GL",
      severity: "high",
      why: "The adjusted bank balance should equal the balance per books. Any difference is either an error in the rec or an unrecorded transaction.",
      run(sheet) {
        const bank = valueRightOf(sheet, ADJUSTED_BANK);
        const books = valueRightOf(sheet, PER_BOOKS);
        if (!bank || !books) return [];
        const diff = bank.value - books.value;
        if (Math.abs(diff) < 0.5) return [];
        return [
          {
            cell: bank.cell.address,
            label: "Unreconciled difference",
            amount: diff,
            detail: `Adjusted bank ${fmtAmount(bank.value)} vs books ${fmtAmount(books.value)} (${books.cell.address})`,
          },
        ];
      },
    },
    rowInsight({
      id: "bank-rec.stale",
      aliases: ["stale"],
      title: "Stale-dated cheques",
      severity: "medium",
      why: "A cheque outstanding for more than six months probably won't be cashed. It may need to be written back to cash and liabilities, or reported as unclaimed property.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const issued = dateOf(row, "issueDate");
        const cleared = dateOf(row, "clearedDate");
        if (issued === null || issued > ctx.yearEnd! - STALE_DAYS) return null;
        if (cleared !== null && cleared <= ctx.yearEnd!) return null;
        return { cellKey: "issueDate", amount: amountOf(row, "amount"), detail: `Written ${fmtDate(issued)}, over 6 months before year-end` };
      },
    }),
    rowInsight({
      id: "bank-rec.cleared-before-ye",
      title: "Cleared before year-end but listed as outstanding",
      severity: "high",
      why: "If the bank cleared it before year-end, it wasn't outstanding at year-end. The rec is wrong or the dates are.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const cleared = dateOf(row, "clearedDate");
        if (cleared === null || cleared > ctx.yearEnd!) return null;
        return { cellKey: "clearedDate", amount: amountOf(row, "amount"), detail: `Cleared ${fmtDate(cleared)}` };
      },
    }),
    rowInsight({
      id: "bank-rec.dated-after-ye",
      title: "Dated after year-end",
      severity: "high",
      why: "A cheque written after year-end can't be outstanding at year-end. Including it understates cash.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const issued = dateOf(row, "issueDate");
        if (issued === null || issued <= ctx.yearEnd!) return null;
        return { cellKey: "issueDate", amount: amountOf(row, "amount"), detail: `Written ${fmtDate(issued)}` };
      },
    }),
    rowInsight({
      id: "bank-rec.large-uncleared",
      title: "Large items not yet cleared",
      severity: "medium",
      why: "Big items that haven't cleared by fieldwork need support, such as the cheque copy or a later bank statement, before you can rely on them.",
      needs: ["threshold"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const amount = amountOf(row, "amount");
        if (amount === null || Math.abs(amount) < ctx.threshold! || dateOf(row, "clearedDate") !== null) return null;
        return { cellKey: "clearedDate", amount, detail: "No cleared date" };
      },
    }),
  ],
};

