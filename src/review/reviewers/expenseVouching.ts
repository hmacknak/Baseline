// Expense vouching: a sample of recorded expenses is vouched back to invoices and other support to show
// each one really happened, was recorded at the right amount, in the right year and the right account,
// and was approved.
//
// Research basis (see docs/PROCEDURES.md):
// - Direction of testing: vouching from the ledger back to documents tests occurrence (overstatement); tracing
//   from documents into the ledger tests completeness (ISA 500 / ISA 315 assertions).
// - Assertions for expenses: occurrence, accuracy, cut-off and classification (ISA 315 (Revised 2019)).
// - Fraud: payments for goods or services not received, including fictitious vendors, are a recognised way
//   management can override controls (ISA 240 application material).
// - Classification: costs meeting IAS 16.7 are capitalised; day-to-day servicing and repairs are expensed (IAS 16.12).

import { amountOf, dateOf, flagOf, isBlank, textOf } from "../table";
import { Reviewer } from "../types";
import { TableSpec, columnsCheck, completenessCheck, detectScore, fmtAmount, fmtDate, rowInsight, sheetTexts, tableFor } from "./procedure";

const spec: TableSpec = {
  minColumns: 3,
  columns: [
    { key: "ref", label: "sample #", kind: "text", names: ["sample #", "sample", "sample no", "item #", "je #", "entry #", "transaction id", "voucher #", "ref"] },
    { key: "date", label: "GL date", kind: "date", names: ["gl date", "posting date", "entry date", "transaction date", "date per gl", "date"] },
    { key: "vendor", label: "vendor", kind: "text", names: ["vendor", "supplier", "payee", "vendor name", "supplier name"] },
    { key: "description", label: "description", kind: "text", names: ["description", "details", "narrative", "memo", "item description"] },
    { key: "account", label: "GL account", kind: "text", names: ["gl account", "expense account", "account", "account name", "gl code", "nominal"] },
    {
      key: "glAmount",
      label: "amount per GL",
      kind: "amount",
      names: ["amount per gl", "gl amount", "amount per ledger", "recorded amount", "book amount", "amount"],
    },
    {
      key: "supportAmount",
      label: "amount per support",
      kind: "amount",
      names: ["amount per invoice", "invoice amount", "per invoice", "amount per support", "support amount", "agreed amount", "amount per document"],
    },
    { key: "invoiceNo", label: "invoice #", kind: "text", names: ["invoice #", "invoice no", "invoice number", "inv #", "document #", "support ref", "support"] },
    { key: "serviceDate", label: "service date", kind: "date", names: ["service date", "service period", "period of service", "delivery date", "date received", "service to", "services to"] },
    { key: "invoiceDate", label: "invoice date", kind: "date", names: ["invoice date", "inv date", "document date"] },
    { key: "approved", label: "approved?", kind: "flag", names: ["approved", "approved?", "approval", "authorised", "authorized", "approved by"] },
    { key: "classOk", label: "correct account?", kind: "flag", names: ["correct account?", "correct account", "classification ok", "classification", "correct gl?", "account ok?"] },
    { key: "conclusion", label: "conclusion", kind: "text", names: ["conclusion", "result", "exception", "comment", "comments"] },
  ],
};

const LABEL = ["vendor", "ref", "description"];

/** Descriptions that suggest a long-lived asset rather than a running cost. */
const CAPITAL_WORDS = /\b(equipment|machine(ry)?|vehicle|truck|forklift|computer|laptop|server|furniture|fit[\s-]?out|renovat\w*|construct\w*|install\w*|upgrade|capital|building|roof|hvac|software licen[cs]e|tooling)\b/i;

function periodStart(yearEnd: number): number {
  return yearEnd - 364;
}

export const expenseVouching: Reviewer = {
  id: "expense-vouching",
  name: "Expense vouching",
  category: "Expenses",
  code: "EV",
  summary: "Vouches a sample of expenses to support: real, right amount, right year, right account, approved.",
  alwaysOn: false,
  detect: sheet => {
    const score = detectScore(sheet, {
      name: /expense|vouch|opex|operating cost/i,
      text: /vouch|expense testing|test of details|sample of expenses/i,
      spec,
      columnsForFull: 5,
    });
    // Vendor/amount columns appear on many sheets (a SURL has them too). Without a name or wording that says
    // "expenses", don't suggest this reviewer.
    const named = /expense|vouch|opex|operating cost/i.test(sheet.name) ||
      sheetTexts(sheet).some(t => /vouch|expense testing|test of details|sample of expenses/i.test(t));
    return named ? score : Math.min(score, 0.5);
  },
  checks: [
    completenessCheck({
      id: "expense-vouching.complete",
      spec,
      labelKeys: LABEL,
      required: ["vendor", "date", "glAmount", ["supportAmount", "invoiceNo"], "conclusion"],
      tableHint: "Add a header row with columns like Sample #, GL date, Vendor, Description, Amount per GL, Amount per invoice, Invoice #, Approved?, Conclusion.",
      assertion: "Occurrence; Accuracy",
      reference: "ISA 500 (sufficient appropriate evidence); ISA 530 (sampling)",
      why: "Each sample item has to show what was recorded, what support it was agreed to, and the result. A sample item with no support reference or no conclusion isn't tested, and an untested item can't be used to conclude on the population.",
    }),
    columnsCheck({ id: "expense-vouching.columns", spec, needed: ["glAmount", "supportAmount", "vendor"] }),
  ],
  insights: [
    rowInsight({
      id: "expense-vouching.amount",
      title: "Recorded amount doesn't agree to support",
      severity: "high",
      assertion: "Accuracy",
      reference: "ISA 315 (accuracy); ISA 530 (projecting misstatements)",
      why: "A difference between the ledger and the invoice is a misstatement found in a sample. Because the sample stands for the whole population, the error has to be projected to the population, not just corrected for this one item, and the cause understood (keying error, wrong invoice, or something deliberate).",
      aliases: ["vouch-amount"],
      spec,
      labelKeys: LABEL,
      test(row) {
        const gl = amountOf(row, "glAmount");
        const doc = amountOf(row, "supportAmount");
        if (gl === null || doc === null || Math.abs(gl - doc) < 0.5) return null;
        return { cellKey: "supportAmount", amount: gl - doc, detail: `Recorded ${fmtAmount(gl)}, support shows ${fmtAmount(doc)}` };
      },
    }),
    rowInsight({
      id: "expense-vouching.cutoff",
      title: "Service falls outside the year",
      severity: "high",
      assertion: "Cut-off",
      reference: "ISA 315 (cut-off); IFRS Conceptual Framework (accrual basis)",
      why: "Expenses belong to the year the goods or services were received, not the year they were paid or invoiced. A service after year-end recorded as this year's expense overstates expenses; one from last year means last year was understated. Both are cut-off errors and often come in clusters around year-end.",
      needs: ["yearEnd"],
      aliases: ["cutoff"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        // The service date decides the year; the invoice date is the next-best evidence when there is none.
        const svc = dateOf(row, "serviceDate");
        const d = svc ?? dateOf(row, "invoiceDate");
        if (d === null) return null;
        const key = svc !== null ? "serviceDate" : "invoiceDate";
        const what = svc !== null ? "Service" : "Invoiced";
        if (d > ctx.yearEnd!) return { cellKey: key, amount: amountOf(row, "glAmount"), detail: `${what} ${fmtDate(d)}, after year-end` };
        if (d < periodStart(ctx.yearEnd!)) {
          return { cellKey: key, amount: amountOf(row, "glAmount"), detail: `${what} ${fmtDate(d)}, before the year started` };
        }
        return null;
      },
    }),
    {
      id: "expense-vouching.duplicates",
      title: "Possible duplicate payments",
      severity: "high",
      assertion: "Occurrence",
      reference: "ISA 240 (payments for goods or services not received)",
      why: "The same vendor, amount and invoice recorded twice is either an error (the invoice was paid twice) or a way money leaves the company for nothing in return. Paying for goods or services not received is one of the recognised ways controls are overridden, so duplicates need to be explained, not just noted.",
      aliases: ["duplicate"],
      run(sheet) {
        const table = tableFor(sheet, spec);
        if (!table) return [];
        const seen = new Map<string, string>();
        const out: Array<{ cell: string; label: string; amount?: number; detail: string }> = [];
        for (const row of table.rows) {
          const vendor = textOf(row, "vendor").toLowerCase();
          const amount = amountOf(row, "supportAmount") ?? amountOf(row, "glAmount");
          if (!vendor || amount === null) continue;
          const inv = textOf(row, "invoiceNo").toLowerCase();
          const key = `${vendor}|${amount.toFixed(2)}|${inv}`;
          const first = seen.get(key);
          if (first) {
            out.push({
              cell: (row.cells.vendor || row.cells.glAmount)!.address,
              label: textOf(row, "vendor"),
              amount,
              detail: `Same vendor, amount${inv ? " and invoice" : ""} as ${first}`,
            });
          } else {
            seen.set(key, `row ${row.row + 1}`);
          }
        }
        return out;
      },
    },
    rowInsight({
      id: "expense-vouching.approval",
      title: "No evidence of approval",
      severity: "medium",
      assertion: "Occurrence",
      reference: "ISA 315 (controls over expenditure)",
      why: "Approval is the control that stops payments nobody asked for. A missing approval doesn't mean the expense is false, but it removes the evidence that someone responsible confirmed it was a genuine business cost, so the support needs a closer look.",
      spec,
      labelKeys: LABEL,
      test(row) {
        // Only when the workpaper has an approval column; blank cells count as "not recorded".
        if (!("approved" in row.cells)) return null;
        if (flagOf(row, "approved") === false) {
          return { cellKey: "approved", amount: amountOf(row, "glAmount"), detail: "Marked as not approved" };
        }
        if (isBlank(row.cells.approved)) return { cellKey: "approved", amount: amountOf(row, "glAmount"), detail: "Approval not recorded" };
        return null;
      },
    }),
    rowInsight({
      id: "expense-vouching.capital",
      title: "Looks like a capital item",
      severity: "medium",
      assertion: "Classification",
      reference: "IAS 16.7; IAS 16.12",
      why: "A cost that gives benefits over more than one year (equipment, vehicles, major upgrades) should be capitalised as property, plant and equipment if it meets IAS 16.7; day-to-day repairs and servicing are expensed (IAS 16.12). Expensing a capital item understates assets and this year's profit; capitalising repairs does the reverse.",
      needs: ["threshold"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const text = `${textOf(row, "description")} ${textOf(row, "account")}`;
        const amount = amountOf(row, "glAmount");
        if (!CAPITAL_WORDS.test(text) || amount === null || Math.abs(amount) < ctx.threshold!) return null;
        return { cellKey: "description", amount, detail: `"${textOf(row, "description") || textOf(row, "account")}"` };
      },
    }),
    rowInsight({
      id: "expense-vouching.classification",
      title: "Recorded in the wrong account",
      severity: "medium",
      assertion: "Classification",
      reference: "ISA 315 (classification)",
      why: "An expense in the wrong account can move a cost between categories the reader cares about, such as cost of sales versus operating expenses, or between expense and asset. Each one is a classification misstatement to record and evaluate.",
      spec,
      labelKeys: LABEL,
      test(row) {
        if (flagOf(row, "classOk") !== false) return null;
        return { cellKey: "classOk", amount: amountOf(row, "glAmount"), detail: "Marked as wrong account" };
      },
    }),
  ],
};
