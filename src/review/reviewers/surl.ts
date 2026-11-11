// Search for unrecorded liabilities (SURL): payments and invoices after year-end, tested for
// whether they relate to the year under audit and were recorded.

import { amountOf, cellAddress, dateOf, flagOf } from "../table";
import { Reviewer } from "../types";
import { TableSpec, columnsCheck, completenessCheck, detectScore, fmtDate, rowInsight, tableFor } from "./procedure";

const spec: TableSpec = {
  minColumns: 3,
  columns: [
    { key: "vendor", label: "vendor", kind: "text", names: ["vendor", "supplier", "payee", "vendor name", "supplier name"] },
    { key: "invoiceNo", label: "invoice #", kind: "text", names: ["invoice #", "invoice no", "invoice number", "inv #", "inv no"] },
    { key: "invoiceDate", label: "invoice date", kind: "date", names: ["invoice date", "inv date", "date of invoice", "invoice dt"] },
    {
      key: "serviceDate",
      label: "service date",
      kind: "date",
      names: ["service date", "service period", "period of service", "services to", "service to", "service end", "goods received", "date received", "receipt date", "delivery date", "service period end"],
    },
    { key: "paymentDate", label: "payment date", kind: "date", names: ["payment date", "date paid", "paid date", "paid", "cheque date", "check date", "disbursement date", "paid on"] },
    { key: "amount", label: "amount", kind: "amount", names: ["amount", "invoice amount", "payment amount", "amount paid", "total", "$"] },
    {
      key: "recorded",
      label: "recorded in AP?",
      kind: "flag",
      names: ["recorded", "recorded in ap", "in ap", "accrued", "recorded?", "in ap?", "recorded in ap?", "liability recorded", "per ap listing"],
    },
    { key: "conclusion", label: "conclusion", kind: "text", names: ["conclusion", "comment", "comments", "result", "exception", "notes"] },
  ],
};

const LABEL = ["vendor", "invoiceNo"];

/** A search window shorter than this after year-end rarely covers a full payment cycle. */
const MIN_WINDOW_DAYS = 30;

// Research basis (see docs/PROCEDURES.md):
// - A liability is a present obligation arising from past events (IFRS Conceptual Framework 4.26). Goods or services
//   received by year-end create that obligation whether or not they have been invoiced; un-invoiced amounts are
//   accruals (IAS 37.11).
// - The audit risk is understatement, so the assertion is completeness, tested by working from evidence outside
//   the payables ledger (later payments and invoices) back into the ledger.
// - The search covers the period after year-end up to as close to the report date as practicable (ISA 560).

export const surl: Reviewer = {
  id: "surl",
  name: "Search for unrecorded liabilities",
  category: "Liabilities",
  code: "UL",
  summary: "Finds post-year-end invoices and payments for this year's goods and services that never made it into payables.",
  alwaysOn: false,
  detect: sheet =>
    detectScore(sheet, {
      name: /unrecorded|\bsurl\b|subsequent (payment|disbursement)/i,
      text: /unrecorded liabilit|subsequent (payments|disbursements)/i,
      spec,
      columnsForFull: 5,
    }),
  checks: [
    completenessCheck({
      id: "surl.complete",
      spec,
      labelKeys: LABEL,
      required: ["vendor", "invoiceDate", "serviceDate", "paymentDate", "amount", "recorded", "conclusion"],
      tableHint: "Add a header row with columns like Vendor, Invoice date, Service date, Payment date, Amount, Recorded in AP?, Conclusion.",
      assertion: "Completeness",
      reference: "ISA 500; ISA 230; PCAOB AS 1105; AS 1215",
      why: "Each selection only proves something if you can show when the service happened, whether it was recorded, and what you concluded. Without the service date you can't say which year the liability belongs to; without a conclusion the reviewer can't tell whether the item was cleared or is an exception.",
    }),
    columnsCheck({ id: "surl.columns", spec, needed: ["invoiceDate", "serviceDate", "amount", "recorded"] }),
  ],
  insights: [
    rowInsight({
      id: "surl.post-ye-invoice",
      title: "Invoices after year-end for this year's services",
      severity: "high",
      assertion: "Completeness",
      reference: "IFRS Conceptual Framework 4.26; IAS 37.11; PCAOB AS 2301",
      why: "The service was received on or before year-end, so the company already owed the money at year-end: that is a liability under the Conceptual Framework, even though the invoice came later. If it isn't recorded or accrued (IAS 37.11), liabilities and expenses are understated. This is exactly the misstatement the search exists to find, and it goes on the summary of uncorrected misstatements if it isn't adjusted.",
      needs: ["yearEnd"],
      aliases: ["unrecorded"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const inv = dateOf(row, "invoiceDate");
        const svc = dateOf(row, "serviceDate");
        if (inv === null || svc === null || inv <= ctx.yearEnd! || svc > ctx.yearEnd!) return null;
        if (flagOf(row, "recorded") === true) return null;
        return {
          cellKey: "serviceDate",
          amount: amountOf(row, "amount"),
          detail: `Invoiced ${fmtDate(inv)} for services on ${fmtDate(svc)}, not marked as recorded`,
        };
      },
    }),
    rowInsight({
      id: "surl.pre-ye-unrecorded",
      title: "Pre-year-end invoices not in payables",
      severity: "high",
      assertion: "Completeness",
      reference: "IAS 37.11 (trade payables); PCAOB AS 2301",
      why: "An invoice dated on or before year-end means the supplier had already billed the company. Unless it was paid before year-end, it should be in the year-end payables listing. A missing one points to a cut-off or recording failure, and there are often more like it.",
      needs: ["yearEnd"],
      aliases: ["unrecorded"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const inv = dateOf(row, "invoiceDate");
        if (inv === null || inv > ctx.yearEnd!) return null;
        if (flagOf(row, "recorded") !== false) return null;
        return {
          cellKey: "recorded",
          amount: amountOf(row, "amount"),
          detail: `Invoice dated ${fmtDate(inv)} marked as not recorded`,
        };
      },
    }),
    rowInsight({
      id: "surl.recorded-next-year",
      title: "Recorded at year-end for next year's services",
      severity: "medium",
      assertion: "Existence; Cut-off",
      reference: "IFRS Conceptual Framework 4.26",
      why: "If the service happened after year-end, there was no obligation at year-end, so recording it as a year-end liability overstates both liabilities and this year's expenses. The search is aimed at understatement, but it is also where an over-accrual (sometimes used to smooth results) becomes visible.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const svc = dateOf(row, "serviceDate");
        if (svc === null || svc <= ctx.yearEnd! || flagOf(row, "recorded") !== true) return null;
        return {
          cellKey: "recorded",
          amount: amountOf(row, "amount"),
          detail: `Service on ${fmtDate(svc)} (after year-end) but marked as recorded at year-end`,
        };
      },
    }),
    rowInsight({
      id: "surl.no-service-date",
      title: "Can't tell which year it belongs to",
      severity: "medium",
      assertion: "Completeness; Cut-off",
      reference: "ISA 500; PCAOB AS 1105",
      why: "The invoice and payment dates don't decide which year a liability belongs to; the date the goods or service were received does. Without it, a post-year-end invoice can't be cleared, so the selection gives no evidence either way.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        if (dateOf(row, "serviceDate") !== null) return null;
        const inv = dateOf(row, "invoiceDate");
        if (inv !== null && inv <= ctx.yearEnd!) return null;
        return {
          cellKey: "serviceDate",
          amount: amountOf(row, "amount"),
          detail: inv !== null ? `Invoiced ${fmtDate(inv)}, no service date` : "No invoice or service date",
        };
      },
    }),
    {
      id: "surl.window",
      title: "Search window looks short",
      severity: "medium",
      assertion: "Completeness",
      reference: "ISA 560; PCAOB AS 2801",
      why: "Suppliers often invoice weeks after delivery, so a search that stops soon after year-end can miss liabilities that are paid later. The window normally runs as close to the audit report date as practicable, and at least one full payment cycle.",
      needs: ["yearEnd"],
      run(sheet, ctx) {
        const table = tableFor(sheet, spec);
        if (!table || table.rows.length < 3) return [];
        let latest: number | null = null;
        let cell = "";
        for (const row of table.rows) {
          const d = dateOf(row, "paymentDate") ?? dateOf(row, "invoiceDate");
          if (d !== null && (latest === null || d > latest)) {
            latest = d;
            cell = cellAddress(table, row, row.cells.paymentDate ? "paymentDate" : "invoiceDate");
          }
        }
        if (latest === null || latest - ctx.yearEnd! >= MIN_WINDOW_DAYS) return [];
        return [
          {
            cell,
            label: "Search window",
            detail: `Latest item tested is ${fmtDate(latest)}, only ${latest - ctx.yearEnd!} days after year-end`,
          },
        ];
      },
    },
  ],
};
