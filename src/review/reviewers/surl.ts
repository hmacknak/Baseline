// Search for unrecorded liabilities (SURL): payments and invoices after year-end, tested for
// whether they relate to the year under audit and were recorded.

import { amountOf, dateOf, flagOf } from "../table";
import { Reviewer } from "../types";
import { TableSpec, columnsCheck, completenessCheck, detectScore, fmtDate, rowInsight } from "./procedure";

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
      names: ["service date", "service period", "period of service", "services to", "goods received", "date received", "receipt date", "delivery date", "service period end"],
    },
    { key: "paymentDate", label: "payment date", kind: "date", names: ["payment date", "date paid", "paid date", "cheque date", "check date", "disbursement date", "paid on"] },
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

export const surl: Reviewer = {
  id: "surl",
  name: "Search for unrecorded liabilities",
  category: "Liabilities",
  code: "UL",
  summary: "Finds post-year-end invoices for this year's services that never made it into AP.",
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
    }),
    columnsCheck({ id: "surl.columns", spec, needed: ["invoiceDate", "serviceDate", "amount", "recorded"] }),
  ],
  insights: [
    rowInsight({
      id: "surl.post-ye-invoice",
      aliases: ["unrecorded"],
      title: "Invoices after year-end for this year's services",
      severity: "high",
      why: "The service happened on or before year-end, so the liability belongs in this year. If it isn't recorded, liabilities and expenses are understated.",
      needs: ["yearEnd"],
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
      aliases: ["unrecorded"],
      title: "Pre-year-end invoices not in AP",
      severity: "high",
      why: "An invoice dated on or before year-end that isn't in the AP listing is an unrecorded liability, unless it was paid before year-end.",
      needs: ["yearEnd"],
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
      id: "surl.no-service-date",
      title: "Can't tell which year it belongs to",
      severity: "medium",
      why: "Without a service or receipt date there's no way to show whether a post-year-end invoice is this year's liability or next year's.",
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
  ],
};
