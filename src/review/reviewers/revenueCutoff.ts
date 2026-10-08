// Revenue cut-off: sales either side of year-end are recorded in the period the goods or services were delivered.

import { amountOf, dateOf, isBlank } from "../table";
import { Reviewer } from "../types";
import { TableSpec, columnsCheck, completenessCheck, detectScore, fmtDate, rowInsight } from "./procedure";

const spec: TableSpec = {
  minColumns: 3,
  columns: [
    { key: "invoiceNo", label: "invoice #", kind: "text", names: ["invoice #", "invoice no", "invoice number", "inv #", "inv no", "document #", "doc #"] },
    { key: "customer", label: "customer", kind: "text", names: ["customer", "client", "customer name", "sold to"] },
    {
      key: "invoiceDate",
      label: "invoice date",
      kind: "date",
      names: ["invoice date", "inv date", "billing date", "date invoiced", "revenue date", "date recorded", "posting date", "gl date"],
    },
    {
      key: "shipDate",
      label: "ship date",
      kind: "date",
      names: ["ship date", "shipping date", "shipped", "date shipped", "delivery date", "delivered", "date delivered", "bol date", "bill of lading date", "service date"],
    },
    { key: "amount", label: "amount", kind: "amount", names: ["amount", "invoice amount", "net amount", "revenue", "sales", "total", "$"] },
    {
      key: "conclusion",
      label: "conclusion",
      kind: "text",
      names: ["conclusion", "comment", "comments", "result", "proper period?", "correct period?", "cut-off ok?", "cut off ok?"],
    },
  ],
};

const LABEL = ["invoiceNo", "customer"];
const NEAR_YE_DAYS = 5;

export const revenueCutoff: Reviewer = {
  id: "revenue-cutoff",
  name: "Revenue cut-off",
  category: "Revenue",
  code: "CO",
  summary: "Catches sales recorded in the wrong year by comparing invoice and shipping dates around year-end.",
  alwaysOn: false,
  detect: sheet =>
    detectScore(sheet, {
      name: /cut-?\s?off|revenue|sales/i,
      text: /cut-?\s?off/i,
      spec,
      columnsForFull: 4,
    }),
  checks: [
    completenessCheck({
      id: "revenue-cutoff.complete",
      spec,
      labelKeys: LABEL,
      required: ["invoiceNo", "invoiceDate", "shipDate", "amount", "conclusion"],
      tableHint: "Add a header row with columns like Invoice #, Customer, Invoice date, Ship date, Amount, Conclusion.",
    }),
    columnsCheck({ id: "revenue-cutoff.columns", spec, needed: ["invoiceDate", "shipDate", "amount"] }),
  ],
  insights: [
    rowInsight({
      id: "revenue-cutoff.early",
      aliases: ["cutoff"],
      title: "Recorded this year, shipped next year",
      severity: "high",
      why: "Revenue was booked before the goods left. If control hadn't passed by year-end, this year's revenue and receivables are overstated.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const inv = dateOf(row, "invoiceDate");
        const ship = dateOf(row, "shipDate");
        if (inv === null || ship === null || inv > ctx.yearEnd! || ship <= ctx.yearEnd!) return null;
        return { cellKey: "shipDate", amount: amountOf(row, "amount"), detail: `Invoiced ${fmtDate(inv)}, shipped ${fmtDate(ship)}` };
      },
    }),
    rowInsight({
      id: "revenue-cutoff.late",
      aliases: ["cutoff"],
      title: "Shipped this year, recorded next year",
      severity: "high",
      why: "The goods left before year-end but the sale was booked after. This year's revenue and receivables are understated.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const inv = dateOf(row, "invoiceDate");
        const ship = dateOf(row, "shipDate");
        if (inv === null || ship === null || ship > ctx.yearEnd! || inv <= ctx.yearEnd!) return null;
        return { cellKey: "invoiceDate", amount: amountOf(row, "amount"), detail: `Shipped ${fmtDate(ship)}, invoiced ${fmtDate(inv)}` };
      },
    }),
    rowInsight({
      id: "revenue-cutoff.credit-notes",
      title: "Credit notes after year-end",
      severity: "medium",
      why: "Credit notes issued soon after year-end often reverse sales booked to hit targets. Check whether the original sale belongs in this year.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const amount = amountOf(row, "amount");
        const inv = dateOf(row, "invoiceDate");
        if (amount === null || amount >= 0 || inv === null || inv <= ctx.yearEnd!) return null;
        return { cellKey: "amount", amount, detail: `Credit dated ${fmtDate(inv)}` };
      },
    }),
    rowInsight({
      id: "revenue-cutoff.large-near-ye",
      title: "Large sales right around year-end with no conclusion",
      severity: "medium",
      why: "Big sales in the last and first few days of the year carry the most cut-off risk. Each needs a clear conclusion.",
      needs: ["yearEnd", "threshold"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const amount = amountOf(row, "amount");
        const inv = dateOf(row, "invoiceDate");
        if (amount === null || inv === null || Math.abs(amount) < ctx.threshold!) return null;
        if (Math.abs(inv - ctx.yearEnd!) > NEAR_YE_DAYS || !isBlank(row.cells.conclusion)) return null;
        return { cellKey: "conclusion", amount, detail: `Invoiced ${fmtDate(inv)}, no conclusion` };
      },
    }),
  ],
};
