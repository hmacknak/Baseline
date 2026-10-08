// AR confirmations: every selected balance is confirmed, differences are explained, and
// non-responses are covered by alternative procedures.

import { amountOf, dateOf, flagOf, isBlank } from "../table";
import { Reviewer } from "../types";
import { TableSpec, columnsCheck, completenessCheck, detectScore, fmtAmount, rowInsight } from "./procedure";

const spec: TableSpec = {
  minColumns: 3,
  columns: [
    { key: "customer", label: "customer", kind: "text", names: ["customer", "client", "debtor", "customer name", "account"] },
    {
      key: "balance",
      label: "balance per ledger",
      kind: "amount",
      names: ["balance", "balance per ledger", "ledger balance", "per ledger", "balance per books", "balance per client", "ar balance", "amount"],
    },
    { key: "sent", label: "date sent", kind: "date", names: ["sent", "date sent", "sent date", "mailed", "date mailed", "sent on"] },
    {
      key: "confirmed",
      label: "confirmed amount",
      kind: "amount",
      names: ["confirmed balance", "confirmed amount", "per confirmation", "amount confirmed", "balance per confirmation", "response amount", "confirmed"],
    },
    { key: "response", label: "response received", kind: "flag", names: ["response", "response received", "received", "date received", "response date", "reply"] },
    { key: "explanation", label: "explanation", kind: "text", names: ["explanation", "reconciling items", "reconciliation", "reason", "comment", "comments"] },
    {
      key: "altProc",
      label: "alternative procedures",
      kind: "text",
      names: ["alternative procedures", "alternative procedure", "alt procedures", "alt proc", "subsequent receipts", "alt procedures performed"],
    },
  ],
};

const LABEL = ["customer"];

function responded(row: Parameters<typeof amountOf>[0]): boolean {
  if (amountOf(row, "confirmed") !== null) return true;
  return flagOf(row, "response") === true || dateOf(row, "response") !== null;
}

export const arConfirmations: Reviewer = {
  id: "ar-confirmations",
  name: "AR confirmations",
  category: "Receivables",
  code: "AC",
  summary: "Checks every confirm was sent, differences are explained, and non-responses have alternative procedures.",
  alwaysOn: false,
  detect: sheet =>
    detectScore(sheet, {
      name: /confirm|\bconf\b|\barc\b/i,
      text: /confirmation/i,
      spec,
      columnsForFull: 4,
    }),
  checks: [
    completenessCheck({
      id: "ar-confirmations.complete",
      spec,
      labelKeys: LABEL,
      required: ["customer", "balance", "sent"],
      tableHint: "Add a header row with columns like Customer, Balance per ledger, Date sent, Confirmed amount, Explanation, Alternative procedures.",
    }),
    columnsCheck({ id: "ar-confirmations.columns", spec, needed: ["balance", "sent", "confirmed"] }),
  ],
  insights: [
    rowInsight({
      id: "ar-confirmations.difference",
      aliases: ["confirm-difference"],
      title: "Differences not explained",
      severity: "high",
      why: "A confirmed amount that doesn't match the ledger is a potential misstatement until it's reconciled, usually by timing (cash in transit) or a dispute.",
      spec,
      labelKeys: LABEL,
      test(row) {
        const bal = amountOf(row, "balance");
        const conf = amountOf(row, "confirmed");
        if (bal === null || conf === null || Math.abs(bal - conf) < 0.5 || !isBlank(row.cells.explanation)) return null;
        return {
          cellKey: "confirmed",
          amount: bal - conf,
          detail: `Ledger ${fmtAmount(bal)} vs confirmed ${fmtAmount(conf)}, no explanation`,
        };
      },
    }),
    rowInsight({
      id: "ar-confirmations.no-response",
      aliases: ["confirm-no-response"],
      title: "No response and no alternative procedures",
      severity: "high",
      why: "An unanswered confirm gives no evidence on its own. Alternative procedures, such as subsequent receipts or shipping documents, are needed to support the balance.",
      spec,
      labelKeys: LABEL,
      test(row) {
        if (dateOf(row, "sent") === null || responded(row) || !isBlank(row.cells.altProc)) return null;
        return { cellKey: "altProc", amount: amountOf(row, "balance"), detail: "Sent, no response, no alternative procedures" };
      },
    }),
    rowInsight({
      id: "ar-confirmations.not-sent",
      title: "Large balances not confirmed",
      severity: "medium",
      why: "Balances over your threshold are usually selected for confirmation. A blank sent date means there's no evidence it went out.",
      needs: ["threshold"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const bal = amountOf(row, "balance");
        if (bal === null || Math.abs(bal) < ctx.threshold! || dateOf(row, "sent") !== null) return null;
        return { cellKey: "sent", amount: bal, detail: "No date sent" };
      },
    }),
  ],
};
