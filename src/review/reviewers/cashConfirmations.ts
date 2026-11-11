// Cash (bank) confirmations: the bank confirms, directly to the auditor, every account the entity holds,
// its balance at year-end, and other arrangements such as loans, guarantees and restrictions.
//
// Research basis (see docs/PROCEDURES.md):
// - ISA 505.7: the auditor maintains control over confirmation requests (what to confirm, who to send to,
//   sending the requests). ISA 505.12-.13: non-responses need alternative procedures, and where a confirmation is
//   necessary the auditor must consider the implications of not getting one. Exceptions must be investigated.
// - PCAOB AS 2310 (periods ending on or after 15 June 2025): cash held by third parties must be confirmed, or
//   evidence obtained directly from a knowledgeable external source. The auditor sends and receives directly
//   (AS 2310.14-.16); intermediaries are allowed only for direct electronic transmission (.17).
// - Assertions: existence and rights (the money is there and belongs to the entity), completeness (all accounts,
//   including nil and closed ones, which can hide loans, guarantees or restrictions).
// - Presentation: overdrafts are financing unless repayable on demand and integral to cash management (IAS 7.8).
//   Foreign-currency balances are translated at the closing rate (IAS 21.23).

import { amountOf, dateOf, isBlank, textOf } from "../table";
import { Reviewer } from "../types";
import { TableSpec, columnsCheck, completenessCheck, detectScore, fmtAmount, fmtDate, rowInsight, tableFor } from "./procedure";

const spec: TableSpec = {
  minColumns: 3,
  columns: [
    { key: "bank", label: "bank", kind: "text", names: ["bank", "bank name", "institution", "financial institution", "counterparty"] },
    {
      key: "account",
      label: "account",
      kind: "text",
      names: ["account", "account #", "account no", "account number", "acct", "acct #", "iban", "account name"],
    },
    { key: "currency", label: "currency", kind: "text", names: ["currency", "ccy", "cur"] },
    {
      key: "books",
      label: "balance per books",
      kind: "amount",
      names: ["balance per books", "per books", "balance per gl", "per gl", "gl balance", "book balance", "balance per ledger", "ledger balance", "balance per tb", "per tb", "per books/gl"],
    },
    {
      key: "confirmed",
      label: "balance per confirmation",
      kind: "amount",
      names: ["balance per confirmation", "per confirmation", "confirmed balance", "confirmed amount", "balance per bank", "per bank", "bank balance", "confirmed"],
    },
    { key: "asAt", label: "confirmation date", kind: "date", names: ["as at", "as of", "confirmation date", "balance date", "date confirmed"] },
    { key: "sent", label: "date sent", kind: "date", names: ["date sent", "sent", "sent date", "request sent", "requested", "date requested"] },
    { key: "received", label: "date received", kind: "date", names: ["date received", "received", "response received", "response date", "replied", "reply"] },
    { key: "channel", label: "sent / received via", kind: "text", names: ["sent by", "sent via", "method", "channel", "platform", "via", "received via", "received from", "response from", "returned by", "source of response"] },
    { key: "explanation", label: "explanation", kind: "text", names: ["explanation", "reconciling items", "reconciliation", "difference explained", "reason"] },
    { key: "conclusion", label: "conclusion", kind: "text", names: ["conclusion", "result", "comment", "comments", "notes"] },
  ],
};

const LABEL = ["bank", "account"];

/** Signs that a response passed through the entity rather than coming straight from the bank. */
const VIA_ENTITY = /\b(client|entity|company|management|cfo|controller|finance (team|dept)|treasur(y|er)|accountant|bookkeeper|forward(ed)?|fwd|scan(ned)?|copy|photocop)/i;

export const cashConfirmations: Reviewer = {
  id: "cash-confirmations",
  name: "Cash confirmations",
  category: "Cash",
  code: "CC",
  summary: "Checks every bank account is confirmed directly with the bank, differences are explained, and no reply goes unchased.",
  alwaysOn: false,
  detect: sheet => {
    const score = detectScore(sheet, {
      name: /cash|bank/i,
      text: /bank confirm|cash confirm|financial institution/i,
      spec,
      columnsForFull: 4,
    });
    // Customer (AR) confirmations look similar; only a bank column makes this a cash confirmation.
    const table = tableFor(sheet, spec);
    return table && table.columns.bank !== undefined ? score : Math.min(score, 0.4);
  },
  checks: [
    completenessCheck({
      id: "cash-confirmations.complete",
      spec,
      labelKeys: LABEL,
      required: ["bank", "account", "books", "sent", "conclusion"],
      tableHint: "Add a header row with columns like Bank, Account, Balance per books, Balance per confirmation, Date sent, Date received, Conclusion.",
      assertion: "Existence; Completeness",
      reference: "ISA 505.7; PCAOB AS 2310",
      why: "A confirmation only counts as evidence if the file shows which account it covers, what the books said, when it was sent and what you concluded. Gaps here are the first thing a reviewer or inspector checks, because cash is where a confirmation is close to mandatory.",
    }),
    columnsCheck({ id: "cash-confirmations.columns", spec, needed: ["bank", "books", "confirmed", "sent"] }),
  ],
  insights: [
    rowInsight({
      id: "cash-confirmations.difference",
      title: "Differences with the bank not explained",
      severity: "high",
      assertion: "Existence; Accuracy",
      reference: "ISA 505 (exceptions); PCAOB AS 2310",
      why: "The bank's figure is independent evidence; the books are the company's own. A difference is a potential misstatement until it's reconciled, usually by items in transit, which should then agree to the bank reconciliation. An unexplained difference in cash is also a classic sign of misappropriation.",
      aliases: ["confirm-difference"],
      spec,
      labelKeys: LABEL,
      test(row) {
        const books = amountOf(row, "books");
        const bank = amountOf(row, "confirmed");
        if (books === null || bank === null || Math.abs(books - bank) < 0.5 || !isBlank(row.cells.explanation)) return null;
        return {
          cellKey: "confirmed",
          amount: books - bank,
          detail: `Books ${fmtAmount(books)} vs bank ${fmtAmount(bank)}, no explanation`,
        };
      },
    }),
    rowInsight({
      id: "cash-confirmations.no-response",
      title: "No reply from the bank",
      severity: "high",
      assertion: "Existence; Rights and obligations",
      reference: "ISA 505.12-.13; PCAOB AS 2310 (cash held by third parties)",
      why: "Without the bank's reply there is no independent evidence the balance exists. Follow up first. For cash, other evidence is weak: a bank statement provided by the company can be altered. Under PCAOB AS 2310 cash must be confirmed or evidenced directly from a knowledgeable external source; under ISA 505.13, if a confirmation is needed and you can't get one, you must assess the effect on the audit.",
      aliases: ["confirm-no-response"],
      spec,
      labelKeys: LABEL,
      test(row) {
        if (dateOf(row, "sent") === null) return null;
        if (amountOf(row, "confirmed") !== null || dateOf(row, "received") !== null || !isBlank(row.cells.received)) return null;
        return { cellKey: "received", amount: amountOf(row, "books"), detail: "Sent, no reply recorded" };
      },
    }),
    rowInsight({
      id: "cash-confirmations.not-direct",
      title: "Not sent or received directly",
      severity: "high",
      assertion: "Existence",
      reference: "ISA 505.7; PCAOB AS 2310.14-.17",
      why: "The auditor must control the confirmation: send it to the bank and get the reply straight back. If it passes through the company, it could be intercepted or altered, which is how several well-known cash frauds stayed hidden. Electronic confirmation platforms are fine because they transmit directly; a reply forwarded by the client is not.",
      aliases: ["confirm-control"],
      spec,
      labelKeys: LABEL,
      test(row) {
        const via = textOf(row, "channel");
        if (!via || !VIA_ENTITY.test(via)) return null;
        return { cellKey: "channel", amount: amountOf(row, "books"), detail: `"${via}"` };
      },
    }),
    rowInsight({
      id: "cash-confirmations.not-sent",
      title: "Accounts not confirmed",
      severity: "medium",
      assertion: "Completeness; Rights and obligations",
      reference: "ISA 505.7; PCAOB AS 2310",
      why: "Every account should be confirmed, including nil-balance and closed ones. The bank's reply also lists loans, guarantees, security and restrictions linked to the account, and that information is often what surfaces an unrecorded liability or a cash balance the company can't actually use.",
      spec,
      labelKeys: LABEL,
      test(row) {
        if (dateOf(row, "sent") !== null || !isBlank(row.cells.sent)) return null;
        const books = amountOf(row, "books");
        return {
          cellKey: "sent",
          amount: books,
          detail: books === 0 ? "Nil-balance account, no confirmation sent" : "No confirmation sent",
        };
      },
    }),
    rowInsight({
      id: "cash-confirmations.date",
      title: "Confirmed at a different date",
      severity: "medium",
      assertion: "Existence; Cut-off",
      reference: "ISA 505 (designing the request)",
      why: "The balance the bank confirms has to be the year-end balance to support the year-end figure. A confirmation at another date needs a roll-forward to year-end before it proves anything.",
      needs: ["yearEnd"],
      spec,
      labelKeys: LABEL,
      test(row, ctx) {
        const asAt = dateOf(row, "asAt");
        if (asAt === null || asAt === ctx.yearEnd) return null;
        return { cellKey: "asAt", amount: amountOf(row, "books"), detail: `Confirmed as at ${fmtDate(asAt)}` };
      },
    }),
    rowInsight({
      id: "cash-confirmations.overdraft",
      title: "Overdrawn accounts",
      severity: "low",
      assertion: "Classification; Presentation",
      reference: "IAS 7.8",
      why: "A negative bank balance is a borrowing. It can only sit within cash and cash equivalents if it is repayable on demand and part of day-to-day cash management (IAS 7.8); otherwise it is a liability and a financing cash flow. Check where it is presented and that it isn't netted against other cash without a right of set-off.",
      spec,
      labelKeys: LABEL,
      test(row) {
        const books = amountOf(row, "books");
        if (books === null || books >= 0) return null;
        return { cellKey: "books", amount: books, detail: "Overdrawn at year-end" };
      },
    }),
    {
      id: "cash-confirmations.foreign",
      title: "Foreign-currency balances",
      severity: "low",
      assertion: "Valuation",
      reference: "IAS 21.23",
      why: "Cash in another currency is a monetary item, so it must be translated at the year-end closing rate (IAS 21.23). Check the rate and that the translated figure, not the foreign-currency figure, is what agrees to the ledger.",
      run(sheet) {
        const table = tableFor(sheet, spec);
        if (!table || table.columns.currency === undefined) return [];
        const counts = new Map<string, number>();
        for (const row of table.rows) {
          const c = textOf(row, "currency").toUpperCase();
          if (c) counts.set(c, (counts.get(c) || 0) + 1);
        }
        if (counts.size < 2) return [];
        let main = "";
        counts.forEach((n, c) => {
          if (!main || n > counts.get(main)!) main = c;
        });
        return table.rows
          .filter(row => {
            const c = textOf(row, "currency").toUpperCase();
            return c && c !== main;
          })
          .map(row => ({
            cell: row.cells.currency!.address,
            label: textOf(row, "bank") || `Row ${row.row + 1}`,
            amount: amountOf(row, "books") ?? undefined,
            detail: `${textOf(row, "currency").toUpperCase()} balance; most accounts are ${main}`,
          }));
      },
    },
  ],
};
