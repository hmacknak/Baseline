# Procedure reviewers: research basis

What each procedure is for, the assertions and risks behind it, how it is normally performed, and what Baseline
checks. Sources are standard-setters (IAASB, PCAOB, IFRS Foundation) and professional bodies. The standards
sites themselves could not be opened from the build environment, so paragraph numbers were confirmed through
published copies of the standards (national adoptions and IFRS Foundation files) and are listed below. Where a
paragraph number could not be confirmed, the standard is cited without one.

Assertions follow ISA 315 (Revised 2019): for transactions, **occurrence, completeness, accuracy, cut-off,
classification**; for balances, **existence, rights and obligations, completeness, accuracy/valuation and
allocation**.

**PCAOB engagements.** PCAOB AS 2110 groups assertions differently: existence or occurrence, completeness,
valuation or allocation, rights and obligations, presentation and disclosure. There is no separate cut-off
assertion, so a cut-off finding maps to occurrence (expenses recorded too early) or completeness (recorded too
late). Each finding therefore cites the PCAOB standard next to the ISA. The accounting references (IAS 37,
IAS 16, IAS 7, IAS 21, Conceptual Framework) are IFRS; for a US GAAP client the matching ASC topics apply and are
not yet shown. Paragraph numbers for PCAOB standards are cited only where they were confirmed.

---

## 1. Search for unrecorded liabilities (SURL)

**Objective.** Find obligations that existed at year-end but were not recorded.

**Why it matters.**
- A liability is a present obligation to transfer an economic resource as a result of past events (IFRS Conceptual
  Framework 4.26). Goods or services received by year-end create that obligation whether or not an invoice has
  arrived.
- Invoiced amounts are trade payables; un-invoiced amounts are accruals (IAS 37.11).
- Management has an incentive to leave costs out: understating liabilities and expenses flatters results. So the
  risk is **understatement**, and the assertion is **completeness**.

**How it is done.**
1. Select payments and invoices from *after* year-end, which is evidence from outside the payables ledger.
2. For each one, establish when the goods or service were received.
3. If that was on or before year-end, check the amount was recorded or accrued.
4. The search runs over the period after year-end up to as close to the auditor's report date as practicable
   (ISA 560). Practice is to cover at least one full payment cycle, longer when risk is higher.

**What Baseline checks.**

| Rule | Assertion | Reference |
|---|---|---|
| Post-year-end invoice for a service on or before year-end, not recorded | Completeness | Conceptual Framework 4.26; IAS 37.11; PCAOB AS 2301 |
| Invoice dated on or before year-end, marked not recorded | Completeness | IAS 37.11 |
| Recorded at year-end but the service was next year (over-accrual) | Existence; cut-off | Conceptual Framework 4.26 |
| No service date, so the year can't be decided | Completeness; cut-off | ISA 500 |
| Search window ends less than 30 days after year-end | Completeness | ISA 560; PCAOB AS 2801 |
| Every selection complete; key columns present | Completeness | ISA 500; ISA 230; PCAOB AS 1105; AS 1215 |

---

## 2. Cash confirmations (bank confirmations)

**Objective.** Get independent evidence, directly from each bank, that cash balances exist, belong to the
entity, and are complete. The same request also brings back loans, guarantees, security and restrictions.

**Why it matters.**
- Cash is the asset most exposed to misappropriation. A confirmation sent and received by the auditor is
  among the most reliable evidence available, but only if the auditor controls it. A confirmation routed
  through the company can be intercepted or altered.
- **ISA 505.7:** the auditor maintains control over confirmation requests (what to confirm, the right confirming
  party, sending the requests).
- **ISA 505.10 and A11:** doubts about the reliability of a response, such as one received indirectly through
  the client, call for further evidence (A14: for example, contact the bank or ask it to re-send directly).
- **ISA 505.12–.13:** non-responses need alternative procedures. Where a confirmation is necessary, the auditor
  must determine the implications of not obtaining it. ISA 505 also requires exceptions to be investigated
  (paragraph not confirmed).
- **PCAOB AS 2310** (periods ending on or after 15 June 2025):
  - Cash held by third parties must be confirmed, or evidence obtained by directly accessing a knowledgeable
    external source.
  - The auditor keeps control so information is less likely to be intercepted or altered (.14). The auditor
    selects items, sends requests and receives responses (.15), directly with the confirming party (.16).
  - Intermediaries are permitted only to facilitate direct electronic transmission, and must be evaluated (.17).
- Nil-balance and closed accounts are confirmed too: the bank's reply can surface loans, guarantees or
  restrictions that aren't in the books (completeness; rights and obligations).
- Presentation:
  - An overdraft is a borrowing. It is part of cash and cash equivalents only if it is repayable on demand and
    integral to cash management (IAS 7.8; IFRS Interpretations Committee agenda decision, June 2018).
  - Foreign-currency cash is a monetary item translated at the closing rate (IAS 21.23).

**How it is done.**
1. List every bank and account, from the ledger and from the previous year's replies.
2. Send requests directly to the banks (often through an electronic confirmation platform), as at year-end.
3. Receive replies directly and agree each balance to the books; investigate and document differences, usually
   reconciling items that should agree to the bank reconciliation.
4. Chase non-replies, and evaluate any extra information in the reply (loans, guarantees, security, restrictions).

**What Baseline checks.**

| Rule | Assertion | Reference |
|---|---|---|
| Books and bank differ with no explanation | Existence; accuracy | ISA 505; AS 2310 |
| Sent, no reply | Existence; rights and obligations | ISA 505.12–.13; AS 2310 |
| Sent or received through the client | Existence | ISA 505.7, .10, A11; AS 2310.14–.17 |
| Account not confirmed (including nil balances) | Completeness; rights and obligations | ISA 505.7; AS 2310 |
| Confirmed at a date other than year-end | Existence; cut-off | ISA 505 |
| Overdrawn account | Classification; presentation | IAS 7.8 |
| Foreign-currency balance | Valuation | IAS 21.23 |
| Every account complete; key columns present | Existence; completeness | ISA 505.7; AS 2310 |

---

## 3. Expense vouching (test of details of expenses)

**Objective.** Show that a sample of recorded expenses really happened, at the right amount, in the right year,
in the right account, and was approved.

**Why it matters.**
- **Direction of testing:** vouching from the ledger back to supporting documents tests **occurrence** (were
  recorded costs real?). Tracing from documents into the ledger tests completeness, which is the SURL's job.
- Assertions: occurrence, accuracy, cut-off and classification (ISA 315).
- **Fraud:** causing the entity to pay for goods or services not received, including through fictitious vendors,
  is a recognised management-override technique (ISA 240 application material). Duplicate payments are the
  innocent version of the same leak.
- **Classification:** a cost that meets IAS 16.7 is capitalised as property, plant and equipment, while
  day-to-day servicing and repairs are expensed (IAS 16.12). Getting this wrong misstates both assets and profit.
- Errors found in a sample stand for the population, so they are **projected**, not just corrected (ISA 530).

**How it is done.**
1. Select a sample from the expense ledger.
2. For each item, agree vendor, amount and date to the invoice and supporting documents.
3. Confirm the goods or service were received in the year, the account is right, and the item was approved.
4. Record exceptions and project them to the population.

**What Baseline checks.**

| Rule | Assertion | Reference |
|---|---|---|
| Ledger amount doesn't agree to support | Accuracy | ISA 315; ISA 530; PCAOB AS 2315 |
| Service (or, failing that, invoice) date outside the year | Cut-off | ISA 315; PCAOB AS 2110 (assertions) |
| Same vendor, amount and invoice twice | Occurrence | ISA 240; PCAOB AS 2401 |
| Approval refused or not recorded | Occurrence | ISA 315 |
| Description suggests a capital item over threshold | Classification | IAS 16.7; IAS 16.12 |
| Marked as recorded in the wrong account | Classification | ISA 315 |
| Every sample item complete; key columns present | Occurrence; accuracy | ISA 500; ISA 530; PCAOB AS 1105; AS 2315 |

---

## Switched off for now

Bank reconciliation, revenue cut-off and AR confirmations reviewers remain in `src/review/reviewers/` with
their tests, but are not offered (`src/config.ts`, `ENABLED_PROCEDURES`).

## Sources

- PCAOB, AS 2310 *The Auditor's Use of Confirmation*: https://pcaobus.org/oversight/standards/auditing-standards/details/AS2310 ;
  adopting release 2023-008: https://assets.pcaobus.org/pcaob-dev/docs/default-source/rulemaking/docket_028/2023-008_confirmation-adopting-release.pdf ;
  Federal Register notice: https://www.govinfo.gov/content/pkg/FR-2023-10-17/pdf/2023-22491.pdf
- ISA 505 *External Confirmations* (as adopted in New Zealand, same numbering): https://xrb.govt.nz/dmsdocument/1605
- ICAEW, *External confirmations in an ISA audit*: https://www.icaew.com/technical/audit-and-assurance/faculty-resources/implementing-isas/external-confirmations-in-an-isa-audit
- AUASB, *Bank Confirmation Requests* (GS 016): https://auasb.gov.au/admin/file/content102/c3/GS_016_Bank_Confirmation_Requests_23-Jun-10.pdf
- ISA 560 *Subsequent Events* (as adopted in New Zealand): https://standards.xrb.govt.nz/standards-navigator/isa-nz-560
- ISA 240 application material (Australian adoption): https://standards.auasb.gov.au/node/227 ;
  ICAEW, *Management override*: https://www.icaew.com/technical/audit-and-assurance/faculty-resources/implementing-isas/management-override
- ACCA, *Assertions* (ISA 315): https://www.accaglobal.com/gb/en/student/exam-support-resources/fundamentals-exams-study-resources/f8/technical-articles/assertions.html
- IFRS Foundation, IAS 7: https://www.ifrs.org/content/dam/ifrs/publications/html-standards/english/2022/issued/ias7.html ;
  agenda decision on short-term loans and credit facilities (2018): https://www.ifrs.org/content/dam/ifrs/supporting-implementation/agenda-decisions/2018/ias-7-agenda-decision-classification-of-short-term-loans-and-credit-facilities-june-18.pdf
- IFRS Foundation, IAS 16: https://ifrs.org/content/dam/ifrs/publications/html-standards/english/2022/issued/ias16.html
- IFRS Foundation, IAS 21: https://ifrs.org/content/dam/ifrs/publications/html-standards/english/2021/issued/ias21.html
- IAS 37 definitions (Australian adoption, same text): https://standards.aasb.gov.au/node/2796
- Conceptual Framework (2018), definition of a liability 4.26 (HKICPA reproduction): https://www.hkicpa.org.hk/-/media/HKICPA-Website/Members-Handbook/volumeII/framework18.pdf
- Institute of Singapore Chartered Accountants, accounts payable audit highlights (search coverage and payment cycles): https://isca.org.sg/docs/default-source/audit-assurance/events/accountspluspayable_highlightsplusofplustc.pdf
