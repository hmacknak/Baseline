// Product switches.

/** Procedure reviewers are on. Only the ones listed in ENABLED_PROCEDURES appear in the catalog. */
export const PROCEDURES_ENABLED = true;

/**
 * Procedure reviewers offered to users. Others (bank reconciliation, revenue cut-off, AR confirmations)
 * remain in src/review/reviewers/ with their tests but are switched off.
 */
export const ENABLED_PROCEDURES = ["surl", "cash-confirmations", "expense-vouching"];
