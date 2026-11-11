// Accounting framework: the same audit risk, cited against IFRS or US GAAP.
// Audit references (ISA, PCAOB AS) don't change; accounting references and any wording that depends on them do.

import { Framework, FrameworkText, ReviewContext } from "./types";

export const FRAMEWORKS: Array<{ id: Framework; name: string }> = [
  { id: "ifrs", name: "IFRS" },
  { id: "us-gaap", name: "US GAAP" },
];

export function frameworkOf(ctx: ReviewContext | undefined): Framework {
  return ctx && ctx.framework === "us-gaap" ? "us-gaap" : "ifrs";
}

/** The "why it matters" text and reference to show for this workbook's framework. */
export function explain(
  item: { why: string; reference?: string; usGaap?: FrameworkText },
  ctx: ReviewContext | undefined,
): { why: string; reference?: string } {
  const over = frameworkOf(ctx) === "us-gaap" ? item.usGaap : undefined;
  return {
    why: (over && over.why) || item.why,
    reference: (over && over.reference) || item.reference,
  };
}
