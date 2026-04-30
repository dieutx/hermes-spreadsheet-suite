import type { CompositePlanData } from "@hermes/contracts";

// Keep composite digest normalization centralized so dry-run, approval, and
// completion compare the same contract shape.
export function normalizeCompositePlanForDigest(
  plan: CompositePlanData
): CompositePlanData {
  return plan;
}
