import type { CompositePlanData } from "@hermes/contracts";

// Composite workflows do not yet have an exact rollback contract, so the gateway
// must normalize them to the non-reversible preview form before digesting them.
export function normalizeCompositePlanForDigest(
  plan: CompositePlanData
): CompositePlanData {
  if (!plan.reversible) {
    return plan;
  }

  return {
    ...plan,
    reversible: false
  };
}
