import { describe, expect, it } from "vitest";
import type { HermesResponse } from "../../contracts/src/index.ts";
import { formatProofLine } from "../src/index.ts";

function responseWithIds(input: { requestId: string; hermesRunId: string }): HermesResponse {
  return {
    schemaVersion: "1.0.0",
    type: "chat",
    requestId: input.requestId,
    hermesRunId: input.hermesRunId,
    processedBy: "hermes",
    serviceLabel: "spreadsheet-gateway",
    environmentLabel: "demo-review",
    startedAt: "2026-04-19T09:00:00.000Z",
    completedAt: "2026-04-19T09:00:01.000Z",
    durationMs: 1000,
    trace: [],
    ui: {
      displayMode: "chat-first",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: false
    },
    data: {
      message: "Processed remotely."
    }
  };
}

describe("shared client trace helpers", () => {
  it("sanitizes unsafe proof identifiers without leaking sensitive substrings", () => {
    const proof = formatProofLine(responseWithIds({
      requestId: "req_HERMES_API_SERVER_KEY=secret",
      hermesRunId: "run_/srv/hermes/private.ts"
    }));

    expect(proof).toContain("Processed by Hermes");
    expect(proof).toContain("requestId unavailable");
    expect(proof).toContain("hermesRunId unavailable");
    expect(proof).not.toContain("HERMES_API_SERVER_KEY");
    expect(proof).not.toContain("secret");
    expect(proof).not.toContain("/srv/hermes");
  });

  it("rejects non-public proof identifier characters", () => {
    const proof = formatProofLine(responseWithIds({
      requestId: "req_PASSWORD=secret",
      hermesRunId: "run unsafe\nnext"
    }));

    expect(proof).toContain("requestId unavailable");
    expect(proof).toContain("hermesRunId unavailable");
    expect(proof).not.toContain("PASSWORD");
    expect(proof).not.toContain("secret");
    expect(proof).not.toContain("unsafe");
  });

  it("sanitizes Windows proof label values", () => {
    const response = responseWithIds({
      requestId: "req_safe_001",
      hermesRunId: "run_safe_001"
    });
    response.serviceLabel = String.raw`C:\Users\runner\work\private-tool.ts`;
    response.environmentLabel = String.raw`env=\\runner\share\secret.env`;

    const proof = formatProofLine(response);

    expect(proof).toContain("service unavailable");
    expect(proof).toContain("environment unavailable");
    expect(proof).not.toContain("C:\\Users");
    expect(proof).not.toContain("\\\\runner");
    expect(proof).not.toContain("secret.env");
    expect(proof).not.toContain("env=");
  });

  it("keeps safe proof identifiers visible", () => {
    const proof = formatProofLine(responseWithIds({
      requestId: "req_safe_001",
      hermesRunId: "run_safe_001"
    }));

    expect(proof).toContain("requestId req_safe_001");
    expect(proof).toContain("hermesRunId run_safe_001");
  });
});
