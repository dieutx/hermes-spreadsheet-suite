import { describe, expect, it } from "vitest";
import { createApprovalToken, verifyApprovalToken } from "../src/lib/approval.ts";

describe("write-back approval tokens", () => {
  it("signs and verifies a request-bound write approval", () => {
    const secret = "demo-secret";
    const token = createApprovalToken({
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      issuedAt: "2026-04-19T09:30:00.000Z",
      secret
    });

    const verified = verifyApprovalToken({
      token,
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      secret
    });

    expect(verified.valid).toBe(true);
  });

  it("adds a random nonce so identical approvals do not reuse the same token", () => {
    const secret = "demo-secret";
    const input = {
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      issuedAt: "2026-04-19T09:30:00.000Z",
      secret
    };

    const firstToken = createApprovalToken(input);
    const secondToken = createApprovalToken(input);

    expect(firstToken).not.toBe(secondToken);

    const [encodedPayload] = firstToken.split(".");
    const payload = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8"));

    expect(payload.nonce).toMatch(/^[A-Za-z0-9_-]{22,}$/);
    expect(
      verifyApprovalToken({
        token: firstToken,
        requestId: input.requestId,
        runId: input.runId,
        planDigest: input.planDigest,
        secret
      }).valid
    ).toBe(true);
  });

  it("rejects a token when the write plan digest changes", () => {
    const secret = "demo-secret";
    const token = createApprovalToken({
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      issuedAt: "2026-04-19T09:30:00.000Z",
      secret
    });

    const verified = verifyApprovalToken({
      token,
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_other",
      secret
    });

    expect(verified.valid).toBe(false);
  });

  it("rejects a token when it is older than the allowed approval TTL", () => {
    const secret = "demo-secret";
    const issuedAt = "2026-04-19T09:30:00.000Z";
    const token = createApprovalToken({
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      issuedAt,
      secret
    });

    const verified = verifyApprovalToken({
      token,
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      secret,
      maxAgeMs: 60_000,
      nowMs: Date.parse("2026-04-19T09:32:00.000Z")
    });

    expect(verified.valid).toBe(false);
    expect(verified.expired).toBe(true);
  });

  it("rejects a token issued in the future", () => {
    const secret = "demo-secret";
    const issuedAt = "2026-04-19T09:35:00.000Z";
    const token = createApprovalToken({
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      issuedAt,
      secret
    });

    const verified = verifyApprovalToken({
      token,
      requestId: "req_123",
      runId: "run_123",
      planDigest: "digest_123",
      secret,
      maxAgeMs: 60_000,
      nowMs: Date.parse("2026-04-19T09:32:00.000Z")
    });

    expect(verified.valid).toBe(false);
    expect(verified.expired).toBe(false);
  });
});
