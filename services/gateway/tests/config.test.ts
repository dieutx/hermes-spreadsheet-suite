import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { getConfig } from "../src/lib/config.ts";

beforeEach(() => {
  process.env.APPROVAL_SECRET = "test-approval-secret";
});

afterEach(() => {
  delete process.env.HERMES_DEBUG_INVALID_RESPONSES;
  delete process.env.APPROVAL_SECRET;
  delete process.env.GATEWAY_PUBLIC_BASE_URL;
  delete process.env.GATEWAY_ALLOWED_ORIGINS;
  delete process.env.HERMES_SERVICE_LABEL;
  delete process.env.HERMES_ENVIRONMENT_LABEL;
  delete process.env.HERMES_AGENT_BASE_URL;
  delete process.env.HERMES_API_SERVER_KEY;
  delete process.env.HERMES_AGENT_API_KEY;
  delete process.env.HERMES_AGENT_MODEL;
  delete process.env.HERMES_AGENT_ID;
  delete process.env.HERMES_AGENT_TIMEOUT_MS;
  delete process.env.HERMES_BASE_URL;
  delete process.env.MAX_UPLOAD_BYTES;
});

describe("gateway config", () => {
  it("reads Hermes Agent API server settings from the explicit env vars", () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test";
    process.env.HERMES_API_SERVER_KEY = "agent-secret";
    process.env.HERMES_AGENT_MODEL = "hermes-agent";

    expect(getConfig()).toMatchObject({
      hermesAgentBaseUrl: "http://agent.test",
      hermesAgentApiKey: "agent-secret",
      hermesAgentModel: "hermes-agent"
    });
  });

  it("keeps HERMES_AGENT_ID as a compatibility fallback for the Hermes model", () => {
    process.env.HERMES_AGENT_ID = "legacy-agent-id";

    expect(getConfig()).toMatchObject({
      hermesAgentModel: "legacy-agent-id"
    });
  });

  it("keeps HERMES_BASE_URL as a legacy fallback for the Hermes Agent base url", () => {
    process.env.HERMES_BASE_URL = "http://legacy-agent.test";

    expect(getConfig()).toMatchObject({
      hermesAgentBaseUrl: "http://legacy-agent.test"
    });
  });

  it("defaults the Hermes Agent API base url to the local API server path", () => {
    expect(getConfig()).toMatchObject({
      hermesAgentBaseUrl: "http://127.0.0.1:8642/v1"
    });
  });

  it("sanitizes public gateway labels before exposing them to clients", () => {
    process.env.HERMES_SERVICE_LABEL = "HERMES_API_SERVER_KEY=secret";
    process.env.HERMES_ENVIRONMENT_LABEL = "/srv/hermes/internal-config.ts";

    expect(getConfig()).toMatchObject({
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev"
    });
  });

  it("defaults the Hermes Agent timeout to 45 seconds", () => {
    expect(getConfig()).toMatchObject({
      hermesAgentTimeoutMs: 45_000
    });
  });

  it("accepts an explicit Hermes Agent timeout override", () => {
    process.env.HERMES_AGENT_TIMEOUT_MS = "120000";

    expect(getConfig()).toMatchObject({
      hermesAgentTimeoutMs: 120_000
    });
  });

  it("uses the configured approval secret on a loopback gateway base url", () => {
    process.env.GATEWAY_PUBLIC_BASE_URL = "http://127.0.0.1:8787";

    expect(getConfig()).toMatchObject({
      approvalSecret: "test-approval-secret",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      allowedCorsOrigins: [
        "https://docs.google.com",
        "https://excel.officeapps.live.com",
        "https://localhost:3000",
        "https://127.0.0.1:3000"
      ],
      saveInvalidHermesDebugArtifacts: false
    });
  });

  it("fails closed when the approval secret is missing even on a loopback gateway base url", () => {
    delete process.env.APPROVAL_SECRET;
    process.env.GATEWAY_PUBLIC_BASE_URL = "http://127.0.0.1:8787";

    expect(() => getConfig()).toThrow(
      "APPROVAL_SECRET must be configured before the gateway can approve writeback plans."
    );
  });

  it("fails closed when the approval secret is blank", () => {
    process.env.APPROVAL_SECRET = "   ";
    process.env.GATEWAY_PUBLIC_BASE_URL = "http://127.0.0.1:8787";

    expect(() => getConfig()).toThrow(
      "APPROVAL_SECRET must be configured before the gateway can approve writeback plans."
    );
  });

  it("fails closed when the default approval secret would be exposed on a non-local base url", () => {
    delete process.env.APPROVAL_SECRET;
    process.env.GATEWAY_PUBLIC_BASE_URL = "https://gateway.example.test/hermes-gateway";

    expect(() => getConfig()).toThrow(
      "APPROVAL_SECRET must be configured before the gateway can approve writeback plans."
    );
  });

  it("enables invalid Hermes debug artifact writes only when explicitly opted in", () => {
    process.env.HERMES_DEBUG_INVALID_RESPONSES = "true";

    expect(getConfig()).toMatchObject({
      saveInvalidHermesDebugArtifacts: true
    });
  });

  it("fails closed when the upload size limit is invalid", () => {
    for (const value of ["0", "-1", "not-a-number", "1.5"]) {
      process.env.MAX_UPLOAD_BYTES = value;

      expect(() => getConfig()).toThrow("MAX_UPLOAD_BYTES must be a positive integer.");
    }
  });

  it("defaults public gateway CORS to the gateway origin when no explicit allowlist is provided", () => {
    process.env.GATEWAY_PUBLIC_BASE_URL = "https://gateway.example.test/hermes-gateway";
    process.env.APPROVAL_SECRET = "secret";

    expect(getConfig()).toMatchObject({
      allowedCorsOrigins: ["https://gateway.example.test"]
    });
  });

  it("accepts an explicit comma-separated CORS allowlist", () => {
    process.env.GATEWAY_ALLOWED_ORIGINS =
      "https://docs.google.com, https://gateway.example.test , https://excel.officeapps.live.com";

    expect(getConfig()).toMatchObject({
      allowedCorsOrigins: [
        "https://docs.google.com",
        "https://gateway.example.test",
        "https://excel.officeapps.live.com"
      ]
    });
  });

  it("fails closed when the explicit CORS allowlist contains invalid origins", () => {
    process.env.GATEWAY_ALLOWED_ORIGINS = "https://docs.google.com,not-a-url";

    expect(() => getConfig()).toThrow(
      "GATEWAY_ALLOWED_ORIGINS contains an invalid origin: not-a-url."
    );
  });

  it("rejects wildcard CORS allowlists for public gateway base urls", () => {
    process.env.GATEWAY_PUBLIC_BASE_URL = "https://gateway.example.test/hermes-gateway";
    process.env.GATEWAY_ALLOWED_ORIGINS = "*";
    process.env.APPROVAL_SECRET = "secret";

    expect(() => getConfig()).toThrow(
      "GATEWAY_ALLOWED_ORIGINS must not contain * when the gateway public base URL is not local."
    );
  });
});
