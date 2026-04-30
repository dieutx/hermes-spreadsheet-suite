import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  allowPrivateNetworkPreflight,
  createApp,
  enforceAllowedOrigin,
  handleGatewayAppError,
  isCorsOriginAllowed
} from "../src/app.ts";

beforeEach(() => {
  process.env.APPROVAL_SECRET = "test-approval-secret";
});

afterEach(() => {
  delete process.env.APPROVAL_SECRET;
  delete process.env.HERMES_SERVICE_LABEL;
  delete process.env.HERMES_ENVIRONMENT_LABEL;
});

describe("gateway app", () => {
  it("allows private-network preflight only for allowed origins", () => {
    const req = {
      headers: {
        origin: "https://docs.google.com",
        "access-control-request-private-network": "true"
      }
    } as any;
    const res = {
      header: vi.fn()
    } as any;
    const next = vi.fn();

    allowPrivateNetworkPreflight(["https://docs.google.com"])(req, res, next);

    expect(res.header).toHaveBeenCalledWith("Access-Control-Allow-Private-Network", "true");
    expect(next).toHaveBeenCalledTimes(1);
  });

  it("rejects private-network preflight for origins outside the allowlist", () => {
    const req = {
      headers: {
        origin: "https://evil.example",
        "access-control-request-private-network": "true"
      }
    } as any;
    const res = {
      header: vi.fn()
    } as any;
    const next = vi.fn();

    allowPrivateNetworkPreflight(["https://docs.google.com"])(req, res, next);

    expect(res.header).not.toHaveBeenCalled();
    expect(next).toHaveBeenCalledTimes(1);
  });

  it("rejects API requests from disallowed browser origins", () => {
    const req = {
      headers: {
        origin: "https://evil.example"
      }
    } as any;
    let statusCode = 200;
    let jsonBody: unknown;
    const res = {
      status(code: number) {
        statusCode = code;
        return this;
      },
      json(payload: unknown) {
        jsonBody = payload;
        return this;
      }
    } as any;
    const next = vi.fn();

    enforceAllowedOrigin(["https://docs.google.com"])(req, res, next);

    expect(statusCode).toBe(403);
    expect(jsonBody).toEqual({
      error: {
        code: "ORIGIN_NOT_ALLOWED",
        message: "This Hermes gateway origin is not allowed.",
        userAction: "Open Hermes from an approved Excel or Google Sheets host, then retry."
      }
    });
    expect(next).not.toHaveBeenCalled();
  });

  it("allows API requests without a browser Origin header", () => {
    const req = {
      headers: {}
    } as any;
    const res = {
      status: vi.fn(),
      json: vi.fn()
    } as any;
    const next = vi.fn();

    enforceAllowedOrigin(["https://docs.google.com"])(req, res, next);

    expect(next).toHaveBeenCalledTimes(1);
    expect(res.status).not.toHaveBeenCalled();
  });

  it("mounts the execution control route surface under /api/execution", () => {
    const { app } = createApp();
    const executionLayer = (app as any)._router.stack.find((layer: any) =>
      layer.name === "router" && layer.regexp?.toString().includes("\\/api\\/execution")
    );

    expect(executionLayer).toBeDefined();
  });

  it("keeps public app metadata labels sanitized for health and proof responses", () => {
    process.env.HERMES_SERVICE_LABEL = "http://internal.example/gateway";
    process.env.HERMES_ENVIRONMENT_LABEL = "APPROVAL_SECRET=secret";

    const { config } = createApp();

    expect(config.serviceLabel).toBe("hermes-gateway-local");
    expect(config.environmentLabel).toBe("local-dev");
  });

  it("returns a JSON invalid-request error for malformed JSON bodies", () => {
    let statusCode = 200;
    let jsonBody: unknown;
    const res = {
      headersSent: false,
      status(code: number) {
        statusCode = code;
        return this;
      },
      json(payload: unknown) {
        jsonBody = payload;
        return this;
      }
    } as any;
    const next = vi.fn();
    const error = Object.assign(new SyntaxError("Unexpected token"), {
      type: "entity.parse.failed"
    });

    handleGatewayAppError(error, {} as any, res, next);

    expect(statusCode).toBe(400);
    expect(jsonBody).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Hermes couldn't read that request body.",
        userAction: "Retry the request. If it keeps failing, refresh the host and try again."
      }
    });
    expect(next).not.toHaveBeenCalled();
  });

  it("treats internal SyntaxError exceptions as gateway failures instead of malformed client JSON", () => {
    let statusCode = 200;
    let jsonBody: unknown;
    const res = {
      headersSent: false,
      status(code: number) {
        statusCode = code;
        return this;
      },
      json(payload: unknown) {
        jsonBody = payload;
        return this;
      }
    } as any;
    const next = vi.fn();
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});

    handleGatewayAppError(new SyntaxError("Unexpected token in internal template"), {} as any, res, next);

    expect(statusCode).toBe(500);
    expect(jsonBody).toEqual({
      error: {
        code: "INTERNAL_ERROR",
        message: "The gateway hit an unexpected error while processing the request.",
        userAction: "Retry the request. If it keeps failing, restart the gateway or check the server logs."
      }
    });
    expect(next).not.toHaveBeenCalled();
    expect(consoleError).toHaveBeenCalledOnce();
  });

  it("returns a JSON invalid-request error for oversized JSON bodies", () => {
    let statusCode = 200;
    let jsonBody: unknown;
    const res = {
      headersSent: false,
      status(code: number) {
        statusCode = code;
        return this;
      },
      json(payload: unknown) {
        jsonBody = payload;
        return this;
      }
    } as any;
    const next = vi.fn();

    handleGatewayAppError({ type: "entity.too.large" }, {} as any, res, next);

    expect(statusCode).toBe(413);
    expect(jsonBody).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "That request is too large for the gateway to process safely.",
        userAction: "Retry with a smaller prompt, fewer attachments, or a smaller spreadsheet selection."
      }
    });
    expect(next).not.toHaveBeenCalled();
  });

  it("allows all browser origins only for wildcard dev CORS configurations", () => {
    expect(isCorsOriginAllowed("https://docs.google.com", ["*"])).toBe(true);
    expect(isCorsOriginAllowed("https://docs.google.com", ["https://gateway.example.test"])).toBe(
      false
    );
  });

  it("allows requests without an Origin header for server-side callers", () => {
    expect(isCorsOriginAllowed(undefined, ["https://gateway.example.test"])).toBe(true);
  });
});
