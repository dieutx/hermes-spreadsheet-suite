import express from "express";
import { describe, expect, it } from "vitest";
import { TraceBus } from "../src/lib/traceBus.ts";
import { createTraceRouter } from "../src/routes/trace.ts";

async function invokeGet(
  router: express.Router,
  runId: string,
  query: Record<string, unknown> = {}
) {
  const layer = router.stack.find((entry) =>
    entry.route?.path === "/:runId" && entry.route.methods.get
  );

  if (!layer) {
    throw new Error("GET /:runId route not found.");
  }

  let statusCode = 200;
  let jsonBody: unknown;

  const req = {
    params: { runId },
    query
  } as unknown as express.Request;
  const res = {
    status(code: number) {
      statusCode = code;
      return this;
    },
    json(payload: unknown) {
      jsonBody = payload;
      return this;
    }
  } as unknown as express.Response;

  await Promise.resolve(layer.route.stack[0]?.handle(req, res));

  return { statusCode, body: jsonBody };
}

describe("trace router", () => {
  it("requires the matching requestId to read a run trace", async () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_trace_001", "req_trace_001");
    traceBus.append("run_trace_001", {
      event: "request_received",
      timestamp: "2026-04-22T00:00:00.000Z"
    });
    const router = createTraceRouter({ traceBus });

    const missingRequestId = await invokeGet(router, "run_trace_001");
    expect(missingRequestId.statusCode).toBe(404);
    expect(missingRequestId.body).toEqual({
      error: {
        code: "RUN_NOT_FOUND",
        message: "That Hermes trace is no longer available.",
        userAction: "Send the request again from the spreadsheet if you need a fresh trace."
      }
    });

    const validRequestId = await invokeGet(router, "run_trace_001", {
      requestId: "req_trace_001"
    });
    expect(validRequestId.statusCode).toBe(200);
    expect((validRequestId.body as any).events).toHaveLength(1);
  });

  it("requires the matching sessionId when the run is session-scoped", async () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_trace_session_001", "req_trace_session_001");
    (traceBus.peekRun("run_trace_session_001") as any).sessionId = "sess_trace_001";
    traceBus.append("run_trace_session_001", {
      event: "request_received",
      timestamp: "2026-04-22T00:00:00.000Z"
    });
    const router = createTraceRouter({ traceBus });

    const missingSessionId = await invokeGet(router, "run_trace_session_001", {
      requestId: "req_trace_session_001"
    });
    expect(missingSessionId.statusCode).toBe(404);

    const wrongSessionId = await invokeGet(router, "run_trace_session_001", {
      requestId: "req_trace_session_001",
      sessionId: "sess_other"
    });
    expect(wrongSessionId.statusCode).toBe(404);

    const matchingSessionId = await invokeGet(router, "run_trace_session_001", {
      requestId: "req_trace_session_001",
      sessionId: "sess_trace_001"
    });
    expect(matchingSessionId.statusCode).toBe(200);
    expect((matchingSessionId.body as any).events).toHaveLength(1);
  });

  it("rejects oversized trace route identifiers", async () => {
    const traceBus = new TraceBus();
    const router = createTraceRouter({ traceBus });

    const response = await invokeGet(router, `run_${"x".repeat(256)}`, {
      requestId: "req_trace_oversized"
    });

    expect(response.statusCode).toBe(400);
    expect(response.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Trace identifiers are invalid.",
        userAction: "Retry the request trace from the current Hermes session."
      }
    });
  });

  it("rejects oversized trace identity query values", async () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_trace_query_bounds", "req_trace_query_bounds");
    (traceBus.peekRun("run_trace_query_bounds") as any).sessionId = "sess_trace_query_bounds";
    const router = createTraceRouter({ traceBus });

    const response = await invokeGet(router, "run_trace_query_bounds", {
      requestId: `req_${"x".repeat(256)}`,
      sessionId: `sess_${"x".repeat(256)}`
    });

    expect(response.statusCode).toBe(400);
    expect(response.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Trace identifiers are invalid.",
        userAction: "Retry the request trace from the current Hermes session."
      }
    });
  });

  it("rejects malformed trace cursors instead of coercing them through Number()", async () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_trace_002", "req_trace_002");
    const router = createTraceRouter({ traceBus });

    const response = await invokeGet(router, "run_trace_002", {
      requestId: "req_trace_002",
      after: "-1"
    });

    expect(response.statusCode).toBe(400);
    expect(response.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Trace cursor is invalid.",
        userAction: "Retry the request trace from the current Hermes session."
      }
    });
  });

  it("returns the shifted nextIndex when older trace events have been trimmed", async () => {
    const traceBus = new TraceBus({ maxEventsPerRun: 2 });
    traceBus.ensureRun("run_trace_003", "req_trace_003");
    traceBus.append("run_trace_003", {
      event: "request_received",
      timestamp: "2026-04-22T00:00:00.000Z"
    });
    traceBus.append("run_trace_003", {
      event: "result_generated",
      timestamp: "2026-04-22T00:00:01.000Z"
    });
    traceBus.append("run_trace_003", {
      event: "completed",
      timestamp: "2026-04-22T00:00:02.000Z"
    });
    const router = createTraceRouter({ traceBus });

    const response = await invokeGet(router, "run_trace_003", {
      requestId: "req_trace_003",
      after: "0"
    });

    expect(response.statusCode).toBe(200);
    expect((response.body as any).events).toHaveLength(2);
    expect((response.body as any).nextIndex).toBe(3);
  });
});
