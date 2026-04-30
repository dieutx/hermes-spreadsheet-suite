import { Router } from "express";
import type { TraceBus } from "../lib/traceBus.js";

const TRACE_ID_MAX_LENGTH = 128;

function parseAfterQuery(value: unknown): number | undefined {
  if (value === undefined) {
    return 0;
  }

  if (typeof value !== "string" || !/^(0|[1-9]\d*)$/.test(value.trim())) {
    return undefined;
  }

  return Number(value);
}

function parseRequiredTraceIdentifier(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const normalized = value.trim();
  if (normalized.length === 0 || normalized.length > TRACE_ID_MAX_LENGTH) {
    return undefined;
  }

  return normalized;
}

function parseOptionalTraceIdentifier(value: unknown): { ok: true; value?: string } | { ok: false } {
  if (value === undefined) {
    return { ok: true };
  }

  if (typeof value !== "string") {
    return { ok: false };
  }

  const normalized = value.trim();
  if (normalized.length === 0) {
    return { ok: true };
  }

  if (normalized.length > TRACE_ID_MAX_LENGTH) {
    return { ok: false };
  }

  return { ok: true, value: normalized };
}

function invalidTraceIdentifierRequest() {
  return {
    error: {
      code: "INVALID_REQUEST",
      message: "Trace identifiers are invalid.",
      userAction: "Retry the request trace from the current Hermes session."
    }
  };
}

export function createTraceRouter(input: {
  traceBus: TraceBus;
}): Router {
  const router = Router();

  router.get("/:runId", (req, res) => {
    const after = parseAfterQuery(req.query.after);
    if (after === undefined) {
      res.status(400).json({
        error: {
          code: "INVALID_REQUEST",
          message: "Trace cursor is invalid.",
          userAction: "Retry the request trace from the current Hermes session."
        }
      });
      return;
    }

    const runId = parseRequiredTraceIdentifier(req.params.runId);
    const requestId = parseOptionalTraceIdentifier(req.query.requestId);
    const sessionId = parseOptionalTraceIdentifier(req.query.sessionId);
    if (!runId || !requestId.ok || !sessionId.ok) {
      res.status(400).json(invalidTraceIdentifierRequest());
      return;
    }

    const run = input.traceBus.peekRun(runId);
    if (
      !run ||
      (run.requestId && requestId.value !== run.requestId) ||
      (run.sessionId && sessionId.value !== run.sessionId)
    ) {
      res.status(404).json({
        error: {
          code: "RUN_NOT_FOUND",
          message: "That Hermes trace is no longer available.",
          userAction: "Send the request again from the spreadsheet if you need a fresh trace."
        }
      });
      return;
    }

    const activeRun = input.traceBus.getRun(runId);
    if (!activeRun) {
      res.status(404).json({
        error: {
          code: "RUN_NOT_FOUND",
          message: "That Hermes trace is no longer available.",
          userAction: "Send the request again from the spreadsheet if you need a fresh trace."
        }
      });
      return;
    }

    const events = input.traceBus.list(runId, after);
    res.json({
      runId,
      requestId: activeRun.requestId,
      hermesRunId: activeRun.hermesRunId,
      status: activeRun.status,
      nextIndex: activeRun.firstEventIndex + activeRun.events.length,
      events
    });
  });

  return router;
}
