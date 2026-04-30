import { Router } from "express";
import type { TraceBus } from "../lib/traceBus.js";

function parseAfterQuery(value: unknown): number | undefined {
  if (value === undefined) {
    return 0;
  }

  if (typeof value !== "string" || !/^(0|[1-9]\d*)$/.test(value.trim())) {
    return undefined;
  }

  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function getRequestIdQueryValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function getSessionIdQueryValue(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
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

    const run = input.traceBus.peekRun(req.params.runId);
    if (
      !run ||
      (run.requestId && getRequestIdQueryValue(req.query.requestId) !== run.requestId) ||
      (run.sessionId && getSessionIdQueryValue(req.query.sessionId) !== run.sessionId)
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

    const activeRun = input.traceBus.getRun(req.params.runId);
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

    const events = input.traceBus.list(req.params.runId, after);
    res.json({
      runId: req.params.runId,
      requestId: activeRun.requestId,
      hermesRunId: activeRun.hermesRunId,
      status: activeRun.status,
      nextIndex: activeRun.firstEventIndex + activeRun.events.length,
      events
    });
  });

  return router;
}
