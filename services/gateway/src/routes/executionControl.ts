import { Router } from "express";
import { z } from "zod";
import {
  CompositePlanDataSchema,
  RedoRequestSchema,
  UndoRequestSchema
} from "@hermes/contracts";
import type { GatewayConfig } from "../lib/config.js";
import { digestCanonicalPlan } from "../lib/approval.js";
import {
  StaleExecutionError,
  UnsupportedExecutionControlError,
  type ExecutionLedger
} from "../lib/executionLedger.js";
import { normalizeCompositePlanForDigest } from "../lib/planNormalization.js";

const DryRunRequestSchema = z.object({
  requestId: z.string().min(1).max(128),
  runId: z.string().min(1).max(128),
  workbookSessionKey: z.string().min(1).max(256).optional(),
  plan: CompositePlanDataSchema
});

const HistoryQuerySchema = z.object({
  workbookSessionKey: z.string().min(1).max(256),
  cursor: z.string().min(1).max(256).regex(/^(0|[1-9]\d*)$/).optional(),
  limit: z.coerce.number().int().positive().max(100).optional()
});

type RouteErrorPayload = {
  error: {
    code: string;
    message: string;
    userAction?: string;
    issues?: Array<{ path: string; message: string }>;
  };
};

function formatIssues(issues: z.ZodIssue[]): Array<{ path: string; message: string }> {
  return issues.map((issue) => ({
    path: issue.path.join("."),
    message: issue.message
  }));
}

function invalidExecutionRequest(
  message: string,
  userAction: string,
  issues?: z.ZodIssue[]
): RouteErrorPayload {
  return {
    error: {
      code: "INVALID_REQUEST",
      message,
      userAction,
      ...(issues ? { issues: formatIssues(issues) } : {})
    }
  };
}

function internalExecutionError(message: string, userAction: string): RouteErrorPayload {
  return {
    error: {
      code: "INTERNAL_ERROR",
      message,
      userAction
    }
  };
}

function summarizeCompositeDryRunStep(step: z.infer<typeof CompositePlanDataSchema>["steps"][number]): string {
  if ("explanation" in step.plan && typeof step.plan.explanation === "string") {
    return step.plan.explanation;
  }

  if ("operation" in step.plan && typeof step.plan.operation === "string") {
    return `Would execute ${step.plan.operation}.`;
  }

  return `Would execute ${step.stepId}.`;
}

function toPredictedSummaries(
  ...values: Array<string | undefined>
): string[] | undefined {
  const summaries = values
    .filter((value): value is string => typeof value === "string" && value.trim().length > 0)
    .map((value) => value.length > 4000 ? value.slice(0, 4000) : value);

  return summaries.length > 0 ? summaries : undefined;
}

function formatExecutionControlError(error: unknown): {
  status: number;
  body: RouteErrorPayload;
} {
  if (error instanceof z.ZodError) {
    return {
      status: 400,
      body: invalidExecutionRequest(
        "That execution-control request is invalid.",
        "Refresh the spreadsheet session and retry the preview, history, undo, or redo action.",
        error.issues
      )
    };
  }

  if (error instanceof UnsupportedExecutionControlError) {
    return {
      status: 501,
      body: {
        error: {
          code: "UNSUPPORTED_OPERATION",
          message: "That undo or redo action is not supported here yet.",
          userAction: "Refresh the sheet history and ask Hermes to prepare a fresh update instead."
        }
      }
    };
  }

  if (error instanceof StaleExecutionError) {
    return {
      status: 409,
      body: {
        error: {
          code: "STALE_EXECUTION",
          message: "That history entry is no longer available for undo or redo.",
          userAction: "Refresh plan history and retry from the latest execution."
        }
      }
    };
  }

  return {
    status: 500,
    body: internalExecutionError(
      "The gateway couldn't complete that execution-control request.",
      "Retry the action. If it keeps failing, refresh the spreadsheet session and try again."
    )
  };
}

export function createExecutionControlRouter(input: {
  executionLedger: ExecutionLedger;
  config?: GatewayConfig;
}) {
  const router = Router();

  router.post("/dry-run", (req, res) => {
    try {
      const parsed = DryRunRequestSchema.parse(req.body);
      const normalizedPlan = normalizeCompositePlanForDigest(parsed.plan);
      const planDigest = digestCanonicalPlan(normalizedPlan);
      const workbookSessionKey = parsed.workbookSessionKey ?? `run::${parsed.runId}`;
      const result = {
        planDigest,
        workbookSessionKey,
        simulated: true,
        steps: normalizedPlan.steps.map((step) => ({
          stepId: step.stepId,
          status: "simulated" as const,
          summary: summarizeCompositeDryRunStep(step),
          predictedAffectedRanges: "affectedRanges" in step.plan ? step.plan.affectedRanges : undefined,
          predictedSummaries: "explanation" in step.plan && typeof step.plan.explanation === "string"
            ? toPredictedSummaries(step.plan.explanation)
            : undefined
        })),
        predictedAffectedRanges: normalizedPlan.affectedRanges,
        predictedSummaries: toPredictedSummaries(normalizedPlan.explanation),
        overwriteRisk: normalizedPlan.overwriteRisk,
        reversible: normalizedPlan.reversible,
        expiresAt: input.executionLedger.isoTimestamp(5 * 60 * 1000)
      };

      input.executionLedger.storeDryRun(result);
      res.json(result);
    } catch (error) {
      const formatted = formatExecutionControlError(error);
      res.status(formatted.status).json(formatted.body);
    }
  });

  router.get("/history", (req, res) => {
    try {
      const parsed = HistoryQuerySchema.parse(req.query);
      res.json(input.executionLedger.listHistory(
        parsed.workbookSessionKey,
        parsed.limit,
        parsed.cursor
      ));
    } catch (error) {
      const formatted = formatExecutionControlError(error);
      res.status(formatted.status).json(formatted.body);
    }
  });

  router.post("/undo", (req, res) => {
    try {
      const parsed = UndoRequestSchema.parse(req.body);
      res.json(input.executionLedger.undoExecution(parsed));
    } catch (error) {
      const formatted = formatExecutionControlError(error);
      res.status(formatted.status).json(formatted.body);
    }
  });

  router.post("/redo", (req, res) => {
    try {
      const parsed = RedoRequestSchema.parse(req.body);
      res.json(input.executionLedger.redoExecution(parsed));
    } catch (error) {
      const formatted = formatExecutionControlError(error);
      res.status(formatted.status).json(formatted.body);
    }
  });

  return router;
}
