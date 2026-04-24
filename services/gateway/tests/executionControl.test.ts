import { describe, expect, it } from "vitest";
import { digestCanonicalPlan } from "../src/lib/approval.ts";
import { normalizeCompositePlanForDigest } from "../src/lib/planNormalization.ts";
import { createExecutionControlRouter } from "../src/routes/executionControl.ts";
import { ExecutionLedger } from "../src/lib/executionLedger.ts";

function invokeExecutionRoute(input: {
  path: "/dry-run" | "/history" | "/undo" | "/redo";
  method: "get" | "post";
  body?: Record<string, unknown>;
  query?: Record<string, unknown>;
}) {
  const router = createExecutionControlRouter({
    executionLedger: new ExecutionLedger()
  }) as any;
  const layer = router.stack.find((candidate: any) => candidate.route?.path === input.path);
  if (!layer) {
    throw new Error(`Expected route ${input.path} to exist.`);
  }

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
  };

  const handler = layer.route.stack.find((entry: any) => entry.method === input.method)?.handle;
  if (!handler) {
    throw new Error(`Expected ${input.method.toUpperCase()} handler for ${input.path}.`);
  }

  handler({ body: input.body ?? {}, query: input.query ?? {} }, res);

  return {
    status: statusCode,
    body: jsonBody
  };
}

function buildDryRunBody(overrides?: Partial<Record<string, unknown>>) {
  return {
    requestId: "req_dry_001",
    runId: "run_dry_001",
    workbookSessionKey: "excel_windows::workbook-123",
    plan: {
      steps: [
        {
          stepId: "step_1",
          dependsOn: [],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Report",
            position: "end",
            explanation: "Create a report sheet.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        }
      ],
      explanation: "Create a report artifact shell.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Report!A1"],
      overwriteRisk: "none",
      confirmationLevel: "standard",
      reversible: false,
      dryRunRecommended: true,
      dryRunRequired: false
    },
    ...overrides
  };
}

function invokeExecutionRouteWithLedger(input: {
  path: "/dry-run" | "/history" | "/undo" | "/redo";
  method: "get" | "post";
  ledger: ExecutionLedger;
  body?: Record<string, unknown>;
  query?: Record<string, unknown>;
}) {
  const router = createExecutionControlRouter({
    executionLedger: input.ledger
  }) as any;
  const layer = router.stack.find((candidate: any) => candidate.route?.path === input.path);
  if (!layer) {
    throw new Error(`Expected route ${input.path} to exist.`);
  }

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
  };

  const handler = layer.route.stack.find((entry: any) => entry.method === input.method)?.handle;
  if (!handler) {
    throw new Error(`Expected ${input.method.toUpperCase()} handler for ${input.path}.`);
  }

  handler({ body: input.body ?? {}, query: input.query ?? {} }, res);

  return {
    status: statusCode,
    body: jsonBody
  };
}

describe("execution control routes", () => {
  it("returns history for a workbook/session and rejects stale undo targets", () => {
    const history = invokeExecutionRoute({
      path: "/history",
      method: "get",
      query: { workbookSessionKey: "excel_windows::workbook-123" }
    });

    expect(history.status).toBe(200);
    expect(history.body).toMatchObject({ entries: [] });

    const undo = invokeExecutionRoute({
      path: "/undo",
      method: "post",
      body: {
        requestId: "req_undo_001",
        workbookSessionKey: "excel_windows::workbook-123",
        executionId: "missing_exec"
      }
    });

    expect(undo.status).toBe(409);
    expect(undo.body).toMatchObject({
      error: {
        code: "STALE_EXECUTION",
        message: "That history entry is no longer available for undo or redo."
      }
    });
  });

  it("rejects malformed undo requests with INVALID_REQUEST instead of stale-execution fallback", () => {
    const undo = invokeExecutionRoute({
      path: "/undo",
      method: "post",
      body: {
        requestId: "req_undo_001",
        workbookSessionKey: "excel_windows::workbook-123"
      }
    });

    expect(undo.status).toBe(400);
    expect(undo.body).toMatchObject({
      error: {
        code: "INVALID_REQUEST"
      }
    });
  });

  it("rejects invalid history cursors with INVALID_REQUEST", () => {
    const history = invokeExecutionRoute({
      path: "/history",
      method: "get",
      query: {
        workbookSessionKey: "excel_windows::workbook-123",
        cursor: "-1"
      }
    });

    expect(history.status).toBe(400);
    expect(history.body).toMatchObject({
      error: {
        code: "INVALID_REQUEST",
        message: "That execution-control request is invalid."
      }
    });
  });

  it("returns INTERNAL_ERROR when the ledger throws an unexpected runtime failure", () => {
    const ledger = {
      listHistory() {
        return { entries: [] };
      },
      storeDryRun() {
        return undefined;
      },
      undoExecution() {
        throw new Error("boom");
      },
      redoExecution() {
        throw new Error("boom");
      }
    } as unknown as ExecutionLedger;

    const undo = invokeExecutionRouteWithLedger({
      path: "/undo",
      method: "post",
      ledger,
      body: {
        requestId: "req_undo_001",
        workbookSessionKey: "excel_windows::workbook-123",
        executionId: "exec_001"
      }
    });

    expect(undo.status).toBe(500);
    expect(undo.body).toMatchObject({
      error: {
        code: "INTERNAL_ERROR",
        message: "The gateway couldn't complete that execution-control request."
      }
    });
  });

  it("rejects dry-run requests that would overflow the DryRunResult contract", () => {
    const dryRun = invokeExecutionRoute({
      path: "/dry-run",
      method: "post",
      body: buildDryRunBody({
        requestId: "r".repeat(129)
      })
    });

    expect(dryRun.status).toBe(400);
    expect(dryRun.body).toMatchObject({
      error: {
        code: "INVALID_REQUEST"
      }
    });
  });

  it("returns INTERNAL_ERROR when post-parse dry-run or history handling fails", () => {
    const ledger = {
      isoTimestamp() {
        return "2026-04-20T10:05:00.000Z";
      },
      listHistory() {
        throw new Error("history boom");
      },
      storeDryRun() {
        throw new Error("dry-run boom");
      },
      undoExecution() {
        throw new Error("undo boom");
      },
      redoExecution() {
        throw new Error("redo boom");
      }
    } as unknown as ExecutionLedger;

    const dryRun = invokeExecutionRouteWithLedger({
      path: "/dry-run",
      method: "post",
      ledger,
      body: buildDryRunBody()
    });
    expect(dryRun.status).toBe(500);
    expect(dryRun.body).toMatchObject({
      error: {
        code: "INTERNAL_ERROR",
        message: "The gateway couldn't complete that execution-control request."
      }
    });

    const history = invokeExecutionRouteWithLedger({
      path: "/history",
      method: "get",
      ledger,
      query: {
        workbookSessionKey: "excel_windows::workbook-123"
      }
    });
    expect(history.status).toBe(500);
    expect(history.body).toMatchObject({
      error: {
        code: "INTERNAL_ERROR",
        message: "The gateway couldn't complete that execution-control request."
      }
    });
  });

  it("uses the injected ledger clock when stamping dry-run expiry", () => {
    const nowMs = Date.parse("2026-04-20T10:00:00.000Z");
    const ledger = new ExecutionLedger({
      now: () => nowMs
    });

    const dryRun = invokeExecutionRouteWithLedger({
      path: "/dry-run",
      method: "post",
      ledger,
      body: buildDryRunBody()
    });

    expect(dryRun.status).toBe(200);
    expect(dryRun.body).toMatchObject({
      reversible: false,
      expiresAt: "2026-04-20T10:05:00.000Z"
    });
  });

  it("rejects undo requests that target an execution from another workbook/session", () => {
    const ledger = new ExecutionLedger();
    ledger.recordCompleted({
      executionId: "exec_001",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_001",
      runId: "run_001",
      planType: "sheet_update",
      planDigest: "digest_001",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated Sales!A1:B2."
    });

    const undo = invokeExecutionRouteWithLedger({
      path: "/undo",
      method: "post",
      ledger,
      body: {
        requestId: "req_undo_001",
        workbookSessionKey: "excel_windows::other-workbook",
        executionId: "exec_001"
      }
    });

    expect(undo.status).toBe(409);
    expect(undo.body).toMatchObject({
      error: {
        code: "STALE_EXECUTION"
      }
    });
  });

  it("stores dry-run results under the canonical composite digest", () => {
    const dryRun = invokeExecutionRoute({
      path: "/dry-run",
      method: "post",
      body: {
        requestId: "req_dry_001",
        runId: "run_dry_001",
        workbookSessionKey: "excel_windows::workbook-123",
        plan: {
          dryRunRequired: false,
          reversible: false,
          confirmationLevel: "standard",
          overwriteRisk: "none",
          affectedRanges: ["Report!A1"],
          requiresConfirmation: true,
          confidence: 0.9,
          explanation: "Create a report artifact shell.",
          steps: [
            {
              plan: {
                position: "end",
                sheetName: "Report",
                confidence: 0.9,
                explanation: "Create a report sheet.",
                requiresConfirmation: true,
                operation: "create_sheet"
              },
              continueOnError: false,
              dependsOn: [],
              stepId: "step_1"
            }
          ],
          dryRunRecommended: true
        }
      }
    });

    expect(dryRun.status).toBe(200);
    expect(dryRun.body).toMatchObject({
      simulated: true,
      reversible: false,
      planDigest: digestCanonicalPlan({
        dryRunRequired: false,
        reversible: false,
        confirmationLevel: "standard",
        overwriteRisk: "none",
        affectedRanges: ["Report!A1"],
        requiresConfirmation: true,
        confidence: 0.9,
        explanation: "Create a report artifact shell.",
        steps: [
          {
            plan: {
              position: "end",
              sheetName: "Report",
              confidence: 0.9,
              explanation: "Create a report sheet.",
              requiresConfirmation: true,
              operation: "create_sheet"
            },
            continueOnError: false,
            dependsOn: [],
            stepId: "step_1"
          }
        ],
        dryRunRecommended: true
      })
    });
  });

  it("normalizes composite dry-run digests when the incoming plan still claims reversible execution", () => {
    const inputPlan = {
      dryRunRequired: true,
      reversible: true,
      confirmationLevel: "standard" as const,
      overwriteRisk: "none" as const,
      affectedRanges: ["Report!A1"],
      requiresConfirmation: true,
      confidence: 0.9,
      explanation: "Create a report artifact shell.",
      steps: [
        {
          plan: {
            position: "end" as const,
            sheetName: "Report",
            confidence: 0.9,
            explanation: "Create a report sheet.",
            requiresConfirmation: true,
            operation: "create_sheet" as const
          },
          continueOnError: false,
          dependsOn: [],
          stepId: "step_1"
        }
      ],
      dryRunRecommended: true
    };

    const dryRun = invokeExecutionRoute({
      path: "/dry-run",
      method: "post",
      body: buildDryRunBody({
        plan: inputPlan
      })
    });

    expect(dryRun.status).toBe(200);
    expect(dryRun.body).toMatchObject({
      reversible: false,
      planDigest: digestCanonicalPlan(normalizeCompositePlanForDigest(inputPlan))
    });
  });

  it("truncates dry-run predicted summaries to the public contract limit", () => {
    const longExplanation = "A".repeat(5001);
    const dryRun = invokeExecutionRoute({
      path: "/dry-run",
      method: "post",
      body: buildDryRunBody({
        plan: {
          ...buildDryRunBody().plan,
          explanation: longExplanation,
          steps: [
            {
              stepId: "step_1",
              dependsOn: [],
              continueOnError: false,
              plan: {
                operation: "create_sheet",
                sheetName: "Report",
                position: "end",
                explanation: longExplanation,
                confidence: 0.9,
                requiresConfirmation: true
              }
            }
          ]
        }
      })
    });

    expect(dryRun.status).toBe(200);
    expect((dryRun.body as any).predictedSummaries[0]).toHaveLength(4000);
    expect((dryRun.body as any).steps[0].predictedSummaries[0]).toHaveLength(4000);
  });

  it("rejects composite dry-run plans whose dependencies appear after the step that depends on them", () => {
    const dryRun = invokeExecutionRoute({
      path: "/dry-run",
      method: "post",
      body: {
        ...buildDryRunBody(),
        plan: {
          ...buildDryRunBody().plan,
          steps: [
            {
              stepId: "step_2",
              dependsOn: ["step_1"],
              continueOnError: false,
              plan: {
                operation: "create_sheet",
                sheetName: "Stage2",
                position: "end",
                explanation: "Create stage 2.",
                confidence: 0.9,
                requiresConfirmation: true
              }
            },
            {
              stepId: "step_1",
              dependsOn: [],
              continueOnError: false,
              plan: {
                operation: "create_sheet",
                sheetName: "Stage1",
                position: "end",
                explanation: "Create stage 1.",
                confidence: 0.9,
                requiresConfirmation: true
              }
            }
          ]
        }
      }
    });

    expect(dryRun.status).toBe(400);
    expect(dryRun.body).toMatchObject({
      error: {
        code: "INVALID_REQUEST"
      }
    });
  });

  it("returns undo and redo control results for eligible reversible lineage points", () => {
    const ledger = new ExecutionLedger();
    ledger.recordCompleted({
      executionId: "exec_001",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_001",
      runId: "run_001",
      planType: "sheet_update",
      planDigest: "digest_001",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated Sales!A1:B2."
    });

    const undo = invokeExecutionRouteWithLedger({
      path: "/undo",
      method: "post",
      ledger,
      body: {
        requestId: "req_undo_001",
        workbookSessionKey: "excel_windows::workbook-123",
        executionId: "exec_001"
      }
    });

    expect(undo.status).toBe(200);
    expect(undo.body).toMatchObject({
      operation: "composite_update",
      summary: "Undid execution exec_001."
    });

    const undoExecutionId = (undo.body as any).executionId;
    const redo = invokeExecutionRouteWithLedger({
      path: "/redo",
      method: "post",
      ledger,
      body: {
        requestId: "req_redo_001",
        workbookSessionKey: "excel_windows::workbook-123",
        executionId: undoExecutionId
      }
    });

    expect(redo.status).toBe(200);
    expect(redo.body).toMatchObject({
      operation: "composite_update",
      summary: `Redid execution ${undoExecutionId}.`
    });
  });
});
