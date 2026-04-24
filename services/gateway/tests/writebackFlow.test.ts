import { describe, expect, it, vi } from "vitest";
import { ExecutionLedger } from "../src/lib/executionLedger.ts";
import { TraceBus } from "../src/lib/traceBus.ts";
import {
  approveWriteback,
  completeWriteback,
  createWritebackRouter
} from "../src/routes/writeback.ts";
import { createApprovalToken, digestCanonicalPlan, digestPlan } from "../src/lib/approval.ts";
import { normalizeCompositePlanForDigest } from "../src/lib/planNormalization.ts";

const testConfig = {
  port: 8787,
  environmentLabel: "review",
  serviceLabel: "hermes-remote-us-1",
  gatewayPublicBaseUrl: "http://127.0.0.1:8787",
  maxUploadBytes: 1000,
  approvalSecret: "secret",
  hermesMode: "mock" as const,
  hermesBaseUrl: undefined,
  skillRegistryPath: ""
};

function setRunResponse(
  traceBus: TraceBus,
  input: {
    runId: string;
    requestId: string;
    type: string;
    traceEvent: string;
    plan: Record<string, unknown>;
  }
) {
  traceBus.ensureRun(input.runId, input.requestId);
  const run = traceBus.getRun(input.runId);
  if (!run) {
    throw new Error("Expected run to exist.");
  }

  run.response = {
    schemaVersion: "1.0.0",
    type: input.type,
    requestId: input.requestId,
    hermesRunId: input.runId,
    processedBy: "hermes",
    serviceLabel: "hermes-remote-us-1",
    environmentLabel: "review",
    startedAt: "2026-04-19T09:30:00.000Z",
    completedAt: "2026-04-19T09:30:01.000Z",
    durationMs: 1000,
    trace: [
      {
        event: input.traceEvent,
        timestamp: "2026-04-19T09:30:01.000Z"
      }
    ],
    ui: {
      displayMode: "structured-preview",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: true
    },
    data: input.plan
  } as any;
}

function invokeWritebackRoute(input: {
  traceBus: TraceBus;
  executionLedger?: ExecutionLedger;
  path: "/approve" | "/complete";
  body: Record<string, unknown>;
}) {
  const router = createWritebackRouter({
    traceBus: input.traceBus,
    executionLedger: input.executionLedger ?? new ExecutionLedger(),
    config: testConfig
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

  layer.route.stack[0]?.handle({ body: input.body }, res);

  return {
    status: statusCode,
    body: jsonBody
  };
}

function expectRouteError(
  response: { status: number; body: unknown },
  status: number,
  code: string,
  message?: string
) {
  expect(response.status).toBe(status);
  expect(response.body).toMatchObject({
    error: {
      code,
      ...(message ? { message } : {})
    }
  });
}

function expectRouteCompletionDetailMismatch(input: {
  traceBus: TraceBus;
  runId: string;
  requestId: string;
  type: string;
  traceEvent: string;
  plan: Record<string, unknown>;
  result: Record<string, unknown>;
  destructiveConfirmation?: {
    confirmed: true;
  };
}) {
  setRunResponse(input.traceBus, {
    runId: input.runId,
    requestId: input.requestId,
    type: input.type,
    traceEvent: input.traceEvent,
    plan: input.plan
  });

  const approval = invokeWritebackRoute({
    traceBus: input.traceBus,
    path: "/approve",
    body: {
      requestId: input.requestId,
      runId: input.runId,
      plan: input.plan,
      destructiveConfirmation: input.destructiveConfirmation
    }
  });
  expect(approval.status).toBe(200);

  const completion = invokeWritebackRoute({
    traceBus: input.traceBus,
    path: "/complete",
    body: {
      requestId: input.requestId,
      runId: input.runId,
      approvalToken: (approval.body as any).approvalToken,
      planDigest: (approval.body as any).planDigest,
      result: input.result
    }
  });

  expectRouteError(
    completion,
    409,
    "STALE_APPROVAL",
    "The approved update no longer matches the current Hermes plan."
  );
}

function buildRangeWriteResult(
  plan: Record<string, any>,
  overrides: Record<string, unknown> = {}
) {
  return {
    kind: "range_write",
    hostPlatform: "excel_windows",
    ...plan,
    writtenRows: plan.shape.rows,
    writtenColumns: plan.shape.columns,
    ...overrides
  };
}

function buildNamedRangeUpdateResult(
  plan: Record<string, any>,
  overrides: Record<string, unknown> = {}
) {
  return {
    kind: "named_range_update",
    hostPlatform: "google_sheets",
    ...plan,
    summary: "Completed named range update.",
    ...overrides
  };
}

describe("writeback confirmation flow", () => {
  it("rejects composite approval when a required dry-run is missing or stale", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      steps: [
        {
          stepId: "cleanup",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:F50",
            operation: "remove_duplicate_rows",
            explanation: "Dedupe rows.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50"],
            overwriteRisk: "medium" as const,
            confirmationLevel: "destructive" as const
          }
        }
      ],
      explanation: "Run cleanup workflow.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "medium" as const,
      confirmationLevel: "destructive" as const,
      reversible: false,
      dryRunRecommended: true,
      dryRunRequired: true
    };

    setRunResponse(traceBus, {
      runId: "run_composite_approve_001",
      requestId: "req_composite_approve_001",
      type: "composite_plan",
      traceEvent: "composite_plan_ready",
      plan
    });

    const response = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_composite_approve_001",
        runId: "run_composite_approve_001",
        workbookSessionKey: "excel_windows::workbook-123",
        destructiveConfirmation: { confirmed: true },
        plan
      }
    });

    expectRouteError(
      response,
      409,
      "STALE_PREVIEW",
      "That workflow preview is stale and must be regenerated before approval."
    );
  });

  it("accepts composite approval after a fresh dry-run even when the original plan claims reversible execution", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      steps: [
        {
          stepId: "setup",
          dependsOn: [],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Report",
            position: "end" as const,
            explanation: "Create a report sheet.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        }
      ],
      explanation: "Create the report sheet.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Report!A1"],
      overwriteRisk: "none" as const,
      confirmationLevel: "standard" as const,
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: true
    };

    executionLedger.storeDryRun({
      planDigest: digestCanonicalPlan(normalizeCompositePlanForDigest(plan)),
      workbookSessionKey: "excel_windows::workbook-123",
      simulated: true,
      steps: [
        {
          stepId: "setup",
          status: "simulated",
          summary: "Would create Report."
        }
      ],
      predictedAffectedRanges: ["Report!A1"],
      predictedSummaries: ["Would create the report sheet."],
      overwriteRisk: "none",
      reversible: false,
      expiresAt: "2099-01-01T00:00:00.000Z"
    });

    setRunResponse(traceBus, {
      runId: "run_composite_approve_002",
      requestId: "req_composite_approve_002",
      type: "composite_plan",
      traceEvent: "composite_plan_ready",
      plan
    });

    const response = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_composite_approve_002",
        runId: "run_composite_approve_002",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });

    expect(response.status).toBe(200);
    expect(response.body).toMatchObject({
      planDigest: digestCanonicalPlan(normalizeCompositePlanForDigest(plan))
    });
  });

  it("records composite completion with step-level status history", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      steps: [
        {
          stepId: "sort",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:F50",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "desc" as const }],
            explanation: "Sort by revenue descending.",
            confidence: 0.91,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50"]
          }
        },
        {
          stepId: "report",
          dependsOn: ["sort"],
          continueOnError: false,
          plan: {
            sourceSheet: "Sales",
            sourceRange: "A1:F50",
            outputMode: "materialize_report" as const,
            targetSheet: "Sales Report",
            targetRange: "A1",
            sections: [
              {
                type: "summary_stats",
                title: "Revenue summary",
                summary: "Top revenue records summarized.",
                sourceRanges: ["Sales!A1:F50"]
              }
            ],
            explanation: "Create a report.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D5"],
            overwriteRisk: "low" as const,
            confirmationLevel: "standard" as const
          }
        }
      ],
      explanation: "Sort then report.",
      confidence: 0.92,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D5"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const,
      reversible: false,
      dryRunRecommended: true,
      dryRunRequired: false
    };

    executionLedger.storeDryRun({
      planDigest: digestPlan(plan),
      workbookSessionKey: "excel_windows::workbook-123",
      simulated: true,
      steps: [
        {
          stepId: "sort",
          status: "simulated",
          summary: "Would sort Sales!A1:F50."
        },
        {
          stepId: "report",
          status: "simulated",
          summary: "Would build Sales Report!A1:D5."
        }
      ],
      predictedAffectedRanges: ["Sales!A1:F50", "Sales Report!A1:D5"],
      predictedSummaries: ["Would run the composite workflow."],
      overwriteRisk: "low",
      reversible: false,
      expiresAt: "2099-01-01T00:00:00.000Z"
    });

    setRunResponse(traceBus, {
      runId: "run_composite_complete_001",
      requestId: "req_composite_complete_001",
      type: "composite_plan",
      traceEvent: "composite_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_composite_complete_001",
        runId: "run_composite_complete_001",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_composite_complete_001",
        runId: "run_composite_complete_001",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "composite_update",
          operation: "composite_update",
          hostPlatform: "excel_windows",
          executionId: (approval.body as any).executionId,
          stepResults: [
            {
              stepId: "sort",
              status: "completed",
              summary: "Sorted Sales!A1:F50."
            },
            {
              stepId: "report",
              status: "completed",
              summary: "Created Sales Report!A1:D5."
            }
          ],
          summary: "Completed the composite workflow."
        }
      }
    });
    expect(completion.status).toBe(200);

    expect(
      executionLedger.listHistory("excel_windows::workbook-123").entries[0]
    ).toMatchObject({
      executionId: (approval.body as any).executionId,
      requestId: "req_composite_complete_001",
      runId: "run_composite_complete_001",
      planType: "composite_plan",
      status: "completed",
      summary: "Completed the composite workflow.",
      stepEntries: [
        {
          stepId: "sort",
          status: "completed",
          summary: "Sorted Sales!A1:F50."
        },
        {
          stepId: "report",
          status: "completed",
          summary: "Created Sales Report!A1:D5."
        }
      ]
    });
  });

  it("does not treat chat-only analysis reports as writeback eligible", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "chat_only" as const,
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Summarize the selected sales range.",
      confidence: 0.92,
      requiresConfirmation: false as const,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "none" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_analysis_chat_only",
      requestId: "req_analysis_chat_only",
      type: "analysis_report_plan",
      traceEvent: "analysis_report_plan_ready",
      plan
    });

    const response = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_analysis_chat_only",
        runId: "run_analysis_chat_only",
        plan
      }
    });

    expect(response.status).toBe(400);
    expect(response.body).toMatchObject({
      error: {
        code: "INVALID_REQUEST"
      }
    });
  });

  it("approves a preview plan and records the confirmed write application", async () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_1", "req_1");
    const config = {
      port: 8787,
      environmentLabel: "review",
      serviceLabel: "hermes-remote-us-1",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      maxUploadBytes: 1000,
      approvalSecret: "secret",
      hermesMode: "mock" as const,
      hermesBaseUrl: undefined,
      skillRegistryPath: ""
    };

    const plan = {
      sourceAttachmentId: "att_1",
      targetSheet: "Sheet3",
      targetRange: "B4:C5",
      headers: ["Name", "Qty"],
      values: [["Widget", 2]],
      confidence: 0.94,
      warnings: [],
      requiresConfirmation: true,
      extractionMode: "real" as const,
      shape: {
        rows: 2,
        columns: 2
      }
    };
    const run = traceBus.getRun("run_1");
    if (!run) {
      throw new Error("Expected run to exist.");
    }
    run.response = {
      schemaVersion: "1.0.0",
      type: "sheet_import_plan",
      requestId: "req_1",
      hermesRunId: "run_1",
      processedBy: "hermes",
      serviceLabel: "hermes-remote-us-1",
      environmentLabel: "review",
      startedAt: "2026-04-19T09:30:00.000Z",
      completedAt: "2026-04-19T09:30:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "sheet_import_plan_ready",
          timestamp: "2026-04-19T09:30:01.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: plan
    };

    const approval = approveWriteback({
      requestId: "req_1",
      runId: "run_1",
      plan,
      traceBus,
      config
    });

    expect(approval.requestId).toBe("req_1");
    expect(approval.runId).toBe("run_1");

    const completion = completeWriteback({
      requestId: "req_1",
      runId: "run_1",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildRangeWriteResult(plan),
      traceBus,
      config
    });
    expect(completion.ok).toBe(true);
    expect(traceBus.getRun("run_1")?.writeback).toMatchObject({
      approvedPlanDigest: approval.planDigest,
      approvalToken: approval.approvalToken,
      completedPlanDigest: approval.planDigest,
      result: {
        hostPlatform: "excel_windows",
        targetSheet: "Sheet3",
        targetRange: "B4:C5",
        writtenRows: 2,
        writtenColumns: 2
      }
    });
  });

  it("rejects writeback completion replay for the same approval token", () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_replay", "req_replay");
    const config = {
      port: 8787,
      environmentLabel: "review",
      serviceLabel: "hermes-remote-us-1",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      maxUploadBytes: 1000,
      approvalSecret: "secret",
      hermesMode: "mock" as const,
      hermesBaseUrl: undefined,
      skillRegistryPath: ""
    };

    const plan = {
      sourceAttachmentId: "att_replay",
      targetSheet: "Sheet3",
      targetRange: "B4:C5",
      headers: ["Name", "Qty"],
      values: [["Widget", 2]],
      confidence: 0.94,
      warnings: [],
      requiresConfirmation: true,
      extractionMode: "real" as const,
      shape: {
        rows: 2,
        columns: 2
      }
    };
    const run = traceBus.getRun("run_replay");
    if (!run) {
      throw new Error("Expected run to exist.");
    }
    run.response = {
      schemaVersion: "1.0.0",
      type: "sheet_import_plan",
      requestId: "req_replay",
      hermesRunId: "run_replay",
      processedBy: "hermes",
      serviceLabel: "hermes-remote-us-1",
      environmentLabel: "review",
      startedAt: "2026-04-19T09:30:00.000Z",
      completedAt: "2026-04-19T09:30:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "sheet_import_plan_ready",
          timestamp: "2026-04-19T09:30:01.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: plan
    };

    const approval = approveWriteback({
      requestId: "req_replay",
      runId: "run_replay",
      plan,
      traceBus,
      config
    });

    completeWriteback({
      requestId: "req_replay",
      runId: "run_replay",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildRangeWriteResult(plan, { hostPlatform: "google_sheets" }),
      traceBus,
      config
    });

    const replayCompletion = completeWriteback({
      requestId: "req_replay",
      runId: "run_replay",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildRangeWriteResult(plan, { hostPlatform: "google_sheets" }),
      traceBus,
      config
    });

    expect(replayCompletion).toEqual({ ok: true });
  });

  it("rejects completion results whose summary exceeds the public history contract limit", () => {
    const traceBus = new TraceBus();
    const plan = {
      operation: "create_sheet" as const,
      sheetName: "Summary Overflow",
      explanation: "Create a sheet for the overflow regression.",
      confidence: 0.95,
      requiresConfirmation: true as const,
      overwriteRisk: "none" as const
    };

    setRunResponse(traceBus, {
      runId: "run_summary_overflow",
      requestId: "req_summary_overflow",
      type: "workbook_structure_update",
      traceEvent: "workbook_structure_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_summary_overflow",
        runId: "run_summary_overflow",
        plan
      }
    });

    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_summary_overflow",
        runId: "run_summary_overflow",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "workbook_structure_update",
          hostPlatform: "excel_windows",
          operation: "create_sheet",
          sheetName: "Summary Overflow",
          positionResolved: 0,
          sheetCount: 1,
          summary: "A".repeat(12001)
        }
      }
    });

    expectRouteError(
      completion,
      400,
      "INVALID_REQUEST",
      "That writeback completion request is invalid."
    );
  });

  it("rejects replayed completion when the consumed approval is retried with a different result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet3",
      targetRange: "B4:C5",
      operation: "replace_range" as const,
      values: [
        ["North", 100],
        ["South", 200]
      ],
      shape: {
        rows: 2,
        columns: 2
      },
      explanation: "Write the regional totals.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sheet3!B4:C5"],
      overwriteRisk: "low"
    };

    setRunResponse(traceBus, {
      runId: "run_replay_mismatch",
      requestId: "req_replay_mismatch",
      type: "sheet_update",
      traceEvent: "sheet_update_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req_replay_mismatch",
      runId: "run_replay_mismatch",
      plan,
      traceBus,
      config: testConfig
    });

    completeWriteback({
      requestId: "req_replay_mismatch",
      runId: "run_replay_mismatch",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildRangeWriteResult(plan, {
        hostPlatform: "google_sheets"
      }),
      traceBus,
      config: testConfig
    });

    expect(() => completeWriteback({
      requestId: "req_replay_mismatch",
      runId: "run_replay_mismatch",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildRangeWriteResult(plan, {
        hostPlatform: "google_sheets",
        writtenColumns: 1
      }),
      traceBus,
      config: testConfig
    })).toThrow("Writeback result does not match the approved plan details.");
  });

  it("does not consume the approval token when execution history persistence fails before completion is recorded", () => {
    class FailingCompletionLedger extends ExecutionLedger {
      override recordCompleted(): void {
        throw new Error("history write failed");
      }
    }

    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "A1:B2",
      operation: "replace_range" as const,
      values: [["North", 120], ["South", 95]],
      explanation: "Write the refreshed sales block.",
      confidence: 0.94,
      requiresConfirmation: true as const,
      overwriteRisk: "low" as const,
      shape: { rows: 2, columns: 2 }
    };

    setRunResponse(traceBus, {
      runId: "run_sheet_retryable_completion",
      requestId: "req_sheet_retryable_completion",
      type: "sheet_update",
      traceEvent: "sheet_update_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_sheet_retryable_completion",
        runId: "run_sheet_retryable_completion",
        plan
      }
    });

    expect(approval.status).toBe(200);

    const failedCompletion = invokeWritebackRoute({
      traceBus,
      executionLedger: new FailingCompletionLedger(),
      path: "/complete",
      body: {
        requestId: "req_sheet_retryable_completion",
        runId: "run_sheet_retryable_completion",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: buildRangeWriteResult(plan)
      }
    });

    expectRouteError(
      failedCompletion,
      500,
      "INTERNAL_ERROR",
      "The gateway couldn't complete that write-back request."
    );
    expect(traceBus.getRun("run_sheet_retryable_completion")?.writeback).toMatchObject({
      approvedPlanDigest: (approval.body as any).planDigest,
      approvalToken: (approval.body as any).approvalToken
    });
    expect(traceBus.getRun("run_sheet_retryable_completion")?.writeback?.completedAt).toBeUndefined();
    expect(traceBus.getRun("run_sheet_retryable_completion")?.writeback?.result).toBeUndefined();

    const retriedCompletion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_sheet_retryable_completion",
        runId: "run_sheet_retryable_completion",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: buildRangeWriteResult(plan)
      }
    });

    expect(retriedCompletion.status).toBe(200);
    expect(retriedCompletion.body).toEqual({ ok: true });
  });

  it("rejects writeback completion when the approval token has expired", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-04-19T09:30:00.000Z"));

    const traceBus = new TraceBus();
    traceBus.ensureRun("run_expired", "req_expired");
    const config = {
      port: 8787,
      environmentLabel: "review",
      serviceLabel: "hermes-remote-us-1",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      maxUploadBytes: 1000,
      approvalSecret: "secret",
      hermesMode: "mock" as const,
      hermesBaseUrl: undefined,
      skillRegistryPath: ""
    };

    const plan = {
      sourceAttachmentId: "att_expired",
      targetSheet: "Sheet3",
      targetRange: "B4:C5",
      headers: ["Name", "Qty"],
      values: [["Widget", 2]],
      confidence: 0.94,
      warnings: [],
      requiresConfirmation: true,
      extractionMode: "real" as const,
      shape: {
        rows: 2,
        columns: 2
      }
    };
    const run = traceBus.getRun("run_expired");
    if (!run) {
      throw new Error("Expected run to exist.");
    }
    run.response = {
      schemaVersion: "1.0.0",
      type: "sheet_import_plan",
      requestId: "req_expired",
      hermesRunId: "run_expired",
      processedBy: "hermes",
      serviceLabel: "hermes-remote-us-1",
      environmentLabel: "review",
      startedAt: "2026-04-19T09:30:00.000Z",
      completedAt: "2026-04-19T09:30:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "sheet_import_plan_ready",
          timestamp: "2026-04-19T09:30:01.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: plan
    };

    const approval = approveWriteback({
      requestId: "req_expired",
      runId: "run_expired",
      plan,
      traceBus,
      config
    });

    vi.setSystemTime(new Date("2026-04-19T09:46:00.000Z"));

    expect(() => completeWriteback({
      requestId: "req_expired",
      runId: "run_expired",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildRangeWriteResult(plan, { hostPlatform: "google_sheets" }),
      traceBus,
      config
    })).toThrow("Approval token expired.");

    vi.useRealTimers();
  });

  it("accepts approval tokens even when requestId contains the token delimiter", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      sourceAttachmentId: "att_pipe",
      targetSheet: "Sheet3",
      targetRange: "B4:C5",
      headers: ["Name", "Qty"],
      values: [["Widget", 2]],
      confidence: 0.94,
      warnings: [],
      requiresConfirmation: true,
      extractionMode: "real" as const,
      shape: {
        rows: 2,
        columns: 2
      }
    };

    setRunResponse(traceBus, {
      runId: "run_pipe",
      requestId: "req|pipe",
      type: "sheet_import_plan",
      traceEvent: "sheet_import_plan_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req|pipe",
      runId: "run_pipe",
      workbookSessionKey: "excel_windows::workbook-123",
      plan,
      traceBus,
      executionLedger,
      config: testConfig
    });

    expect(() => completeWriteback({
      requestId: "req|pipe",
      runId: "run_pipe",
      workbookSessionKey: "excel_windows::workbook-123",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildRangeWriteResult(plan),
      traceBus,
      executionLedger,
      config: testConfig
    })).not.toThrow();
  });

  it("rejects approval when the requestId does not match the stored run", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      explanation: "Build a sales pivot by region.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_request_id_approval_mismatch",
      requestId: "req_expected",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan
    });

    const response = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_other",
        runId: "run_request_id_approval_mismatch",
        plan
      }
    });

    expectRouteError(
      response,
      409,
      "STALE_REQUEST",
      "This approval no longer matches the current Hermes request."
    );
  });

  it("rejects completion when the requestId does not match the stored run", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      explanation: "Build a sales pivot by region.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_request_id_completion_mismatch",
      requestId: "req_expected",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_expected",
        runId: "run_request_id_completion_mismatch",
        plan
      }
    });

    const response = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_other",
        runId: "run_request_id_completion_mismatch",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "excel_windows",
          ...plan,
          summary: "Created pivot table on Sales Pivot!A1."
        }
      }
    });

    expectRouteError(
      response,
      409,
      "STALE_REQUEST",
      "This approval no longer matches the current Hermes request."
    );
  });

  it("rejects completion for an unknown run", () => {
    const traceBus = new TraceBus();
    const config = {
      port: 8787,
      environmentLabel: "review",
      serviceLabel: "hermes-remote-us-1",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      maxUploadBytes: 1000,
      approvalSecret: "secret",
      hermesMode: "mock" as const,
      hermesBaseUrl: undefined,
      skillRegistryPath: ""
    };
    const plan = {
      sourceAttachmentId: "att_missing",
      targetSheet: "Sheet3",
      targetRange: "B4:C5",
      headers: ["Name", "Qty"],
      values: [["Widget", 2]],
      confidence: 0.94,
      warnings: [],
      requiresConfirmation: true,
      extractionMode: "real" as const,
      shape: {
        rows: 2,
        columns: 2
      }
    };
    const planDigest = digestPlan(plan);
    const approvalToken = createApprovalToken({
      requestId: "req_missing",
      runId: "run_missing",
      planDigest,
      issuedAt: "2026-04-19T09:30:00.000Z",
      secret: config.approvalSecret
    });

    const response = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_missing",
        runId: "run_missing",
        approvalToken,
        planDigest,
        result: buildRangeWriteResult(plan)
      }
    });

    expectRouteError(
      response,
      404,
      "RUN_NOT_FOUND",
      "That Hermes request is no longer available."
    );
  });

  it("rejects completion when the workbook session does not match the approved session", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      explanation: "Build a sales pivot by region.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_session_mismatch",
      requestId: "req_session_mismatch",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_session_mismatch",
        runId: "run_session_mismatch",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const response = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_session_mismatch",
        runId: "run_session_mismatch",
        workbookSessionKey: "excel_windows::workbook-456",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "excel_windows",
          ...plan,
          summary: "Created pivot table on Sales Pivot!A1."
        }
      }
    });

    expectRouteError(
      response,
      409,
      "INVALID_APPROVAL",
      "This approval is no longer valid for the current update."
    );
  });

  it("rejects approval for a plan that does not match the stored Hermes response", () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_mismatch", "req_mismatch");
    const run = traceBus.getRun("run_mismatch");
    if (!run) {
      throw new Error("Expected run to exist.");
    }

    run.response = {
      schemaVersion: "1.0.0",
      type: "sheet_import_plan",
      requestId: "req_mismatch",
      hermesRunId: "run_mismatch",
      processedBy: "hermes",
      serviceLabel: "hermes-remote-us-1",
      environmentLabel: "review",
      startedAt: "2026-04-19T09:30:00.000Z",
      completedAt: "2026-04-19T09:30:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "sheet_import_plan_ready",
          timestamp: "2026-04-19T09:30:01.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceAttachmentId: "att_real",
        targetSheet: "Sheet3",
        targetRange: "B4:C5",
        headers: ["Name", "Qty"],
        values: [["Widget", 2]],
        confidence: 0.94,
        warnings: [],
        requiresConfirmation: true,
        extractionMode: "real",
        shape: {
          rows: 2,
          columns: 2
        }
      }
    };

    const config = {
      port: 8787,
      environmentLabel: "review",
      serviceLabel: "hermes-remote-us-1",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      maxUploadBytes: 1000,
      approvalSecret: "secret",
      hermesMode: "mock" as const,
      hermesBaseUrl: undefined,
      skillRegistryPath: ""
    };

    expect(() => approveWriteback({
      requestId: "req_mismatch",
      runId: "run_mismatch",
      plan: {
        sourceAttachmentId: "att_other",
        targetSheet: "Sheet4",
        targetRange: "A1:B2",
        headers: ["Other", "Plan"],
        values: [["Wrong", 9]],
        confidence: 0.5,
        warnings: [],
        requiresConfirmation: true,
        extractionMode: "real",
        shape: {
          rows: 2,
          columns: 2
        }
      },
      traceBus,
      config
    })).toThrow("Approved plan does not match the stored Hermes response.");
  });

  it("accepts pivot table writeback requests through /approve and /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      columnGroups: ["Quarter"],
      valueAggregations: [
        { field: "Revenue", aggregation: "sum" },
        { field: "Deals", aggregation: "count" }
      ],
      filters: [{ field: "Status", operator: "equal_to", value: "Closed Won" }],
      sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
      explanation: "Build a sales pivot by region and rep.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_pivot",
      requestId: "req_pivot",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_pivot",
        runId: "run_pivot",
        plan
      }
    });
    expect(approval.status).toBe(200);
    expect(approval.body).toMatchObject({
      requestId: "req_pivot",
      runId: "run_pivot"
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_pivot",
        runId: "run_pivot",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "excel_windows",
          ...plan,
          summary: "Created pivot table on Sales Pivot!A1."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
    expect(traceBus.getRun("run_pivot")?.writeback?.result).toMatchObject({
      kind: "pivot_table_update",
      targetSheet: "Sales Pivot",
      targetRange: "A1"
    });
  });

  it("records pivot history as non-reversible and not undo-eligible", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" as const }],
      explanation: "Build a sales pivot by region.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_pivot_history_001",
      requestId: "req_pivot_history_001",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_pivot_history_001",
        runId: "run_pivot_history_001",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_pivot_history_001",
        runId: "run_pivot_history_001",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "excel_windows",
          ...plan,
          summary: "Created pivot table on Sales Pivot!A1."
        }
      }
    });
    expect(completion.status).toBe(200);

    expect(
      executionLedger.listHistory("excel_windows::workbook-123").entries[0]
    ).toMatchObject({
      planType: "pivot_table_plan",
      reversible: false,
      undoEligible: false
    });
  });

  it("keeps ordinary sheet writes non-undoable until the host confirms an exact rollback snapshot", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      targetSheet: "Sales",
      targetRange: "A1:B2",
      operation: "replace_range" as const,
      values: [["Region", "Revenue"], ["North", 1200]],
      explanation: "Write sample sales data.",
      confidence: 0.92,
      requiresConfirmation: true as const,
      shape: { rows: 2, columns: 2 }
    };

    setRunResponse(traceBus, {
      runId: "run_sheet_write_history_001",
      requestId: "req_sheet_write_history_001",
      type: "sheet_update",
      traceEvent: "sheet_update_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_sheet_write_history_001",
        runId: "run_sheet_write_history_001",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_sheet_write_history_001",
        runId: "run_sheet_write_history_001",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: buildRangeWriteResult(plan)
      }
    });
    expect(completion.status).toBe(200);

    expect(
      executionLedger.listHistory("excel_windows::workbook-123").entries[0]
    ).toMatchObject({
      planType: "sheet_update",
      reversible: true,
      undoEligible: false
    });
  });

  it("records ordinary sheet writes as undo-eligible when the host confirms an exact rollback snapshot", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      targetSheet: "Sales",
      targetRange: "A1:B2",
      operation: "replace_range" as const,
      values: [["Region", "Revenue"], ["North", 1200]],
      explanation: "Write sample sales data.",
      confidence: 0.92,
      requiresConfirmation: true as const,
      shape: { rows: 2, columns: 2 }
    };

    setRunResponse(traceBus, {
      runId: "run_sheet_write_history_002",
      requestId: "req_sheet_write_history_002",
      type: "sheet_update",
      traceEvent: "sheet_update_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_sheet_write_history_002",
        runId: "run_sheet_write_history_002",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_sheet_write_history_002",
        runId: "run_sheet_write_history_002",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: buildRangeWriteResult(plan, { undoReady: true })
      }
    });
    expect(completion.status).toBe(200);

    expect(
      executionLedger.listHistory("excel_windows::workbook-123").entries[0]
    ).toMatchObject({
      planType: "sheet_update",
      reversible: true,
      undoEligible: true
    });
  });

  it("rejects composite completion when stepResults are duplicated and a step is missing", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      steps: [
        {
          stepId: "first",
          dependsOn: [],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 1",
            position: "end" as const,
            explanation: "Create Stage 1.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        },
        {
          stepId: "second",
          dependsOn: ["first"],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 2",
            position: "end" as const,
            explanation: "Create Stage 2.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        }
      ],
      explanation: "Two-step workbook setup.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      overwriteRisk: "none" as const,
      confirmationLevel: "standard" as const,
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    };

    executionLedger.storeDryRun({
      planDigest: digestCanonicalPlan(plan),
      workbookSessionKey: "excel_windows::workbook-123",
      simulated: true,
      steps: [
        { stepId: "first", status: "simulated", summary: "Would create Stage 1." },
        { stepId: "second", status: "simulated", summary: "Would create Stage 2." }
      ],
      predictedAffectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      predictedSummaries: ["Would run the two-step workbook setup."],
      overwriteRisk: "none",
      reversible: true,
      expiresAt: "2099-01-01T00:00:00.000Z"
    });

    setRunResponse(traceBus, {
      runId: "run_composite_dupe",
      requestId: "req_composite_dupe",
      type: "composite_plan",
      traceEvent: "composite_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_composite_dupe",
        runId: "run_composite_dupe",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_composite_dupe",
        runId: "run_composite_dupe",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "composite_update",
          operation: "composite_update",
          hostPlatform: "excel_windows",
          executionId: (approval.body as any).executionId,
          stepResults: [
            {
              stepId: "first",
              status: "completed",
              summary: "Created Stage 1."
            },
            {
              stepId: "first",
              status: "completed",
              summary: "Created Stage 1 again."
            }
          ],
          summary: "Incorrectly reported composite completion."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects composite completion when step statuses violate dependency or halt semantics", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      steps: [
        {
          stepId: "first",
          dependsOn: [],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 1",
            position: "end" as const,
            explanation: "Create Stage 1.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        },
        {
          stepId: "second",
          dependsOn: ["first"],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 2",
            position: "end" as const,
            explanation: "Create Stage 2.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        }
      ],
      explanation: "Two-step workbook setup.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      overwriteRisk: "none" as const,
      confirmationLevel: "standard" as const,
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    };

    executionLedger.storeDryRun({
      planDigest: digestPlan(plan),
      workbookSessionKey: "excel_windows::workbook-123",
      simulated: true,
      steps: [
        { stepId: "first", status: "simulated", summary: "Would create Stage 1." },
        { stepId: "second", status: "simulated", summary: "Would create Stage 2." }
      ],
      predictedAffectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      predictedSummaries: ["Would run the two-step workbook setup."],
      overwriteRisk: "none",
      reversible: true,
      expiresAt: "2099-01-01T00:00:00.000Z"
    });

    setRunResponse(traceBus, {
      runId: "run_composite_status_violation",
      requestId: "req_composite_status_violation",
      type: "composite_plan",
      traceEvent: "composite_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_composite_status_violation",
        runId: "run_composite_status_violation",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_composite_status_violation",
        runId: "run_composite_status_violation",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "composite_update",
          operation: "composite_update",
          hostPlatform: "excel_windows",
          executionId: (approval.body as any).executionId,
          stepResults: [
            {
              stepId: "first",
              status: "failed",
              summary: "Stage 1 failed."
            },
            {
              stepId: "second",
              status: "completed",
              summary: "Stage 2 incorrectly claimed success."
            }
          ],
          summary: "Incorrectly reported composite completion."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("records composite history as failed and non-undoable when any workflow step fails or is skipped", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      steps: [
        {
          stepId: "first",
          dependsOn: [],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 1",
            position: "end" as const,
            explanation: "Create Stage 1.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        },
        {
          stepId: "second",
          dependsOn: ["first"],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 2",
            position: "end" as const,
            explanation: "Create Stage 2.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        }
      ],
      explanation: "Two-step workbook setup.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      overwriteRisk: "none" as const,
      confirmationLevel: "standard" as const,
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    };

    executionLedger.storeDryRun({
      planDigest: digestPlan(plan),
      workbookSessionKey: "excel_windows::workbook-123",
      simulated: true,
      steps: [
        { stepId: "first", status: "simulated", summary: "Would create Stage 1." },
        { stepId: "second", status: "simulated", summary: "Would create Stage 2." }
      ],
      predictedAffectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      predictedSummaries: ["Would run the two-step workbook setup."],
      overwriteRisk: "none",
      reversible: true,
      expiresAt: "2099-01-01T00:00:00.000Z"
    });

    setRunResponse(traceBus, {
      runId: "run_composite_partial_failure",
      requestId: "req_composite_partial_failure",
      type: "composite_plan",
      traceEvent: "composite_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_composite_partial_failure",
        runId: "run_composite_partial_failure",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_composite_partial_failure",
        runId: "run_composite_partial_failure",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "composite_update",
          operation: "composite_update",
          hostPlatform: "excel_windows",
          executionId: (approval.body as any).executionId,
          stepResults: [
            {
              stepId: "first",
              status: "failed",
              summary: "Stage 1 failed."
            },
            {
              stepId: "second",
              status: "skipped",
              summary: "Skipped because an earlier workflow step failed."
            }
          ],
          summary: "Workflow finished with failures."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(
      executionLedger.listHistory("excel_windows::workbook-123").entries[0]
    ).toMatchObject({
      executionId: (approval.body as any).executionId,
      planType: "composite_plan",
      status: "failed",
      undoEligible: false,
      stepEntries: [
        {
          stepId: "first",
          status: "failed"
        },
        {
          stepId: "second",
          status: "skipped"
        }
      ]
    });
  });

  it("records successful composite history as non-undoable until hosts provide an exact composite rollback snapshot", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      steps: [
        {
          stepId: "first",
          dependsOn: [],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 1",
            position: "end" as const,
            explanation: "Create Stage 1.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        },
        {
          stepId: "second",
          dependsOn: ["first"],
          continueOnError: false,
          plan: {
            operation: "create_sheet",
            sheetName: "Stage 2",
            position: "end" as const,
            explanation: "Create Stage 2.",
            confidence: 0.9,
            requiresConfirmation: true
          }
        }
      ],
      explanation: "Two-step workbook setup.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      overwriteRisk: "none" as const,
      confirmationLevel: "standard" as const,
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    };

    executionLedger.storeDryRun({
      planDigest: digestPlan(plan),
      workbookSessionKey: "excel_windows::workbook-123",
      simulated: true,
      steps: [
        { stepId: "first", status: "simulated", summary: "Would create Stage 1." },
        { stepId: "second", status: "simulated", summary: "Would create Stage 2." }
      ],
      predictedAffectedRanges: ["Stage 1!A1", "Stage 2!A1"],
      predictedSummaries: ["Would run the two-step workbook setup."],
      overwriteRisk: "none",
      reversible: true,
      expiresAt: "2099-01-01T00:00:00.000Z"
    });

    setRunResponse(traceBus, {
      runId: "run_composite_success_no_undo",
      requestId: "req_composite_success_no_undo",
      type: "composite_plan",
      traceEvent: "composite_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/approve",
      body: {
        requestId: "req_composite_success_no_undo",
        runId: "run_composite_success_no_undo",
        workbookSessionKey: "excel_windows::workbook-123",
        plan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      executionLedger,
      path: "/complete",
      body: {
        requestId: "req_composite_success_no_undo",
        runId: "run_composite_success_no_undo",
        workbookSessionKey: "excel_windows::workbook-123",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "composite_update",
          operation: "composite_update",
          hostPlatform: "excel_windows",
          executionId: (approval.body as any).executionId,
          stepResults: [
            {
              stepId: "first",
              status: "completed",
              summary: "Created Stage 1."
            },
            {
              stepId: "second",
              status: "completed",
              summary: "Created Stage 2."
            }
          ],
          summary: "Workflow finished successfully."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(
      executionLedger.listHistory("excel_windows::workbook-123").entries[0]
    ).toMatchObject({
      executionId: (approval.body as any).executionId,
      planType: "composite_plan",
      status: "completed",
      reversible: false,
      undoEligible: false,
      redoEligible: false
    });
  });

  it("rejects pivot table completion results with the wrong result kind through /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      explanation: "Build a sales pivot by region and rep.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_pivot_wrong_kind",
      requestId: "req_pivot_wrong_kind",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_pivot_wrong_kind",
        runId: "run_pivot_wrong_kind",
        plan
      }
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_pivot_wrong_kind",
        runId: "run_pivot_wrong_kind",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "chart_update",
          operation: "chart_update",
          hostPlatform: "excel_windows",
          sourceSheet: "Sales",
          sourceRange: "A1:F50",
          targetSheet: "Sales Pivot",
          targetRange: "A1",
          chartType: "bar",
          series: [{ field: "Revenue", label: "Revenue" }],
          explanation: "Wrong update family.",
          confidence: 0.5,
          requiresConfirmation: true,
          affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
          overwriteRisk: "low",
          confirmationLevel: "standard",
          summary: "Wrong update family."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects pivot table completion results with same-family target mismatches through /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      explanation: "Build a sales pivot by region and rep.",
      confidence: 0.9,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_pivot_wrong_target",
      requestId: "req_pivot_wrong_target",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_pivot_wrong_target",
        runId: "run_pivot_wrong_target",
        plan
      }
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_pivot_wrong_target",
        runId: "run_pivot_wrong_target",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "excel_windows",
          ...plan,
          targetSheet: "Wrong Pivot",
          targetRange: "B2",
          summary: "Created pivot table on the wrong target."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects pivot table completion results when the applied pivot configuration differs on the approved target", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_pivot_wrong_config",
      requestId: "req_pivot_wrong_config",
      type: "pivot_table_plan",
      traceEvent: "pivot_table_plan_ready",
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        columnGroups: ["Quarter"],
        valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
        filters: [{ field: "Status", operator: "equal_to", value: "Closed Won" }],
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      result: {
        kind: "pivot_table_update",
        operation: "pivot_table_update",
        hostPlatform: "google_sheets",
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region"],
        columnGroups: ["Quarter"],
        valueAggregations: [{ field: "Revenue", aggregation: "average" }],
        filters: [{ field: "Status", operator: "equal_to", value: "Closed Won" }],
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        summary: "Created a pivot table with the wrong grouping and aggregation."
      }
    });
  });

  it("accepts chart writeback requests through /approve and /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:C20",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line" as const,
      categoryField: "Month",
      series: [
        { field: "Revenue", label: "Revenue" },
        { field: "Margin", label: "Margin" }
      ],
      title: "Revenue vs Margin",
      legendPosition: "bottom" as const,
      explanation: "Chart monthly revenue and margin.",
      confidence: 0.93,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_chart",
      requestId: "req_chart",
      type: "chart_plan",
      traceEvent: "chart_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_chart",
        runId: "run_chart",
        plan
      }
    });
    expect(approval.status).toBe(200);
    expect(approval.body).toMatchObject({
      requestId: "req_chart",
      runId: "run_chart"
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_chart",
        runId: "run_chart",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "chart_update",
          operation: "chart_update",
          hostPlatform: "google_sheets",
          ...plan,
          summary: "Created line chart on Sales Chart!A1."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
    expect(traceBus.getRun("run_chart")?.writeback?.result).toMatchObject({
      kind: "chart_update",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line"
    });
  });

  it("rejects chart completion results with the wrong result kind through /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:C20",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line" as const,
      series: [{ field: "Revenue", label: "Revenue" }],
      explanation: "Chart monthly revenue.",
      confidence: 0.93,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_chart_wrong_kind",
      requestId: "req_chart_wrong_kind",
      type: "chart_plan",
      traceEvent: "chart_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_chart_wrong_kind",
        runId: "run_chart_wrong_kind",
        plan
      }
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_chart_wrong_kind",
        runId: "run_chart_wrong_kind",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "google_sheets",
          sourceSheet: "Sales",
          sourceRange: "A1:C20",
          targetSheet: "Sales Chart",
          targetRange: "A1",
          rowGroups: ["Month"],
          valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
          explanation: "Wrong update family.",
          confidence: 0.5,
          requiresConfirmation: true,
          affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
          overwriteRisk: "low",
          confirmationLevel: "standard",
          summary: "Wrong update family."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects chart completion results with same-family target mismatches through /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:C20",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line" as const,
      series: [{ field: "Revenue", label: "Revenue" }],
      explanation: "Chart monthly revenue.",
      confidence: 0.93,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_chart_wrong_target",
      requestId: "req_chart_wrong_target",
      type: "chart_plan",
      traceEvent: "chart_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_chart_wrong_target",
        runId: "run_chart_wrong_target",
        plan
      }
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_chart_wrong_target",
        runId: "run_chart_wrong_target",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
      result: {
        kind: "chart_update",
        operation: "chart_update",
        hostPlatform: "google_sheets",
        ...plan,
        targetSheet: "Other Chart",
        targetRange: "C5",
        chartType: "bar",
        summary: "Created the wrong chart."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects chart completion results when the applied data semantics differ on the approved target", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_chart_wrong_config",
      requestId: "req_chart_wrong_config",
      type: "chart_plan",
      traceEvent: "chart_plan_ready",
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Month",
        series: [
          { field: "Revenue", label: "Revenue" },
          { field: "Margin", label: "Margin" }
        ],
        title: "Revenue vs Margin",
        legendPosition: "bottom",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      result: {
        kind: "chart_update",
        operation: "chart_update",
        hostPlatform: "google_sheets",
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Quarter",
        series: [{ field: "Revenue", label: "Revenue" }],
        title: "Revenue Only",
        legendPosition: "hidden",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        summary: "Created a chart with the wrong category, series, and title."
      }
    });
  });

  it("accepts materialized analysis report writeback requests through /approve and /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report" as const,
      targetSheet: "Sales Report",
      targetRange: "A1",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Write a sales report sheet.",
      confidence: 0.91,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_analysis_report",
      requestId: "req_analysis_report",
      type: "analysis_report_plan",
      traceEvent: "analysis_report_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_analysis_report",
        runId: "run_analysis_report",
        plan
      }
    });
    expect(approval.status).toBe(200);
    expect(approval.body).toMatchObject({
      requestId: "req_analysis_report",
      runId: "run_analysis_report"
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_analysis_report",
        runId: "run_analysis_report",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "analysis_report_update",
          hostPlatform: "excel_macos",
          sourceSheet: "Sales",
          sourceRange: "A1:F50",
          outputMode: "materialize_report",
          targetSheet: "Sales Report",
          targetRange: "A1:D5",
          sections: plan.sections,
          explanation: plan.explanation,
          confidence: plan.confidence,
          requiresConfirmation: plan.requiresConfirmation,
          affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D5"],
          overwriteRisk: plan.overwriteRisk,
          confirmationLevel: plan.confirmationLevel,
          summary: "Created analysis report on Sales Report!A1:D5."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
    expect(traceBus.getRun("run_analysis_report")?.writeback?.result).toMatchObject({
      kind: "analysis_report_update",
      targetSheet: "Sales Report",
      targetRange: "A1:D5"
    });
  });

  it("accepts resolved materialized analysis report approvals and completions when the stored Hermes plan still uses the anchor range", () => {
    const traceBus = new TraceBus();
    const storedPlan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report" as const,
      targetSheet: "Sales Report",
      targetRange: "A1",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        },
        {
          type: "group_breakdown",
          title: "By region",
          summary: "West leads closed-won revenue.",
          sourceRanges: ["Sales!A1:F50", "Sales!H1:J20"]
        }
      ],
      explanation: "Write a sales report sheet.",
      confidence: 0.91,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };
    const resolvedApprovalPlan = {
      ...storedPlan,
      targetRange: "A1:D6"
    };

    setRunResponse(traceBus, {
      runId: "run_analysis_report_resolved",
      requestId: "req_analysis_report_resolved",
      type: "analysis_report_plan",
      traceEvent: "analysis_report_plan_ready",
      plan: storedPlan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_analysis_report_resolved",
        runId: "run_analysis_report_resolved",
        plan: resolvedApprovalPlan
      }
    });
    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_analysis_report_resolved",
        runId: "run_analysis_report_resolved",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "analysis_report_update",
          hostPlatform: "excel_windows",
          sourceSheet: "Sales",
          sourceRange: "A1:F50",
          outputMode: "materialize_report",
          targetSheet: "Sales Report",
          targetRange: "A1:D6",
          sections: storedPlan.sections,
          explanation: storedPlan.explanation,
          confidence: storedPlan.confidence,
          requiresConfirmation: storedPlan.requiresConfirmation,
          affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"],
          overwriteRisk: storedPlan.overwriteRisk,
          confirmationLevel: storedPlan.confirmationLevel,
          summary: "Created analysis report on Sales Report!A1:D6."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
    expect(traceBus.getRun("run_analysis_report_resolved")?.writeback?.result).toMatchObject({
      kind: "analysis_report_update",
      targetSheet: "Sales Report",
      targetRange: "A1:D6"
    });
  });

  it("rejects invalid materialized analysis report target ranges at request validation instead of throwing a 500", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report" as const,
      targetSheet: "Sales Report",
      targetRange: "A1",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Write a sales report sheet.",
      confidence: 0.91,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_analysis_report_invalid_anchor",
      requestId: "req_analysis_report_invalid_anchor",
      type: "analysis_report_plan",
      traceEvent: "analysis_report_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_analysis_report_invalid_anchor",
        runId: "run_analysis_report_invalid_anchor",
        plan: {
          ...plan,
          targetRange: "A:A"
        }
      }
    });

    expectRouteError(
      approval,
      400,
      "INVALID_REQUEST",
      "That writeback approval request is invalid."
    );
    expect(approval.body).toMatchObject({
      error: {
        issues: expect.arrayContaining([
          {
            path: "plan.targetRange",
            message: "targetRange must be a valid A1 range."
          }
        ])
      }
    });
  });

  it("rejects materialized analysis report completion results with the wrong result kind through /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report" as const,
      targetSheet: "Sales Report",
      targetRange: "A1",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Write a sales report sheet.",
      confidence: 0.91,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_analysis_report_wrong_kind",
      requestId: "req_analysis_report_wrong_kind",
      type: "analysis_report_plan",
      traceEvent: "analysis_report_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_analysis_report_wrong_kind",
        runId: "run_analysis_report_wrong_kind",
        plan
      }
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_analysis_report_wrong_kind",
        runId: "run_analysis_report_wrong_kind",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "excel_macos",
          sourceSheet: "Sales",
          sourceRange: "A1:F50",
          targetSheet: "Sales Report",
          targetRange: "A1",
          rowGroups: ["Region"],
          valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
          explanation: "Wrong update family.",
          confidence: 0.5,
          requiresConfirmation: true,
          affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
          overwriteRisk: "low",
          confirmationLevel: "standard",
          summary: "Wrong update family."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects materialized analysis report completion results with same-family target mismatches through /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report" as const,
      targetSheet: "Sales Report",
      targetRange: "A1",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Write a sales report sheet.",
      confidence: 0.91,
      requiresConfirmation: true as const,
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_analysis_report_wrong_target",
      requestId: "req_analysis_report_wrong_target",
      type: "analysis_report_plan",
      traceEvent: "analysis_report_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_analysis_report_wrong_target",
        runId: "run_analysis_report_wrong_target",
        plan
      }
    });

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_analysis_report_wrong_target",
        runId: "run_analysis_report_wrong_target",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "analysis_report_update",
          hostPlatform: "excel_macos",
          sourceSheet: "Sales",
          sourceRange: "A1:F50",
          outputMode: "materialize_report",
          targetSheet: "Other Report",
          targetRange: "B4",
          sections: plan.sections,
          explanation: plan.explanation,
          confidence: plan.confidence,
          requiresConfirmation: plan.requiresConfirmation,
          affectedRanges: ["Sales!A1:F50", "Other Report!B4"],
          overwriteRisk: plan.overwriteRisk,
          confirmationLevel: plan.confirmationLevel,
          summary: "Created analysis report on the wrong sheet."
        }
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("does not treat chat-only analysis reports as writeback eligible on /complete", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "chat_only" as const,
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Summarize the selected sales range.",
      confidence: 0.92,
      requiresConfirmation: false as const,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "none" as const,
      confirmationLevel: "standard" as const
    };
    const planDigest = digestPlan(plan);
    const approvalToken = createApprovalToken({
      requestId: "req_analysis_chat_only_complete",
      runId: "run_analysis_chat_only_complete",
      planDigest,
      issuedAt: "2026-04-19T09:30:00.000Z",
      secret: testConfig.approvalSecret
    });

    setRunResponse(traceBus, {
      runId: "run_analysis_chat_only_complete",
      requestId: "req_analysis_chat_only_complete",
      type: "analysis_report_plan",
      traceEvent: "analysis_report_plan_ready",
      plan
    });

    const response = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_analysis_chat_only_complete",
        runId: "run_analysis_chat_only_complete",
        approvalToken,
        planDigest,
        result: {
          kind: "analysis_report_update",
          hostPlatform: "excel_windows",
          sourceSheet: "Sales",
          sourceRange: "A1:F50",
          outputMode: "materialize_report",
          targetSheet: "Sales Report",
          targetRange: "A1",
          sections: plan.sections,
          explanation: plan.explanation,
          confidence: plan.confidence,
          requiresConfirmation: true,
          affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
          overwriteRisk: "none",
          confirmationLevel: "standard",
          summary: "Should not be accepted."
        }
      }
    });

    expectRouteError(
      response,
      409,
      "APPROVAL_NOT_FOUND",
      "This update is no longer awaiting approval."
    );
  });

  it("approves and completes workbook structure updates with a structure result payload", () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_structure", "req_structure");
    const run = traceBus.getRun("run_structure");
    if (!run) {
      throw new Error("Expected run to exist.");
    }

    const plan = {
      operation: "create_sheet",
      sheetName: "New Sheet",
      position: "end",
      explanation: "Create a new sheet at the end of the workbook.",
      confidence: 0.95,
      requiresConfirmation: true,
      overwriteRisk: "none"
    };

    run.response = {
      schemaVersion: "1.0.0",
      type: "workbook_structure_update",
      requestId: "req_structure",
      hermesRunId: "run_structure",
      processedBy: "hermes",
      serviceLabel: "hermes-remote-us-1",
      environmentLabel: "review",
      startedAt: "2026-04-19T09:30:00.000Z",
      completedAt: "2026-04-19T09:30:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "result_generated",
          timestamp: "2026-04-19T09:30:01.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: plan
    } as any;

    const config = {
      port: 8787,
      environmentLabel: "review",
      serviceLabel: "hermes-remote-us-1",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      maxUploadBytes: 1000,
      approvalSecret: "secret",
      hermesMode: "mock" as const,
      hermesBaseUrl: undefined,
      skillRegistryPath: ""
    };

    const approval = approveWriteback({
      requestId: "req_structure",
      runId: "run_structure",
      plan: plan as any,
      traceBus,
      config
    });

    const completion = completeWriteback({
      requestId: "req_structure",
      runId: "run_structure",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: {
        kind: "workbook_structure_update",
        hostPlatform: "google_sheets",
        sheetName: "New Sheet",
        operation: "create_sheet",
        positionResolved: 1,
        sheetCount: 2,
        summary: "Created sheet New Sheet at the end of the workbook."
      } as any,
      traceBus,
      config
    });

    expect(completion.ok).toBe(true);
    expect(traceBus.getRun("run_structure")?.writeback?.result).toMatchObject({
      kind: "workbook_structure_update",
      sheetName: "New Sheet",
      operation: "create_sheet"
    });
  });

  it("rejects workbook structure completion results with same-family identity mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_structure_wrong_details",
      requestId: "req_structure_wrong_details",
      type: "workbook_structure_update",
      traceEvent: "result_generated",
      plan: {
        operation: "create_sheet",
        sheetName: "New Sheet",
        position: "end",
        explanation: "Create a new sheet at the end of the workbook.",
        confidence: 0.95,
        requiresConfirmation: true,
        overwriteRisk: "none"
      },
      result: {
        kind: "workbook_structure_update",
        hostPlatform: "google_sheets",
        sheetName: "Other Sheet",
        operation: "rename_sheet",
        newSheetName: "Wrong Name",
        summary: "Wrong workbook structure result."
      }
    });
  });

  it("accepts workbook rename completions that report the new sheet name", () => {
    const traceBus = new TraceBus();
    const plan = {
      operation: "rename_sheet" as const,
      sheetName: "Old Sheet",
      newSheetName: "New Sheet",
      explanation: "Rename the sheet.",
      confidence: 0.95,
      requiresConfirmation: true as const,
      overwriteRisk: "none" as const
    };

    setRunResponse(traceBus, {
      runId: "run_structure_rename",
      requestId: "req_structure_rename",
      type: "workbook_structure_update",
      traceEvent: "result_generated",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_structure_rename",
        runId: "run_structure_rename",
        plan
      }
    });

    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_structure_rename",
        runId: "run_structure_rename",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "workbook_structure_update",
          hostPlatform: "google_sheets",
          sheetName: "Old Sheet",
          operation: "rename_sheet",
          newSheetName: "New Sheet",
          summary: "Renamed sheet Old Sheet to New Sheet."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
  });

  it("accepts workbook duplicate completions that report the duplicated sheet name", () => {
    const traceBus = new TraceBus();
    const plan = {
      operation: "duplicate_sheet" as const,
      sheetName: "Template",
      newSheetName: "Template Copy",
      position: "end" as const,
      explanation: "Duplicate the template sheet.",
      confidence: 0.95,
      requiresConfirmation: true as const,
      overwriteRisk: "none" as const
    };

    setRunResponse(traceBus, {
      runId: "run_structure_duplicate",
      requestId: "req_structure_duplicate",
      type: "workbook_structure_update",
      traceEvent: "result_generated",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_structure_duplicate",
        runId: "run_structure_duplicate",
        plan
      }
    });

    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_structure_duplicate",
        runId: "run_structure_duplicate",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "workbook_structure_update",
          hostPlatform: "excel_windows",
          sheetName: "Template",
          operation: "duplicate_sheet",
          newSheetName: "Template Copy",
          positionResolved: 3,
          sheetCount: 4,
          summary: "Duplicated sheet Template to Template Copy."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
  });

  it("rejects workbook structure completions when the resolved sheet position differs from the approved plan", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_structure_wrong_position",
      requestId: "req_structure_wrong_position",
      type: "workbook_structure_update",
      traceEvent: "result_generated",
      plan: {
        operation: "create_sheet",
        sheetName: "New Sheet",
        position: "end",
        explanation: "Create a new sheet at the end of the workbook.",
        confidence: 0.95,
        requiresConfirmation: true,
        overwriteRisk: "none"
      },
      result: {
        kind: "workbook_structure_update",
        hostPlatform: "google_sheets",
        operation: "create_sheet",
        sheetName: "New Sheet",
        positionResolved: 0,
        sheetCount: 4,
        summary: "Created sheet New Sheet in the wrong position."
      }
    });
  });

  it("approves and completes range format updates with a range write result payload", () => {
    const traceBus = new TraceBus();
    traceBus.ensureRun("run_format", "req_format");
    const run = traceBus.getRun("run_format");
    if (!run) {
      throw new Error("Expected run to exist.");
    }

    const plan = {
      targetSheet: "Sheet1",
      targetRange: "A1:B2",
      format: {
        backgroundColor: "#4472C4",
        bold: true
      },
      explanation: "Format the selected header range.",
      confidence: 0.92,
      requiresConfirmation: true,
      overwriteRisk: "low"
    };

    run.response = {
      schemaVersion: "1.0.0",
      type: "range_format_update",
      requestId: "req_format",
      hermesRunId: "run_format",
      processedBy: "hermes",
      serviceLabel: "hermes-remote-us-1",
      environmentLabel: "review",
      startedAt: "2026-04-19T09:30:00.000Z",
      completedAt: "2026-04-19T09:30:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "result_generated",
          timestamp: "2026-04-19T09:30:01.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: plan
    } as any;

    const config = {
      port: 8787,
      environmentLabel: "review",
      serviceLabel: "hermes-remote-us-1",
      gatewayPublicBaseUrl: "http://127.0.0.1:8787",
      maxUploadBytes: 1000,
      approvalSecret: "secret",
      hermesMode: "mock" as const,
      hermesBaseUrl: undefined,
      skillRegistryPath: ""
    };

    const approval = approveWriteback({
      requestId: "req_format",
      runId: "run_format",
      plan: plan as any,
      traceBus,
      config
    });

    const completion = completeWriteback({
      requestId: "req_format",
      runId: "run_format",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: {
        kind: "range_format_update",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        format: {
          backgroundColor: "#4472C4",
          bold: true
        },
        explanation: "Format the selected header range.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low",
        summary: "Applied formatting to Sheet1!A1:B2."
      } as any,
      traceBus,
      config
    });

    expect(completion.ok).toBe(true);
    expect(traceBus.getRun("run_format")?.writeback?.result).toMatchObject({
      kind: "range_format_update",
      targetSheet: "Sheet1",
      targetRange: "A1:B2",
      format: {
        backgroundColor: "#4472C4",
        bold: true
      }
    });
  });

  it("rejects range write completion results with same-family target mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_format_wrong_details",
      requestId: "req_format_wrong_details",
      type: "range_format_update",
      traceEvent: "result_generated",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        format: {
          backgroundColor: "#4472C4",
          bold: true
        },
        explanation: "Format the selected header range.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low"
      },
      result: {
        kind: "range_format_update",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        format: {
          backgroundColor: "#ffffff",
          bold: false
        },
        explanation: "Wrong formatting semantics.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low",
        summary: "Applied the wrong formatting."
      }
    });
  });

  it("approves and completes a conditional format plan with a typed result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:D12",
      explanation: "Clear existing conditional rules before applying the new business rule.",
      confidence: 0.97,
      requiresConfirmation: true,
      affectedRanges: ["B2:D12"],
      replacesExistingRules: true,
      managementMode: "clear_on_target"
    };

    setRunResponse(traceBus, {
      runId: "run_conditional_format_typed",
      requestId: "req_conditional_format_typed",
      type: "conditional_format_plan",
      traceEvent: "conditional_format_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_conditional_format_typed",
        runId: "run_conditional_format_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_conditional_format_typed",
        runId: "run_conditional_format_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "conditional_format_update",
          hostPlatform: "excel_windows",
          ...plan,
          summary: "Cleared and replaced conditional formatting on Sheet1!B2:D12."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_conditional_format_typed")?.writeback?.result).toMatchObject({
      kind: "conditional_format_update",
      targetSheet: "Sheet1",
      targetRange: "B2:D12",
      managementMode: "clear_on_target"
    });
  });

  it("rejects a conditional format completion when the result kind does not match the approved family", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:D12",
      explanation: "Clear existing conditional rules before applying the new business rule.",
      confidence: 0.97,
      requiresConfirmation: true,
      affectedRanges: ["B2:D12"],
      replacesExistingRules: true,
      managementMode: "clear_on_target"
    };

    setRunResponse(traceBus, {
      runId: "run_conditional_format_kind_mismatch",
      requestId: "req_conditional_format_kind_mismatch",
      type: "conditional_format_plan",
      traceEvent: "conditional_format_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_conditional_format_kind_mismatch",
        runId: "run_conditional_format_kind_mismatch",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_conditional_format_kind_mismatch",
        runId: "run_conditional_format_kind_mismatch",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: buildRangeWriteResult({
          targetSheet: "Sheet1",
          targetRange: "A1",
          operation: "replace_range",
          values: [["urgent"]],
          explanation: "Wrong family completion payload.",
          confidence: 0.5,
          requiresConfirmation: true,
          shape: { rows: 1, columns: 1 }
        })
      }
    });

    expectRouteError(
      completionResponse,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects a conditional format completion with same-family detail mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_conditional_format_wrong_details",
      requestId: "req_conditional_format_wrong_details",
      type: "conditional_format_plan",
      traceEvent: "conditional_format_plan_ready",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:D12",
        explanation: "Highlight rows that contain urgent text.",
        confidence: 0.97,
        requiresConfirmation: true,
        affectedRanges: ["B2:D12"],
        replacesExistingRules: false,
        managementMode: "add",
        ruleType: "text_contains",
        text: "urgent",
        style: {
          backgroundColor: "#FFF2CC",
          bold: true
        }
      },
      result: {
        kind: "conditional_format_update",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        targetRange: "B2:D12",
        explanation: "Highlight rows that contain urgent text.",
        confidence: 0.97,
        requiresConfirmation: true,
        affectedRanges: ["B2:D12"],
        replacesExistingRules: false,
        managementMode: "add",
        ruleType: "text_contains",
        text: "blocked",
        style: {
          backgroundColor: "#FFF2CC",
          bold: true
        },
        summary: "Wrong conditional formatting target."
      }
    });
  });

  it("rejects conditional format completion replay for the same approval token", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:D12",
      explanation: "Clear existing conditional rules before applying the new business rule.",
      confidence: 0.97,
      requiresConfirmation: true,
      affectedRanges: ["B2:D12"],
      replacesExistingRules: true,
      managementMode: "clear_on_target"
    };

    setRunResponse(traceBus, {
      runId: "run_conditional_format_replay",
      requestId: "req_conditional_format_replay",
      type: "conditional_format_plan",
      traceEvent: "conditional_format_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_conditional_format_replay",
        runId: "run_conditional_format_replay",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionBody = {
      requestId: "req_conditional_format_replay",
      runId: "run_conditional_format_replay",
      approvalToken: approvedBody.approvalToken,
      planDigest: approvedBody.planDigest,
      result: {
        kind: "conditional_format_update",
        hostPlatform: "excel_windows",
        ...plan,
        summary: "Cleared and replaced conditional formatting on Sheet1!B2:D12."
      }
    };

    const firstCompletion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: completionBody
    });

    expect(firstCompletion.status).toBe(200);

    const replayCompletion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: completionBody
    });

    expect(replayCompletion.status).toBe(200);
    expect(replayCompletion.body).toEqual({ ok: true });
  });

  it("reuses the same approval token for a repeated pre-completion approval in the same workbook session", () => {
    const traceBus = new TraceBus();
    const executionLedger = new ExecutionLedger();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:D12",
      explanation: "Clear existing conditional rules before applying the new business rule.",
      confidence: 0.97,
      requiresConfirmation: true,
      affectedRanges: ["B2:D12"],
      replacesExistingRules: true,
      managementMode: "clear_on_target"
    };

    setRunResponse(traceBus, {
      runId: "run_conditional_format_superseded_token",
      requestId: "req_conditional_format_superseded_token",
      type: "conditional_format_plan",
      traceEvent: "conditional_format_plan_ready",
      plan
    });

    vi.useFakeTimers();
    try {
      vi.setSystemTime(new Date("2026-04-19T09:30:00.000Z"));
      const firstApproval = invokeWritebackRoute({
        traceBus,
        executionLedger,
        path: "/approve",
        body: {
          requestId: "req_conditional_format_superseded_token",
          runId: "run_conditional_format_superseded_token",
          workbookSessionKey: "google_sheets::sheet-123",
          plan
        }
      });
      expect(firstApproval.status).toBe(200);

      vi.setSystemTime(new Date("2026-04-19T09:30:01.000Z"));
      const secondApproval = invokeWritebackRoute({
        traceBus,
        executionLedger,
        path: "/approve",
        body: {
          requestId: "req_conditional_format_superseded_token",
          runId: "run_conditional_format_superseded_token",
          workbookSessionKey: "google_sheets::sheet-123",
          plan
        }
      });
      expect(secondApproval.status).toBe(200);

      const firstApprovedBody = firstApproval.body as Record<string, string>;
      const secondApprovedBody = secondApproval.body as Record<string, string>;
      expect(secondApprovedBody.approvalToken).toBe(firstApprovedBody.approvalToken);
      expect(secondApprovedBody.executionId).toBe(firstApprovedBody.executionId);
      expect(secondApprovedBody.approvedAt).toBe(firstApprovedBody.approvedAt);

      const completionWithReusedToken = invokeWritebackRoute({
        traceBus,
        executionLedger,
        path: "/complete",
        body: {
          requestId: "req_conditional_format_superseded_token",
          runId: "run_conditional_format_superseded_token",
          workbookSessionKey: "google_sheets::sheet-123",
          approvalToken: firstApprovedBody.approvalToken,
          planDigest: firstApprovedBody.planDigest,
          result: {
            kind: "conditional_format_update",
            hostPlatform: "excel_windows",
            ...plan,
            summary: "Cleared and replaced conditional formatting on Sheet1!B2:D12."
          }
        }
      });

      expect(completionWithReusedToken.status).toBe(200);

      expect(traceBus.getRun("run_conditional_format_superseded_token")?.writeback).toMatchObject({
        approvalToken: secondApprovedBody.approvalToken,
        approvedPlanDigest: secondApprovedBody.planDigest
      });
      expect(traceBus.getRun("run_conditional_format_superseded_token")?.writeback?.completedAt).toBeDefined();
      expect(executionLedger.listHistory("google_sheets::sheet-123").entries).toHaveLength(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it("rejects a second pre-completion approval from a different workbook session", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:D12",
      explanation: "Clear existing conditional rules before applying the new business rule.",
      confidence: 0.97,
      requiresConfirmation: true,
      affectedRanges: ["B2:D12"],
      replacesExistingRules: true,
      managementMode: "clear_on_target"
    };

    setRunResponse(traceBus, {
      runId: "run_conditional_format_pending_conflict",
      requestId: "req_conditional_format_pending_conflict",
      type: "conditional_format_plan",
      traceEvent: "conditional_format_plan_ready",
      plan
    });

    const firstApproval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_conditional_format_pending_conflict",
        runId: "run_conditional_format_pending_conflict",
        workbookSessionKey: "google_sheets::sheet-123",
        plan
      }
    });
    expect(firstApproval.status).toBe(200);

    const conflictingApproval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_conditional_format_pending_conflict",
        runId: "run_conditional_format_pending_conflict",
        workbookSessionKey: "excel_windows::workbook-456",
        plan
      }
    });

    expectRouteError(
      conflictingApproval,
      409,
      "APPROVAL_ALREADY_PENDING",
      "This update is already awaiting completion in another spreadsheet session."
    );
  });

  it("rejects re-approval after writeback completion and preserves the original audit state", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:D12",
      explanation: "Clear existing conditional rules before applying the new business rule.",
      confidence: 0.97,
      requiresConfirmation: true,
      affectedRanges: ["B2:D12"],
      replacesExistingRules: true,
      managementMode: "clear_on_target"
    };

    setRunResponse(traceBus, {
      runId: "run_conditional_format_reapprove_after_complete",
      requestId: "req_conditional_format_reapprove_after_complete",
      type: "conditional_format_plan",
      traceEvent: "conditional_format_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_conditional_format_reapprove_after_complete",
        runId: "run_conditional_format_reapprove_after_complete",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionBody = {
      requestId: "req_conditional_format_reapprove_after_complete",
      runId: "run_conditional_format_reapprove_after_complete",
      approvalToken: approvedBody.approvalToken,
      planDigest: approvedBody.planDigest,
      result: {
        kind: "conditional_format_update",
        hostPlatform: "excel_windows",
        ...plan,
        summary: "Cleared and replaced conditional formatting on Sheet1!B2:D12."
      }
    };

    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: completionBody
    });

    expect(completionResponse.status).toBe(200);

    const writebackBeforeReplay = traceBus.getRun("run_conditional_format_reapprove_after_complete")
      ?.writeback;
    expect(writebackBeforeReplay?.completedAt).toBeDefined();
    expect(writebackBeforeReplay?.result).toMatchObject({
      kind: "conditional_format_update",
      targetSheet: "Sheet1",
      targetRange: "B2:D12"
    });

    const replayApproval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_conditional_format_reapprove_after_complete",
        runId: "run_conditional_format_reapprove_after_complete",
        plan
      }
    });

    expectRouteError(
      replayApproval,
      409,
      "ALREADY_COMPLETED",
      "This update was already applied."
    );

    expect(traceBus.getRun("run_conditional_format_reapprove_after_complete")?.writeback).toEqual(
      writebackBeforeReplay
    );
  });

  it("approves and completes a data validation plan with a validation result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Restrict the status column to approved values.",
      confidence: 0.95,
      requiresConfirmation: true,
      replacesExistingValidation: true,
      ruleType: "list",
      values: ["Open", "Closed"],
    };

    setRunResponse(traceBus, {
      runId: "run_validation_typed",
      requestId: "req_validation_typed",
      type: "data_validation_plan",
      traceEvent: "data_validation_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_validation_typed",
        runId: "run_validation_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_validation_typed",
        runId: "run_validation_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "data_validation_update",
          hostPlatform: "excel_windows",
          ...plan,
          summary: "Applied validation to Sheet1!B2:B20."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_validation_typed")?.writeback?.result).toMatchObject({
      kind: "data_validation_update",
      targetSheet: "Sheet1",
      targetRange: "B2:B20"
    });
  });

  it("rejects data validation completion results with same-family target mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_validation_wrong_details",
      requestId: "req_validation_wrong_details",
      type: "data_validation_plan",
      traceEvent: "data_validation_plan_ready",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict the status column to approved values.",
        confidence: 0.95,
        requiresConfirmation: true,
        replacesExistingValidation: true,
        ruleType: "list",
        values: ["Open", "Closed"]
      },
      result: {
        kind: "data_validation_update",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict the status column to approved values.",
        confidence: 0.95,
        requiresConfirmation: true,
        replacesExistingValidation: true,
        ruleType: "list",
        values: ["Pending", "Done"],
        summary: "Applied validation to the wrong range."
      }
    });
  });

  it("rejects data validation completion results when the applied rule differs on the approved target", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_validation_wrong_rule",
      requestId: "req_validation_wrong_rule",
      type: "data_validation_plan",
      traceEvent: "data_validation_plan_ready",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict the status column to approved values.",
        confidence: 0.95,
        requiresConfirmation: true,
        replacesExistingValidation: true,
        ruleType: "list",
        values: ["Open", "Closed"]
      },
      result: {
        kind: "data_validation_update",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict the status column to approved values.",
        confidence: 0.95,
        requiresConfirmation: true,
        replacesExistingValidation: true,
        ruleType: "list",
        values: ["Pending", "Done"],
        summary: "Applied the wrong validation rule on the approved range."
      }
    });
  });

  it("approves and completes a named range update with a named-range result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      operation: "retarget",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      name: "InputRange",
      explanation: "Retarget the named range to the input block.",
      confidence: 0.93,
      requiresConfirmation: true
    };

    setRunResponse(traceBus, {
      runId: "run_named_range_typed",
      requestId: "req_named_range_typed",
      type: "named_range_update",
      traceEvent: "named_range_update_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_named_range_typed",
        runId: "run_named_range_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_named_range_typed",
        runId: "run_named_range_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: buildNamedRangeUpdateResult(plan, {
          summary: "Retargeted InputRange to Sheet1!B2:D20."
        })
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_named_range_typed")?.writeback?.result).toMatchObject({
      kind: "named_range_update",
      operation: "retarget",
      name: "InputRange"
    });
  });

  it("rejects named range completion results with same-family identity mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_named_range_wrong_details",
      requestId: "req_named_range_wrong_details",
      type: "named_range_update",
      traceEvent: "named_range_update_ready",
      plan: {
        operation: "retarget",
        scope: "sheet",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        name: "InputRange",
        explanation: "Retarget the named range to the input block.",
        confidence: 0.93,
        requiresConfirmation: true
      },
      result: buildNamedRangeUpdateResult({
        operation: "retarget",
        scope: "sheet",
        sheetName: "Sheet2",
        targetSheet: "Sheet2",
        targetRange: "C3:E21",
        name: "OtherRange",
        explanation: "Retarget the named range to the wrong block.",
        confidence: 0.93,
        requiresConfirmation: true
      }, {
        hostPlatform: "excel_windows",
        summary: "Wrong named range completion."
      })
    });
  });

  it("rejects named range completions that change scope or sheet identity while keeping the same name", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_named_range_wrong_scope",
      requestId: "req_named_range_wrong_scope",
      type: "named_range_update",
      traceEvent: "named_range_update_ready",
      plan: {
        operation: "retarget",
        scope: "sheet",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        name: "InputRange",
        explanation: "Retarget the sheet-scoped named range.",
        confidence: 0.93,
        requiresConfirmation: true
      },
      result: buildNamedRangeUpdateResult({
        operation: "retarget",
        scope: "workbook",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        name: "InputRange",
        explanation: "Retarget the workbook-scoped named range.",
        confidence: 0.93,
        requiresConfirmation: true
      }, {
        hostPlatform: "excel_windows",
        summary: "Wrong named range scope completion."
      })
    });
  });

  it("requires named range rename completions to report the approved new name", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_named_range_rename_wrong_details",
      requestId: "req_named_range_rename_wrong_details",
      type: "named_range_update",
      traceEvent: "named_range_update_ready",
      plan: {
        operation: "rename",
        scope: "workbook",
        name: "InputRange",
        newName: "RevenueInputRange",
        explanation: "Rename the named range for clarity.",
        confidence: 0.92,
        requiresConfirmation: true
      },
      result: buildNamedRangeUpdateResult({
        operation: "rename",
        scope: "workbook",
        name: "InputRange",
        newName: "RevenueInputRange_v2",
        explanation: "Rename the named range for clarity.",
        confidence: 0.92,
        requiresConfirmation: true
      }, {
        hostPlatform: "excel_windows",
        summary: "Wrong named range rename completion."
      })
    });
  });

  it("rejects data validation completion replay for the same approval token", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Restrict the status column to approved values.",
      confidence: 0.95,
      requiresConfirmation: true,
      replacesExistingValidation: true,
      ruleType: "list",
      values: ["Open", "Closed"]
    };

    setRunResponse(traceBus, {
      runId: "run_validation_replay",
      requestId: "req_validation_replay",
      type: "data_validation_plan",
      traceEvent: "data_validation_plan_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req_validation_replay",
      runId: "run_validation_replay",
      plan: plan as any,
      traceBus,
      config: testConfig
    });

    completeWriteback({
      requestId: "req_validation_replay",
      runId: "run_validation_replay",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: {
        kind: "data_validation_update",
        hostPlatform: "excel_windows",
        ...plan,
        summary: "Applied validation to Sheet1!B2:B20."
      },
      traceBus,
      config: testConfig
    });

    const replayCompletion = completeWriteback({
      requestId: "req_validation_replay",
      runId: "run_validation_replay",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: {
        kind: "data_validation_update",
        hostPlatform: "excel_windows",
        ...plan,
        summary: "Applied validation to Sheet1!B2:B20."
      },
      traceBus,
      config: testConfig
    });

    expect(replayCompletion).toEqual({ ok: true });
  });

  it("rejects named range completion when result kind does not match the approved plan family", () => {
    const traceBus = new TraceBus();
    const plan = {
      operation: "retarget",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      name: "InputRange",
      explanation: "Retarget the named range to the input block.",
      confidence: 0.93,
      requiresConfirmation: true
    };

    setRunResponse(traceBus, {
      runId: "run_named_range_kind_mismatch",
      requestId: "req_named_range_kind_mismatch",
      type: "named_range_update",
      traceEvent: "named_range_update_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req_named_range_kind_mismatch",
      runId: "run_named_range_kind_mismatch",
      plan: plan as any,
      traceBus,
      config: testConfig
    });

    expect(() => completeWriteback({
      requestId: "req_named_range_kind_mismatch",
      runId: "run_named_range_kind_mismatch",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: {
        kind: "data_validation_update",
        hostPlatform: "google_sheets",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        summary: "Applied validation to Sheet1!B2:D20."
      },
      traceBus,
      config: testConfig
    })).toThrow("Writeback result does not match the approved plan family.");
  });

  it("rejects named range completion replay for the same approval token", () => {
    const traceBus = new TraceBus();
    const plan = {
      operation: "retarget",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      name: "InputRange",
      explanation: "Retarget the named range to the input block.",
      confidence: 0.93,
      requiresConfirmation: true
    };

    setRunResponse(traceBus, {
      runId: "run_named_range_replay",
      requestId: "req_named_range_replay",
      type: "named_range_update",
      traceEvent: "named_range_update_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req_named_range_replay",
      runId: "run_named_range_replay",
      plan: plan as any,
      traceBus,
      config: testConfig
    });

    completeWriteback({
      requestId: "req_named_range_replay",
      runId: "run_named_range_replay",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildNamedRangeUpdateResult(plan, {
        summary: "Retargeted InputRange to Sheet1!B2:D20."
      }),
      traceBus,
      config: testConfig
    });

    const replayCompletion = completeWriteback({
      requestId: "req_named_range_replay",
      runId: "run_named_range_replay",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildNamedRangeUpdateResult(plan, {
        summary: "Retargeted InputRange to Sheet1!B2:D20."
      }),
      traceBus,
      config: testConfig
    });

    expect(replayCompletion).toEqual({ ok: true });
  });

  it("rejects data validation completion when result kind does not match the approved plan family", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Restrict the status column to approved values.",
      confidence: 0.95,
      requiresConfirmation: true,
      replacesExistingValidation: true,
      ruleType: "list",
      values: ["Open", "Closed"]
    };

    setRunResponse(traceBus, {
      runId: "run_validation_kind_mismatch",
      requestId: "req_validation_kind_mismatch",
      type: "data_validation_plan",
      traceEvent: "data_validation_plan_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req_validation_kind_mismatch",
      runId: "run_validation_kind_mismatch",
      plan: plan as any,
      traceBus,
      config: testConfig
    });

    expect(() => completeWriteback({
      requestId: "req_validation_kind_mismatch",
      runId: "run_validation_kind_mismatch",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: buildNamedRangeUpdateResult({
        operation: "retarget",
        scope: "sheet",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        name: "InputRange",
        explanation: "Retarget the named range to the input block.",
        confidence: 0.93,
        requiresConfirmation: true
      }, {
        summary: "Retargeted InputRange to Sheet1!B2:D20."
      }),
      traceBus,
      config: testConfig
    })).toThrow("Writeback result does not match the approved plan family.");
  });

  it("rejects destructive row deletion approval without a second confirmation payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      operation: "delete_rows",
      startIndex: 7,
      count: 2,
      explanation: "Delete two empty rows.",
      confidence: 0.91,
      requiresConfirmation: true,
      confirmationLevel: "destructive"
    };

    setRunResponse(traceBus, {
      runId: "run_delete_rows",
      requestId: "req_delete_rows",
      type: "sheet_structure_update",
      traceEvent: "sheet_structure_update_ready",
      plan
    });

    expect(() => approveWriteback({
      requestId: "req_delete_rows",
      runId: "run_delete_rows",
      plan: plan as any,
      traceBus,
      config: testConfig
    })).toThrow("Destructive confirmation required.");
  });

  it("approves destructive row deletion with second confirmation and persists the confirmation flag", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      operation: "delete_rows",
      startIndex: 7,
      count: 2,
      explanation: "Delete two empty rows.",
      confidence: 0.91,
      requiresConfirmation: true,
      confirmationLevel: "destructive"
    };

    setRunResponse(traceBus, {
      runId: "run_delete_rows_confirmed",
      requestId: "req_delete_rows_confirmed",
      type: "sheet_structure_update",
      traceEvent: "sheet_structure_update_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req_delete_rows_confirmed",
      runId: "run_delete_rows_confirmed",
      plan: plan as any,
      destructiveConfirmation: { confirmed: true },
      traceBus,
      config: testConfig
    });

    const completion = completeWriteback({
      requestId: "req_delete_rows_confirmed",
      runId: "run_delete_rows_confirmed",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: {
        kind: "sheet_structure_update",
        hostPlatform: "google_sheets",
        targetSheet: "Sheet1",
        operation: "delete_rows",
        startIndex: 7,
        count: 2,
        summary: "Deleted two empty rows."
      },
      traceBus,
      config: testConfig
    });

    expect(completion.ok).toBe(true);
    expect(traceBus.getRun("run_delete_rows_confirmed")?.writeback).toMatchObject({
      destructiveConfirmation: { confirmed: true },
      result: {
        kind: "sheet_structure_update",
        operation: "delete_rows"
      }
    });
  });

  it("approves and completes sheet structure updates with a typed result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      operation: "insert_rows",
      startIndex: 4,
      count: 2,
      explanation: "Insert two spacer rows above totals.",
      confidence: 0.9,
      requiresConfirmation: true,
      confirmationLevel: "standard"
    };

    setRunResponse(traceBus, {
      runId: "run_structure_typed",
      requestId: "req_structure_typed",
      type: "sheet_structure_update",
      traceEvent: "sheet_structure_update_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_structure_typed",
        runId: "run_structure_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_structure_typed",
        runId: "run_structure_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "sheet_structure_update",
          hostPlatform: "excel_windows",
          targetSheet: "Sheet1",
          operation: "insert_rows",
          startIndex: 4,
          count: 2,
          summary: "Inserted two spacer rows above totals."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_structure_typed")?.writeback?.result).toMatchObject({
      kind: "sheet_structure_update",
      targetSheet: "Sheet1",
      operation: "insert_rows"
    });
  });

  it("rejects sheet structure completion results with same-family detail mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_sheet_structure_wrong_details",
      requestId: "req_sheet_structure_wrong_details",
      type: "sheet_structure_update",
      traceEvent: "sheet_structure_update_ready",
      plan: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 4,
        count: 2,
        explanation: "Insert two spacer rows above totals.",
        confidence: 0.9,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      },
      result: {
        kind: "sheet_structure_update",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 6,
        count: 2,
        summary: "Wrong sheet structure completion."
      }
    });
  });

  it("rejects sheet structure freeze completions when frozen pane details drift from the approved plan", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_sheet_structure_freeze_wrong_details",
      requestId: "req_sheet_structure_freeze_wrong_details",
      type: "sheet_structure_update",
      traceEvent: "sheet_structure_update_ready",
      plan: {
        targetSheet: "Sheet1",
        operation: "freeze_panes",
        frozenRows: 1,
        frozenColumns: 2,
        explanation: "Freeze the header rows and key columns.",
        confidence: 0.9,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      },
      result: {
        kind: "sheet_structure_update",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        operation: "freeze_panes",
        frozenRows: 1,
        frozenColumns: 1,
        summary: "Wrong freeze pane completion."
      }
    });
  });

  it("approves and completes range sort plans with a typed result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      keys: [
        { columnRef: "Status", direction: "asc" },
        { columnRef: "Due Date", direction: "desc" }
      ],
      explanation: "Sort open items first, then latest due date.",
      confidence: 0.94,
      requiresConfirmation: true
    };

    setRunResponse(traceBus, {
      runId: "run_sort_typed",
      requestId: "req_sort_typed",
      type: "range_sort_plan",
      traceEvent: "range_sort_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_sort_typed",
        runId: "run_sort_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_sort_typed",
        runId: "run_sort_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "range_sort",
          hostPlatform: "google_sheets",
          ...plan,
          summary: "Sorted Sheet1!A1:F25 by Status asc, Due Date desc."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_sort_typed")?.writeback?.result).toMatchObject({
      kind: "range_sort",
      targetSheet: "Sheet1",
      targetRange: "A1:F25"
    });
  });

  it("rejects range sort completion results with same-family target mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_sort_wrong_details",
      requestId: "req_sort_wrong_details",
      type: "range_sort_plan",
      traceEvent: "range_sort_plan_ready",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Status", direction: "asc" },
          { columnRef: "Due Date", direction: "desc" }
        ],
        explanation: "Sort open items first, then latest due date.",
        confidence: 0.94,
        requiresConfirmation: true
      },
      result: {
        kind: "range_sort",
        hostPlatform: "google_sheets",
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Priority", direction: "asc" },
          { columnRef: "Created At", direction: "desc" }
        ],
        explanation: "Sort open items first, then latest due date.",
        confidence: 0.94,
        requiresConfirmation: true,
        summary: "Wrong sort target."
      }
    });
  });

  it("rejects range sort completion results when the applied sort keys differ on the approved target", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_sort_wrong_keys",
      requestId: "req_sort_wrong_keys",
      type: "range_sort_plan",
      traceEvent: "range_sort_plan_ready",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Status", direction: "asc" },
          { columnRef: "Due Date", direction: "desc" }
        ],
        explanation: "Sort open items first, then latest due date.",
        confidence: 0.94,
        requiresConfirmation: true
      },
      result: {
        kind: "range_sort",
        hostPlatform: "google_sheets",
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Priority", direction: "asc" },
          { columnRef: "Due Date", direction: "asc" }
        ],
        explanation: "Sort open items first, then latest due date.",
        confidence: 0.94,
        requiresConfirmation: true,
        summary: "Sorted the approved target with the wrong keys."
      }
    });
  });

  it("rejects completion when result kind does not match the approved plan family", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      operation: "insert_rows",
      startIndex: 4,
      count: 2,
      explanation: "Insert two spacer rows above totals.",
      confidence: 0.9,
      requiresConfirmation: true,
      confirmationLevel: "standard"
    };

    setRunResponse(traceBus, {
      runId: "run_result_kind_mismatch",
      requestId: "req_result_kind_mismatch",
      type: "sheet_structure_update",
      traceEvent: "sheet_structure_update_ready",
      plan
    });

    const approval = approveWriteback({
      requestId: "req_result_kind_mismatch",
      runId: "run_result_kind_mismatch",
      plan: plan as any,
      traceBus,
      config: testConfig
    });

    expect(() => completeWriteback({
      requestId: "req_result_kind_mismatch",
      runId: "run_result_kind_mismatch",
      approvalToken: approval.approvalToken,
      planDigest: approval.planDigest,
      result: {
        kind: "range_sort",
        hostPlatform: "google_sheets",
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        summary: "Sorted Sheet1!A1:F25."
      },
      traceBus,
      config: testConfig
    })).toThrow("Writeback result does not match the approved plan family.");
  });

  it("rejects partial range-write completions that do not cover the approved rectangle", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "A1:B2",
      operation: "replace_range" as const,
      values: [["North", 120], ["South", 95]],
      explanation: "Write the refreshed sales block.",
      confidence: 0.94,
      requiresConfirmation: true as const,
      overwriteRisk: "low" as const,
      shape: { rows: 2, columns: 2 }
    };

    setRunResponse(traceBus, {
      runId: "run_partial_range_write",
      requestId: "req_partial_range_write",
      type: "sheet_update",
      traceEvent: "sheet_update_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_partial_range_write",
        runId: "run_partial_range_write",
        plan
      }
    });

    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_partial_range_write",
        runId: "run_partial_range_write",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: buildRangeWriteResult(plan, {
          writtenRows: 1
        })
      }
    });

    expectRouteError(
      completion,
      409,
      "STALE_APPROVAL",
      "The approved update no longer matches the current Hermes plan."
    );
  });

  it("rejects range-write completions that reuse the approved rectangle with different write semantics", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_wrong_range_write_semantics",
      requestId: "req_wrong_range_write_semantics",
      type: "sheet_update",
      traceEvent: "sheet_update_plan_ready",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        operation: "replace_range",
        values: [["Region", "Revenue"], ["North", 1200]],
        explanation: "Write the approved sales block.",
        confidence: 0.92,
        requiresConfirmation: true,
        shape: { rows: 2, columns: 2 }
      },
      result: buildRangeWriteResult({
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        operation: "set_formulas",
        formulas: [["=UPPER(\"Region\")", "=UPPER(\"Revenue\")"], ["=LOWER(\"North\")", "=1200"]],
        explanation: "Write formulas instead of literal values.",
        confidence: 0.92,
        requiresConfirmation: true,
        shape: { rows: 2, columns: 2 }
      }, {
        hostPlatform: "google_sheets"
      })
    });
  });

  it("approves and completes range filter plans with a typed result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      conditions: [
        { columnRef: "Status", operator: "equals", value: "Open" }
      ],
      combiner: "and",
      clearExistingFilters: true,
      explanation: "Show only open items.",
      confidence: 0.93,
      requiresConfirmation: true
    };

    setRunResponse(traceBus, {
      runId: "run_filter_typed",
      requestId: "req_filter_typed",
      type: "range_filter_plan",
      traceEvent: "range_filter_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_filter_typed",
        runId: "run_filter_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_filter_typed",
        runId: "run_filter_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "range_filter",
          hostPlatform: "google_sheets",
          ...plan,
          summary: "Filtered Sheet1!A1:F25 to rows where Status equals Open."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_filter_typed")?.writeback?.result).toMatchObject({
      kind: "range_filter",
      targetSheet: "Sheet1",
      targetRange: "A1:F25"
    });
  });

  it("rejects range filter completion results with same-family detail mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_filter_wrong_details",
      requestId: "req_filter_wrong_details",
      type: "range_filter_plan",
      traceEvent: "range_filter_plan_ready",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Show only open items.",
        confidence: 0.93,
        requiresConfirmation: true
      },
      result: {
        kind: "range_filter",
        hostPlatform: "google_sheets",
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Closed" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Show only open items.",
        confidence: 0.93,
        requiresConfirmation: true,
        summary: "Applied the wrong filter rule."
      }
    });
  });

  it("approves and completes a range transfer plan with a typed result payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "RawData",
      sourceRange: "A2:C10",
      targetSheet: "Report",
      targetRange: "B5:D13",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the cleaned rows into the report sheet.",
      confidence: 0.96,
      requiresConfirmation: true,
      affectedRanges: ["RawData!A2:C10", "Report!B5:D13"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    };

    setRunResponse(traceBus, {
      runId: "run_range_transfer_typed",
      requestId: "req_range_transfer_typed",
      type: "range_transfer_plan",
      traceEvent: "range_transfer_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_range_transfer_typed",
        runId: "run_range_transfer_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_range_transfer_typed",
        runId: "run_range_transfer_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "range_transfer_update",
          operation: "range_transfer_update",
          hostPlatform: "excel_windows",
          sourceSheet: "RawData",
          sourceRange: "A2:C10",
          targetSheet: "Report",
          targetRange: "B5:D13",
          transferOperation: "copy",
          summary: "Copied RawData!A2:C10 to Report!B5:D13."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_range_transfer_typed")?.writeback?.result).toMatchObject({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:C10",
      targetSheet: "Report",
      targetRange: "B5:D13",
      transferOperation: "copy"
    });
  });

  it("accepts copy transfer completions that report the resolved target rectangle from a single-cell anchor", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Report",
      targetRange: "F2",
      operation: "copy" as const,
      pasteMode: "values" as const,
      transpose: false,
      explanation: "Copy the input block into the report anchor.",
      confidence: 0.95,
      requiresConfirmation: true as const,
      affectedRanges: ["RawData!A2:B3", "Report!F2:G3"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_range_transfer_anchor_copy",
      requestId: "req_range_transfer_anchor_copy",
      type: "range_transfer_plan",
      traceEvent: "range_transfer_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_range_transfer_anchor_copy",
        runId: "run_range_transfer_anchor_copy",
        plan
      }
    });

    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_range_transfer_anchor_copy",
        runId: "run_range_transfer_anchor_copy",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "range_transfer_update",
          operation: "range_transfer_update",
          hostPlatform: "excel_windows",
          sourceSheet: "RawData",
          sourceRange: "A2:B3",
          targetSheet: "Report",
          targetRange: "F2:G3",
          transferOperation: "copy",
          summary: "Copied RawData!A2:B3 to Report!F2:G3."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
  });

  it("accepts append transfer completions that report the resolved rows written inside the approved target block", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Report",
      targetRange: "F2:G8",
      operation: "append" as const,
      pasteMode: "values" as const,
      transpose: false,
      explanation: "Append the input block into the report staging area.",
      confidence: 0.95,
      requiresConfirmation: true as const,
      affectedRanges: ["RawData!A2:B3", "Report!F2:G8"],
      overwriteRisk: "low" as const,
      confirmationLevel: "standard" as const
    };

    setRunResponse(traceBus, {
      runId: "run_range_transfer_append_actual_range",
      requestId: "req_range_transfer_append_actual_range",
      type: "range_transfer_plan",
      traceEvent: "range_transfer_plan_ready",
      plan
    });

    const approval = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_range_transfer_append_actual_range",
        runId: "run_range_transfer_append_actual_range",
        plan
      }
    });

    expect(approval.status).toBe(200);

    const completion = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_range_transfer_append_actual_range",
        runId: "run_range_transfer_append_actual_range",
        approvalToken: (approval.body as any).approvalToken,
        planDigest: (approval.body as any).planDigest,
        result: {
          kind: "range_transfer_update",
          operation: "range_transfer_update",
          hostPlatform: "google_sheets",
          sourceSheet: "RawData",
          sourceRange: "A2:B3",
          targetSheet: "Report",
          targetRange: "F5:G6",
          transferOperation: "append",
          summary: "Appended RawData!A2:B3 into Report!F5:G6."
        }
      }
    });

    expect(completion.status).toBe(200);
    expect(completion.body).toEqual({ ok: true });
  });

  it("rejects range transfer completion results with same-family detail mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_range_transfer_wrong_details",
      requestId: "req_range_transfer_wrong_details",
      type: "range_transfer_plan",
      traceEvent: "range_transfer_plan_ready",
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:C10",
        targetSheet: "Report",
        targetRange: "B5:D13",
        operation: "copy",
        pasteMode: "values",
        transpose: false,
        explanation: "Copy the cleaned rows into the report sheet.",
        confidence: 0.96,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:C10", "Report!B5:D13"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      result: {
        kind: "range_transfer_update",
        operation: "range_transfer_update",
        hostPlatform: "excel_windows",
        sourceSheet: "OtherData",
        sourceRange: "D2:F10",
        targetSheet: "Archive",
        targetRange: "A1:C9",
        transferOperation: "move",
        summary: "Wrong transfer completion."
      }
    });
  });

  it("approves a structurally equivalent range transfer plan when request parsing normalizes key order", () => {
    const traceBus = new TraceBus();
    const storedPlan = {
      overwriteRisk: "low",
      confirmationLevel: "standard",
      affectedRanges: ["RawData!A2:C10", "Report!B5:D13"],
      requiresConfirmation: true,
      confidence: 0.96,
      explanation: "Copy the cleaned rows into the report sheet.",
      transpose: false,
      pasteMode: "values",
      targetRange: "B5:D13",
      targetSheet: "Report",
      operation: "copy",
      sourceRange: "A2:C10",
      sourceSheet: "RawData"
    };
    const approvedPlan = {
      sourceSheet: "RawData",
      sourceRange: "A2:C10",
      operation: "copy",
      targetSheet: "Report",
      targetRange: "B5:D13",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the cleaned rows into the report sheet.",
      confidence: 0.96,
      requiresConfirmation: true,
      affectedRanges: ["RawData!A2:C10", "Report!B5:D13"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    };

    expect(digestPlan(storedPlan)).not.toBe(digestPlan(approvedPlan));

    setRunResponse(traceBus, {
      runId: "run_range_transfer_digest_canonicalized",
      requestId: "req_range_transfer_digest_canonicalized",
      type: "range_transfer_plan",
      traceEvent: "range_transfer_plan_ready",
      plan: storedPlan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_range_transfer_digest_canonicalized",
        runId: "run_range_transfer_digest_canonicalized",
        plan: approvedPlan
      }
    });

    expect(approvalResponse.status).toBe(200);
    expect(traceBus.getRun("run_range_transfer_digest_canonicalized")?.writeback).toMatchObject({
      approvedPlanDigest: digestPlan({
        affectedRanges: ["RawData!A2:C10", "Report!B5:D13"],
        confidence: 0.96,
        confirmationLevel: "standard",
        explanation: "Copy the cleaned rows into the report sheet.",
        operation: "copy",
        overwriteRisk: "low",
        pasteMode: "values",
        requiresConfirmation: true,
        sourceRange: "A2:C10",
        sourceSheet: "RawData",
        targetRange: "B5:D13",
        targetSheet: "Report",
        transpose: false
      })
    });
  });

  it("rejects destructive range transfer approval without a second confirmation payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceSheet: "Staging",
      sourceRange: "A2:D12",
      targetSheet: "Archive",
      targetRange: "A2:D12",
      operation: "move",
      pasteMode: "values",
      transpose: false,
      explanation: "Move finalized rows into the archive sheet.",
      confidence: 0.95,
      requiresConfirmation: true,
      affectedRanges: ["Staging!A2:D12", "Archive!A2:D12"],
      overwriteRisk: "high",
      confirmationLevel: "destructive"
    };

    setRunResponse(traceBus, {
      runId: "run_range_transfer_destructive_missing_confirm",
      requestId: "req_range_transfer_destructive_missing_confirm",
      type: "range_transfer_plan",
      traceEvent: "range_transfer_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_range_transfer_destructive_missing_confirm",
        runId: "run_range_transfer_destructive_missing_confirm",
        plan
      }
    });

    expectRouteError(
      approvalResponse,
      400,
      "DESTRUCTIVE_CONFIRMATION_REQUIRED",
      "This update needs an explicit destructive confirmation before it can run."
    );
  });

  it("approves and completes an external data plan with typed completion state", () => {
    const traceBus = new TraceBus();
    const plan = {
      sourceType: "market_data",
      provider: "googlefinance",
      query: {
        symbol: "CURRENCY:BTCUSD",
        attribute: "price"
      },
      targetSheet: "Market Data",
      targetRange: "B2",
      formula: '=GOOGLEFINANCE("CURRENCY:BTCUSD","price")',
      explanation: "Anchor the latest BTC price in B2.",
      confidence: 0.92,
      requiresConfirmation: true,
      affectedRanges: ["Market Data!B2"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    };

    setRunResponse(traceBus, {
      runId: "run_external_data_typed",
      requestId: "req_external_data_typed",
      type: "external_data_plan",
      traceEvent: "external_data_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_external_data_typed",
        runId: "run_external_data_typed",
        plan
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_external_data_typed",
        runId: "run_external_data_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "external_data_update",
          hostPlatform: "google_sheets",
          sourceType: "market_data",
          provider: "googlefinance",
          query: {
            symbol: "CURRENCY:BTCUSD",
            attribute: "price"
          },
          targetSheet: "Market Data",
          targetRange: "B2",
          formula: '=GOOGLEFINANCE("CURRENCY:BTCUSD","price")',
          explanation: "Anchor the latest BTC price in B2.",
          confidence: 0.92,
          requiresConfirmation: true,
          affectedRanges: ["Market Data!B2"],
          overwriteRisk: "low",
          confirmationLevel: "standard",
          summary: "Anchored market data for CURRENCY:BTCUSD in Market Data!B2."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_external_data_typed")?.writeback).toMatchObject({
      result: {
        kind: "external_data_update",
        sourceType: "market_data",
        provider: "googlefinance",
        query: {
          symbol: "CURRENCY:BTCUSD",
          attribute: "price"
        },
        targetSheet: "Market Data",
        targetRange: "B2"
      }
    });
  });

  it("rejects external data completion results with same-family detail mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_external_data_wrong_details",
      requestId: "req_external_data_wrong_details",
      type: "external_data_plan",
      traceEvent: "external_data_plan_ready",
      plan: {
        sourceType: "web_table_import",
        provider: "importhtml",
        sourceUrl: "https://example.com/prices",
        selectorType: "table",
        selector: 1,
        targetSheet: "Imported Data",
        targetRange: "A1",
        formula: '=IMPORTHTML("https://example.com/prices","table",1)',
        explanation: "Import the first public table.",
        confidence: 0.87,
        requiresConfirmation: true,
        affectedRanges: ["Imported Data!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      result: {
        kind: "external_data_update",
        hostPlatform: "google_sheets",
        sourceType: "web_table_import",
        provider: "importhtml",
        sourceUrl: "https://example.com/prices",
        selectorType: "table",
        selector: 2,
        targetSheet: "Imported Data",
        targetRange: "A1",
        formula: '=IMPORTHTML("https://example.com/prices","table",2)',
        explanation: "Imported the wrong table.",
        confidence: 0.87,
        requiresConfirmation: true,
        affectedRanges: ["Imported Data!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        summary: "Wrong external data completion."
      }
    });
  });

  it("approves and completes a destructive data cleanup plan with typed completion state", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Contacts",
      targetRange: "A2:F100",
      operation: "remove_duplicate_rows",
      keyColumns: ["A", "C"],
      explanation: "Remove duplicate contacts based on name and email.",
      confidence: 0.94,
      requiresConfirmation: true,
      affectedRanges: ["Contacts!A2:F100"],
      overwriteRisk: "high",
      confirmationLevel: "destructive"
    };

    setRunResponse(traceBus, {
      runId: "run_data_cleanup_typed",
      requestId: "req_data_cleanup_typed",
      type: "data_cleanup_plan",
      traceEvent: "data_cleanup_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_data_cleanup_typed",
        runId: "run_data_cleanup_typed",
        plan,
        destructiveConfirmation: {
          confirmed: true
        }
      }
    });

    expect(approvalResponse.status).toBe(200);

    const approvedBody = approvalResponse.body as Record<string, string>;
    const completionResponse = invokeWritebackRoute({
      traceBus,
      path: "/complete",
      body: {
        requestId: "req_data_cleanup_typed",
        runId: "run_data_cleanup_typed",
        approvalToken: approvedBody.approvalToken,
        planDigest: approvedBody.planDigest,
        result: {
          kind: "data_cleanup_update",
          operation: "remove_duplicate_rows",
          hostPlatform: "google_sheets",
          targetSheet: "Contacts",
          targetRange: "A2:F100",
          keyColumns: ["A", "C"],
          explanation: "Remove duplicate contacts based on name and email.",
          confidence: 0.94,
          requiresConfirmation: true,
          affectedRanges: ["Contacts!A2:F100"],
          overwriteRisk: "high",
          confirmationLevel: "destructive",
          summary: "Removed duplicate rows from Contacts!A2:F100."
        }
      }
    });

    expect(completionResponse.status).toBe(200);
    expect(traceBus.getRun("run_data_cleanup_typed")?.writeback).toMatchObject({
      destructiveConfirmation: { confirmed: true },
      result: {
        kind: "data_cleanup_update",
        operation: "remove_duplicate_rows",
        targetSheet: "Contacts",
        targetRange: "A2:F100",
        keyColumns: ["A", "C"]
      }
    });
  });

  it("rejects data cleanup completion results with same-family detail mismatches through /complete", () => {
    expectRouteCompletionDetailMismatch({
      traceBus: new TraceBus(),
      runId: "run_data_cleanup_wrong_details",
      requestId: "req_data_cleanup_wrong_details",
      type: "data_cleanup_plan",
      traceEvent: "data_cleanup_plan_ready",
      destructiveConfirmation: { confirmed: true },
      plan: {
        targetSheet: "Contacts",
        targetRange: "A2:F100",
        operation: "remove_duplicate_rows",
        keyColumns: ["A", "C"],
        explanation: "Remove duplicate contacts based on name and email.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:F100"],
        overwriteRisk: "high",
        confirmationLevel: "destructive"
      },
      result: {
        kind: "data_cleanup_update",
        operation: "remove_duplicate_rows",
        hostPlatform: "google_sheets",
        targetSheet: "Contacts",
        targetRange: "A2:F100",
        keyColumns: ["A"],
        explanation: "Wrong cleanup semantics.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:F100"],
        overwriteRisk: "high",
        confirmationLevel: "destructive",
        summary: "Wrong cleanup completion."
      }
    });
  });

  it("rejects destructive data cleanup approval without a second confirmation payload", () => {
    const traceBus = new TraceBus();
    const plan = {
      targetSheet: "Contacts",
      targetRange: "A2:F100",
      operation: "remove_blank_rows",
      keyColumns: ["A"],
      explanation: "Remove empty rows before export.",
      confidence: 0.93,
      requiresConfirmation: true,
      affectedRanges: ["Contacts!A2:F100"],
      overwriteRisk: "medium",
      confirmationLevel: "destructive"
    };

    setRunResponse(traceBus, {
      runId: "run_data_cleanup_missing_confirm",
      requestId: "req_data_cleanup_missing_confirm",
      type: "data_cleanup_plan",
      traceEvent: "data_cleanup_plan_ready",
      plan
    });

    const approvalResponse = invokeWritebackRoute({
      traceBus,
      path: "/approve",
      body: {
        requestId: "req_data_cleanup_missing_confirm",
        runId: "run_data_cleanup_missing_confirm",
        plan
      }
    });

    expectRouteError(
      approvalResponse,
      400,
      "DESTRUCTIVE_CONFIRMATION_REQUIRED",
      "This update needs an explicit destructive confirmation before it can run."
    );
  });
});
