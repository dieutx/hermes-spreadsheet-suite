import { afterEach, describe, expect, it, vi } from "vitest";
import type {
  HermesRequest,
  HermesResponse,
  SpreadsheetContext
} from "../../contracts/src/index.ts";
import {
  buildHermesRequest,
  buildDryRunPreview,
  buildPlanHistoryPreview,
  buildStructuredPreview,
  buildDataCleanupPreview,
  buildRangeTransferPreview,
  buildSheetUpdatePreview,
  buildWriteMatrix,
  createGatewayClient,
  filterSupportedImageFiles,
  formatProofLine,
  formatTraceEvent,
  formatTraceTimeline,
  formatWritebackStatusLine,
  getFollowUpSuggestions,
  getResponseBodyText,
  getResponseConfidence,
  getResponseMetaLine,
  getRequiresConfirmation,
  getResponseWarnings,
  getStructuredPreview,
  isWritePlanResponse,
  formatDryRunSummary,
  formatHistoryEntrySummary,
  summarizeLatestTrace
} from "../src/index.ts";
import type { WritebackResult, WritePlan } from "../src/index.ts";

type ResponseFixture<T extends HermesResponse["type"]> =
  Partial<Extract<HermesResponse, { type: T }>>;

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function baseContext(overrides?: Partial<SpreadsheetContext>): SpreadsheetContext {
  return {
    selection: {
      range: "A1:F2",
      headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"],
      values: [
        ["2026-04-01", "Audio", "Cable", "North", 1, 15.5],
        ["2026-04-02", "Audio", "Adapter", "South", 2, 31]
      ],
      formulas: [
        [null, null, null, null, null, null],
        [null, null, null, null, null, null]
      ]
    },
    activeCell: {
      a1Notation: "A1",
      displayValue: "Date",
      value: "Date"
    },
    ...overrides
  };
}

function baseResponse(overrides: Partial<HermesResponse>): HermesResponse {
  return {
    schemaVersion: "1.0.0",
    type: "chat",
    requestId: "req_001",
    hermesRunId: "run_001",
    processedBy: "hermes",
    serviceLabel: "spreadsheet-gateway",
    environmentLabel: "demo-review",
    startedAt: "2026-04-19T09:00:00.000Z",
    completedAt: "2026-04-19T09:00:01.000Z",
    durationMs: 1000,
    skillsUsed: ["SelectionExplainerSkill"],
    downstreamProvider: {
      label: "openai",
      model: "gpt-5.4"
    },
    warnings: [],
    trace: [
      { event: "request_received", timestamp: "2026-04-19T09:00:00.000Z" },
      { event: "result_generated", timestamp: "2026-04-19T09:00:01.000Z" },
      { event: "completed", timestamp: "2026-04-19T09:00:01.000Z" }
    ],
    ui: {
      displayMode: "chat-first",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: false
    },
    data: {
      message: "Processed remotely.",
      confidence: 0.93
    },
    ...overrides
  } as HermesResponse;
}

describe("shared client request helpers", () => {
  it("truncates oversized request text so long chats do not break the Step 2 contract", () => {
    const oversizedMessage = "a".repeat(16_050);
    const request = buildHermesRequest({
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0",
        sessionId: "sess_001"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Sheet1"
      },
      userMessage: oversizedMessage,
      conversation: [
        { role: "assistant", content: oversizedMessage },
        { role: "user", content: "short" }
      ],
      context: baseContext(),
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true,
        supportsImageInputs: true,
        supportsWriteBackExecution: true
      },
      reviewer: {
        reviewerSafeMode: false,
        forceExtractionMode: null
      },
      confirmation: {
        state: "none"
      }
    });

    expect(request.userMessage).toHaveLength(16_000);
    expect(request.userMessage.endsWith("...")).toBe(true);
    expect(request.conversation[0]?.content).toHaveLength(16_000);
    expect(request.conversation[0]?.content.endsWith("...")).toBe(true);
  });

  it("builds the exact Step 2 request envelope for Google Sheets", () => {
    const request = buildHermesRequest({
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0",
        sessionId: "sess_001"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        workbookId: "sheet_123",
        activeSheet: "Sheet1",
        selectedRange: "A1:F2",
        locale: "en-US",
        timeZone: "Asia/Ho_Chi_Minh"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: baseContext(),
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true,
        supportsImageInputs: true,
        supportsWriteBackExecution: true
      },
      reviewer: {
        reviewerSafeMode: false,
        forceExtractionMode: null
      },
      confirmation: {
        state: "none"
      }
    });

    expect(request.schemaVersion).toBe("1.0.0");
    expect(request.requestId).toMatch(/^req_/);
    expect(request.source.channel).toBe("google_sheets");
    expect(request.host.platform).toBe("google_sheets");
    expect(request.context.selection?.range).toBe("A1:F2");
    expect(request.context.attachments).toBeUndefined();
  });

  it("builds the exact Step 2 request envelope for Excel macOS with attachments and reviewer flags", () => {
    const request = buildHermesRequest({
      requestId: "req_excel_001",
      source: {
        channel: "excel_macos",
        clientVersion: "16.97",
        sessionId: "sess_excel_001"
      },
      host: {
        platform: "excel_macos",
        workbookTitle: "Budget.xlsx",
        activeSheet: "Sheet3",
        selectedRange: "B4:C6",
        locale: "en-US",
        timeZone: "Asia/Ho_Chi_Minh"
      },
      userMessage: "Extract the attached table and put it into Sheet3 starting at B4.",
      conversation: [
        { role: "user", content: "Extract the attached table and put it into Sheet3 starting at B4." }
      ],
      context: {
        ...baseContext({
          selection: {
            headers: ["Name", "Qty"],
            values: []
          }
        }),
        attachments: [
          {
            id: "att_001",
            type: "image",
            mimeType: "image/png",
            source: "upload",
            fileName: "table.png",
            uploadToken: "upl_001"
          }
        ]
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true,
        supportsImageInputs: true,
        supportsWriteBackExecution: true
      },
      reviewer: {
        reviewerSafeMode: true,
        forceExtractionMode: "demo"
      },
      confirmation: {
        state: "none"
      }
    });

    expect(request.requestId).toBe("req_excel_001");
    expect(request.source.channel).toBe("excel_macos");
    expect(request.host.platform).toBe("excel_macos");
    expect(request.context.attachments?.[0]?.source).toBe("upload");
    expect(request.reviewer).toEqual({
      reviewerSafeMode: true,
      forceExtractionMode: "demo"
    });
  });

  it("strips non-contract conversation fields before building the request payload", () => {
    const request = buildHermesRequest({
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0",
        sessionId: "sess_001"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Sheet1"
      },
      userMessage: "Continue.",
      conversation: [
        {
          role: "assistant",
          content: "Existing reply",
          runId: "run_001",
          requestId: "req_001",
          response: { type: "chat" },
          trace: [{ event: "completed" }]
        } as any,
        {
          role: "user",
          content: "Continue.",
          selectedRange: "A1:F10"
        } as any
      ],
      context: baseContext(),
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true,
        supportsImageInputs: true,
        supportsWriteBackExecution: true
      },
      reviewer: {
        reviewerSafeMode: false,
        forceExtractionMode: null
      },
      confirmation: {
        state: "none"
      }
    });

    expect(request.conversation).toEqual([
      {
        role: "assistant",
        content: "Existing reply"
      },
      {
        role: "user",
        content: "Continue."
      }
    ]);
  });
});

describe("shared client render helpers", () => {
  it("surfaces structured JSON error recovery guidance from gateway responses", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({
      error: {
        message: "I can't access that uploaded file anymore.",
        userAction: "Reattach the file, then retry."
      }
    }), {
      status: 400,
      headers: {
        "content-type": "application/json"
      }
    }));
    vi.stubGlobal("fetch", fetchMock);

    const client = createGatewayClient("http://localhost:18787");

    await expect(client.startRun({} as HermesRequest)).rejects.toThrow(
      "I can't access that uploaded file anymore.\n\nReattach the file, then retry."
    );
  });

  it("fails with a user-facing invalid-JSON message when the gateway returns malformed success JSON", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response("{broken", {
      status: 200,
      headers: {
        "content-type": "application/json"
      }
    })));

    const client = createGatewayClient("http://localhost:18787");

    await expect(client.pollRun("run_123")).rejects.toThrow(
      "The Hermes service returned a response the client could not use.\n\n" +
      "Retry the request, then reload the client if it keeps happening."
    );
  });

  it("sanitizes unexpected HTML error pages from the gateway into a user-facing transport error", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response("<!doctype html><html><body>502</body></html>", {
      status: 502,
      headers: {
        "content-type": "text/html"
      }
    })));

    const client = createGatewayClient("http://localhost:18787");

    await expect(client.pollRun("run_123")).rejects.toThrow(
      "The Hermes service returned an unexpected error page.\n\n" +
      "Retry the request, then check the Hermes gateway if it keeps happening."
    );
  });

  it("sanitizes sensitive raw text gateway errors from shared-client responses", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(
      "Error: failed at /root/hermes/src/server.ts:42\\nAPPROVAL_SECRET=super-secret",
      {
        status: 500,
        headers: {
          "content-type": "text/plain"
        }
      }
    )));

    const client = createGatewayClient("http://localhost:18787");

    await expect(client.pollRun("run_123")).rejects.toThrow(
      "Hermes gateway request failed with HTTP 500."
    );
  });

  it("fails closed before fetch when the gateway client base url is invalid or non-http", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { "content-type": "application/json" }
    }));
    vi.stubGlobal("fetch", fetchMock);

    const client = createGatewayClient("javascript:alert(1)");

    await expect(client.startRun({} as HermesRequest)).rejects.toThrow(
      "Hermes gateway URL is not configured."
    );
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("surfaces legacy string-shaped JSON errors from gateway responses", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({
      error: "That Hermes request is no longer available.",
      userAction: "Send the request again from the spreadsheet if you need a fresh result."
    }), {
      status: 404,
      headers: {
        "content-type": "application/json"
      }
    })));

    const client = createGatewayClient("http://localhost:18787");

    await expect(client.pollRun("run_123")).rejects.toThrow(
      "That Hermes request is no longer available.\n\nSend the request again from the spreadsheet if you need a fresh result."
    );
  });

  it("includes sessionId when polling runs and traces", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({
      runId: "run_123",
      requestId: "req_123",
      status: "processing",
      nextIndex: 0,
      events: [],
      startedAt: "2026-04-22T00:00:00.000Z"
    }), {
      status: 200,
      headers: {
        "content-type": "application/json"
      }
    }));
    vi.stubGlobal("fetch", fetchMock);

    const client = createGatewayClient("http://localhost:18787");
    await (client.pollRun as any)("run_123", "req_123", "sess_123");
    await (client.pollTrace as any)("run_123", 0, "req_123", "sess_123");

    expect(fetchMock.mock.calls[0]?.[0]).toBe(
      "http://localhost:18787/api/requests/run_123?requestId=req_123&sessionId=sess_123"
    );
    expect(fetchMock.mock.calls[1]?.[0]).toBe(
      "http://localhost:18787/api/trace/run_123?after=0&requestId=req_123&sessionId=sess_123"
    );
  });

  it("encodes run identifiers in request and trace polling paths", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({
      runId: "run/../unsafe?x=1",
      requestId: "req_123",
      status: "processing",
      nextIndex: 0,
      events: [],
      startedAt: "2026-04-22T00:00:00.000Z"
    }), {
      status: 200,
      headers: {
        "content-type": "application/json"
      }
    }));
    vi.stubGlobal("fetch", fetchMock);

    const client = createGatewayClient("http://localhost:18787");
    await (client.pollRun as any)("run/../unsafe?x=1", "req_123", "sess_123");
    await (client.pollTrace as any)("run/../unsafe?x=1", 0, "req_123", "sess_123");

    expect(fetchMock.mock.calls[0]?.[0]).toBe(
      "http://localhost:18787/api/requests/run%2F..%2Funsafe%3Fx%3D1?requestId=req_123&sessionId=sess_123"
    );
    expect(fetchMock.mock.calls[1]?.[0]).toBe(
      "http://localhost:18787/api/trace/run%2F..%2Funsafe%3Fx%3D1?after=0&requestId=req_123&sessionId=sess_123"
    );
  });

  it("preserves executionId from writeback approval responses and forwards workbook-scoped approval fields", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({
      requestId: "req_approve_001",
      runId: "run_approve_001",
      executionId: "exec_approve_001",
      approvalToken: "token_approve_001",
      planDigest: "digest_approve_001",
      approvedAt: "2026-04-22T11:30:00.000Z"
    }), {
      status: 200,
      headers: {
        "content-type": "application/json"
      }
    }));
    vi.stubGlobal("fetch", fetchMock);

    const client = createGatewayClient("http://localhost:18787/");
    const input: Parameters<typeof client.approveWrite>[0] = {
      requestId: "req_approve_001",
      runId: "run_approve_001",
      workbookSessionKey: "excel_windows::workbook-123",
      destructiveConfirmation: { confirmed: true },
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        values: [
          ["Date", "Revenue"],
          ["2026-04-01", 42]
        ],
        confidence: 0.92,
        requiresConfirmation: true,
        warnings: [],
        confirmationLevel: "destructive",
        explanation: "Write a small summary table."
      }
    };

    const approval = await client.approveWrite(input);

    expect(approval.executionId).toBe("exec_approve_001");
    expect(fetchMock).toHaveBeenCalledWith(
      "http://localhost:18787/api/writeback/approve",
      expect.objectContaining({
        method: "POST",
        headers: { "content-type": "application/json" }
      })
    );
    expect(JSON.parse(String(fetchMock.mock.calls[0][1]?.body))).toMatchObject({
      requestId: "req_approve_001",
      runId: "run_approve_001",
      workbookSessionKey: "excel_windows::workbook-123",
      destructiveConfirmation: { confirmed: true }
    });
  });

  it("rejects malformed writeback approval responses that omit executionId", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({
      requestId: "req_approve_001",
      runId: "run_approve_001",
      approvalToken: "token_approve_001",
      planDigest: "digest_approve_001",
      approvedAt: "2026-04-22T11:30:00.000Z"
    }), {
      status: 200,
      headers: {
        "content-type": "application/json"
      }
    })));

    const client = createGatewayClient("http://localhost:18787");

    await expect(client.approveWrite({
      requestId: "req_approve_001",
      runId: "run_approve_001",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        values: [
          ["Date", "Revenue"],
          ["2026-04-01", 42]
        ],
        confidence: 0.92,
        requiresConfirmation: true,
        warnings: [],
        confirmationLevel: "standard",
        explanation: "Write a small summary table."
      }
    })).rejects.toThrow(
      "The Hermes service returned a writeback approval response the client could not use.\n\nRetry the approval, then reload the client if it keeps happening."
    );
  });

  it("rejects malformed writeback completion responses that do not confirm success", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({
      ok: false
    }), {
      status: 200,
      headers: {
        "content-type": "application/json"
      }
    })));

    const client = createGatewayClient("http://localhost:18787");

    await expect(client.completeWrite({
      requestId: "req_complete_001",
      runId: "run_complete_001",
      approvalToken: "token_complete_001",
      planDigest: "digest_complete_001",
      result: {
        kind: "range_write",
        hostPlatform: "excel_windows",
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        operation: "replace_range",
        values: [["Date", "Revenue"], ["2026-04-01", 42]],
        explanation: "Write a small summary table.",
        confidence: 0.92,
        requiresConfirmation: true,
        shape: {
          rows: 2,
          columns: 2
        },
        writtenRows: 2,
        writtenColumns: 2
      }
    })).rejects.toThrow(
      "The Hermes service returned a writeback completion response the client could not use.\n\nRetry the writeback completion, then reload the client if it keeps happening."
    );
  });

  it("renders a composite preview with step flags", () => {
    const response = {
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_sort",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:F50",
              hasHeader: true,
              keys: [{ columnRef: "Revenue", direction: "desc" }],
              explanation: "Sort by revenue.",
              confidence: 0.91,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Run the workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    };

    expect(buildStructuredPreview(response as any)).toMatchObject({
      kind: "composite_plan",
      stepCount: 1,
      destructiveConfirmationRequired: false,
      reversible: true,
      dryRunRequired: false,
      steps: [
        expect.objectContaining({
          stepId: "step_sort",
          destructive: false,
          reversible: true,
          skippedIfDependenciesFail: false
        })
      ]
    });
  });

  it("renders composite update bodies and dry-run/history summaries", () => {
    const response = baseResponse({
      type: "composite_update",
      processedBy: "host",
      data: {
        operation: "composite_update",
        executionId: "exec_001",
        stepResults: [
          {
            stepId: "step_sort",
            status: "completed",
            summary: "Sorted Sales!A1:F50."
          }
        ],
        summary: "Completed 1-step composite execution."
      }
    });

    expect(getResponseBodyText(response)).toBe("Completed 1-step composite execution.");
    const dryRunResult = {
      simulated: true,
      planDigest: "digest_001",
      workbookSessionKey: "excel_windows::workbook-123",
      steps: [
        {
          stepId: "step_sort",
          status: "simulated",
          summary: "Will sort rows."
        }
      ],
      predictedAffectedRanges: ["Sales!A1:F50"],
      predictedSummaries: ["Will sort Sales!A1:F50 by Revenue descending."],
      overwriteRisk: "low",
      reversible: true,
      expiresAt: "2026-04-20T13:00:00.000Z"
    } as any;
    expect(formatDryRunSummary(dryRunResult)).toContain("Will sort Sales!A1:F50");
    expect(buildDryRunPreview(dryRunResult)).toMatchObject({
      kind: "dry_run_result",
      simulated: true,
      steps: [
        expect.objectContaining({
          stepId: "step_sort",
          status: "simulated"
        })
      ]
    });
    const historyEntry = {
      status: "completed",
      summary: "Completed 2-step composite execution.",
      timestamp: "2026-04-20T12:00:00.000Z",
      undoEligible: true,
      redoEligible: false
    } as any;
    expect(formatHistoryEntrySummary(historyEntry)).toBe("Completed 2-step composite execution.");
    expect(buildPlanHistoryPreview({
      entries: [historyEntry]
    } as any)).toMatchObject({
      kind: "plan_history_page",
      entries: [
        expect.objectContaining({
          summary: "Completed 2-step composite execution."
        })
      ]
    });
  });

  it("sanitizes technical dry-run unsupported reasons before rendering them", () => {
    const dryRunResult = {
      simulated: false,
      planDigest: "digest_unsupported",
      workbookSessionKey: "excel_windows::workbook-123",
      steps: [],
      predictedAffectedRanges: [],
      predictedSummaries: [],
      overwriteRisk: "low",
      reversible: false,
      expiresAt: "2026-04-20T13:00:00.000Z",
      unsupportedReason: "Google Sheets host cannot find chart field in header row: Revenue."
    } as any;

    expect(formatDryRunSummary(dryRunResult)).toBe(
      "This preview needs a table with clear, matching column headers."
    );
    expect(buildDryRunPreview(dryRunResult)).toMatchObject({
      kind: "dry_run_result",
      unsupportedReason: "This preview needs a table with clear, matching column headers.",
      summary: "This preview needs a table with clear, matching column headers."
    });
  });

  it("renders composite previews for sheet import steps without assuming explanation exists", () => {
    const response = baseResponse({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_import",
            dependsOn: [],
            continueOnError: false,
            plan: {
              sourceAttachmentId: "att_100",
              targetSheet: "Imported Table",
              targetRange: "A1:C4",
              headers: ["Date", "Item", "Amount"],
              values: [
                ["2026-04-01", "Cable", 15.5],
                ["2026-04-02", "Adapter", 23],
                ["2026-04-03", "Speaker", 91]
              ],
              confidence: 0.89,
              warnings: [],
              requiresConfirmation: true,
              extractionMode: "real",
              shape: {
                rows: 4,
                columns: 3
              }
            }
          }
        ],
        explanation: "Import and prepare the extracted table.",
        confidence: 0.88,
        requiresConfirmation: true,
        affectedRanges: ["Imported Table!A1:C4"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });

    expect(buildStructuredPreview(response as any)).toMatchObject({
      kind: "composite_plan",
      steps: [
        expect.objectContaining({
          stepId: "step_import",
          summary: "Import extracted data into Imported Table!A1:C4."
        })
      ]
    });
  });

  it("formats Wave 6 control trace labels", () => {
    expect(formatTraceEvent({
      event: "composite_plan_ready",
      timestamp: "2026-04-20T09:00:01.000Z"
    } as any)).toBe("Composite plan ready");
    expect(formatTraceEvent({
      event: "composite_update_ready",
      timestamp: "2026-04-20T09:00:02.000Z"
    } as any)).toBe("Composite update ready");
    expect(formatTraceEvent({
      event: "dry_run_requested",
      timestamp: "2026-04-20T09:00:03.000Z"
    } as any)).toBe("Dry run requested");
    expect(formatTraceEvent({
      event: "history_requested",
      timestamp: "2026-04-20T09:00:04.000Z"
    } as any)).toBe("History requested");
    expect(formatTraceEvent({
      event: "external_data_plan_ready",
      timestamp: "2026-04-20T09:00:05.000Z"
    } as any)).toBe("External data plan ready");
  });

  it("renders safe proof and trace metadata without hidden reasoning fields", () => {
    const response = baseResponse({
      requestId: "req_safe_001",
      hermesRunId: "run_safe_001"
    });

    const proof = formatProofLine(response);
    expect(proof).toContain("Processed by Hermes");
    expect(proof).toContain("requestId req_safe_001");
    expect(proof).toContain("hermesRunId run_safe_001");
    expect(proof).toContain("service spreadsheet-gateway");
    expect(proof).toContain("environment demo-review");
    expect(proof).toContain("1000ms");
    expect(proof).not.toContain("prompt");
    expect(proof).not.toContain("stack");

    expect(summarizeLatestTrace(response.trace)).toBe("Completed");
    expect(formatTraceTimeline(response.trace)).toBe(
      "Request received by Hermes -> Result generated -> Completed"
    );
  });

  it("redacts unsafe proof metadata in shared-client meta lines", () => {
    const response = baseResponse({
      skillsUsed: [
        "SelectionExplainerSkill",
        "/srv/hermes/private-tool.ts",
        "HERMES_API_SERVER_KEY=secret"
      ],
      downstreamProvider: {
        label: "https://internal.example/provider",
        model: "gpt-5 HERMES_API_SERVER_KEY=secret"
      }
    });

    const metaLine = getResponseMetaLine(response);

    expect(metaLine).toContain("skills SelectionExplainerSkill");
    expect(metaLine).not.toContain("HERMES_API_SERVER_KEY");
    expect(metaLine).not.toContain("/srv/hermes");
    expect(metaLine).not.toContain("internal.example");
    expect(metaLine).not.toContain("provider https://internal");
  });

  it("formats the wave 2 trace labels", () => {
    expect(formatTraceTimeline([
      { event: "data_validation_plan_ready", timestamp: "2026-04-20T09:00:01.000Z" },
      { event: "named_range_update_ready", timestamp: "2026-04-20T09:00:02.000Z" }
    ])).toBe("Data validation plan ready -> Named range update ready");
  });

  it("formats the wave 3 conditional format trace label", () => {
    expect(formatTraceTimeline([
      { event: "conditional_format_plan_ready", timestamp: "2026-04-20T09:00:03.000Z" }
    ])).toBe("Conditional format plan ready");
  });

  it("formats the wave 4 transfer and cleanup trace labels", () => {
    expect(formatTraceTimeline([
      { event: "range_transfer_plan_ready", timestamp: "2026-04-20T09:00:04.000Z" },
      { event: "data_cleanup_plan_ready", timestamp: "2026-04-20T09:00:05.000Z" }
    ])).toBe("Range transfer plan ready -> Data cleanup plan ready");
  });

  it("treats chat-only analysis reports as non-write plans without confirmation", () => {
    const response = baseResponse({
      type: "analysis_report_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        outputMode: "chat_only",
        sections: [
          {
            type: "summary_stats",
            title: "Revenue summary",
            summary: "Average revenue is 12,500.",
            sourceRanges: ["Sales!A1:F50"]
          }
        ],
        explanation: "Summarize the selected range.",
        confidence: 0.92,
        requiresConfirmation: false,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "none",
        confirmationLevel: "standard"
      }
    } satisfies ResponseFixture<"analysis_report_plan">);

    expect(isWritePlanResponse(response)).toBe(false);
    expect(getRequiresConfirmation(response)).toBe(false);
  });

  it("treats materialized analysis reports as write plans", () => {
    const response = baseResponse({
      type: "analysis_report_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        outputMode: "materialize_report",
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
        explanation: "Write a report sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    } satisfies ResponseFixture<"analysis_report_plan">);

    expect(isWritePlanResponse(response)).toBe(true);
    expect(getRequiresConfirmation(response)).toBe(true);
  });

  it("renders a pivot preview with explicit group and value metadata", () => {
    const response = baseResponse({
      type: "pivot_table_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
        explanation: "Build a pivot.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    } satisfies ResponseFixture<"pivot_table_plan">);

    expect(getStructuredPreview(response)).toMatchObject({
      kind: "pivot_table_plan",
      sourceSheet: "Sales",
      targetSheet: "Sales Pivot",
      rowGroups: ["Region", "Rep"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }]
    });
    expect(getResponseBodyText(response)).toContain("pivot");
    expect(isWritePlanResponse(response)).toBe(true);
  });

  it("renders external data plans as confirmable write previews", () => {
    const response = baseResponse({
      type: "external_data_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
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
      }
    } satisfies ResponseFixture<"external_data_plan">);

    expect(getStructuredPreview(response)).toEqual({
      kind: "external_data_plan",
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
      summary: "Will anchor GOOGLEFINANCE market data for CURRENCY:BTCUSD at Market Data!B2.",
      details: [
        "Source type: market_data.",
        "Provider: googlefinance.",
        "Target sheet: Market Data.",
        "Target range: B2.",
        "Symbol: CURRENCY:BTCUSD.",
        "Attribute: price.",
        "Formula: =GOOGLEFINANCE(\"CURRENCY:BTCUSD\",\"price\").",
        "Affected ranges: Market Data!B2.",
        "Overwrite risk: low.",
        "Confirmation level: standard."
      ]
    });
    expect(getResponseBodyText(response)).toBe(
      "Will anchor GOOGLEFINANCE market data for CURRENCY:BTCUSD at Market Data!B2."
    );
    expect(getResponseConfidence(response)).toBe(0.92);
    expect(getRequiresConfirmation(response)).toBe(true);
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseMetaLine(response)).toContain("confidence 92%");
    expect(getResponseMetaLine(response)).toContain("confirmation required");
  });

  it("renders chart update body text directly from the summary", () => {
    const response = baseResponse({
      type: "chart_update",
      processedBy: "host",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        operation: "chart_update",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        summary: "Created line chart on Sales Chart!A1."
      }
    } satisfies ResponseFixture<"chart_update">);

    expect(getResponseBodyText(response)).toBe("Created line chart on Sales Chart!A1.");
  });

  it("exposes Wave 6 gateway client methods", async () => {
    const client = createGatewayClient("http://localhost:18787");

    expect(typeof client.dryRunPlan).toBe("function");
    expect(typeof client.listPlanHistory).toBe("function");
    expect(typeof client.prepareUndoExecution).toBe("function");
    expect(typeof client.undoExecution).toBe("function");
    expect(typeof client.prepareRedoExecution).toBe("function");
    expect(typeof client.redoExecution).toBe("function");
  });

  it("sends session ids on execution-control gateway client calls", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ entries: [] }), {
      status: 200,
      headers: { "content-type": "application/json" }
    }));
    vi.stubGlobal("fetch", fetchMock);
    const client = createGatewayClient("http://localhost:18787");

    await client.dryRunPlan({
      requestId: "req_dry_001",
      runId: "run_dry_001",
      sessionId: "sess_client_001",
      plan: { steps: [] }
    } as any);
    await client.listPlanHistory({
      workbookSessionKey: "excel_windows::workbook-123",
      sessionId: "sess_client_001",
      limit: 5
    } as any);
    await client.prepareUndoExecution({
      requestId: "req_undo_001",
      workbookSessionKey: "excel_windows::workbook-123",
      sessionId: "sess_client_001",
      executionId: "exec_001"
    });
    await client.redoExecution({
      requestId: "req_redo_001",
      workbookSessionKey: "excel_windows::workbook-123",
      sessionId: "sess_client_001",
      executionId: "exec_undo_001"
    });

    expect(JSON.parse(String(fetchMock.mock.calls[0][1]?.body))).toMatchObject({
      sessionId: "sess_client_001"
    });
    expect(String(fetchMock.mock.calls[1][0])).toBe(
      "http://localhost:18787/api/execution/history?workbookSessionKey=excel_windows%3A%3Aworkbook-123&sessionId=sess_client_001&limit=5"
    );
    expect(JSON.parse(String(fetchMock.mock.calls[2][1]?.body))).toMatchObject({
      sessionId: "sess_client_001"
    });
    expect(JSON.parse(String(fetchMock.mock.calls[3][1]?.body))).toMatchObject({
      sessionId: "sess_client_001"
    });
  });

  it("renders chat responses as message-first with follow-up suggestions", () => {
    const response = baseResponse({
      type: "chat",
      data: {
        message: "The selection is a sales summary table.",
        followUpSuggestions: ["Explain Revenue", "Suggest a formula"],
        confidence: 0.98
      }
    });

    expect(getResponseBodyText(response)).toBe("The selection is a sales summary table.");
    expect(getFollowUpSuggestions(response)).toEqual(["Explain Revenue", "Suggest a formula"]);
    expect(getStructuredPreview(response)).toBeNull();
    expect(getResponseMetaLine(response)).toContain("confidence 98%");
  });

  it("renders formula responses with a formula preview card", () => {
    const response = baseResponse({
      type: "formula",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        intent: "suggest",
        targetCell: "F12",
        formula: "=SUMIFS(F2:F11, D2:D11, \"North\")",
        formulaLanguage: "google_sheets",
        explanation: "This sums Revenue where Region equals North.",
        confidence: 0.95,
        requiresConfirmation: true
      }
    });

    const preview = getStructuredPreview(response);
    expect(preview).toEqual({
      kind: "formula",
      intent: "suggest",
      formula: "=SUMIFS(F2:F11, D2:D11, \"North\")",
      formulaLanguage: "google_sheets",
      targetCell: "F12",
      explanation: "This sums Revenue where Region equals North.",
      alternateFormulas: undefined
    });
    expect(getResponseMetaLine(response)).toContain("confirmation required");
  });

  it("preserves sheet_import_plan header/value/shape semantics in preview rendering", () => {
    const response = baseResponse({
      type: "sheet_import_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceAttachmentId: "att_100",
        targetSheet: "Imported Table",
        targetRange: "A1:C4",
        headers: ["Date", "Item", "Amount"],
        values: [
          ["2026-04-01", "Cable", 15.5],
          ["2026-04-02", "Adapter", 23],
          ["2026-04-03", "Speaker", 91]
        ],
        confidence: 0.89,
        warnings: [],
        requiresConfirmation: true,
        extractionMode: "real",
        shape: {
          rows: 4,
          columns: 3
        }
      }
    });

    const preview = getStructuredPreview(response);
    if (!preview || preview.kind !== "sheet_import_plan") {
      throw new Error("Expected a sheet_import_plan preview.");
    }

    expect(preview.targetRange).toBe("A1:C4");
    expect(preview.table.headers).toEqual(["Date", "Item", "Amount"]);
    expect(preview.table.rows).toEqual([
      ["2026-04-01", "Cable", 15.5],
      ["2026-04-02", "Adapter", 23],
      ["2026-04-03", "Speaker", 91]
    ]);
    expect(preview.shape.rows).toBe(4);
    expect(preview.shape.columns).toBe(3);
    expect(buildWriteMatrix(response.data)).toEqual([
      ["Date", "Item", "Amount"],
      ["2026-04-01", "Cable", 15.5],
      ["2026-04-02", "Adapter", 23],
      ["2026-04-03", "Speaker", 91]
    ]);
    expect(isWritePlanResponse(response)).toBe(true);
  });

  it("renders sheet_update previews using the exact full target rectangle", () => {
    const response = baseResponse({
      type: "sheet_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet3",
        targetRange: "B4:C5",
        operation: "replace_range",
        values: [
          ["North", 12],
          ["South", 20]
        ],
        explanation: "Prepared a safe two-row update.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low",
        shape: {
          rows: 2,
          columns: 2
        }
      }
    });

    const preview = getStructuredPreview(response);
    if (!preview || preview.kind !== "sheet_update") {
      throw new Error("Expected a sheet_update preview.");
    }

    expect(preview.targetRange).toBe("B4:C5");
    expect(preview.matrixKind).toBe("values");
    expect(preview.table.rows).toEqual([
      ["North", 12],
      ["South", 20]
    ]);
    expect(buildWriteMatrix(response.data)).toEqual([
      ["North", 12],
      ["South", 20]
    ]);
  });

  it("preserves plan shape for an explicitly present formulas matrix", () => {
    const preview = buildSheetUpdatePreview({
      targetSheet: "Sheet3",
      targetRange: "B4:C5",
      operation: "set_formulas",
      formulas: [["=SUM(B4:B4)"]],
      explanation: "Prepared a formulas-only update with an explicit matrix.",
      confidence: 0.92,
      requiresConfirmation: true,
      overwriteRisk: "low",
      shape: {
        rows: 1,
        columns: 1
      }
    } as never);

    expect(preview.matrixKind).toBe("formulas");
    expect(preview.table.shape).toEqual({
      rows: 1,
      columns: 1
    });
  });

  it("renders mixed sheet_update previews without omitting any matrix sections", () => {
    const response = baseResponse({
      type: "sheet_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet3",
        targetRange: "B4:C5",
        operation: "mixed_update",
        values: [
          ["North", 12],
          ["South", 20]
        ],
        formulas: [
          ["=SUM(B4:B5)", "=SUM(C4:C5)"]
        ],
        notes: [
          ["North summary", "South summary"]
        ],
        explanation: "Prepared a mixed update with values, formulas, and notes.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low",
        shape: {
          rows: 2,
          columns: 2
        }
      }
    });

    const preview = getStructuredPreview(response);
    if (!preview || preview.kind !== "sheet_update") {
      throw new Error("Expected a sheet_update preview.");
    }

    expect(preview.matrixKind).toBe("mixed_update");
    expect(preview.table.rows).toEqual([
      ["values"],
      ["North", 12],
      ["South", 20],
      ["formulas"],
      ["=SUM(B4:B5)", "=SUM(C4:C5)"],
      ["notes"],
      ["North summary", "South summary"]
    ]);
    expect(buildWriteMatrix(response.data)).toEqual([
      ["values"],
      ["North", 12],
      ["South", 20],
      ["formulas"],
      ["=SUM(B4:B5)", "=SUM(C4:C5)"],
      ["notes"],
      ["North summary", "South summary"]
    ]);
  });

  it("renders workbook structure updates as confirmable write plans", () => {
    const response = baseResponse({
      type: "workbook_structure_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "create_sheet",
        sheetName: "New Sheet",
        position: "end",
        explanation: "Create a new sheet at the end of the workbook.",
        confidence: 0.94,
        requiresConfirmation: true,
        overwriteRisk: "none"
      }
    } satisfies ResponseFixture<"workbook_structure_update">);

    const preview = getStructuredPreview(response);
    expect(preview).toEqual({
      kind: "workbook_structure_update",
      operation: "create_sheet",
      sheetName: "New Sheet",
      position: "end",
      explanation: "Create a new sheet at the end of the workbook.",
      overwriteRisk: "none"
    });
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseBodyText(response)).toContain("New Sheet");
  });

  it("builds a sheet structure summary preview", () => {
    const response = baseResponse({
      type: "sheet_structure_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 7,
        count: 3,
        explanation: "Insert three rows above the totals block.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    } satisfies ResponseFixture<"sheet_structure_update">);

    const preview = getStructuredPreview(response);
    expect(preview).toMatchObject({
      kind: "sheet_structure_update",
      targetSheet: "Sheet1",
      operation: "insert_rows",
      confirmationLevel: "standard",
      summary: "Insert three rows above the totals block."
    });
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseBodyText(response)).toContain("sheet structure update");
  });

  it("builds a range sort summary preview", () => {
    const response = baseResponse({
      type: "range_sort_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
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
      }
    } satisfies ResponseFixture<"range_sort_plan">);

    const preview = getStructuredPreview(response);
    expect(preview).toEqual({
      kind: "range_sort_plan",
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      keys: [
        { columnRef: "Status", direction: "asc" },
        { columnRef: "Due Date", direction: "desc" }
      ],
      summary: "Sort open items first, then latest due date."
    });
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseBodyText(response)).toContain("sort plan");
  });

  it("builds a range filter summary preview", () => {
    const response = baseResponse({
      type: "range_filter_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "Revenue", operator: "greaterThan", value: 1000 }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Filter open rows with revenue above 1000.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    } satisfies ResponseFixture<"range_filter_plan">);

    const preview = getStructuredPreview(response);
    expect(preview).toEqual({
      kind: "range_filter_plan",
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      conditions: [
        { columnRef: "Status", operator: "equals", value: "Open" },
        { columnRef: "Revenue", operator: "greaterThan", value: 1000 }
      ],
      combiner: "and",
      clearExistingFilters: true,
      summary: "Filter open rows with revenue above 1000."
    });
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseBodyText(response)).toContain("filter plan");
  });

  it("renders a non-lossy conditional formatting preview", () => {
    const response = baseResponse({
      type: "conditional_format_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "replace_all_on_target",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc",
          textColor: "#990000",
          bold: true
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: true
      }
    } satisfies ResponseFixture<"conditional_format_plan">);

    const preview = getStructuredPreview(response);
    if (!preview || preview.kind !== "conditional_format_plan") {
      throw new Error("Expected a conditional_format_plan preview.");
    }

    expect(preview.summary).toBe("Will replace all conditional formatting on Sheet1!B2:B20.");
    expect(preview.summary).toContain("Sheet1!B2:B20");
    expect(preview.details.join(" ")).toContain("text contains \"overdue\"");
    expect(preview.details.join(" ")).toContain("background #ffcccc");
    expect(preview.targetSheet).toBe("Sheet1");
    expect(preview.targetRange).toBe("B2:B20");
    expect(preview.managementMode).toBe("replace_all_on_target");
    expect(preview.ruleType).toBe("text_contains");
    expect(preview.affectedRanges).toEqual(["Sheet1!B2:B20"]);
    expect(preview.replacesExistingRules).toBe(true);
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseBodyText(response)).toContain("conditional formatting");
  });

  it("renders a clear_on_target conditional formatting preview without rule payload", () => {
    const response = baseResponse({
      type: "conditional_format_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "clear_on_target",
        explanation: "Clear conditional formatting from the target range.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: true
      }
    } satisfies ResponseFixture<"conditional_format_plan">);

    const preview = getStructuredPreview(response);
    if (!preview || preview.kind !== "conditional_format_plan") {
      throw new Error("Expected a conditional_format_plan preview.");
    }

    expect(preview.summary).toBe("Will clear conditional formatting on Sheet1!B2:B20.");
    expect(preview.details).toContain("Rule: clear existing conditional formatting.");
    expect(preview.details).toContain("Style: not applicable.");
    expect("ruleType" in preview).toBe(false);
    expect("style" in preview).toBe(false);
  });

  it("covers the wave 1 plan and result shapes at runtime", () => {
    const sheetStructurePlan = {
      targetSheet: "Sheet1",
      operation: "insert_rows",
      startIndex: 7,
      count: 3,
      explanation: "Insert three rows above the totals block.",
      confidence: 0.91,
      requiresConfirmation: true,
      confirmationLevel: "standard"
    } satisfies WritePlan;

    const rangeSortPlan = {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      keys: [{ columnRef: "Status", direction: "asc" }],
      explanation: "Sort open items first.",
      confidence: 0.94,
      requiresConfirmation: true
    } satisfies WritePlan;

    const rangeFilterPlan = {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      conditions: [{ columnRef: "Status", operator: "equals", value: "Open" }],
      combiner: "and",
      clearExistingFilters: true,
      explanation: "Filter open rows.",
      confidence: 0.9,
      requiresConfirmation: true
    } satisfies WritePlan;

    const sheetStructureResult = {
      kind: "sheet_structure_update",
      hostPlatform: "google_sheets",
      operation: "insert_rows",
      targetSheet: "Sheet1",
      summary: "Inserted three rows above the totals block."
    } satisfies WritebackResult;

    const rangeSortResult = {
      kind: "range_sort",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      summary: "Sorted Sheet1!A1:F25."
    } satisfies WritebackResult;

    const rangeFilterResult = {
      kind: "range_filter",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      conditions: [
        { columnRef: "Status", operator: "equals", value: "Open" }
      ],
      combiner: "and",
      clearExistingFilters: true,
      explanation: "Filter open rows.",
      confidence: 0.89,
      requiresConfirmation: true,
      summary: "Applied filter to Sheet1!A1:F25."
    } satisfies WritebackResult;

    expect(sheetStructurePlan.operation).toBe("insert_rows");
    expect(rangeSortPlan.keys).toHaveLength(1);
    expect(rangeFilterPlan.conditions).toHaveLength(1);
    expect(sheetStructureResult.kind).toBe("sheet_structure_update");
    expect(rangeSortResult.kind).toBe("range_sort");
    expect(rangeFilterResult.kind).toBe("range_filter");
  });

  it("covers the wave 3 plan and result shapes at runtime", () => {
    const conditionalFormatPlan = {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "single_color",
      comparator: "greater_than",
      value: 10,
      style: {
        backgroundColor: "#ffcccc"
      },
      explanation: "Highlight high values.",
      confidence: 0.94,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false
    } satisfies WritePlan;

    const conditionalFormatResult = {
      kind: "conditional_format_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "single_color",
      comparator: "greater_than",
      value: 10,
      style: {
        backgroundColor: "#ffcccc"
      },
      explanation: "Highlight high values.",
      confidence: 0.94,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false,
      summary: "Added conditional formatting to Sheet1!B2:B20."
    } satisfies WritebackResult;

    expect(conditionalFormatPlan.ruleType).toBe("single_color");
    expect(conditionalFormatResult.kind).toBe("conditional_format_update");
  });

  it("covers the wave 4 plan and result shapes at runtime", () => {
    const rangeTransferPlan = {
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Sheet2",
      targetRange: "A1",
      transferOperation: "copy",
      explanation: "Copy the source block into Sheet2.",
      confidence: 0.93,
      requiresConfirmation: true
    } satisfies WritePlan;

    const dataCleanupPlan = {
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      operation: "remove_duplicate_rows",
      explanation: "Remove duplicate rows from the working range.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F100"],
      overwriteRisk: "high",
      confirmationLevel: "destructive",
      keyColumns: ["A", "C"]
    } satisfies WritePlan;

    const rangeTransferResult = {
      kind: "range_transfer_update",
      hostPlatform: "google_sheets",
      operation: "range_transfer_update",
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Sheet2",
      targetRange: "A1",
      transferOperation: "copy",
      summary: "Copied Sheet1!A1:D20 to Sheet2!A1."
    } satisfies WritebackResult;

    const dataCleanupResult = {
      kind: "data_cleanup_update",
      hostPlatform: "excel_windows",
      operation: "remove_duplicate_rows",
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      explanation: "Remove duplicates.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F100"],
      overwriteRisk: "high",
      confirmationLevel: "destructive",
      keyColumns: ["A", "C"],
      summary: "Removed duplicate rows from Sheet1!A2:F100."
    } satisfies WritebackResult;

    expect(rangeTransferPlan.transferOperation).toBe("copy");
    expect(dataCleanupPlan.operation).toBe("remove_duplicate_rows");
    expect(rangeTransferResult.operation).toBe("range_transfer_update");
    expect(dataCleanupResult.operation).toBe("remove_duplicate_rows");
  });

  it("renders a non-lossy validation preview and typed metadata", () => {
    const response = baseResponse({
      type: "data_validation_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        values: ["Open", "Closed"],
        showDropdown: true,
        allowBlank: false,
        invalidDataBehavior: "reject",
        helpText: "Choose a valid status.",
        explanation: "Restrict the status column.",
        confidence: 0.95,
        requiresConfirmation: true,
        replacesExistingValidation: true
      }
    } satisfies ResponseFixture<"data_validation_plan">);

    const preview = buildStructuredPreview(response);
    expect(preview).toEqual({
      kind: "data_validation_plan",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      ruleType: "list",
      values: ["Open", "Closed"],
      showDropdown: true,
      allowBlank: false,
      invalidDataBehavior: "reject",
      helpText: "Choose a valid status.",
      explanation: "Restrict the status column.",
      confidence: 0.95,
      requiresConfirmation: true,
      replacesExistingValidation: true
    });
    expect(getResponseBodyText(response)).toContain("validation plan");
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseMetaLine(response)).toContain("confidence 95%");
    expect(getResponseMetaLine(response)).toContain("confirmation required");
  });

  it("renders a non-lossy named range preview and typed metadata", () => {
    const response = baseResponse({
      type: "named_range_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "retarget",
        name: "InputRange",
        scope: "sheet",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Retarget the named range.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    } satisfies ResponseFixture<"named_range_update">);

    const preview = buildStructuredPreview(response);
    expect(preview).toEqual({
      kind: "named_range_update",
      operation: "retarget",
      name: "InputRange",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      explanation: "Retarget the named range.",
      confidence: 0.9,
      requiresConfirmation: true
    });
    expect(getResponseBodyText(response)).toContain("named range update");
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseMetaLine(response)).toContain("confidence 90%");
    expect(getResponseMetaLine(response)).toContain("confirmation required");
  });

  it("renders a non-lossy range transfer preview", () => {
    const preview = buildRangeTransferPreview({
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Sheet2",
      targetRange: "A1",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the table to the archive sheet.",
      confidence: 0.96,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A1:D20", "Sheet2!A1:D20"],
      overwriteRisk: "none",
      confirmationLevel: "standard"
    } satisfies WritePlan);

    expect(preview).toEqual({
      kind: "range_transfer_plan",
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Sheet2",
      targetRange: "A1",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the table to the archive sheet.",
      confidence: 0.96,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A1:D20", "Sheet2!A1:D20"],
      overwriteRisk: "none",
      confirmationLevel: "standard",
      summary: "Will copy values from Sheet1!A1:D20 to Sheet2!A1.",
      details: [
        "Source sheet: Sheet1.",
        "Source range: A1:D20.",
        "Target sheet: Sheet2.",
        "Target range: A1.",
        "Operation: copy.",
        "Paste mode: values.",
        "Transpose: off.",
        "This will leave the source unchanged.",
        "Affected ranges: Sheet1!A1:D20, Sheet2!A1:D20.",
        "Overwrite risk: none.",
        "Confirmation level: standard."
      ]
    });
  });

  it("renders a non-lossy cleanup preview", () => {
    const preview = buildDataCleanupPreview({
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      operation: "remove_duplicate_rows",
      keyColumns: ["A", "C"],
      explanation: "Remove duplicate rows from the dataset.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F100"],
      overwriteRisk: "medium",
      confirmationLevel: "destructive"
    } satisfies WritePlan);

    expect(preview).toEqual({
      kind: "data_cleanup_plan",
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      operation: "remove_duplicate_rows",
      keyColumns: ["A", "C"],
      explanation: "Remove duplicate rows from the dataset.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F100"],
      overwriteRisk: "medium",
      confirmationLevel: "destructive",
      summary: "Will remove duplicate rows in Sheet1!A2:F100 using key columns A and C.",
      details: [
        "Target sheet: Sheet1.",
        "Target range: A2:F100.",
        "Operation: remove duplicate rows.",
        "Key columns: A and C.",
        "This will remove rows from the target range.",
        "Affected ranges: Sheet1!A2:F100.",
        "Overwrite risk: medium.",
        "Confirmation level: destructive."
      ]
    });
  });

  it("renders a range_transfer_plan response through shared client helpers", () => {
    const response = baseResponse({
      type: "range_transfer_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceSheet: "Sheet1",
        sourceRange: "A1:D20",
        targetSheet: "Sheet2",
        targetRange: "A1",
        operation: "copy",
        pasteMode: "values",
        transpose: false,
        explanation: "Copy the table to the archive sheet.",
        confidence: 0.96,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!A1:D20", "Sheet2!A1:D20"],
        overwriteRisk: "none",
        confirmationLevel: "standard"
      }
    } satisfies ResponseFixture<"range_transfer_plan">);

    expect(getStructuredPreview(response)).toEqual({
      kind: "range_transfer_plan",
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Sheet2",
      targetRange: "A1",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the table to the archive sheet.",
      confidence: 0.96,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A1:D20", "Sheet2!A1:D20"],
      overwriteRisk: "none",
      confirmationLevel: "standard",
      summary: "Will copy values from Sheet1!A1:D20 to Sheet2!A1.",
      details: [
        "Source sheet: Sheet1.",
        "Source range: A1:D20.",
        "Target sheet: Sheet2.",
        "Target range: A1.",
        "Operation: copy.",
        "Paste mode: values.",
        "Transpose: off.",
        "This will leave the source unchanged.",
        "Affected ranges: Sheet1!A1:D20, Sheet2!A1:D20.",
        "Overwrite risk: none.",
        "Confirmation level: standard."
      ]
    });
    expect(getResponseBodyText(response)).toBe(
      "Will copy values from Sheet1!A1:D20 to Sheet2!A1."
    );
    expect(getResponseConfidence(response)).toBe(0.96);
    expect(getRequiresConfirmation(response)).toBe(true);
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseMetaLine(response)).toContain("confidence 96%");
    expect(getResponseMetaLine(response)).toContain("confirmation required");
  });

  it("renders a data_cleanup_plan response through shared client helpers", () => {
    const response = baseResponse({
      type: "data_cleanup_plan",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        targetRange: "A2:F100",
        operation: "remove_duplicate_rows",
        keyColumns: ["A", "C"],
        explanation: "Remove duplicate rows from the dataset.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!A2:F100"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive"
      }
    } satisfies ResponseFixture<"data_cleanup_plan">);

    expect(getStructuredPreview(response)).toEqual({
      kind: "data_cleanup_plan",
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      operation: "remove_duplicate_rows",
      keyColumns: ["A", "C"],
      explanation: "Remove duplicate rows from the dataset.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F100"],
      overwriteRisk: "medium",
      confirmationLevel: "destructive",
      summary: "Will remove duplicate rows in Sheet1!A2:F100 using key columns A and C.",
      details: [
        "Target sheet: Sheet1.",
        "Target range: A2:F100.",
        "Operation: remove duplicate rows.",
        "Key columns: A and C.",
        "This will remove rows from the target range.",
        "Affected ranges: Sheet1!A2:F100.",
        "Overwrite risk: medium.",
        "Confirmation level: destructive."
      ]
    });
    expect(getResponseBodyText(response)).toBe(
      "Will remove duplicate rows in Sheet1!A2:F100 using key columns A and C."
    );
    expect(getResponseConfidence(response)).toBe(0.91);
    expect(getRequiresConfirmation(response)).toBe(true);
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseMetaLine(response)).toContain("confidence 91%");
    expect(getResponseMetaLine(response)).toContain("confirmation required");
  });

  it("renders a structured preview and typed completion line for conditional_format_update", () => {
    const response = baseResponse({
      type: "conditional_format_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "conditional_format_update",
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        summary: "Added conditional formatting to Sheet1!B2:B20."
      }
    } satisfies ResponseFixture<"conditional_format_update">);

    const preview = getStructuredPreview(response);
    if (!preview || preview.kind !== "conditional_format_update") {
      throw new Error("Expected a conditional_format_update preview.");
    }

    expect(preview).toEqual({
      kind: "conditional_format_update",
      operation: "conditional_format_update",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      summary: "Added conditional formatting to Sheet1!B2:B20."
    });
    expect(formatWritebackStatusLine({
      kind: "conditional_format_update",
      hostPlatform: "excel_windows",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "text_contains",
      text: "overdue",
      style: {
        backgroundColor: "#ffcccc"
      },
      explanation: "Highlight overdue items.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false,
      summary: "Added conditional formatting to Sheet1!B2:B20."
    })).toBe("Added conditional formatting to Sheet1!B2:B20.");
    expect(formatWritebackStatusLine({
      kind: "conditional_format_update",
      hostPlatform: "excel_windows",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target",
      ruleType: "text_contains",
      text: "overdue",
      style: {
        backgroundColor: "#ffcccc"
      },
      explanation: "Highlight overdue items.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: true,
      summary: "Replaced conditional formatting on Sheet1!B2:B20."
    })).toBe("Replaced conditional formatting on Sheet1!B2:B20.");
    expect(formatWritebackStatusLine({
      kind: "conditional_format_update",
      hostPlatform: "excel_windows",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "clear_on_target",
      explanation: "Clear existing rules on the target range.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: true,
      summary: "Cleared conditional formatting on Sheet1!B2:B20."
    })).toBe("Cleared conditional formatting on Sheet1!B2:B20.");
  });

  it("formats range write completion lines with row, column, and target formatting", () => {
    expect(formatWritebackStatusLine({
      kind: "range_write",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "B2:D3",
      operation: "replace_range",
      values: [["Open", "North", 1200], ["Closed", "South", 900]],
      explanation: "Write the current status table.",
      confidence: 0.94,
      requiresConfirmation: true,
      shape: {
        rows: 2,
        columns: 3
      },
      writtenRows: 2,
      writtenColumns: 3
    } satisfies WritebackResult)).toBe("Wrote 2 rows x 3 columns to Sheet1!B2:D3.");
  });

  it("formats typed completion lines for range transfer and cleanup updates", () => {
    expect(formatWritebackStatusLine({
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
      summary: "Inserted GOOGLEFINANCE formula into Market Data!B2."
    } satisfies WritebackResult)).toBe("Inserted GOOGLEFINANCE formula into Market Data!B2.");

    expect(formatWritebackStatusLine({
      kind: "range_transfer_update",
      hostPlatform: "google_sheets",
      operation: "range_transfer_update",
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Sheet2",
      targetRange: "A1",
      transferOperation: "copy",
      summary: "Copied Sheet1!A1:D20 to Sheet2!A1."
    } satisfies WritebackResult)).toBe("Copied Sheet1!A1:D20 to Sheet2!A1.");

    expect(formatWritebackStatusLine({
      kind: "data_cleanup_update",
      hostPlatform: "excel_windows",
      operation: "remove_duplicate_rows",
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      explanation: "Remove duplicates.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F100"],
      overwriteRisk: "high",
      confirmationLevel: "destructive",
      keyColumns: ["A", "C"],
      summary: "Removed duplicate rows from Sheet1!A2:F100."
    } satisfies WritebackResult)).toBe("Removed duplicate rows from Sheet1!A2:F100.");
  });

  it("renders a range_transfer_update response through structured preview and body text", () => {
    const response = baseResponse({
      type: "range_transfer_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "range_transfer_update",
        sourceSheet: "Sheet1",
        sourceRange: "A1:D20",
        targetSheet: "Sheet2",
        targetRange: "A1",
        transferOperation: "copy",
        summary: "Copied Sheet1!A1:D20 to Sheet2!A1."
      }
    } satisfies ResponseFixture<"range_transfer_update">);

    expect(getStructuredPreview(response)).toEqual({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Sheet2",
      targetRange: "A1",
      transferOperation: "copy",
      summary: "Copied Sheet1!A1:D20 to Sheet2!A1."
    });
    expect(getResponseBodyText(response)).toBe("Copied Sheet1!A1:D20 to Sheet2!A1.");
  });

  it("renders a data_cleanup_update response through structured preview and body text", () => {
    const response = baseResponse({
      type: "data_cleanup_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "data_cleanup_update",
        targetSheet: "Sheet1",
        targetRange: "A2:F100",
        cleanupOperation: "remove_duplicate_rows",
        summary: "Removed duplicate rows from Sheet1!A2:F100."
      }
    } satisfies ResponseFixture<"data_cleanup_update">);

    expect(getStructuredPreview(response)).toEqual({
      kind: "data_cleanup_update",
      operation: "data_cleanup_update",
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      cleanupOperation: "remove_duplicate_rows",
      summary: "Removed duplicate rows from Sheet1!A2:F100."
    });
    expect(getResponseBodyText(response)).toBe("Removed duplicate rows from Sheet1!A2:F100.");
  });

  it("covers the wave 2 plan and result shapes at runtime", () => {
    const dataValidationPlan = {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      ruleType: "list",
      values: ["Open", "Closed"],
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Restrict the status column.",
      confidence: 0.95,
      requiresConfirmation: true
    } satisfies WritePlan;

    const namedRangeUpdatePlan = {
      operation: "retarget",
      name: "InputRange",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      explanation: "Retarget the named range.",
      confidence: 0.9,
      requiresConfirmation: true
    } satisfies WritePlan;

    const dataValidationResult = {
      kind: "data_validation_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      summary: "Applied validation to Sheet1!B2:B20."
    } satisfies WritebackResult;

    const namedRangeUpdateResult = {
      kind: "named_range_update",
      hostPlatform: "google_sheets",
      operation: "retarget",
      name: "InputRange",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      explanation: "Retarget the named range.",
      confidence: 0.9,
      requiresConfirmation: true,
      summary: "Retargeted InputRange to Sheet1!B2:D20."
    } satisfies WritebackResult;

    expect(dataValidationPlan.ruleType).toBe("list");
    expect(namedRangeUpdatePlan.operation).toBe("retarget");
    expect(dataValidationResult.kind).toBe("data_validation_update");
    expect(namedRangeUpdateResult.kind).toBe("named_range_update");
  });

  it("renders range format updates as confirmable write plans", () => {
    const response = baseResponse({
      type: "range_format_update",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:J10",
        format: {
          backgroundColor: "#fff2cc",
          textColor: "#1f1f1f",
          bold: true,
          horizontalAlignment: "center",
          verticalAlignment: "middle",
          wrapStrategy: "wrap",
          numberFormat: "0.00",
          columnWidth: 96,
          rowHeight: 24
        },
        explanation: "Apply square-table formatting.",
        confidence: 0.9,
        requiresConfirmation: true,
        overwriteRisk: "low"
      }
    } satisfies ResponseFixture<"range_format_update">);

    const preview = getStructuredPreview(response);
    expect(preview).toEqual({
      kind: "range_format_update",
      targetSheet: "Sheet1",
      targetRange: "A1:J10",
      format: {
        backgroundColor: "#fff2cc",
        textColor: "#1f1f1f",
        bold: true,
        horizontalAlignment: "center",
        verticalAlignment: "middle",
        wrapStrategy: "wrap",
        numberFormat: "0.00",
        columnWidth: 96,
        rowHeight: 24
      },
      explanation: "Apply square-table formatting.",
      overwriteRisk: "low"
    });
    expect(isWritePlanResponse(response)).toBe(true);
    expect(getResponseBodyText(response)).toContain("A1:J10");
  });

  it("renders unavailable extraction safely without pretending content was extracted", () => {
    const response = baseResponse({
      type: "error",
      ui: {
        displayMode: "error",
        showTrace: true,
        showWarnings: true,
        showConfidence: false,
        showRequiresConfirmation: false
      },
      warnings: [
        {
          code: "EXTRACTION_UNAVAILABLE",
          message: "Real extraction is unavailable in the current reviewer-safe runtime.",
          severity: "warning"
        }
      ],
      trace: [
        { event: "request_received", timestamp: "2026-04-19T09:00:00.000Z" },
        { event: "attachment_received", timestamp: "2026-04-19T09:00:00.100Z" },
        { event: "failed", timestamp: "2026-04-19T09:00:00.200Z" }
      ],
      data: {
        code: "EXTRACTION_UNAVAILABLE",
        message: "Real image extraction is unavailable in the current reviewer-safe runtime.",
        retryable: false,
        userAction: "Switch to a runtime with real extraction or disable reviewer-safe forced unavailable mode."
      }
    });

    expect(getResponseBodyText(response)).toContain("unavailable");
    expect(getResponseBodyText(response)).toContain("Switch to a runtime with real extraction");
    expect(getStructuredPreview(response)).toBeNull();
    expect(getResponseWarnings(response)[0]?.code).toBe("EXTRACTION_UNAVAILABLE");
  });

  it("renders generic error user actions instead of dropping the recovery guidance", () => {
    const response = baseResponse({
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare a valid spreadsheet response for that request.",
        retryable: true,
        userAction: "Try again with the target sheet or range, or split the request into smaller steps."
      }
    });

    expect(getResponseBodyText(response)).toBe(
      [
        "I couldn't prepare a valid spreadsheet response for that request.",
        "Try again with the target sheet or range, or split the request into smaller steps."
      ].join("\n\n")
    );
  });

  it("keeps demo extraction explicitly labeled through warnings and metadata", () => {
    const response = baseResponse({
      type: "sheet_import_plan",
      environmentLabel: "demo-review",
      warnings: [
        {
          code: "DEMO_OUTPUT",
          message: "This is a demo import preview and not a real extraction of the uploaded image.",
          severity: "warning"
        }
      ],
      data: {
        sourceAttachmentId: "att_101",
        targetSheet: "Demo Import",
        targetRange: "A1:B3",
        headers: ["Column A", "Column B"],
        values: [
          ["Sample 1", "Sample 2"],
          ["Sample 3", "Sample 4"]
        ],
        confidence: 0.2,
        warnings: [
          {
            code: "DEMO_OUTPUT",
            message: "This is a demo import preview and not a real extraction of the uploaded image.",
            severity: "warning"
          }
        ],
        requiresConfirmation: true,
        extractionMode: "demo",
        shape: {
          rows: 3,
          columns: 2
        }
      }
    });

    expect(getResponseMetaLine(response)).toContain("extraction demo");
    expect(getResponseWarnings(response).some((warning) => /demo/i.test(warning.message))).toBe(true);
  });
});

describe("shared client attachment helpers", () => {
  it("filters attachments to the MVP image set", () => {
    const files = filterSupportedImageFiles([
      { name: "table.png", type: "image/png" },
      { name: "table.jpg", type: "image/jpg" },
      { name: "doc.pdf", type: "application/pdf" }
    ]);

    expect(files.map((file) => file.name)).toEqual(["table.png", "table.jpg"]);
  });
});
