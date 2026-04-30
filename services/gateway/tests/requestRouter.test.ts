import express from "express";
import { describe, expect, it, vi } from "vitest";
import { AttachmentStore } from "../src/lib/store.ts";
import { createRequestRouter } from "../src/routes/requests.ts";
import { TraceBus } from "../src/lib/traceBus.ts";

function validRequestBody() {
  return {
    schemaVersion: "1.0.0",
    requestId: "req_route_001",
    source: {
      channel: "google_sheets",
      clientVersion: "0.1.0",
      sessionId: "sess_123"
    },
    host: {
      platform: "google_sheets",
      workbookTitle: "Revenue Demo",
      workbookId: "sheet_route_001",
      activeSheet: "Sheet1",
      selectedRange: "A1:F1"
    },
    userMessage: "Explain the current selection",
    conversation: [
      { role: "user", content: "Explain the current selection" }
    ],
    context: {
      selection: {
        range: "A1:F1",
        headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"],
        values: [["2026-04-01", "Audio", "Cable", "North", 1, 15.5]]
      }
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
  };
}

function createTestRouter(options?: {
  hermesClient?: { processRequest: (...args: any[]) => Promise<void> };
  attachmentStore?: AttachmentStore;
  traceBus?: TraceBus;
}) {
  const traceBus = options?.traceBus ?? new TraceBus();
  const attachmentStore = options?.attachmentStore ?? new AttachmentStore();
  const hermesClient = options?.hermesClient ?? {
    processRequest: vi.fn(async () => undefined)
  };

  return {
    router: createRequestRouter({
      traceBus,
      hermesClient,
      attachmentStore,
      config: {
        port: 8787,
        environmentLabel: "review",
        serviceLabel: "spreadsheet-gateway",
        gatewayPublicBaseUrl: "http://127.0.0.1:8787",
        maxUploadBytes: 8_000_000,
        approvalSecret: "secret",
        hermesAgentBaseUrl: "http://agent.test",
        hermesAgentApiKey: undefined,
        hermesAgentModel: undefined,
        skillRegistryPath: ""
      }
    }),
    traceBus,
    hermesClient,
    attachmentStore
  };
}

async function invokePost(router: express.Router, body: unknown) {
  const layer = router.stack.find((entry) =>
    entry.route?.path === "/" && entry.route.methods.post
  );

  if (!layer) {
    throw new Error("POST / route not found.");
  }

  let statusCode = 200;
  let jsonBody: unknown;

  const req = { body } as express.Request;
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

describe("request router", () => {
  it("accepts a valid Step 2 backend request and preserves requestId and reviewer fields", async () => {
    const hermesClient = {
      processRequest: vi.fn(async () => undefined)
    };
    const { router } = createTestRouter({ hermesClient });

    const response = await invokePost(router, validRequestBody());

    expect(response.statusCode).toBe(202);
    expect((response.body as any).requestId).toBe("req_route_001");
    expect((response.body as any).runId).toMatch(/^run_/);
    expect(hermesClient.processRequest).toHaveBeenCalledTimes(1);
    expect(hermesClient.processRequest).toHaveBeenCalledWith(
      expect.objectContaining({
        request: expect.objectContaining({
          requestId: "req_route_001",
          reviewer: {
            reviewerSafeMode: true,
            forceExtractionMode: "demo"
          }
        })
      })
    );
  });

  it("requires the matching sessionId to poll a session-scoped request run", async () => {
    const { router } = createTestRouter();
    const start = await invokePost(router, validRequestBody());
    const runId = (start.body as any).runId;

    const missingSessionId = await invokeGet(router, runId, {
      requestId: "req_route_001"
    });
    expect(missingSessionId.statusCode).toBe(404);

    const wrongSessionId = await invokeGet(router, runId, {
      requestId: "req_route_001",
      sessionId: "sess_other"
    });
    expect(wrongSessionId.statusCode).toBe(404);

    const matchingSessionId = await invokeGet(router, runId, {
      requestId: "req_route_001",
      sessionId: "sess_123"
    });
    expect(matchingSessionId.statusCode).toBe(200);
    expect((matchingSessionId.body as any).runId).toBe(runId);
  });

  it("uses the injected TraceBus clock for request trace timestamps", async () => {
    let nowMs = Date.UTC(2026, 3, 23, 1, 2, 3);
    const traceBus = new TraceBus({
      now: () => nowMs
    });
    const { router } = createTestRouter({
      traceBus,
      hermesClient: {
        processRequest: vi.fn(async () => undefined)
      }
    });

    const response = await invokePost(router, validRequestBody());
    expect(response.statusCode).toBe(202);

    const events = traceBus.list((response.body as any).runId, 0);
    expect(events).toEqual([
      { event: "request_received", timestamp: "2026-04-23T01:02:03.000Z" },
      {
        event: "spreadsheet_context_received",
        timestamp: "2026-04-23T01:02:03.000Z",
        details: {
          range: "A1:F1",
          sheet: "Sheet1"
        }
      }
    ]);
  });

  it("rejects an invalid backend request envelope", async () => {
    const { router } = createTestRouter();

    const response = await invokePost(router, {
      schemaVersion: "v1",
      requestId: "req_bad_001",
      userMessage: "Explain this"
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).requestId).toBe("req_bad_001");
    expect((response.body as any).error.code).toBe("INVALID_REQUEST");
    expect((response.body as any).error.message).toBe(
      "Hermes couldn't prepare a valid request from the current spreadsheet state."
    );
    expect((response.body as any).error.userAction).toBe(
      "Refresh the sheet state and try again. If it keeps failing, retry with a smaller prompt or reselect the relevant range."
    );
  });

  it("does not reflect oversized request ids from invalid request envelopes", async () => {
    const { router } = createTestRouter();
    const oversizedRequestId = "R".repeat(129);

    const response = await invokePost(router, {
      schemaVersion: "v1",
      requestId: oversizedRequestId,
      userMessage: "Explain this"
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).requestId).toBeUndefined();
    expect(JSON.stringify(response.body)).not.toContain(oversizedRequestId);
    expect((response.body as any).error.code).toBe("INVALID_REQUEST");
  });

  it("rejects non-MVP attachment types in the request envelope", async () => {
    const { router } = createTestRouter();

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        attachments: [
          {
            id: "att_pdf_001",
            type: "image",
            mimeType: "application/pdf",
            source: "upload"
          }
        ]
      }
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).error.code).toBe("INVALID_REQUEST");
  });

  it("accepts legacy requests with null optional fields after normalization", async () => {
    const hermesClient = {
      processRequest: vi.fn(async () => undefined)
    };
    const { router } = createTestRouter({ hermesClient });

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        selection: {
          ...validRequestBody().context.selection,
          headers: null
        },
        activeCell: {
          a1Notation: "F6",
          displayValue: "#N/A",
          value: "#N/A",
          formula: null,
          note: null
        }
      },
      confirmation: {
        state: "none"
      }
    });

    expect(response.statusCode).toBe(202);
    expect(hermesClient.processRequest).toHaveBeenCalledWith(
      expect.objectContaining({
        request: expect.objectContaining({
          source: {
            channel: "google_sheets",
            clientVersion: "0.1.0",
            sessionId: "sess_123"
          },
          host: {
            platform: "google_sheets",
            workbookTitle: "Revenue Demo",
            workbookId: "sheet_route_001",
            activeSheet: "Sheet1",
            selectedRange: "A1:F1"
          },
          context: {
            selection: {
              range: "A1:F1",
              values: [["2026-04-01", "Audio", "Cable", "North", 1, 15.5]]
            },
            activeCell: {
              a1Notation: "F6",
              displayValue: "#N/A",
              value: "#N/A"
            }
          },
          confirmation: {
            state: "none"
          }
        })
      })
    );
  });

  it("uses currentRegion in trace details when the explicit selection is only a single cell", async () => {
    const hermesClient = {
      processRequest: vi.fn(async () => undefined)
    };
    const { router, traceBus } = createTestRouter({ hermesClient });

    const response = await invokePost(router, {
      ...validRequestBody(),
      host: {
        ...validRequestBody().host,
        selectedRange: "J6"
      },
      context: {
        selection: {
          range: "J6",
          values: [[42]]
        },
        currentRegion: {
          range: "A1:F11",
          headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"]
        }
      }
    });

    const runId = (response.body as any).runId;
    const traceEvent = traceBus.list(runId, 0).find((event) => event.event === "spreadsheet_context_received");

    expect(response.statusCode).toBe(202);
    expect(traceEvent).toMatchObject({
      event: "spreadsheet_context_received",
      details: {
        range: "A1:F11",
        sheet: "Sheet1"
      }
    });
  });

  it("still rejects nulls for non-whitelisted fields after normalization", async () => {
    const { router } = createTestRouter();

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        selection: {
          range: "A1:F6",
          headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"],
          values: null
        }
      }
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).error.code).toBe("INVALID_REQUEST");
    expect((response.body as any).error.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          path: "context.selection.values"
        })
      ])
    );
  });

  it("rejects attachment references that are not available in the gateway store", async () => {
    const { router } = createTestRouter();

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        attachments: [
          {
            id: "att_missing_001",
            type: "image",
            mimeType: "image/png",
            source: "upload",
            storageRef: "blob://att_missing_001"
          }
        ]
      }
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).error.code).toBe("ATTACHMENT_UNAVAILABLE");
    expect((response.body as any).error.message).toBe("I can't access that uploaded file anymore.");
    expect((response.body as any).error.userAction).toBe(
      "Reattach the file, then tell me the target sheet or range if you want me to paste or import its contents."
    );
  });

  it("replaces client-supplied attachment identity fields with canonical store metadata", async () => {
    const hermesClient = {
      processRequest: vi.fn(async () => undefined)
    };
    const attachmentStore = new AttachmentStore();
    const storedAttachment = attachmentStore.save({
      buffer: Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      mimeType: "image/png",
      fileName: "real-table.png",
      size: 8,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/att_real/content?uploadToken=upl_real",
      sessionId: "sess_123",
      workbookId: "sheet_route_001"
    });
    const { router } = createTestRouter({ hermesClient, attachmentStore });

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        attachments: [
          {
            id: storedAttachment.id,
            type: "image",
            mimeType: "image/webp",
            fileName: "tampered.webp",
            size: 999,
            source: "drag_drop",
            previewUrl: "https://evil.invalid/preview",
            uploadToken: storedAttachment.uploadToken,
            storageRef: "blob://fake",
            extractedText: "keep auxiliary metadata"
          }
        ]
      }
    });

    expect(response.statusCode).toBe(202);
    expect(hermesClient.processRequest).toHaveBeenCalledWith(
      expect.objectContaining({
        request: expect.objectContaining({
          context: expect.objectContaining({
            attachments: [
              expect.objectContaining({
                id: storedAttachment.id,
                mimeType: "image/png",
                fileName: "real-table.png",
                size: 8,
                source: "upload",
                previewUrl: storedAttachment.previewUrl,
                uploadToken: storedAttachment.uploadToken,
                storageRef: storedAttachment.storageRef
              })
            ]
          })
        })
      })
    );
    const forwardedAttachment = (hermesClient.processRequest.mock.calls[0]?.[0] as any)
      ?.request?.context?.attachments?.[0];
    expect(forwardedAttachment?.extractedText).toBeUndefined();
    expect(forwardedAttachment?.metadata).toBeUndefined();
    expect(forwardedAttachment?.extractedTables).toBeUndefined();
  });

  it("rejects attachment references when the upload token does not match the stored attachment", async () => {
    const attachmentStore = new AttachmentStore();
    const storedAttachment = attachmentStore.save({
      buffer: Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      mimeType: "image/png",
      fileName: "real-table.png",
      size: 8,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/att_real/content?uploadToken=upl_real",
      sessionId: "sess_123",
      workbookId: "sheet_route_001"
    });
    const { router } = createTestRouter({ attachmentStore });

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        attachments: [
          {
            id: storedAttachment.id,
            type: "image",
            mimeType: "image/png",
            fileName: "real-table.png",
            size: 8,
            source: "upload",
            previewUrl: storedAttachment.previewUrl,
            uploadToken: "upl_fake",
            storageRef: storedAttachment.storageRef
          }
        ]
      }
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).error.code).toBe("ATTACHMENT_UNAVAILABLE");
  });

  it("rejects attachment references from a different Hermes session", async () => {
    const attachmentStore = new AttachmentStore();
    const storedAttachment = attachmentStore.save({
      buffer: Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      mimeType: "image/png",
      fileName: "real-table.png",
      size: 8,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/att_real/content?uploadToken=upl_real",
      sessionId: "sess_original",
      workbookId: "sheet_route_001"
    });
    const { router } = createTestRouter({ attachmentStore });

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        attachments: [
          {
            id: storedAttachment.id,
            type: "image",
            mimeType: "image/png",
            fileName: "real-table.png",
            size: 8,
            source: "upload",
            previewUrl: storedAttachment.previewUrl,
            uploadToken: storedAttachment.uploadToken,
            storageRef: storedAttachment.storageRef
          }
        ]
      }
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).error.code).toBe("ATTACHMENT_UNAVAILABLE");
  });

  it("rejects attachment references from a different workbook even inside the same Hermes session", async () => {
    const attachmentStore = new AttachmentStore();
    const storedAttachment = attachmentStore.save({
      buffer: Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      mimeType: "image/png",
      fileName: "real-table.png",
      size: 8,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/att_real/content?uploadToken=upl_real",
      sessionId: "sess_123",
      workbookId: "sheet_other"
    });
    const { router } = createTestRouter({ attachmentStore });

    const response = await invokePost(router, {
      ...validRequestBody(),
      context: {
        attachments: [
          {
            id: storedAttachment.id,
            type: "image",
            mimeType: "image/png",
            fileName: "real-table.png",
            size: 8,
            source: "upload",
            previewUrl: storedAttachment.previewUrl,
            uploadToken: storedAttachment.uploadToken,
            storageRef: storedAttachment.storageRef
          }
        ]
      }
    });

    expect(response.statusCode).toBe(400);
    expect((response.body as any).error.code).toBe("ATTACHMENT_UNAVAILABLE");
  });

  it("requires the matching requestId to read a stored run payload", async () => {
    const { router, traceBus } = createTestRouter();
    traceBus.setResponse("run_status_001", {
      schemaVersion: "1.0.0",
      type: "chat",
      requestId: "req_status_001",
      hermesRunId: "hermes_run_status_001",
      processedBy: "hermes",
      serviceLabel: "spreadsheet-gateway",
      environmentLabel: "review",
      startedAt: "2026-04-22T00:00:00.000Z",
      completedAt: "2026-04-22T00:00:01.000Z",
      durationMs: 1000,
      trace: [],
      ui: {
        displayMode: "chat-first",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        message: "Done"
      }
    });

    const missingRequestId = await invokeGet(router, "run_status_001");
    expect(missingRequestId.statusCode).toBe(404);
    expect(missingRequestId.body).toEqual({
      error: {
        code: "RUN_NOT_FOUND",
        message: "That Hermes request is no longer available.",
        userAction: "Send the request again from the spreadsheet if you need a fresh result."
      }
    });

    const validRequestId = await invokeGet(router, "run_status_001", {
      requestId: "req_status_001"
    });
    expect(validRequestId.statusCode).toBe(200);
    expect((validRequestId.body as any).response?.data?.message).toBe("Done");
  });

  it("can omit response trace from stored run payloads when includeTrace=0 is requested", async () => {
    const { router, traceBus } = createTestRouter();
    traceBus.setResponse("run_status_trimmed_001", {
      schemaVersion: "1.0.0",
      type: "chat",
      requestId: "req_status_trimmed_001",
      hermesRunId: "hermes_run_status_trimmed_001",
      processedBy: "hermes",
      serviceLabel: "spreadsheet-gateway",
      environmentLabel: "review",
      startedAt: "2026-04-22T00:00:00.000Z",
      completedAt: "2026-04-22T00:00:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "response_completed",
          timestamp: "2026-04-22T00:00:01.000Z"
        }
      ],
      ui: {
        displayMode: "chat-first",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        message: "Done"
      }
    });

    const response = await invokeGet(router, "run_status_trimmed_001", {
      requestId: "req_status_trimmed_001",
      includeTrace: "0"
    });

    expect(response.statusCode).toBe(200);
    expect((response.body as any).response?.trace).toEqual([]);
    expect((response.body as any).response?.ui?.showTrace).toBe(false);
    expect((response.body as any).response?.data?.message).toBe("Done");
  });
});
