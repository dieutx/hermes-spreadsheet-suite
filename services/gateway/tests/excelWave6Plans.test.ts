import { afterEach, describe, expect, it, vi } from "vitest";

const TASKPANE_MODULE_URL = new URL(
  "../../../apps/excel-addin/src/taskpane/taskpane.js",
  import.meta.url
).href;
let taskpaneUuidCounter = 0;

function createElementStub() {
  const listeners = new Map<string, Array<(event?: unknown) => void>>();

  return {
    innerHTML: "",
    value: "",
    scrollTop: 0,
    scrollHeight: 0,
    clientHeight: 0,
    children: [],
    addEventListener(type: string, listener: (event?: unknown) => void) {
      listeners.set(type, [...(listeners.get(type) ?? []), listener]);
    },
    dispatch(type: string, event?: unknown) {
      for (const listener of listeners.get(type) ?? []) {
        listener(event);
      }
    },
    querySelectorAll() {
      return [];
    },
    focus() {},
    closest() {
      return null;
    },
    setAttribute() {},
    getAttribute() {
      return null;
    }
  };
}

async function loadTaskpaneModule(
  excelContext: Record<string, unknown>,
  options?: {
    documentUrl?: string;
    documentSettings?: Map<string, unknown>;
    disableDocumentSettings?: boolean;
    disableRandomUUID?: boolean;
    fetchImpl?: ReturnType<typeof vi.fn>;
    locationSearch?: string;
    localStorageSeed?: Record<string, string>;
    throwOnLocalStorageAccess?: boolean;
    sessionStorageSeed?: Record<string, string>;
    addinSetStartupBehavior?: ReturnType<typeof vi.fn>;
  }
) {
  vi.resetModules();

  const elements = new Map([
    ["app", createElementStub()],
    ["messages", createElementStub()],
    ["prompt", createElementStub()],
    ["send-button", createElementStub()],
    ["file-input", createElementStub()],
    ["attachment-strip", createElementStub()]
  ]);

  const storage = new Map<string, string>();
  for (const [key, value] of Object.entries(options?.localStorageSeed ?? {})) {
    storage.set(key, value);
  }
  const sessionStorageBacking = new Map<string, string>();
  for (const [key, value] of Object.entries(options?.sessionStorageSeed ?? {})) {
    sessionStorageBacking.set(key, value);
  }
  const documentSettings = options?.documentSettings ?? new Map<string, unknown>();
  const localStorage = {
    getItem(key: string) {
      if (options?.throwOnLocalStorageAccess) {
        throw new Error("localStorage blocked");
      }
      return storage.get(key) ?? null;
    },
    setItem(key: string, value: string) {
      if (options?.throwOnLocalStorageAccess) {
        throw new Error("localStorage blocked");
      }
      storage.set(key, value);
    },
    removeItem(key: string) {
      if (options?.throwOnLocalStorageAccess) {
        throw new Error("localStorage blocked");
      }
      storage.delete(key);
    }
  };
  const sessionStorage = {
    getItem(key: string) {
      return sessionStorageBacking.get(key) ?? null;
    },
    setItem(key: string, value: string) {
      sessionStorageBacking.set(key, value);
    },
    removeItem(key: string) {
      sessionStorageBacking.delete(key);
    }
  };

  vi.stubGlobal("window", {
    location: { search: options?.locationSearch ?? "" },
    localStorage,
    sessionStorage,
    addEventListener() {},
    setInterval,
    clearInterval,
    setTimeout,
    clearTimeout
  });
  vi.stubGlobal("document", {
    getElementById(id: string) {
      return elements.get(id) ?? createElementStub();
    }
  });
  vi.stubGlobal("fetch", options?.fetchImpl ?? vi.fn());
  vi.stubGlobal("crypto", options?.disableRandomUUID
    ? {}
    : {
        randomUUID() {
          taskpaneUuidCounter += 1;
          return `test-uuid-${taskpaneUuidCounter}`;
        }
      });
  vi.stubGlobal("Office", {
    PlatformType: { Mac: "Mac" },
    AsyncResultStatus: { Succeeded: "succeeded" },
    context: {
      platform: "PC",
      diagnostics: { version: "test-client" },
      document: {
        url: options?.documentUrl ?? "https://example.test/Budget.xlsx",
        settings: options?.disableDocumentSettings
          ? undefined
          : {
              get(key: string) {
                return documentSettings.get(key) ?? null;
              },
              set(key: string, value: unknown) {
                documentSettings.set(key, value);
              },
              remove(key: string) {
                documentSettings.delete(key);
              },
              saveAsync(callback?: (result: { status: string; error?: unknown }) => void) {
                callback?.({ status: "succeeded" });
              }
            }
      },
      displayLanguage: "en-US"
    },
    addin: options?.addinSetStartupBehavior
      ? {
          setStartupBehavior: options.addinSetStartupBehavior
        }
      : undefined,
    StartupBehavior: {
      load: "load"
    },
    onReady() {}
  });
  vi.stubGlobal("Excel", {
    run: async (callback: (context: Record<string, unknown>) => unknown) => callback(excelContext),
    RangeCopyType: {
      formulas: "Formulas",
      formats: "Formats"
    },
    SheetVisibility: {
      visible: "visible",
      hidden: "hidden",
      veryHidden: "veryHidden"
    },
    WorksheetPositionType: {
      beginning: "beginning",
      end: "end",
      before: "before"
    }
  });

  return import(`${TASKPANE_MODULE_URL}?t=${Date.now()}_${Math.random()}`);
}

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("Excel wave 6 composite plans and execution controls", () => {
  it("sanitizes host execution failures into user-facing guidance", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    expect(taskpane.sanitizeHostExecutionError(
      new Error("The approved targetRange does not match the proposed shape.")
    )).toBe(
      "The spreadsheet changed, so the approved destination no longer matches the intended shape.\n\n" +
      "Refresh the spreadsheet state and run the request again."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Excel host requires append targetRange to match the full destination rectangle.")
    )).toBe(
      "The chosen destination range cannot accept this write safely.\n\n" +
      "Choose a clean target range or ask Hermes to write into a blank area."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Target sheet not found: Sales Pivot")
    )).toBe(
      "Sheet \"Sales Pivot\" was not found.\n\nCreate or select that sheet, then retry."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Excel host does not support filter combiners other than and.")
    )).toBe(
      "This spreadsheet app cannot combine those filter conditions in one exact step.\n\n" +
      "Use a single filter rule per column, or split the filter into smaller steps."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Cannot hide the only visible worksheet.")
    )).toBe(
      "At least one worksheet must stay visible.\n\n" +
      "Keep another sheet visible or unhide one first, then retry."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Excel host does not support named ranges on this scope.")
    )).toBe(
      "This named range action is not supported in this spreadsheet app.\n\n" +
      "Use a workbook-level named range or ask Hermes for a simpler named range update."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Hermes Agent returned a response body that does not match the structured gateway contract.")
    )).toBe(
      "The Hermes service returned a response the add-in could not use.\n\n" +
      "Retry the request. If it keeps happening, reload the add-in or check the Hermes gateway."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Excel host does not support exact conditional-format mapping for ruleType color_scale.")
    )).toBe(
      "This conditional formatting step is not supported here.\n\n" +
      "Try a simpler highlight rule, or ask Hermes for a preview-only result first."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("The requested resource doesn't exist.")
    )).toBe(
      "Hermes could not read the current workbook or selection.\n\n" +
      "Select a normal worksheet cell, reload the add-in, and retry. If it keeps happening, reopen the workbook in Excel and try again."
    );

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Unhandled failure at /srv/hermes/services/gateway/src/routes/writeback.ts:42 client_secret=secret_123")
    )).toBe("Write-back failed.");

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Request failed at https://internal.example/api with HERMES_API_SERVER_KEY=secret_123"),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(taskpane.sanitizeHostExecutionError(
      new Error("Writeback failed for qa_HERMES_API_SERVER_KEY")
    )).toBe("Write-back failed.");
  });

  it("formats gateway request issue paths into a visible request-details summary", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    expect(taskpane.appendGatewayIssueSummary(
      "Hermes couldn't prepare a valid request from the current spreadsheet state.\n\nRefresh the sheet state and try again.",
      [
        {
          path: "context.selection.headers",
          message: "headers.length must match selection.range width."
        }
      ]
    )).toBe(
      "Hermes couldn't prepare a valid request from the current spreadsheet state.\n\n" +
      "Refresh the sheet state and try again.\n\n" +
      "Request details:\n" +
      "context.selection.headers: headers.length must match selection.range width."
    );
  });

  it("scrolls the Excel message list to the latest message after render work completes", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });
    const messagesElement = document.getElementById("messages") as {
      scrollTop: number;
      scrollHeight: number;
      lastElementChild?: { scrollIntoView?: ReturnType<typeof vi.fn> };
    };
    const scrollIntoView = vi.fn();
    const raf = vi.fn((callback: FrameRequestCallback) => {
      callback(0);
      return 1;
    });

    messagesElement.scrollTop = 0;
    messagesElement.scrollHeight = 480;
    messagesElement.lastElementChild = { scrollIntoView };
    (window as typeof window & { requestAnimationFrame?: typeof raf }).requestAnimationFrame = raf;

    taskpane.scrollMessagesToBottom();

    expect(messagesElement.scrollTop).toBe(480);
    expect(raf).toHaveBeenCalled();
    expect(scrollIntoView).toHaveBeenCalled();
  });

  it("caps stored Excel messages and per-message trace history to keep long sessions responsive", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const messages = Array.from({ length: 130 }, (_, index) => ({
      role: index % 2 === 0 ? "user" : "assistant",
      content: `message_${index}`
    }));
    const traces = Array.from({ length: 260 }, (_, index) => ({
      event: "result_generated",
      timestamp: `2026-04-22T00:00:${String(index % 60).padStart(2, "0")}.000Z`,
      label: `trace_${index}`
    }));

    const trimmedMessages = taskpane.pruneStoredMessages(messages);
    const trimmedTrace = taskpane.trimMessageTraceEvents(traces);

    expect(trimmedMessages).toHaveLength(100);
    expect(trimmedMessages[0].content).toBe("message_30");
    expect(trimmedMessages.at(-1)?.content).toBe("message_129");
    expect(trimmedTrace).toHaveLength(200);
    expect(trimmedTrace[0].label).toBe("trace_60");
    expect(trimmedTrace.at(-1)?.label).toBe("trace_259");
  });

  it("does not overlap Excel poll requests when a previous poll is still in flight", async () => {
    vi.useFakeTimers();
    const fetchMock = vi.fn(() => new Promise(() => {}));
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      fetchImpl: fetchMock
    });

    void taskpane.pollRun({
      runId: "run_poll_001",
      requestId: "req_poll_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    });

    await vi.advanceTimersByTimeAsync(900);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(String(fetchMock.mock.calls[0]?.[0] || "")).toContain("sessionId=sess_");

    await vi.advanceTimersByTimeAsync(5000);
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("encodes Excel run identifiers in request and trace polling paths", async () => {
    vi.useFakeTimers();
    const fetchMock = vi.fn(async (url: unknown) => new Response(JSON.stringify(
      String(url).includes("/api/trace/")
        ? {
            runId: "run/../unsafe?x=1",
            requestId: "req_poll_path_001",
            status: "processing",
            nextIndex: 0,
            events: []
          }
        : {
            runId: "run/../unsafe?x=1",
            requestId: "req_poll_path_001",
            status: "processing"
          }
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    }));
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      fetchImpl: fetchMock
    });

    await taskpane.pollRun({
      runId: "run/../unsafe?x=1",
      requestId: "req_poll_path_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    });

    await vi.advanceTimersByTimeAsync(900);

    expect(String(fetchMock.mock.calls[0]?.[0] || "")).toContain(
      "/api/trace/run%2F..%2Funsafe%3Fx%3D1?"
    );
    expect(String(fetchMock.mock.calls[1]?.[0] || "")).toContain(
      "/api/requests/run%2F..%2Funsafe%3Fx%3D1?"
    );
  });

  it("loads the Excel taskpane without crypto.randomUUID and still creates request ids", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      disableRandomUUID: true
    });

    const request = taskpane.buildRequestEnvelope({
      userMessage: "Explain this selection",
      conversation: [{ role: "user", content: "Explain this selection" }],
      snapshot: {
        source: {
          channel: "excel_windows",
          clientVersion: "test-client",
          sessionId: "sess_test"
        },
        host: {
          platform: "excel_windows",
          workbookTitle: "Budget.xlsx",
          activeSheet: "Sheet1"
        },
        context: {}
      },
      attachments: []
    });

    expect(request.requestId).toMatch(/^req_/);
  });

  it("loads the Excel taskpane when localStorage access is blocked", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      throwOnLocalStorageAccess: true
    });

    const request = taskpane.buildRequestEnvelope({
      userMessage: "Summarize this table",
      conversation: [{ role: "user", content: "Summarize this table" }],
      snapshot: {
        source: {
          channel: "excel_windows",
          clientVersion: "test-client",
          sessionId: "sess_test"
        },
        host: {
          platform: "excel_windows",
          workbookTitle: "Budget.xlsx",
          activeSheet: "Sheet1"
        },
        context: {}
      },
      attachments: []
    });

    expect(request.requestId).toMatch(/^req_/);
    expect(request.source.sessionId).toMatch(/^sess_/);
  });

  it("uses the persisted localStorage gateway override when no query-string gateway is present", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { "content-type": "application/json" }
    }));
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      fetchImpl: fetchMock,
      localStorageSeed: {
        hermesGatewayBaseUrl: "http://gateway-from-storage.test"
      }
    });

    await taskpane.listExecutionHistory({
      workbookSessionKey: "excel_windows::workbook-123",
      limit: 5
    });

    expect(String(fetchMock.mock.calls[0][0])).toContain(
      "http://gateway-from-storage.test/api/execution/history?workbookSessionKey="
    );
  });

  it("fails closed before fetch when the Excel gateway override is invalid or non-http", async () => {
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { "content-type": "application/json" }
    }));
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      fetchImpl: fetchMock,
      locationSearch: "?gateway=javascript%3Aalert(1)"
    });

    await expect(taskpane.listExecutionHistory({
      workbookSessionKey: "excel_windows::workbook-123",
      limit: 5
    })).rejects.toThrow("Hermes gateway URL is not configured.");
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("sanitizes raw text gateway failures before display in Excel", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    await expect(taskpane.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return "ReferenceError at /srv/hermes/services/gateway/src/app.ts:99 HERMES_API_SERVER_KEY=secret_123";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(taskpane.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return "Gateway failed for qa_HERMES_API_SERVER_KEY";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");
  });

  it("sanitizes JSON gateway failures before display in Excel", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    await expect(taskpane.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        return {
          error: {
            message: "ReferenceError at /srv/hermes/services/gateway/src/app.ts:99 HERMES_API_SERVER_KEY=secret_123",
            userAction: "Inspect https://internal.example/debug for stack trace details."
          }
        };
      },
      async text() {
        return "";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");
  });

  it("sanitizes malformed successful gateway JSON before display in Excel", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    await expect(taskpane.parseGatewayJsonResponse({
      ok: true,
      status: 200,
      async json() {
        throw new Error("Unexpected token at /srv/hermes/services/gateway/src/app.ts:99");
      }
    })).rejects.toThrow(
      "The Hermes service returned a response Excel could not use.\n\n" +
      "Retry the request, then reload the client if it keeps happening."
    );
  });

  it("routes natural-language undo prompts to execution control instead of sending them through the model", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_001"],
      executions: {
        exec_001: {
          baseExecutionId: "exec_001"
        }
      },
      bases: {
        exec_001: {
          baseExecutionId: "exec_001",
          targetSheet: "Sheet8",
          targetRange: "A1",
          beforeCells: [[{ kind: "value", value: { type: "string", value: "before" } }]],
          afterCells: [[{ kind: "value", value: { type: "string", value: "after" } }]]
        }
      }
    });

    const fetchMock = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        if (url.includes("/api/execution/history")) {
          return {
            entries: [
              {
                executionId: "exec_001",
                requestId: "req_001",
                runId: "run_001",
                planType: "sheet_update",
                planDigest: "digest_001",
                status: "completed",
                timestamp: "2099-01-01T00:00:00.000Z",
                reversible: true,
                undoEligible: true,
                redoEligible: false,
                summary: "Write applied to Sheet8!A1."
              }
            ]
          };
        }

        if (url.endsWith("/api/execution/undo/prepare")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "excel_windows",
            executionId: "exec_undo_preview_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }

        if (url.endsWith("/api/execution/undo")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "excel_windows",
            executionId: "exec_undo_001",
            stepResults: [],
            summary: "Undid Sheet8!A1."
          };
        }

        throw new Error(`Unexpected fetch URL: ${url}`);
      }
    }));

    const appliedCells: Array<{ kind: string; value: unknown }> = [];
    const targetRange = {
      rowCount: 1,
      columnCount: 1,
      load() {},
      getCell() {
        return {
          set formulas(value: unknown) {
            appliedCells.push({ kind: "formula", value });
          },
          set values(value: unknown) {
            appliedCells.push({ kind: "value", value });
          }
        };
      }
    };

    const taskpane = await loadTaskpaneModule(
      {
        sync: vi.fn(async () => {}),
        workbook: {
          worksheets: {
            getItem() {
              return {
                getRange() {
                  return targetRange;
                }
              };
            }
          }
        }
      },
      {
        fetchImpl: fetchMock,
        documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
        localStorageSeed: {
          [snapshotStoreKey]: localSnapshotStore
        }
      }
    );

    document.getElementById("prompt").value = "undo";
    await taskpane.sendPrompt();

    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(String(fetchMock.mock.calls[0][0])).toContain(
      `workbookSessionKey=${encodeURIComponent(workbookSessionKey)}`
    );
    expect(String(fetchMock.mock.calls[0][0])).toContain("sessionId=sess_");
    expect(String(fetchMock.mock.calls[0][0])).toContain("limit=20");
    expect(fetchMock.mock.calls[1][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(JSON.parse(String(fetchMock.mock.calls[1][1]?.body))).toMatchObject({
      executionId: "exec_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey
    });
    expect(fetchMock.mock.calls[2][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(JSON.parse(String(fetchMock.mock.calls[2][1]?.body))).toMatchObject({
      executionId: "exec_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey
    });
    expect(appliedCells).toEqual([
      { kind: "value", value: [["before"]] }
    ]);
    expect(document.getElementById("messages").innerHTML).toContain("Undid Sheet8!A1.");
  });

  it("handles bare affirmations locally instead of sending an under-specified follow-up back through the model", async () => {
    const fetchMock = vi.fn();
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      fetchImpl: fetchMock
    });

    taskpane.appendStoredMessage({
      role: "assistant",
      content: "If you want, ask me to restore a specific range by describing what should be put back in Sheet8!A1:K33."
    });
    document.getElementById("prompt").value = "yep";

    await taskpane.sendPrompt();

    expect(fetchMock).not.toHaveBeenCalled();
    expect(document.getElementById("messages").innerHTML).toContain(
      "I still need the exact range, cell, sheet, or action you want me to apply."
    );
  });

  it("renders message status lines so confirm-path host errors are visible in the sidebar", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    taskpane.appendStoredMessage({
      role: "assistant",
      content: "Prepared a chart preview for Sheet1!A78.",
      statusLine: "Write-back failed."
    });
    taskpane.renderMessages();

    expect(document.getElementById("messages").innerHTML).toContain("Write-back failed.");
  });

  it("redacts unsafe proof metadata in Excel meta lines", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const metaLine = taskpane.getResponseMetaLine({
      type: "chat",
      skillsUsed: [
        "SelectionExplainerSkill",
        "/srv/hermes/private-tool.ts",
        "HERMES_API_SERVER_KEY=secret"
      ],
      downstreamProvider: {
        label: "https://internal.example/provider",
        model: "gpt-5 HERMES_API_SERVER_KEY=secret"
      },
      ui: {
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        message: "Processed remotely.",
        confidence: 0.9
      }
    });

    expect(metaLine).toContain("skills SelectionExplainerSkill");
    expect(metaLine).not.toContain("HERMES_API_SERVER_KEY");
    expect(metaLine).not.toContain("/srv/hermes");
    expect(metaLine).not.toContain("internal.example");
    expect(metaLine).not.toContain("provider https://internal");

    const embeddedMetaLine = taskpane.getResponseMetaLine({
      type: "chat",
      skillsUsed: [
        "SelectionExplainerSkill",
        "qa_HERMES_API_SERVER_KEY"
      ],
      downstreamProvider: {
        label: "Hermes Gateway",
        model: "gpt-5 qa_HERMES_AGENT_BASE_URL"
      },
      ui: {
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        message: "Processed remotely.",
        confidence: 0.9
      }
    });

    expect(embeddedMetaLine).toContain("skills SelectionExplainerSkill");
    expect(embeddedMetaLine).toContain("provider Hermes Gateway");
    expect(embeddedMetaLine).not.toContain("qa_HERMES");
  });

  it("escapes quotes in Excel preview action attributes", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const html = taskpane.renderStructuredPreview({
      type: "workbook_structure_update",
      data: {
        operation: "create_sheet",
        sheetName: "Report",
        position: "end",
        explanation: "Create a report sheet.",
        confidence: 0.9,
        requiresConfirmation: true,
        overwriteRisk: "none"
      }
    }, {
      runId: 'run_001" autofocus onfocus="alert(1)',
      requestId: 'req_001" onclick="alert(2)'
    });

    expect(html).toContain("run_001&quot; autofocus onfocus=&quot;alert(1)");
    expect(html).toContain("req_001&quot; onclick=&quot;alert(2)");
    expect(html).not.toContain('data-confirm-run="run_001" autofocus');
    expect(html).not.toContain('onclick="alert(2)');
  });

  it("retries Excel writeback completion without reapplying the local change", async () => {
    let completeAttempts = 0;
    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return new Response(JSON.stringify({
          approvalToken: "approval-token",
          planDigest: "plan-digest",
          executionId: "exec_retry_001"
        }), {
          status: 200,
          headers: { "content-type": "application/json" }
        });
      }

      if (url === "http://gateway.test/api/writeback/complete") {
        completeAttempts += 1;
        if (completeAttempts === 1) {
          return new Response(JSON.stringify({
            error: {
              message: "Temporary completion failure."
            }
          }), {
            status: 500,
            headers: { "content-type": "application/json" }
          });
        }

        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { "content-type": "application/json" }
        });
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });

    const worksheets = {
      items: [
        { name: "Sheet1", position: 0, visibility: "visible" }
      ],
      load: vi.fn(),
      add: vi.fn((name: string) => ({ name, position: 1, visibility: "visible" }))
    };
    const sync = vi.fn(async () => {});
    const taskpane = await loadTaskpaneModule({
      sync,
      workbook: { worksheets }
    }, {
      fetchImpl: fetchMock,
      locationSearch: "?gateway=http%3A%2F%2Fgateway.test"
    });

    const message = {
      role: "assistant",
      requestId: "req_confirm_retry_001",
      runId: "run_confirm_retry_001",
      response: {
        type: "workbook_structure_update",
        data: {
          operation: "create_sheet",
          sheetName: "Retry Demo",
          explanation: "Create a new sheet named Retry Demo.",
          confidence: 0.99,
          requiresConfirmation: true
        }
      },
      content: "Prepared a workbook update to create sheet Retry Demo.",
      statusLine: "",
      traceIndex: 0,
      trace: []
    };

    await taskpane.executeWritePlanMessage(message);

    expect(worksheets.add).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      "http://gateway.test/api/writeback/approve",
      "http://gateway.test/api/writeback/complete"
    ]);
    expect(message.statusLine).toBe(
      "Applied locally. Retry confirm to finish syncing Hermes history."
    );
    expect(message.pendingCompletion).toMatchObject({
      requestId: "req_confirm_retry_001",
      runId: "run_confirm_retry_001",
      workbookSessionKey: expect.stringMatching(/^excel_windows::/),
      approvalToken: "approval-token",
      planDigest: "plan-digest"
    });

    await taskpane.executeWritePlanMessage(message);

    expect(worksheets.add).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      "http://gateway.test/api/writeback/approve",
      "http://gateway.test/api/writeback/complete",
      "http://gateway.test/api/writeback/complete"
    ]);
    expect(message).toMatchObject({
      content: "Created sheet Retry Demo.",
      response: null,
      statusLine: ""
    });
    expect(message.pendingCompletion).toBeUndefined();
  });

  it("keeps Excel writebacks pending when completion returns malformed success JSON", async () => {
    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return new Response(JSON.stringify({
          approvalToken: "approval-token",
          planDigest: "plan-digest",
          executionId: "exec_malformed_ack_001"
        }), {
          status: 200,
          headers: { "content-type": "application/json" }
        });
      }

      if (url === "http://gateway.test/api/writeback/complete") {
        return new Response(JSON.stringify({ ok: false }), {
          status: 200,
          headers: { "content-type": "application/json" }
        });
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });

    const worksheets = {
      items: [
        { name: "Sheet1", position: 0, visibility: "visible" }
      ],
      load: vi.fn(),
      add: vi.fn((name: string) => ({ name, position: 1, visibility: "visible" }))
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    }, {
      fetchImpl: fetchMock,
      locationSearch: "?gateway=http%3A%2F%2Fgateway.test"
    });

    const message = {
      role: "assistant",
      requestId: "req_malformed_ack_001",
      runId: "run_malformed_ack_001",
      response: {
        type: "workbook_structure_update",
        data: {
          operation: "create_sheet",
          sheetName: "Malformed Ack",
          explanation: "Create a new sheet named Malformed Ack.",
          confidence: 0.99,
          requiresConfirmation: true
        }
      },
      content: "Prepared a workbook update to create sheet Malformed Ack.",
      statusLine: "",
      traceIndex: 0,
      trace: []
    };

    await taskpane.executeWritePlanMessage(message);

    expect(worksheets.add).toHaveBeenCalledTimes(1);
    expect(message.statusLine).toBe(
      "Applied locally. Retry confirm to finish syncing Hermes history."
    );
    expect(message.pendingCompletion).toMatchObject({
      requestId: "req_malformed_ack_001",
      runId: "run_malformed_ack_001",
      approvalToken: "approval-token",
      planDigest: "plan-digest"
    });
    expect(message.response).toBeTruthy();
  });

  it("does not execute Excel writes when approval responses omit required tokens", async () => {
    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return new Response(JSON.stringify({
          planDigest: "plan-digest",
          executionId: "exec_missing_token_001"
        }), {
          status: 200,
          headers: { "content-type": "application/json" }
        });
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });

    const worksheets = {
      items: [
        { name: "Sheet1", position: 0, visibility: "visible" }
      ],
      load: vi.fn(),
      add: vi.fn((name: string) => ({ name, position: 1, visibility: "visible" }))
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    }, {
      fetchImpl: fetchMock,
      locationSearch: "?gateway=http%3A%2F%2Fgateway.test"
    });

    const message = {
      role: "assistant",
      requestId: "req_missing_token_001",
      runId: "run_missing_token_001",
      response: {
        type: "workbook_structure_update",
        data: {
          operation: "create_sheet",
          sheetName: "Missing Token",
          explanation: "Create a new sheet named Missing Token.",
          confidence: 0.99,
          requiresConfirmation: true
        }
      },
      content: "Prepared a workbook update to create sheet Missing Token.",
      statusLine: "",
      traceIndex: 0,
      trace: []
    };

    await taskpane.executeWritePlanMessage(message);

    expect(worksheets.add).not.toHaveBeenCalled();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(message.pendingCompletion).toBeUndefined();
    expect(message.statusLine).toContain("writeback approval response");
  });

  it("keeps polling the run when the trace has already expired", async () => {
    vi.useFakeTimers();
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({
        error: {
          code: "RUN_NOT_FOUND",
          message: "That Hermes trace is no longer available.",
          userAction: "Send the request again from the spreadsheet if you need a fresh trace."
        }
      }), {
        status: 404,
        headers: { "content-type": "application/json" }
      }))
      .mockResolvedValueOnce(new Response(JSON.stringify({
        runId: "run_poll_trace_gone_001",
        requestId: "req_poll_trace_gone_001",
        status: "completed",
        response: {
          schemaVersion: "1.0.0",
          requestId: "req_poll_trace_gone_001",
          hermesRunId: "run_poll_trace_gone_001",
          processedBy: "hermes",
          serviceLabel: "hermes-gateway-local",
          environmentLabel: "local-dev",
          startedAt: "2026-04-22T00:00:00.000Z",
          completedAt: "2026-04-22T00:00:01.000Z",
          durationMs: 1000,
          skillsUsed: [],
          downstreamProvider: null,
          warnings: [],
          trace: [],
          ui: {
            displayMode: "inline",
            showTrace: true,
            showWarnings: true,
            showConfidence: true,
            showRequiresConfirmation: false
          },
          type: "chat",
          data: {
            message: "Done."
          }
        }
      }), {
        status: 200,
        headers: { "content-type": "application/json" }
      }));

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      fetchImpl: fetchMock
    });

    const message = {
      runId: "run_poll_trace_gone_001",
      requestId: "req_poll_trace_gone_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    };

    void taskpane.pollRun(message);

    await vi.advanceTimersByTimeAsync(900);

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(message.content).toBe("Done.");
    expect(message.statusLine).not.toContain("trace is no longer available");
    expect(message.statusLine).not.toContain("Request failed");
    expect(message.response?.type).toBe("chat");
  });

  it("keeps polling the run when Excel trace polling hits bandwidth quota and disables live trace polling", async () => {
    vi.useFakeTimers();
    const fetchMock = vi.fn()
      .mockRejectedValueOnce(
        new Error(
          "Exception: Bandwidth quota exceeded: https://gateway.test/api/trace/run_poll_quota_001?after=0. Try reducing the rate of data transfer."
        )
      )
      .mockResolvedValueOnce(new Response(JSON.stringify({
        runId: "run_poll_quota_001",
        requestId: "req_poll_quota_001",
        status: "completed",
        response: {
          schemaVersion: "1.0.0",
          requestId: "req_poll_quota_001",
          hermesRunId: "run_poll_quota_001",
          processedBy: "hermes",
          serviceLabel: "hermes-gateway-local",
          environmentLabel: "local-dev",
          startedAt: "2026-04-22T00:00:00.000Z",
          completedAt: "2026-04-22T00:00:01.000Z",
          durationMs: 1000,
          skillsUsed: [],
          downstreamProvider: null,
          warnings: [],
          trace: [],
          ui: {
            displayMode: "inline",
            showTrace: true,
            showWarnings: true,
            showConfidence: true,
            showRequiresConfirmation: false
          },
          type: "chat",
          data: {
            message: "Done."
          }
        }
      }), {
        status: 200,
        headers: { "content-type": "application/json" }
      }));

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      fetchImpl: fetchMock
    });

    const message = {
      runId: "run_poll_quota_001",
      requestId: "req_poll_quota_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    };

    void taskpane.pollRun(message);

    await vi.advanceTimersByTimeAsync(900);

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(String(fetchMock.mock.calls[0]?.[0] || "")).toContain("sessionId=sess_");
    expect(String(fetchMock.mock.calls[1]?.[0] || "")).toContain("sessionId=sess_");
    expect(message.content).toBe("Done.");
    expect(message.statusLine).not.toContain("Request failed");
    expect(message.tracePollingDisabled).toBe(true);
    expect(message.response?.type).toBe("chat");
  });

  it("reports Excel-generated duplicate sheet names when no new name is approved", async () => {
    const copiedSheet = {
      name: "Template (2)",
      position: 1,
      visibility: "visible",
      load: vi.fn()
    };
    const sourceSheet = {
      name: "Template",
      position: 0,
      visibility: "visible",
      copy: vi.fn(() => copiedSheet)
    };
    const worksheets = {
      items: [sourceSheet],
      load: vi.fn()
    };
    const sync = vi.fn(async () => {});
    const taskpane = await loadTaskpaneModule({
      sync,
      workbook: { worksheets }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        operation: "duplicate_sheet",
        sheetName: "Template",
        position: "end",
        explanation: "Duplicate the template sheet.",
        confidence: 0.95,
        requiresConfirmation: true,
        overwriteRisk: "low"
      },
      requestId: "req_duplicate_generated_name_excel_001",
      runId: "run_duplicate_generated_name_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "workbook_structure_update",
      operation: "duplicate_sheet",
      sheetName: "Template",
      newSheetName: "Template (2)",
      positionResolved: 1,
      sheetCount: 2,
      summary: "Duplicated sheet Template."
    });

    expect(sourceSheet.copy).toHaveBeenCalledWith("end");
    expect(copiedSheet.load).toHaveBeenCalledWith("name");
  });

  it("truncates oversized prompt and conversation content before building the Excel request envelope", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });
    const oversized = "x".repeat(16_050);

    const request = taskpane.buildRequestEnvelope({
      userMessage: oversized,
      conversation: [
        { role: "assistant", content: oversized },
        { role: "user", content: "short" }
      ],
      snapshot: {
        source: {
          channel: "excel_windows",
          clientVersion: "test-client",
          sessionId: "sess_test"
        },
        host: {
          platform: "excel_windows",
          workbookTitle: "Budget.xlsx",
          activeSheet: "Sheet1"
        },
        context: {}
      },
      attachments: []
    });

    expect(request.userMessage).toHaveLength(16_000);
    expect(request.userMessage.endsWith("...")).toBe(true);
    expect(request.conversation[0].content).toHaveLength(16_000);
    expect(request.conversation[0].content.endsWith("...")).toBe(true);
  });

  it("truncates oversized Excel spreadsheet-context strings before building the host snapshot", async () => {
    const longHeader = "H".repeat(300);
    const longValue = "V".repeat(4500);
    const longFormula = `=${"A".repeat(17000)}`;
    const selectionRange = {
      address: "Sheet1!A1:B2",
      values: [
        [longHeader, "Revenue"],
        [longValue, 123]
      ],
      formulas: [
        ["", ""],
        [longFormula, ""]
      ],
      load: vi.fn()
    };
    const currentRegion = {
      address: "Sheet1!A1:B2",
      values: selectionRange.values,
      formulas: selectionRange.formulas,
      load: vi.fn()
    };
    const activeCell = {
      address: "Sheet1!A2",
      values: [[longValue]],
      formulas: [[longFormula]],
      load: vi.fn(),
      getSurroundingRegion: vi.fn(() => currentRegion)
    };
    const headerRange = {
      values: [[longHeader, "Revenue"]],
      load: vi.fn()
    };
    const sheet = {
      name: "Sheet1",
      load: vi.fn(),
      getRange: vi.fn(() => headerRange)
    };

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        getSelectedRange() {
          return selectionRange;
        },
        getActiveCell() {
          return activeCell;
        },
        worksheets: {
          getActiveWorksheet() {
            return sheet;
          }
        }
      }
    });

    const snapshot = await taskpane.getSpreadsheetSnapshot("Explain the current selection");

    expect(snapshot.context.selection.headers?.[0]).toHaveLength(256);
    expect(snapshot.context.selection.headers?.[0].endsWith("…")).toBe(true);
    expect(String(snapshot.context.selection.values?.[1]?.[0])).toHaveLength(4000);
    expect(String(snapshot.context.selection.values?.[1]?.[0]).endsWith("…")).toBe(true);
    expect(snapshot.context.selection.formulas?.[1]?.[0]).toHaveLength(16000);
    expect(snapshot.context.selection.formulas?.[1]?.[0]?.endsWith("…")).toBe(true);
    expect(String(snapshot.context.activeCell?.displayValue)).toHaveLength(4000);
    expect(String(snapshot.context.activeCell?.displayValue).endsWith("…")).toBe(true);
    expect(snapshot.context.activeCell?.formula).toHaveLength(16000);
    expect(snapshot.context.activeCell?.formula?.endsWith("…")).toBe(true);
  });

  it("includes Excel active and referenced cell notes in workbook context", async () => {
    const selectionRange = {
      address: "Sheet1!A1:B2",
      values: [
        ["Status", "Revenue"],
        ["Open", 123]
      ],
      formulas: [
        ["", ""],
        ["", ""]
      ],
      load: vi.fn()
    };
    const currentRegion = {
      address: "Sheet1!A1:B2",
      values: selectionRange.values,
      formulas: selectionRange.formulas,
      load: vi.fn()
    };
    const activeCell = {
      address: "Sheet1!A2",
      values: [["Open"]],
      formulas: [[""]],
      load: vi.fn(),
      getSurroundingRegion: vi.fn(() => currentRegion)
    };
    const referencedCell = {
      address: "Sheet1!H11",
      values: [["Blocked"]],
      formulas: [[""]],
      load: vi.fn()
    };
    const notes = {
      getItemOrNullObject: vi.fn((address: string) => ({
        isNullObject: false,
        content: address === "A2" ? "Active note" : "Referenced note",
        load: vi.fn()
      }))
    };
    const sheet = {
      name: "Sheet1",
      notes,
      load: vi.fn(),
      getRange: vi.fn((address: string) => {
        if (address === "H11") {
          return referencedCell;
        }
        return selectionRange;
      })
    };

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        getSelectedRange() {
          return selectionRange;
        },
        getActiveCell() {
          return activeCell;
        },
        worksheets: {
          getActiveWorksheet() {
            return sheet;
          }
        }
      }
    });

    const snapshot = await taskpane.getSpreadsheetSnapshot("Explain H11 relative to the active cell");

    expect(snapshot.context.activeCell?.note).toBe("Active note");
    expect(snapshot.context.referencedCells?.[0]).toMatchObject({
      a1Notation: "H11",
      displayValue: "Blocked",
      note: "Referenced note"
    });
    expect(notes.getItemOrNullObject).toHaveBeenCalledWith("A2");
    expect(notes.getItemOrNullObject).toHaveBeenCalledWith("H11");
  });

  it("keeps following the bottom after delayed layout growth when pinned", async () => {
    vi.useFakeTimers();
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });
    const messagesElement = document.getElementById("messages") as {
      scrollTop: number;
      scrollHeight: number;
      clientHeight: number;
      lastElementChild?: { scrollIntoView?: ReturnType<typeof vi.fn> };
    };
    const scrollIntoView = vi.fn();

    messagesElement.scrollTop = 120;
    messagesElement.scrollHeight = 480;
    messagesElement.clientHeight = 320;
    messagesElement.lastElementChild = { scrollIntoView };

    taskpane.scheduleMessagesAutoScroll(true);
    expect(messagesElement.scrollTop).toBe(480);

    messagesElement.scrollHeight = 960;
    vi.runAllTimers();

    expect(messagesElement.scrollTop).toBe(960);
    expect(scrollIntoView).toHaveBeenCalled();
    vi.useRealTimers();
  });

  it("keeps following the bottom when Excel layout settles after the earlier follow-up window", async () => {
    vi.useFakeTimers();
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });
    const messagesElement = document.getElementById("messages") as {
      scrollTop: number;
      scrollHeight: number;
      clientHeight: number;
      lastElementChild?: { scrollIntoView?: ReturnType<typeof vi.fn> };
    };
    const scrollIntoView = vi.fn();

    messagesElement.scrollTop = 120;
    messagesElement.scrollHeight = 480;
    messagesElement.clientHeight = 320;
    messagesElement.lastElementChild = { scrollIntoView };

    taskpane.scheduleMessagesAutoScroll(true);
    expect(messagesElement.scrollTop).toBe(480);

    vi.advanceTimersByTime(500);
    messagesElement.scrollHeight = 1040;
    vi.advanceTimersByTime(200);

    expect(messagesElement.scrollTop).toBe(1040);
    expect(scrollIntoView).toHaveBeenCalled();
    vi.useRealTimers();
  });

  it("does not yank the viewport back down after the user scrolls away", async () => {
    vi.useFakeTimers();
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });
    const messagesElement = document.getElementById("messages") as {
      scrollTop: number;
      scrollHeight: number;
      clientHeight: number;
      lastElementChild?: { scrollIntoView?: ReturnType<typeof vi.fn> };
      dispatch: (type: string, event?: unknown) => void;
    };
    const scrollIntoView = vi.fn();

    messagesElement.scrollHeight = 900;
    messagesElement.clientHeight = 300;
    messagesElement.lastElementChild = { scrollIntoView };
    taskpane.bindMessageAutoScrollObservers();

    messagesElement.scrollTop = 80;
    messagesElement.dispatch("scroll");

    taskpane.scheduleMessagesAutoScroll();
    expect(messagesElement.scrollTop).toBe(80);

    vi.runAllTimers();
    expect(messagesElement.scrollTop).toBe(80);
    expect(scrollIntoView).not.toHaveBeenCalled();
    vi.useRealTimers();
  });

  it("does not auto-enable document auto-open by default for ordinary workbooks", async () => {
    const documentSettings = new Map<string, unknown>();
    const setStartupBehavior = vi.fn(async () => undefined);
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      documentSettings,
      addinSetStartupBehavior: setStartupBehavior
    });

    await taskpane.ensureDemoStartupDefaults();

    expect(setStartupBehavior).not.toHaveBeenCalled();
    expect(documentSettings.get("Office.AutoShowTaskpaneWithDocument")).toBeUndefined();
    expect(documentSettings.get("Hermes.EnableAutoOpen")).toBeUndefined();
  });

  it("persists document auto-open only when explicitly opted in for demo usage", async () => {
    const documentSettings = new Map<string, unknown>();
    const setStartupBehavior = vi.fn(async () => undefined);
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, {
      documentSettings,
      locationSearch: "?enableDocumentAutoOpen=1",
      addinSetStartupBehavior: setStartupBehavior
    });

    await taskpane.ensureDemoStartupDefaults();

    expect(setStartupBehavior).toHaveBeenCalledWith("load");
    expect(documentSettings.get("Office.AutoShowTaskpaneWithDocument")).toBe(true);
    expect(documentSettings.get("Hermes.EnableAutoOpen")).toBe(true);
  });

  it("does not load the full selected range matrix when the selected range exceeds the inline threshold", async () => {
    const selectionHeaderValues = [[
      "Date", "Category", "Product", "Region", "Units",
      "Revenue", "Rep", "Channel", "Segment", "Discount",
      "COGS", "Margin", "City", "State", "Country",
      "Quarter", "Month", "Year", "Customer", "Order ID"
    ]];
    const selectionRange = {
      address: "Sheet1!A1:T500",
      values: [[123]],
      formulas: [[""]],
      load: vi.fn()
    };
    const selectionHeaderRange = {
      values: selectionHeaderValues,
      load: vi.fn()
    };
    const currentRegion = {
      address: "Sheet1!A1:F10",
      values: [[
        "Date", "Category", "Product", "Region", "Units", "Revenue"
      ]],
      formulas: [["", "", "", "", "", ""]],
      load: vi.fn()
    };
    const activeCell = {
      address: "Sheet1!J6",
      values: [[123]],
      formulas: [[""]],
      load: vi.fn(),
      getSurroundingRegion: vi.fn(() => currentRegion)
    };
    const sheet = {
      name: "Sheet1",
      load: vi.fn(),
      getRange: vi.fn((a1: string) => a1 === "A1:T1" ? selectionHeaderRange : selectionHeaderRange)
    };

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        getSelectedRange() {
          return selectionRange;
        },
        getActiveCell() {
          return activeCell;
        },
        worksheets: {
          getActiveWorksheet() {
            return sheet;
          }
        }
      }
    });

    const snapshot = await taskpane.getSpreadsheetSnapshot("Explain the current selection");

    expect(selectionRange.load).toHaveBeenCalledWith(["address"]);
    expect(selectionRange.load).not.toHaveBeenCalledWith(["values", "formulas"]);
    expect(sheet.getRange).toHaveBeenCalledWith("A1:T1");
    expect(selectionHeaderRange.load).toHaveBeenCalledWith(["values"]);
    expect(snapshot.context.selection).toMatchObject({
      range: "A1:T500",
      headers: selectionHeaderValues[0]
    });
    expect(snapshot.context.selection.values).toBeUndefined();
    expect(snapshot.context.selection.formulas).toBeUndefined();
  });

  it("does not load the full currentRegion matrix when the current table exceeds the inline threshold", async () => {
    const headerValues = [[
      "Date", "Category", "Product", "Region", "Units",
      "Revenue", "Rep", "Channel", "Segment", "Discount",
      "COGS", "Margin", "City", "State", "Country",
      "Quarter", "Month", "Year", "Customer", "Order ID"
    ]];
    const selectionRange = {
      address: "Sheet1!J6",
      values: [[123]],
      formulas: [[""]],
      load: vi.fn()
    };
    const currentRegion = {
      address: "Sheet1!A1:T500",
      values: [[123]],
      formulas: [[""]],
      load: vi.fn()
    };
    const headerRange = {
      values: headerValues,
      load: vi.fn()
    };
    const activeCell = {
      address: "Sheet1!J6",
      values: [[123]],
      formulas: [[""]],
      load: vi.fn(),
      getSurroundingRegion: vi.fn(() => currentRegion)
    };
    const sheet = {
      name: "Sheet1",
      load: vi.fn(),
      getRange: vi.fn(() => headerRange)
    };

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        getSelectedRange() {
          return selectionRange;
        },
        getActiveCell() {
          return activeCell;
        },
        worksheets: {
          getActiveWorksheet() {
            return sheet;
          }
        }
      }
    });

    const snapshot = await taskpane.getSpreadsheetSnapshot("Explain the current selection");

    expect(currentRegion.load).toHaveBeenCalledWith(["address"]);
    expect(currentRegion.load).not.toHaveBeenCalledWith(["values", "formulas"]);
    expect(sheet.getRange).toHaveBeenCalledWith("A1:T1");
    expect(headerRange.load).toHaveBeenCalledWith(["values"]);
    expect(snapshot.context.currentRegion).toMatchObject({
      range: "A1:T500",
      headers: headerValues[0]
    });
    expect(snapshot.context.currentRegion.values).toBeUndefined();
    expect(snapshot.context.currentRegion.formulas).toBeUndefined();
    expect(snapshot.context.currentRegionArtifactTarget).toBe("A502");
    expect(snapshot.context.currentRegionAppendTarget).toBe("A501:T501");
  });

  it("falls back to worksheet A1 when Excel cannot resolve the current selection resource", async () => {
    let syncCount = 0;
    const fallbackRange = {
      address: "Sheet1!A1",
      values: [["EEID"]],
      formulas: [[""]],
      load: vi.fn()
    };
    const selectionRange = {
      address: "Sheet1!C7",
      values: [["stale"]],
      formulas: [[""]],
      load: vi.fn()
    };
    const activeCell = {
      address: "Sheet1!C7",
      values: [["stale"]],
      formulas: [[""]],
      load: vi.fn(),
      getSurroundingRegion: vi.fn(() => selectionRange)
    };
    const sheet = {
      name: "Sheet1",
      load: vi.fn(),
      getRange: vi.fn((a1: string) => {
        if (a1 === "A1") {
          return fallbackRange;
        }
        return fallbackRange;
      })
    };

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {
        syncCount += 1;
        if (syncCount === 2) {
          throw new Error("The requested resource doesn't exist.");
        }
      }),
      workbook: {
        getSelectedRange() {
          return selectionRange;
        },
        getActiveCell() {
          return activeCell;
        },
        worksheets: {
          getActiveWorksheet() {
            return sheet;
          }
        }
      }
    });

    const snapshot = await taskpane.getSpreadsheetSnapshot("Explain the current selection");

    expect(sheet.getRange).toHaveBeenCalledWith("A1");
    expect(snapshot.host.activeSheet).toBe("Sheet1");
    expect(snapshot.context.selection.range).toBe("A1");
    expect(snapshot.context.activeCell?.a1Notation).toBe("A1");
  });

  it("preserves the selected range when only the active-cell resource is unavailable", async () => {
    let syncCount = 0;
    const selectionRange = {
      address: "Sheet1!C7:D8",
      values: [["foo", "bar"], ["baz", "qux"]],
      formulas: [["", ""], ["", ""]],
      load: vi.fn()
    };
    const fallbackActiveCell = {
      address: "Sheet1!C7",
      values: [["foo"]],
      formulas: [[""]],
      load: vi.fn(),
      getSurroundingRegion: vi.fn(() => selectionRange)
    };
    const activeCell = {
      address: "Sheet1!C8",
      values: [["qux"]],
      formulas: [[""]],
      load: vi.fn(),
      getSurroundingRegion: vi.fn(() => selectionRange)
    };
    const headerRange = {
      values: [["foo", "bar"]],
      load: vi.fn()
    };
    const sheet = {
      name: "Sheet1",
      load: vi.fn(),
      getRange: vi.fn((a1: string) => {
        if (a1 === "C7") {
          return fallbackActiveCell;
        }
        return headerRange;
      })
    };

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {
        syncCount += 1;
        if (syncCount === 3) {
          throw new Error("The requested resource doesn't exist.");
        }
      }),
      workbook: {
        getSelectedRange() {
          return selectionRange;
        },
        getActiveCell() {
          return activeCell;
        },
        worksheets: {
          getActiveWorksheet() {
            return sheet;
          }
        }
      }
    });

    const snapshot = await taskpane.getSpreadsheetSnapshot("Explain the current selection");

    expect(sheet.getRange).toHaveBeenCalledWith("C7");
    expect(snapshot.context.selection.range).toBe("C7:D8");
    expect(snapshot.context.activeCell?.a1Notation).toBe("C7");
  });

  it("treats unsupported Excel preview plans as non-confirmable write plans", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const unsupportedValidation = {
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "time",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Validate times only.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    };
    const unsupportedConditional = {
      type: "conditional_format_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          numberFormat: "0.00"
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    };
    const unsupportedCleanup = {
      type: "data_cleanup_plan",
      data: {
        targetSheet: "Contacts",
        targetRange: "A2:A20",
        operation: "standardize_format",
        formatType: "date_text",
        formatPattern: "locale-sensitive-fuzzy",
        explanation: "Normalize date strings with a fuzzy locale format.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    };
    const unsupportedRangeFormat = {
      type: "range_format_update",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:C10",
        format: {
          wrapStrategy: "clip"
        },
        explanation: "Clip overflowing text in the sales table.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    };

    expect(taskpane.isWritePlanResponse(unsupportedValidation)).toBe(false);
    expect(taskpane.renderStructuredPreview(unsupportedValidation, {
      runId: "run_validation_unsupported_excel_001",
      requestId: "req_validation_unsupported_excel_001"
    })).toContain("This Excel runtime can't apply that validation rule.");

    expect(taskpane.isWritePlanResponse(unsupportedConditional)).toBe(false);
    expect(taskpane.renderStructuredPreview(unsupportedConditional, {
      runId: "run_conditional_unsupported_excel_001",
      requestId: "req_conditional_unsupported_excel_001"
    })).toContain("This Excel runtime can't apply that conditional formatting style exactly.");

    expect(taskpane.isWritePlanResponse(unsupportedCleanup)).toBe(false);
    const cleanupHtml = taskpane.renderStructuredPreview(unsupportedCleanup, {
      runId: "run_cleanup_unsupported_excel_001",
      requestId: "req_cleanup_unsupported_excel_001"
    });
    expect(cleanupHtml).toContain("This Excel runtime only supports exact year-first date text patterns");
    expect(cleanupHtml).not.toContain("Confirm Cleanup");

    expect(taskpane.isWritePlanResponse(unsupportedRangeFormat)).toBe(false);
    const rangeFormatHtml = taskpane.renderStructuredPreview(unsupportedRangeFormat, {
      runId: "run_range_format_clip_excel_001",
      requestId: "req_range_format_clip_excel_001"
    });
    expect(rangeFormatHtml).toContain("This Excel runtime can't clip overflowing text exactly");
    expect(rangeFormatHtml).not.toContain("Confirm Format Update");
  });

  it("fails closed for Excel range format clip wrapping", async () => {
    const wrapTextSet = vi.fn();
    const targetFormat = {};
    Object.defineProperty(targetFormat, "wrapText", {
      configurable: true,
      set(value) {
        wrapTextSet(value);
      }
    });
    const targetRange = {
      rowCount: 10,
      columnCount: 3,
      values: [["Region", "Revenue", "Notes"]],
      formulas: [["", "", ""]],
      load: vi.fn(),
      format: targetFormat
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange)
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sales",
        targetRange: "A1:C10",
        format: {
          wrapStrategy: "clip"
        },
        explanation: "Clip overflowing text in the sales table.",
        confidence: 0.91,
        requiresConfirmation: true
      },
      requestId: "req_range_format_clip_apply_excel_001",
      runId: "run_range_format_clip_apply_excel_001",
      approvalToken: "token"
    })).rejects.toThrow("Excel host cannot clip overflowing text exactly.");

    expect(wrapTextSet).not.toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Excel range format writes", async () => {
    let targetRange: any;
    targetRange = {
      rowCount: 1,
      columnCount: 1,
      values: [["Revenue"]],
      formulas: [[""]],
      numberFormat: [["General"]],
      load: vi.fn(),
      getCell: vi.fn(() => targetRange),
      format: {
        fill: { color: "#FFFFFF" },
        font: {
          color: "#000000",
          name: "Calibri",
          size: 11,
          bold: false,
          italic: false,
          underline: "None",
          strikethrough: false
        },
        horizontalAlignment: "Left",
        verticalAlignment: "Bottom",
        wrapText: false,
        columnWidth: 64,
        rowHeight: 18
      }
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange)
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    });

    const result = await taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sales",
        targetRange: "B2",
        format: {
          backgroundColor: "#FFF2CC",
          bold: true,
          numberFormat: "$#,##0"
        },
        explanation: "Format revenue.",
        confidence: 0.91,
        requiresConfirmation: true
      },
      requestId: "req_range_format_snapshot_excel_001",
      runId: "run_range_format_snapshot_excel_001",
      approvalToken: "token",
      executionId: "exec_range_format_snapshot_excel_001"
    });

    expect(result).toMatchObject({
      kind: "range_format_update",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_range_format_snapshot_excel_001",
        kind: "range_format",
        targetSheet: "Sales",
        targetRange: "B2",
        shape: {
          rows: 1,
          columns: 1
        },
        beforeFormatCells: [[{
          backgroundColor: "#FFFFFF",
          bold: false,
          numberFormat: [["General"]]
        }]],
        afterFormatCells: [[{
          backgroundColor: "#FFF2CC",
          bold: true,
          numberFormat: [["$#,##0"]]
        }]]
      }
    });
  });

  it("restores Excel range format snapshots before committing undo", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_range_format_001"],
      executions: {
        exec_range_format_001: {
          baseExecutionId: "exec_range_format_001"
        }
      },
      bases: {
        exec_range_format_001: {
          baseExecutionId: "exec_range_format_001",
          kind: "range_format",
          targetSheet: "Sales",
          targetRange: "B2",
          shape: {
            rows: 1,
            columns: 1
          },
          beforeFormatCells: [[{
            backgroundColor: "#FFFFFF",
            bold: false,
            numberFormat: [["General"]]
          }]],
          afterFormatCells: [[{
            backgroundColor: "#FFF2CC",
            bold: true,
            numberFormat: [["$#,##0"]]
          }]]
        }
      }
    });
    const fetchMock = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "range_format_update",
          operation: "range_format_update",
          hostPlatform: "excel_windows",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));
    let targetRange: any;
    targetRange = {
      rowCount: 1,
      columnCount: 1,
      numberFormat: [["$#,##0"]],
      load: vi.fn(),
      getCell: vi.fn(() => targetRange),
      format: {
        fill: { color: "#FFF2CC" },
        font: {
          bold: true
        }
      }
    };
    const taskpane = await loadTaskpaneModule(
      {
        sync: vi.fn(async () => {}),
        workbook: {
          worksheets: {
            getItem() {
              return {
                getRange() {
                  return targetRange;
                }
              };
            }
          }
        }
      },
      {
        fetchImpl: fetchMock,
        documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
        localStorageSeed: {
          [snapshotStoreKey]: localSnapshotStore
        }
      }
    );

    await taskpane.undoExecution("exec_range_format_001");

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(fetchMock.mock.calls[1][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(targetRange.format.fill.color).toBe("#FFFFFF");
    expect(targetRange.format.font.bold).toBe(false);
    expect(targetRange.numberFormat).toEqual([["General"]]);
  });

  it("restores Excel data validation snapshots before committing undo", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const beforeValidation = {
      rule: {
        decimal: {
          operator: "GreaterThanOrEqualTo",
          formula1: 0
        }
      },
      ignoreBlanks: true,
      prompt: {
        title: "Old validation",
        message: "Old prompt"
      },
      errorAlert: {
        title: "Old error",
        message: "Old error message",
        style: "warning",
        showAlert: true
      }
    };
    const afterValidation = {
      rule: {
        wholeNumber: {
          operator: "Between",
          formula1: 1,
          formula2: 10
        }
      },
      ignoreBlanks: false,
      prompt: {
        title: "Entry guidance",
        message: "Enter a whole number from 1 to 10."
      },
      errorAlert: {
        title: "Invalid entry",
        message: "Only whole numbers from 1 to 10 are allowed.",
        style: "stop",
        showAlert: true
      }
    };
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_validation_001"],
      executions: {
        exec_validation_001: {
          baseExecutionId: "exec_validation_001"
        }
      },
      bases: {
        exec_validation_001: {
          baseExecutionId: "exec_validation_001",
          kind: "data_validation",
          targetSheet: "Sales",
          targetRange: "B2:B20",
          shape: {
            rows: 19,
            columns: 1
          },
          beforeValidation,
          afterValidation
        }
      }
    });
    const fetchMock = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "data_validation_update",
          operation: "data_validation_update",
          hostPlatform: "excel_windows",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));
    let currentRule = afterValidation.rule;
    let currentIgnoreBlanks = afterValidation.ignoreBlanks;
    let currentPrompt = afterValidation.prompt;
    let currentErrorAlert = afterValidation.errorAlert;
    const dataValidation = {};
    Object.defineProperty(dataValidation, "rule", {
      configurable: true,
      get() {
        return currentRule;
      },
      set(rule) {
        currentRule = rule;
      }
    });
    Object.defineProperty(dataValidation, "ignoreBlanks", {
      configurable: true,
      get() {
        return currentIgnoreBlanks;
      },
      set(value) {
        currentIgnoreBlanks = value;
      }
    });
    Object.defineProperty(dataValidation, "prompt", {
      configurable: true,
      get() {
        return currentPrompt;
      },
      set(value) {
        currentPrompt = value;
      }
    });
    Object.defineProperty(dataValidation, "errorAlert", {
      configurable: true,
      get() {
        return currentErrorAlert;
      },
      set(value) {
        currentErrorAlert = value;
      }
    });
    const targetRange = {
      rowCount: 19,
      columnCount: 1,
      load: vi.fn(),
      dataValidation
    };
    const taskpane = await loadTaskpaneModule(
      {
        sync: vi.fn(async () => {}),
        workbook: {
          worksheets: {
            getItem() {
              return {
                getRange() {
                  return targetRange;
                }
              };
            }
          }
        }
      },
      {
        fetchImpl: fetchMock,
        documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
        localStorageSeed: {
          [snapshotStoreKey]: localSnapshotStore
        }
      }
    );

    await taskpane.undoExecution("exec_validation_001");

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(fetchMock.mock.calls[1][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(currentRule).toEqual(beforeValidation.rule);
    expect(currentIgnoreBlanks).toBe(true);
    expect(currentPrompt).toEqual(beforeValidation.prompt);
    expect(currentErrorAlert).toEqual(beforeValidation.errorAlert);
  });

  it("restores Excel named range snapshots before committing undo", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_named_range_001"],
      executions: {
        exec_named_range_001: {
          baseExecutionId: "exec_named_range_001"
        }
      },
      bases: {
        exec_named_range_001: {
          baseExecutionId: "exec_named_range_001",
          kind: "named_range",
          scope: "workbook",
          before: {
            exists: true,
            name: "OldRange",
            reference: "Sheet1!A1:A10"
          },
          after: {
            exists: true,
            name: "NewRange",
            reference: "Sheet1!A1:A10"
          }
        }
      }
    });
    const fetchMock = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "named_range_update",
          operation: "rename",
          hostPlatform: "excel_windows",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));
    const namedRange = {
      name: "NewRange",
      reference: "Sheet1!A1:A10",
      load: vi.fn(),
      delete: vi.fn()
    };
    const workbookNames = {
      getItem: vi.fn((name: string) => {
        expect(name).toBe("NewRange");
        return namedRange;
      }),
      add: vi.fn()
    };
    const taskpane = await loadTaskpaneModule(
      {
        sync: vi.fn(async () => {}),
        workbook: {
          names: workbookNames,
          worksheets: {
            getItem: vi.fn()
          }
        }
      },
      {
        fetchImpl: fetchMock,
        documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
        localStorageSeed: {
          [snapshotStoreKey]: localSnapshotStore
        }
      }
    );

    await taskpane.undoExecution("exec_named_range_001");

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(fetchMock.mock.calls[1][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(namedRange.name).toBe("OldRange");
    expect(namedRange.reference).toBe("Sheet1!A1:A10");
    expect(workbookNames.add).not.toHaveBeenCalled();
    expect(namedRange.delete).not.toHaveBeenCalled();
  });

  it("renders and applies native Excel table plans", async () => {
    const table = {
      name: "",
      style: "",
      showBandedRows: false,
      showBandedColumns: false,
      showFilterButton: false,
      showTotals: false
    };
    const targetRange = {
      rowCount: 50,
      columnCount: 6,
      load: vi.fn()
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange),
      tables: {
        add: vi.fn(() => table)
      }
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    });
    const response = {
      type: "table_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        styleName: "TableStyleMedium2",
        showBandedRows: true,
        showBandedColumns: false,
        showFilterButton: true,
        showTotalsRow: false,
        explanation: "Convert the sales range into a native table.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(true);
    expect(taskpane.getStructuredPreview(response)).toMatchObject({
      kind: "table_plan",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      name: "SalesTable"
    });
    expect(taskpane.renderStructuredPreview(response, {
      runId: "run_table_preview_excel_001",
      requestId: "req_table_preview_excel_001"
    })).toContain("Confirm Table");

    await expect(taskpane.applyWritePlan({
      plan: response.data,
      requestId: "req_table_apply_excel_001",
      runId: "run_table_apply_excel_001",
      approvalToken: "token",
      executionId: "exec_table_apply_excel_001"
    })).resolves.toMatchObject({
      kind: "table_update",
      operation: "table_update",
      hostPlatform: "excel_windows",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      name: "SalesTable",
      hasHeaders: true,
      summary: "Created table SalesTable on Sales!A1:F50."
    });
    expect(worksheet.tables.add).toHaveBeenCalledWith(targetRange, true);
    expect(table).toMatchObject({
      name: "SalesTable",
      style: "TableStyleMedium2",
      showBandedRows: true,
      showBandedColumns: false,
      showFilterButton: true,
      showTotals: false
    });
  });

  it("renders advisory formula-debug previews with intent metadata without treating them as write plans", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const response = {
      type: "formula",
      data: {
        intent: "fix",
        targetCell: "H11",
        formula: "=SUMIF(B:B,\"north\",F:F)",
        formulaLanguage: "excel",
        explanation: "Use the Region column as the criteria range and Revenue as the sum range.",
        confidence: 0.94
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(false);
    expect(taskpane.getStructuredPreview(response)).toMatchObject({
      kind: "formula",
      intent: "fix",
      targetCell: "H11"
    });

    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_formula_debug_001",
      requestId: "req_formula_debug_001"
    });

    expect(html).toContain("Formula");
    expect(html).toContain("fix");
    expect(html).toContain("H11");
    expect(html).toContain('=SUMIF(B:B,"north",F:F)');
  });

  it("renders external data imports as preview-only plans on Excel hosts", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const response = {
      type: "external_data_plan",
      data: {
        sourceType: "web_table_import",
        provider: "importhtml",
        sourceUrl: "https://example.com/prices",
        selectorType: "table",
        selector: 1,
        targetSheet: "Imported Data",
        targetRange: "A1",
        formula: '=IMPORTHTML("https://example.com/prices","table",1)',
        explanation: "Anchor the first public table in A1.",
        confidence: 0.87,
        requiresConfirmation: true,
        affectedRanges: ["Imported Data!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(false);
    expect(taskpane.getStructuredPreview(response)).toMatchObject({
      kind: "external_data_plan",
      provider: "importhtml",
      targetSheet: "Imported Data",
      targetRange: "A1"
    });

    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_external_data_excel_001",
      requestId: "req_external_data_excel_001"
    });

    expect(html).toContain("IMPORTHTML");
    expect(html).toContain("can't create first-class external data imports yet");
    expect(html).not.toContain("Confirm External Data");
  });

  it("preserves contract metadata across Excel structured previews", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const shared = {
      explanation: "Preview the planned workbook change.",
      confidence: 0.86,
      requiresConfirmation: true
    };
    const affectedRanges = ["Sales!A1:C10"];
    const warnings = [
      { code: "ocr_confidence", message: "One value may need review.", severity: "low" }
    ];
    const cases = [
      {
        name: "workbook_structure_update",
        response: {
          type: "workbook_structure_update",
          data: {
            operation: "create_sheet",
            sheetName: "Audit",
            overwriteRisk: "low",
            ...shared
          }
        },
        expected: {
          confidence: shared.confidence,
          requiresConfirmation: true
        }
      },
      {
        name: "sheet_structure_update",
        response: {
          type: "sheet_structure_update",
          data: {
            targetSheet: "Sales",
            operation: "hide_rows",
            startIndex: 2,
            count: 3,
            affectedRanges,
            overwriteRisk: "low",
            confirmationLevel: "standard",
            ...shared
          }
        },
        expected: {
          confidence: shared.confidence,
          requiresConfirmation: true,
          affectedRanges,
          overwriteRisk: "low"
        }
      },
      {
        name: "range_sort_plan",
        response: {
          type: "range_sort_plan",
          data: {
            targetSheet: "Sales",
            targetRange: "A1:C10",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "desc" }],
            affectedRanges,
            ...shared
          }
        },
        expected: {
          confidence: shared.confidence,
          requiresConfirmation: true,
          affectedRanges
        }
      },
      {
        name: "range_filter_plan",
        response: {
          type: "range_filter_plan",
          data: {
            targetSheet: "Sales",
            targetRange: "A1:C10",
            hasHeader: true,
            conditions: [{ columnRef: "Status", operator: "equals", value: "Closed Won" }],
            combiner: "and",
            clearExistingFilters: true,
            affectedRanges,
            ...shared
          }
        },
        expected: {
          confidence: shared.confidence,
          requiresConfirmation: true,
          affectedRanges
        }
      },
      {
        name: "range_format_update",
        response: {
          type: "range_format_update",
          data: {
            targetSheet: "Sales",
            targetRange: "A1:C10",
            format: { bold: true },
            overwriteRisk: "low",
            ...shared
          }
        },
        expected: {
          confidence: shared.confidence,
          requiresConfirmation: true
        }
      },
      {
        name: "conditional_format_plan",
        response: {
          type: "conditional_format_plan",
          data: {
            targetSheet: "Sales",
            targetRange: "C2:C10",
            managementMode: "add",
            ruleType: "color_scale",
            points: [
              { type: "min", color: "#63BE7B" },
              { type: "max", color: "#F8696B" }
            ],
            affectedRanges,
            replacesExistingRules: false,
            ...shared
          }
        },
        expected: {
          points: [
            { type: "min", color: "#63BE7B" },
            { type: "max", color: "#F8696B" }
          ],
          affectedRanges
        }
      },
      {
        name: "data_validation_plan",
        response: {
          type: "data_validation_plan",
          data: {
            targetSheet: "Sales",
            targetRange: "B2:B10",
            ruleType: "list",
            values: ["Open", "Closed"],
            showDropdown: false,
            allowBlank: false,
            invalidDataBehavior: "reject",
            affectedRanges,
            replacesExistingValidation: true,
            ...shared
          }
        },
        expected: {
          showDropdown: false,
          affectedRanges
        }
      },
      {
        name: "named_range_update",
        response: {
          type: "named_range_update",
          data: {
            operation: "retarget",
            scope: "workbook",
            name: "SalesData",
            targetSheet: "Sales",
            targetRange: "A1:C10",
            affectedRanges,
            overwriteRisk: "low",
            ...shared
          }
        },
        expected: {
          affectedRanges,
          overwriteRisk: "low"
        }
      },
      {
        name: "sheet_update",
        response: {
          type: "sheet_update",
          data: {
            targetSheet: "Sales",
            targetRange: "A1:B2",
            operation: "replace_range",
            values: [["Region", "Revenue"], ["West", 100]],
            shape: { rows: 2, columns: 2 },
            overwriteRisk: "medium",
            ...shared
          }
        },
        expected: {
          explanation: shared.explanation,
          confidence: shared.confidence,
          requiresConfirmation: true
        }
      },
      {
        name: "sheet_import_plan",
        response: {
          type: "sheet_import_plan",
          data: {
            sourceAttachmentId: "att_123",
            targetSheet: "Import",
            targetRange: "A1:B2",
            headers: ["Region", "Revenue"],
            values: [["West", 100]],
            extractionMode: "table",
            shape: { rows: 2, columns: 2 },
            confidence: 0.77,
            warnings,
            requiresConfirmation: true
          }
        },
        expected: {
          sourceAttachmentId: "att_123",
          confidence: 0.77,
          warnings,
          requiresConfirmation: true
        }
      }
    ];

    for (const testCase of cases) {
      expect(taskpane.getStructuredPreview(testCase.response), testCase.name).toMatchObject(testCase.expected);
    }
  });

  it("renders a composite preview with dry-run and destructive flags", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

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
          },
          {
            stepId: "step_cleanup",
            dependsOn: ["step_sort"],
            continueOnError: true,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:F50",
              operation: "remove_duplicate_rows",
              explanation: "Remove duplicate rows after sorting.",
              confidence: 0.85,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50"],
              overwriteRisk: "medium",
              confirmationLevel: "destructive"
            }
          }
        ],
        explanation: "Sort the sales table and then remove duplicate rows.",
        confidence: 0.89,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: true
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(true);
    expect(taskpane.getRequiresConfirmation(response)).toBe(true);
    expect(taskpane.getStructuredPreview(response)).toMatchObject({
      kind: "composite_plan",
      stepCount: 2,
      dryRunRequired: true,
      reversible: false,
      confirmationLevel: "destructive"
    });

    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_composite_preview_001",
      requestId: "req_composite_preview_001"
    });

    expect(html).toContain("Confirm Workflow");
    expect(html).toContain("Will run 2 workflow steps.");
    expect(html).toContain("dry run required");
    expect(html).toContain("step_sort");
    expect(html).toContain("step_cleanup");
    expect(html).toContain("Remove duplicate rows after sorting.");
  });

  it("flags unsupported composite steps before the user confirms the workflow", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const response = {
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_validation",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sheet1",
              targetRange: "B2:B20",
              ruleType: "list",
              allowBlank: false,
              invalidDataBehavior: "reject",
              explanation: "Apply a supported validation rule.",
              confidence: 0.9,
              requiresConfirmation: true
            }
          },
          {
            stepId: "step_cleanup",
            dependsOn: ["step_validation"],
            continueOnError: false,
            plan: {
              targetSheet: "Contacts",
              targetRange: "A2:A20",
              operation: "standardize_format",
              formatType: "date_text",
              formatPattern: "locale-sensitive-fuzzy",
              explanation: "Normalize date strings with a fuzzy locale format.",
              confidence: 0.5,
              requiresConfirmation: true
            }
          }
        ],
        explanation: "Apply validation, then normalize the imported dates.",
        confidence: 0.7,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20", "Contacts!A2:A20"],
        overwriteRisk: "medium",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(false);
    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_composite_unsupported_excel_001",
      requestId: "req_composite_unsupported_excel_001"
    });
    expect(html).toContain("Some workflow steps can't run in this Excel runtime yet.");
    expect(html).toContain("This Excel runtime only supports exact year-first date text patterns");
    expect(html).not.toContain("Confirm Workflow");
  });

  it("flags unsupported filter child steps inside composite previews", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const response = {
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_filter",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:F50",
              hasHeader: true,
              conditions: [
                { columnRef: "Status", operator: "equals", value: "Open" },
                { columnRef: "Amount", operator: "greaterThan", value: 1000 }
              ],
              combiner: "or",
              clearExistingFilters: true,
              explanation: "Keep open rows or large deals.",
              confidence: 0.86,
              requiresConfirmation: true
            }
          },
          {
            stepId: "step_validation",
            dependsOn: ["step_filter"],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "G2:G50",
              ruleType: "whole_number",
              comparator: "greater_than",
              value: 0,
              allowBlank: false,
              invalidDataBehavior: "reject",
              explanation: "Keep forecast amounts positive.",
              confidence: 0.9,
              requiresConfirmation: true
            }
          }
        ],
        explanation: "Filter the sales table, then validate forecast values.",
        confidence: 0.78,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales!G2:G50"],
        overwriteRisk: "medium",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(false);
    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_composite_filter_unsupported_excel_001",
      requestId: "req_composite_filter_unsupported_excel_001"
    });
    expect(html).toContain("Some workflow steps can't run in this Excel runtime yet.");
    expect(html).toContain("can't combine those filter conditions exactly");
    expect(html).not.toContain("Confirm Workflow");
  });

  it("marks destructive structural child steps as non-reversible in composite previews", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const preview = taskpane.getStructuredPreview({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_delete_rows",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              operation: "delete_rows",
              startIndex: 4,
              count: 2,
              explanation: "Delete two stale rows.",
              confidence: 0.82,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A5:F6"],
              overwriteRisk: "high",
              confirmationLevel: "destructive"
            }
          }
        ],
        explanation: "Delete stale rows.",
        confidence: 0.82,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A5:F6"],
        overwriteRisk: "high",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });

    expect(preview).toMatchObject({
      kind: "composite_plan",
      steps: [
        {
          stepId: "step_delete_rows",
          destructive: true,
          reversible: false
        }
      ]
    });
  });

  it("requires destructive confirmation when a composite child step is destructive even if the top-level plan is standard", async () => {
    const confirmMock = vi.fn(() => true);
    vi.stubGlobal("confirm", confirmMock);

    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    expect(taskpane.buildWriteApprovalRequest({
      requestId: "req_composite_destructive_001",
      runId: "run_composite_destructive_001",
      workbookSessionKey: "excel_windows::workbook-123",
      plan: {
        steps: [
          {
            stepId: "step_delete_rows",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              operation: "delete_rows",
              startIndex: 4,
              count: 2,
              explanation: "Delete two stale rows.",
              confidence: 0.82,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A5:F6"],
              overwriteRisk: "high",
              confirmationLevel: "destructive"
            }
          }
        ],
        explanation: "Delete stale rows.",
        confidence: 0.82,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A5:F6"],
        overwriteRisk: "high",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    })).toMatchObject({
      requestId: "req_composite_destructive_001",
      runId: "run_composite_destructive_001",
      workbookSessionKey: "excel_windows::workbook-123",
      destructiveConfirmation: {
        confirmed: true
      }
    });

    expect(confirmMock).toHaveBeenCalledTimes(1);
  });

  it("maps dry-run, history, undo, and redo through the gateway client using workbook/session scope", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_001"],
      executions: {
        exec_001: {
          baseExecutionId: "exec_001"
        }
      },
      bases: {
        exec_001: {
          baseExecutionId: "exec_001",
          targetSheet: "Sales",
          targetRange: "A1",
          beforeCells: [[{ kind: "value", value: { type: "string", value: "before" } }]],
          afterCells: [[{ kind: "value", value: { type: "string", value: "after" } }]]
        }
      }
    });
    const fetchMock = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        if (url.endsWith("/api/execution/dry-run")) {
          return {
            planDigest: "digest_001",
            workbookSessionKey: "excel_windows::budget-xlsx",
            simulated: true,
            predictedAffectedRanges: ["Sales!A1:F50"],
            predictedSummaries: ["Would sort and filter Sales!A1:F50."],
            overwriteRisk: "low",
            reversible: true,
            expiresAt: "2099-01-01T00:00:00.000Z"
          };
        }

        if (url.includes("/api/execution/history")) {
          return {
            entries: [
              {
                executionId: "exec_001",
                requestId: "req_001",
                runId: "run_001",
                planType: "composite_plan",
                planDigest: "digest_001",
                status: "completed",
                timestamp: "2099-01-01T00:00:00.000Z",
                reversible: true,
                undoEligible: true,
                redoEligible: false,
                summary: "Completed workflow."
              }
            ]
          };
        }

        if (url.endsWith("/api/execution/undo")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "excel_windows",
            executionId: "exec_undo_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }

        if (url.endsWith("/api/execution/redo/prepare")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "excel_windows",
            executionId: "exec_redo_preview_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }

        if (url.endsWith("/api/execution/redo")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "excel_windows",
            executionId: "exec_redo_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }

        return {
          kind: "composite_update",
          operation: "composite_update",
          hostPlatform: "excel_windows",
          executionId: "exec_001",
          stepResults: [],
          summary: init?.body || ""
        };
      }
    }));

    const appliedCells: Array<{ kind: string; value: unknown }> = [];
    const targetRange = {
      rowCount: 1,
      columnCount: 1,
      load() {},
      getCell() {
        return {
          set formulas(value: unknown) {
            appliedCells.push({ kind: "formula", value });
          },
          set values(value: unknown) {
            appliedCells.push({ kind: "value", value });
          }
        };
      }
    };

    const taskpane = await loadTaskpaneModule(
      {
        sync: vi.fn(async () => {}),
        workbook: {
          worksheets: {
            getItem() {
              return {
                getRange() {
                  return targetRange;
                }
              };
            }
          }
        }
      },
      {
        fetchImpl: fetchMock,
        documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
        localStorageSeed: {
          [snapshotStoreKey]: localSnapshotStore
        }
      }
    );

    await taskpane.dryRunCompositePlan({
      requestId: "req_composite_dry_run_001",
      runId: "run_composite_dry_run_001",
      plan: {
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
        explanation: "Sort the sales table.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });

    await taskpane.listExecutionHistory({ limit: 5 });
    await taskpane.undoExecution("exec_001");
    await taskpane.redoExecution("exec_undo_001");

    expect(fetchMock).toHaveBeenCalledTimes(6);
    expect(fetchMock.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/dry-run");
    expect(JSON.parse(String(fetchMock.mock.calls[0][1]?.body))).toMatchObject({
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey
    });
    expect(String(fetchMock.mock.calls[1][0])).toContain(
      `workbookSessionKey=${encodeURIComponent(workbookSessionKey)}`
    );
    expect(String(fetchMock.mock.calls[1][0])).toContain("sessionId=sess_");
    expect(String(fetchMock.mock.calls[1][0])).toContain("limit=5");
    expect(fetchMock.mock.calls[2][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(JSON.parse(String(fetchMock.mock.calls[2][1]?.body))).toMatchObject({
      executionId: "exec_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey
    });
    expect(fetchMock.mock.calls[3][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(JSON.parse(String(fetchMock.mock.calls[3][1]?.body))).toMatchObject({
      executionId: "exec_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey
    });
    expect(fetchMock.mock.calls[4][0]).toBe("http://127.0.0.1:8787/api/execution/redo/prepare");
    expect(JSON.parse(String(fetchMock.mock.calls[4][1]?.body))).toMatchObject({
      executionId: "exec_undo_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey
    });
    expect(fetchMock.mock.calls[5][0]).toBe("http://127.0.0.1:8787/api/execution/redo");
    expect(JSON.parse(String(fetchMock.mock.calls[5][1]?.body))).toMatchObject({
      executionId: "exec_undo_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey
    });
    expect(appliedCells).toEqual([
      { kind: "value", value: [["before"]] },
      { kind: "value", value: [["after"]] }
    ]);
  });

  it("fails undo and redo before calling the gateway when the local snapshot shape is stale", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_001", "exec_undo_001"],
      executions: {
        exec_001: {
          baseExecutionId: "exec_001"
        },
        exec_undo_001: {
          baseExecutionId: "exec_001"
        }
      },
      bases: {
        exec_001: {
          baseExecutionId: "exec_001",
          targetSheet: "Sales",
          targetRange: "A1",
          beforeCells: [[{ kind: "value", value: { type: "string", value: "before" } }]],
          afterCells: [[{ kind: "value", value: { type: "string", value: "after" } }]]
        }
      }
    });
    const fetchMock = vi.fn();
    const targetRange = {
      rowCount: 2,
      columnCount: 1,
      load() {},
      getCell() {
        return {
          set formulas(_value: unknown) {},
          set values(_value: unknown) {}
        };
      }
    };

    const taskpane = await loadTaskpaneModule(
      {
        sync: vi.fn(async () => {}),
        workbook: {
          worksheets: {
            getItem() {
              return {
                getRange() {
                  return targetRange;
                }
              };
            }
          }
        }
      },
      {
        fetchImpl: fetchMock,
        documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
        localStorageSeed: {
          [snapshotStoreKey]: localSnapshotStore
        }
      }
    );

    await expect(taskpane.undoExecution("exec_001")).rejects.toThrow(
      "The saved undo snapshot no longer matches the current range shape."
    );
    await expect(taskpane.redoExecution("exec_undo_001")).rejects.toThrow(
      "The saved undo snapshot no longer matches the current range shape."
    );
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("prepares undo and redo but does not commit the gateway when local snapshot apply fails", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_001", "exec_undo_001"],
      executions: {
        exec_001: {
          baseExecutionId: "exec_001"
        },
        exec_undo_001: {
          baseExecutionId: "exec_001"
        }
      },
      bases: {
        exec_001: {
          baseExecutionId: "exec_001",
          targetSheet: "Sales",
          targetRange: "A1",
          beforeCells: [[{ kind: "value", value: { type: "string", value: "before" } }]],
          afterCells: [[{ kind: "value", value: { type: "string", value: "after" } }]]
        }
      }
    });

    async function loadFailingRestoreTaskpane() {
      let syncCalls = 0;
      const fetchMock = vi.fn(async (url: string, init?: { body?: string }) => ({
        ok: true,
        async json() {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "excel_windows",
            executionId: String(url).includes("/redo/") ? "exec_redo_preview_001" : "exec_undo_preview_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }
      }));
      const targetRange = {
        rowCount: 1,
        columnCount: 1,
        load() {},
        getCell() {
          return {
            set formulas(_value: unknown) {},
            set values(_value: unknown) {}
          };
        }
      };
      const taskpane = await loadTaskpaneModule(
        {
          sync: vi.fn(async () => {
            syncCalls += 1;
            if (syncCalls === 3) {
              throw new Error("Office final sync failed.");
            }
          }),
          workbook: {
            worksheets: {
              getItem() {
                return {
                  getRange() {
                    return targetRange;
                  }
                };
              }
            }
          }
        },
        {
          fetchImpl: fetchMock,
          documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
          localStorageSeed: {
            [snapshotStoreKey]: localSnapshotStore
          }
        }
      );

      return { taskpane, fetchMock };
    }

    const undo = await loadFailingRestoreTaskpane();
    await expect(undo.taskpane.undoExecution("exec_001")).rejects.toThrow("Office final sync failed.");
    expect(undo.fetchMock).toHaveBeenCalledTimes(1);
    expect(undo.fetchMock.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(undo.fetchMock.mock.calls.some(([url]) => String(url).endsWith("/api/execution/undo"))).toBe(false);

    const redo = await loadFailingRestoreTaskpane();
    await expect(redo.taskpane.redoExecution("exec_undo_001")).rejects.toThrow("Office final sync failed.");
    expect(redo.fetchMock).toHaveBeenCalledTimes(1);
    expect(redo.fetchMock.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/redo/prepare");
    expect(redo.fetchMock.mock.calls.some(([url]) => String(url).endsWith("/api/execution/redo"))).toBe(false);
  });

  it("fails undo and redo before calling the gateway when the local snapshot store cannot persist redo lineage", async () => {
    const workbookSessionId = "workbook-123";
    const workbookSessionKey = `excel_windows::${workbookSessionId}`;
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;
    const localSnapshotStore = JSON.stringify({
      version: 1,
      order: ["exec_001", "exec_undo_001"],
      executions: {
        exec_001: {
          baseExecutionId: "exec_001"
        },
        exec_undo_001: {
          baseExecutionId: "exec_001"
        }
      },
      bases: {
        exec_001: {
          baseExecutionId: "exec_001",
          targetSheet: "Sales",
          targetRange: "A1",
          beforeCells: [[{ kind: "value", value: { type: "string", value: "before" } }]],
          afterCells: [[{ kind: "value", value: { type: "string", value: "after" } }]]
        }
      }
    });
    const fetchMock = vi.fn();
    const targetRange = {
      rowCount: 1,
      columnCount: 1,
      load() {},
      getCell() {
        return {
          set formulas(_value: unknown) {},
          set values(_value: unknown) {}
        };
      }
    };

    const taskpane = await loadTaskpaneModule(
      {
        sync: vi.fn(async () => {}),
        workbook: {
          worksheets: {
            getItem() {
              return {
                getRange() {
                  return targetRange;
                }
              };
            }
          }
        }
      },
      {
        fetchImpl: fetchMock,
        documentSettings: new Map([["Hermes.WorkbookSessionId", workbookSessionId]]),
        localStorageSeed: {
          [snapshotStoreKey]: localSnapshotStore
        }
      }
    );

    window.localStorage.setItem = vi.fn(() => {
      throw new Error("quota exceeded");
    });

    await expect(taskpane.undoExecution("exec_001")).rejects.toThrow(
      "That history entry is no longer available in this spreadsheet session."
    );
    await expect(taskpane.redoExecution("exec_undo_001")).rejects.toThrow(
      "That history entry is no longer available in this spreadsheet session."
    );
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("uses a workbook identity that does not collide for same-name files at different URLs", async () => {
    const fetchA = vi.fn(async (url: string) => ({
      ok: true,
      async json() {
        return url;
      }
    }));
    const fetchB = vi.fn(async (url: string) => ({
      ok: true,
      async json() {
        return url;
      }
    }));

    const taskpaneA = await loadTaskpaneModule(
      { sync: vi.fn(async () => {}) },
      {
        documentUrl: "https://tenant-a.example/Shared/Budget.xlsx",
        fetchImpl: fetchA
      }
    );
    await taskpaneA.listExecutionHistory({ limit: 1 });

    const taskpaneB = await loadTaskpaneModule(
      { sync: vi.fn(async () => {}) },
      {
        documentUrl: "https://tenant-b.example/Shared/Budget.xlsx",
        fetchImpl: fetchB
      }
    );
    await taskpaneB.listExecutionHistory({ limit: 1 });

    expect(fetchA.mock.calls[0][0]).not.toBe(fetchB.mock.calls[0][0]);
    expect(fetchA.mock.calls[0][0]).not.toContain("excel_windows%3A%3Abudget-xlsx");
    expect(fetchB.mock.calls[0][0]).not.toContain("excel_windows%3A%3Abudget-xlsx");
  });

  it("uses distinct workbook identities for unsaved workbooks when document settings are unavailable", async () => {
    const fetchA = vi.fn(async (url: string) => ({
      ok: true,
      async json() {
        return url;
      }
    }));
    const fetchB = vi.fn(async (url: string) => ({
      ok: true,
      async json() {
        return url;
      }
    }));

    const taskpaneA = await loadTaskpaneModule(
      { sync: vi.fn(async () => {}) },
      {
        documentUrl: "",
        disableDocumentSettings: true,
        localStorageSeed: { hermesSessionId: "sess_shared_001" },
        fetchImpl: fetchA
      }
    );
    await taskpaneA.listExecutionHistory({ limit: 1 });

    const taskpaneB = await loadTaskpaneModule(
      { sync: vi.fn(async () => {}) },
      {
        documentUrl: "",
        disableDocumentSettings: true,
        localStorageSeed: { hermesSessionId: "sess_shared_001" },
        fetchImpl: fetchB
      }
    );
    await taskpaneB.listExecutionHistory({ limit: 1 });

    expect(fetchA.mock.calls[0][0]).not.toBe(fetchB.mock.calls[0][0]);
    expect(fetchA.mock.calls[0][0]).toContain("workbookSessionKey=excel_windows%3A%3Alocal_");
    expect(fetchB.mock.calls[0][0]).toContain("workbookSessionKey=excel_windows%3A%3Alocal_");
  });

  it("applies composite update status summaries through the normal message completion path", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const message = {
      content: "Thinking...",
      response: {
        type: "composite_plan",
        data: {
          steps: [],
          explanation: "Run workflow.",
          confidence: 0.75,
          requiresConfirmation: true,
          affectedRanges: [],
          overwriteRisk: "none",
          confirmationLevel: "standard",
          reversible: true,
          dryRunRecommended: false,
          dryRunRequired: false
        }
      },
      statusLine: "Waiting"
    };

    taskpane.applyWritebackResultToMessage(message, {
      kind: "composite_update",
      operation: "composite_update",
      hostPlatform: "excel_windows",
      executionId: "exec_001",
      stepResults: [
        { stepId: "step_sort", status: "completed", summary: "Sorted table." },
        { stepId: "step_filter", status: "skipped", summary: "Skipped after dependency failure." }
      ],
      summary: "Completed workflow with 2 steps."
    });

    expect(message.content).toBe(
      "Completed workflow with 2 steps. Completed: Sorted table. Skipped: Skipped after dependency failure."
    );
    expect(message.response).toBeNull();
    expect(message.statusLine).toBe("");
  });

  it("returns a valid composite_update result when a workflow step fails closed and later steps are skipped", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {}
      }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        steps: [
          {
            stepId: "step_pivot",
            dependsOn: [],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:F50",
              targetSheet: "Sales Pivot",
              targetRange: "A1:B2",
              rowGroups: ["Region"],
              valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
              filters: [{ field: "Region", operator: "not_equal_to", value: "APAC" }],
              sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
              explanation: "Build a pivot first.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1:B2"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "step_chart",
            dependsOn: ["step_pivot"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:C20",
              targetSheet: "Sales Chart",
              targetRange: "A1",
              chartType: "line",
              categoryField: "Month",
              series: [{ field: "Revenue", label: "Revenue" }],
              explanation: "Chart the pivot output.",
              confidence: 0.88,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Build a pivot and then chart it.",
        confidence: 0.86,
        requiresConfirmation: true,
        affectedRanges: ["Sales Pivot!A1:B2", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      },
      requestId: "req_composite_apply_excel_001",
      runId: "run_composite_apply_excel_001",
      approvalToken: "token",
      executionId: "exec_composite_apply_001"
    })).resolves.toMatchObject({
      kind: "composite_update",
      operation: "composite_update",
      executionId: "exec_composite_apply_001",
      stepResults: [
        {
          stepId: "step_pivot",
          status: "failed",
          summary:
            "This action needs a valid destination cell or anchor.\n\n" +
            "Choose a single target cell or a valid destination range, then retry."
        },
        {
          stepId: "step_chart",
          status: "skipped",
          summary: "Skipped because an earlier workflow step failed."
        }
      ]
    });
  });

  it("returns child writeback proof for completed Excel composite steps", async () => {
    const targetRange = {
      rowCount: 1,
      columnCount: 2,
      values: [["", ""]],
      formulas: [["", ""]],
      load: vi.fn(),
      getCell: vi.fn()
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange)
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        steps: [
          {
            stepId: "step_write",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:B1",
              operation: "replace_range",
              values: [["Region", "Revenue"]],
              explanation: "Write headers.",
              confidence: 0.9,
              requiresConfirmation: true,
              overwriteRisk: "low",
              shape: {
                rows: 1,
                columns: 2
              }
            }
          }
        ],
        explanation: "Write headers.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:B1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      },
      requestId: "req_composite_success_excel_001",
      runId: "run_composite_success_excel_001",
      approvalToken: "token",
      executionId: "exec_composite_success_excel_001"
    })).resolves.toMatchObject({
      kind: "composite_update",
      operation: "composite_update",
      executionId: "exec_composite_success_excel_001",
      stepResults: [
        {
          stepId: "step_write",
          status: "completed",
          summary: "Wrote values to Sales!A1:B1.",
          result: {
            kind: "range_write",
            hostPlatform: "excel_windows",
            targetSheet: "Sales",
            targetRange: "A1:B1",
            operation: "replace_range",
            values: [["Region", "Revenue"]],
            writtenRows: 1,
            writtenColumns: 2
          }
        }
      ]
    });
    expect(targetRange.values).toEqual([["Region", "Revenue"]]);
  });

  it("attaches a composite undo snapshot when every Excel step captures a local snapshot", async () => {
    const targetRange = {
      rowCount: 1,
      columnCount: 2,
      values: [["", ""]],
      formulas: [["", ""]],
      load: vi.fn(),
      getCell: vi.fn()
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange)
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    });

    const result = await taskpane.applyWritePlan({
      plan: {
        steps: [
          {
            stepId: "step_write",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:B1",
              operation: "replace_range",
              values: [["Region", "Revenue"]],
              explanation: "Write headers.",
              confidence: 0.9,
              requiresConfirmation: true,
              overwriteRisk: "low",
              shape: {
                rows: 1,
                columns: 2
              }
            }
          }
        ],
        explanation: "Write headers.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:B1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      },
      requestId: "req_composite_snapshot_excel_001",
      runId: "run_composite_snapshot_excel_001",
      approvalToken: "token",
      executionId: "exec_composite_snapshot_excel_001"
    });

    expect(result).toMatchObject({
      kind: "composite_update",
      undoReady: true,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_composite_snapshot_excel_001",
        entries: [
          {
            targetSheet: "Sales",
            targetRange: "A1:B1",
            beforeCells: [
              [
                { kind: "value", value: { type: "string", value: "" } },
                { kind: "value", value: { type: "string", value: "" } }
              ]
            ],
            afterCells: [
              [
                { kind: "value", value: { type: "string", value: "Region" } },
                { kind: "value", value: { type: "string", value: "Revenue" } }
              ]
            ]
          }
        ]
      }
    });
  });

  it("fails closed when a composite step appears before an unsatisfied dependency in Excel execution order", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {}
      }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        steps: [
          {
            stepId: "step_chart",
            dependsOn: ["step_pivot"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:C20",
              targetSheet: "Sales Chart",
              targetRange: "A1",
              chartType: "line",
              categoryField: "Month",
              series: [{ field: "Revenue", label: "Revenue" }],
              explanation: "Chart before the pivot exists.",
              confidence: 0.88,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "step_pivot",
            dependsOn: [],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:F50",
              targetSheet: "Sales Pivot",
              targetRange: "A1",
              rowGroups: ["Region"],
              valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
              explanation: "Build the pivot second.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Malformed workflow order.",
        confidence: 0.7,
        requiresConfirmation: true,
        affectedRanges: ["Sales Pivot!A1", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      },
      requestId: "req_composite_order_excel_001",
      runId: "run_composite_order_excel_001",
      approvalToken: "token",
      executionId: "exec_composite_order_excel_001"
    })).resolves.toMatchObject({
      kind: "composite_update",
      operation: "composite_update",
      executionId: "exec_composite_order_excel_001",
      stepResults: [
        {
          stepId: "step_chart",
          status: "failed",
          summary: "Dependency step_pivot has not completed before this step."
        },
        {
          stepId: "step_pivot",
          status: "skipped",
          summary: "Skipped because an earlier workflow step failed."
        }
      ]
    });
  });
});
