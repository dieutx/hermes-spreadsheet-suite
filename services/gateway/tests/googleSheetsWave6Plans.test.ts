import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it, vi } from "vitest";

const require = createRequire(import.meta.url);
const TEST_FILE_PATH = fileURLToPath(import.meta.url);
const TEST_DIR = path.dirname(TEST_FILE_PATH);
const REPO_ROOT = path.resolve(TEST_DIR, "../../..");
const CODE_PATH = path.join(REPO_ROOT, "apps/google-sheets-addon/src/Code.gs");
const SIDEBAR_PATH = path.join(REPO_ROOT, "apps/google-sheets-addon/html/Sidebar.js.html");
const codeScript = fs.readFileSync(CODE_PATH, "utf8");
const sidebarHtml = fs.readFileSync(SIDEBAR_PATH, "utf8");
const sidebarScript = sidebarHtml.match(/<script>([\s\S]*)<\/script>/)?.[1] || "";

function loadCodeModule(options: {
  spreadsheet?: Record<string, unknown>;
  flush?: ReturnType<typeof vi.fn>;
  scriptProperties?: Record<string, string>;
  deploymentOverrides?: Record<string, unknown>;
  userProperties?: Record<string, string>;
  utilities?: {
    base64Decode: (value: string) => unknown;
    newBlob: (bytes: unknown, mimeType: string, fileName: string) => unknown;
  };
  urlFetchApp?: {
    fetch: ReturnType<typeof vi.fn>;
  };
} = {}) {
  const flush = options.flush || vi.fn();
  const scriptProperties = new Map<string, string>(Object.entries(options.scriptProperties || {}));
  const userProperties = new Map<string, string>(Object.entries(options.userProperties || {}));
  const documentProperties = new Map<string, string>();
  const createHtmlOutputFromFile = vi.fn((filename: string) => ({
    getContent() {
      return `<included:${filename}>`;
    }
  }));
  const showSidebar = vi.fn();
  const createTemplateFromFile = vi.fn((filename: string) => ({
    evaluate() {
      return {
        setTitle(title: string) {
          return {
            __templateFilename: filename,
            __title: title
          };
        }
      };
    }
  }));
  const context = {
    console,
    module: { exports: {} },
    exports: {},
    SpreadsheetApp: {
      getActive() {
        return options.spreadsheet;
      },
      getActiveSpreadsheet() {
        return options.spreadsheet || {
          getId() {
            return "sheet_test_001";
          }
        };
      },
      flush,
      newFilterCriteria() {
        let matchedValue: string | null = null;
        return {
          whenTextEqualTo: vi.fn(function() {
            matchedValue = arguments[0] == null ? null : String(arguments[0]);
            return this;
          }),
          build: vi.fn(function() {
            return { type: "text_equal_to", value: matchedValue };
          })
        };
      },
      PivotTableSummarizeFunction: {
        SUM: "SUM",
        COUNTA: "COUNTA",
        AVERAGE: "AVERAGE",
        MIN: "MIN",
        MAX: "MAX"
      },
      getUi() {
        return {
          createMenu() {
            return {
              addItem() {
                return this;
              },
              addToUi() {}
            };
          },
          showSidebar
        };
      },
      WrapStrategy: {
        WRAP: "WRAP",
        CLIP: "CLIP",
        OVERFLOW: "OVERFLOW"
      }
    },
    HtmlService: {
      createHtmlOutputFromFile,
      createTemplateFromFile
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty(key: string) {
            return scriptProperties.get(key) ?? null;
          },
          setProperty(key: string, value: string) {
            scriptProperties.set(key, value);
          }
        };
      },
      getUserProperties() {
        return {
          getProperty(key: string) {
            return userProperties.get(key) ?? null;
          },
          setProperty(key: string, value: string) {
            userProperties.set(key, value);
          }
        };
      },
      getDocumentProperties() {
        return {
          getProperty(key: string) {
            return documentProperties.get(key) ?? null;
          },
          setProperty(key: string, value: string) {
            documentProperties.set(key, value);
          },
          deleteProperty(key: string) {
            documentProperties.delete(key);
          }
        };
      }
    },
    Session: {
      getActiveUserLocale() {
        return "en-US";
      },
      getScriptTimeZone() {
        return "America/Los_Angeles";
      }
    },
    Utilities: options.utilities || {
      base64Decode(value: string) {
        return Buffer.from(value, "base64");
      },
      newBlob(bytes: unknown, mimeType: string, fileName: string) {
        return { bytes, mimeType, fileName };
      }
    },
    UrlFetchApp: options.urlFetchApp || {
      fetch: vi.fn()
    },
    Charts: {
      ChartType: {
        LINE: "LINE",
        BAR: "BAR",
        COLUMN: "COLUMN",
        AREA: "AREA",
        PIE: "PIE",
        SCATTER: "SCATTER"
      }
    },
    ...(options.deploymentOverrides
      ? {
          getHermesDeploymentOverrides() {
            return options.deploymentOverrides;
          }
        }
      : {}),
    ...require(path.join(REPO_ROOT, "apps/google-sheets-addon/src/Wave1Plans.js"))
  };

  vm.runInNewContext(codeScript, context, { filename: CODE_PATH });

  return {
    ...context.module.exports,
    applyWritePlan: context.applyWritePlan,
    validateExecutionCellSnapshot: context.validateExecutionCellSnapshot,
    getSpreadsheetSnapshot: context.getSpreadsheetSnapshot,
    getWorkbookSessionKey: context.getWorkbookSessionKey,
    getRuntimeConfig: context.getRuntimeConfig,
    showHermesSidebar: context.showHermesSidebar,
    __htmlServiceMocks: {
      createHtmlOutputFromFile,
      createTemplateFromFile,
      showSidebar
    },
    uploadGatewayImageAttachment: context.uploadGatewayImageAttachment,
    extractGatewayErrorMessage: context.extractGatewayErrorMessage_,
    sanitizeHostExecutionError: context.sanitizeHostExecutionError_,
    flush
  };
}

function createSidebarElementStub() {
  const listeners = new Map<string, (event?: unknown) => unknown>();
  return {
    innerHTML: "",
    value: "",
    scrollTop: 0,
    scrollHeight: 0,
    addEventListener(eventName: string, handler: (event?: unknown) => unknown) {
      listeners.set(eventName, handler);
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
    },
    trigger(eventName: string, event?: unknown) {
      const handler = listeners.get(eventName);
      if (!handler) {
        throw new Error(`No listener registered for ${eventName}`);
      }
      return handler(event);
    }
  };
}

function createRangeStub(options: {
  a1Notation: string;
  row: number;
  column: number;
  numRows: number;
  numColumns: number;
  values?: unknown[][];
  displayValues?: string[][];
  formulas?: string[][];
}) {
  let currentValues = (options.values || []).map((row) => [...row]);
  let currentDisplayValues = (options.displayValues || []).map((row) => [...row]);
  let currentFormulas = (options.formulas || []).map((row) => [...row]);

  return {
    getA1Notation() {
      return options.a1Notation;
    },
    getRow() {
      return options.row;
    },
    getColumn() {
      return options.column;
    },
    getNumRows() {
      return options.numRows;
    },
    getNumColumns() {
      return options.numColumns;
    },
    setValues: vi.fn((nextValues: unknown[][]) => {
      currentValues = nextValues.map((row) => [...row]);
      currentDisplayValues = nextValues.map((row) => row.map((value) => value == null ? "" : String(value)));
      currentFormulas = nextValues.map((row) => row.map(() => ""));
    }),
    setFormula: vi.fn((nextFormula: string) => {
      currentValues = [[nextFormula]];
      currentDisplayValues = [[nextFormula]];
      currentFormulas = [[nextFormula]];
    }),
    getValues: vi.fn(() => currentValues.map((row) => [...row])),
    getDisplayValues: vi.fn(() =>
      (currentDisplayValues.length > 0 ? currentDisplayValues : currentValues.map((row) => row.map((value) => value == null ? "" : String(value))))
        .map((row) => [...row])
    ),
    getFormulas: vi.fn(() =>
      (currentFormulas.length > 0 ? currentFormulas : currentValues.map((row) => row.map(() => "")))
        .map((row) => [...row])
    ),
    getResizedRange: vi.fn()
  };
}

function loadSidebarContext(options: { disableRandomUUID?: boolean; throwOnStorageAccess?: boolean } = {}) {
  const scriptWithoutBootstrap = sidebarScript.replace(/\n\s*initialize\(\);\s*$/, "\n");
  let uuidCounter = 0;
  let successHandler: ((value: unknown) => unknown) | null = null;
  let failureHandler: ((error: unknown) => unknown) | null = null;
  const elements = new Map([
    ["app", createSidebarElementStub()],
    ["messages", createSidebarElementStub()],
    ["prompt", createSidebarElementStub()],
    ["send-button", createSidebarElementStub()],
    ["file-input", createSidebarElementStub()],
    ["attachment-strip", createSidebarElementStub()]
  ]);

  const context = {
    console,
    window: {
      localStorage: {
        getItem() {
          if (options.throwOnStorageAccess) {
            throw new Error("localStorage blocked");
          }
          return null;
        },
        setItem() {
          if (options.throwOnStorageAccess) {
            throw new Error("localStorage blocked");
          }
        },
        removeItem() {
          if (options.throwOnStorageAccess) {
            throw new Error("localStorage blocked");
          }
        }
      },
      setInterval,
      clearInterval,
      setTimeout,
      clearTimeout,
      addEventListener() {}
    },
    crypto: options.disableRandomUUID
      ? {}
      : {
          randomUUID() {
            uuidCounter += 1;
            return `test-uuid-${uuidCounter}`;
          }
        },
    document: {
      getElementById(id: string) {
        return elements.get(id) || createSidebarElementStub();
      }
    },
    google: {
      script: {
        run: {
          withSuccessHandler(handler: (value: unknown) => unknown) {
            successHandler = handler;
            return this;
          },
          withFailureHandler(handler: (error: unknown) => unknown) {
            failureHandler = handler;
            return this;
          },
          getRuntimeConfig() {
            successHandler?.({
              gatewayBaseUrl: "http://127.0.0.1:8787",
              clientVersion: "google-sheets-addon-dev",
              reviewerSafeMode: false,
              forceExtractionMode: null
            });
          },
          consumePrefillPrompt() {
            successHandler?.(null);
          }
        }
      }
    },
    fetch: vi.fn(),
    URL: {
      createObjectURL() {
        return "blob:test";
      },
      revokeObjectURL() {}
    },
    FormData: class FormData {
      set() {}
    }
  };
  void failureHandler;

  vm.runInNewContext(scriptWithoutBootstrap, context, { filename: SIDEBAR_PATH });
  vm.runInNewContext(
    "this.__sidebarTestHooks = { state, elements, renderMessages, sendPrompt, pollRun, sanitizeConversation, sanitizeHostExecutionError, buildRequestEnvelope, pruneStoredMessages, trimMessageTraceEvents, ensureRuntimeConfig, parseGatewayJsonResponse, initialize };",
    context,
    { filename: `${SIDEBAR_PATH}#test-hooks` }
  );
  return context;
}

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

describe("Google Sheets wave 6 composite plans and execution controls", () => {
  it("sanitizes host execution failures in both Sidebar and Code.gs", () => {
    const sidebar = loadSidebarContext();
    const code = loadCodeModule();

    const expectedTargetRangeMessage =
      "The spreadsheet changed, so the approved destination no longer matches the intended shape.\n\n" +
      "Refresh the spreadsheet state and run the request again.";
    const expectedAppendMessage =
      "The chosen destination range cannot accept this write safely.\n\n" +
      "Choose a clean target range or ask Hermes to write into a blank area.";

    expect(sidebar.sanitizeHostExecutionError(
      new Error("The approved targetRange does not match the proposed shape.")
    )).toBe(expectedTargetRangeMessage);
    expect(code.sanitizeHostExecutionError(
      new Error("The approved targetRange does not match the transfer shape.")
    )).toBe(expectedTargetRangeMessage);
    expect(code.sanitizeHostExecutionError(
      new Error("Google Sheets host cannot append exactly within the approved target range.")
    )).toBe(expectedAppendMessage);

    expect(sidebar.sanitizeHostExecutionError(
      new Error("Invalid date literal: 2026-13-40")
    )).toBe(
      "The date \"2026-13-40\" is not valid.\n\n" +
      "Use a real calendar date such as 2026-04-22, then retry."
    );

    expect(code.sanitizeHostExecutionError(
      new Error("Google Sheets host cannot find chart field in header row: Revenue.")
    )).toBe(
      "Column \"Revenue\" was not found in the header row.\n\n" +
      "Select the full table with headers, or use the exact column name in the request and retry."
    );

    expect(code.sanitizeHostExecutionError(
      new Error("Google Sheets grid filters cannot represent combiner \"or\" exactly for multiple conditions.")
    )).toBe(
      "This spreadsheet app cannot combine those filter conditions in one exact step.\n\n" +
      "Use a single filter rule per column, or split the filter into smaller steps."
    );

    expect(sidebar.sanitizeHostExecutionError(
      new Error("Hermes gateway URL is not configured.")
    )).toBe(
      "The Hermes connection is not configured for this sheet.\n\n" +
      "Set the Hermes gateway URL, reload the sidebar, and retry."
    );

    expect(sidebar.sanitizeHostExecutionError(
      new Error("The requested resource doesn't exist.")
    )).toBe(
      "Hermes could not read the current sheet or selection.\n\n" +
      "Select a normal sheet range, reload the sidebar, and retry. If it keeps happening, reopen the spreadsheet and try again."
    );

    expect(code.sanitizeHostExecutionError(
      new Error("List validation requires values, sourceRange, or namedRangeName.")
    )).toBe(
      "This validation setup cannot be represented safely here.\n\n" +
      "Try a simpler dropdown, checkbox, or date rule, then retry."
    );
  });

  it("fails fast when Google Sheets image upload is pointed at a loopback gateway", () => {
    const fetch = vi.fn();
    const code = loadCodeModule({
      scriptProperties: {
        HERMES_GATEWAY_URL: "http://127.0.0.1:8787"
      },
      urlFetchApp: { fetch }
    });

    expect(() => code.uploadGatewayImageAttachment({
      fileName: "table.png",
      mimeType: "image/png",
      base64Data: "aGVsbG8=",
      source: "upload",
      sessionId: "sess_upload_001"
    })).toThrow(
      "Google Sheets image upload requires a public Hermes gateway URL.\n\n" +
      "Set HERMES_GATEWAY_URL to a reachable HTTPS or public address, then retry the upload."
    );
    expect(fetch).not.toHaveBeenCalled();
  });

  it("fails fast when Google Sheets image upload is pointed at a private or loopback IPv6 gateway", () => {
    const fetch = vi.fn();
    const code = loadCodeModule({
      scriptProperties: {
        HERMES_GATEWAY_URL: "https://[::ffff:127.0.0.1]:8787"
      },
      urlFetchApp: { fetch }
    });

    expect(() => code.uploadGatewayImageAttachment({
      fileName: "table.png",
      mimeType: "image/png",
      base64Data: "aGVsbG8=",
      source: "upload",
      sessionId: "sess_upload_ipv6_001"
    })).toThrow(
      "Google Sheets image upload requires a public Hermes gateway URL.\n\n" +
      "Set HERMES_GATEWAY_URL to a reachable HTTPS or public address, then retry the upload."
    );
    expect(fetch).not.toHaveBeenCalled();
  });

  it("prefers generated deployment overrides over script properties in Code.gs runtime config", () => {
    const code = loadCodeModule({
      scriptProperties: {
        HERMES_GATEWAY_URL: "https://stale.example.test",
        HERMES_CLIENT_VERSION: "google-sheets-addon-stale",
        HERMES_REVIEWER_SAFE_MODE: "false",
        HERMES_FORCE_EXTRACTION_MODE: "real"
      },
      deploymentOverrides: {
        gatewayBaseUrl: "https://gateway.example.test",
        clientVersion: "google-sheets-addon-live-demo",
        reviewerSafeMode: true,
        forceExtractionMode: "demo"
      }
    });

    expect(code.getRuntimeConfig()).toEqual({
      gatewayBaseUrl: "https://gateway.example.test",
      clientVersion: "google-sheets-addon-live-demo",
      reviewerSafeMode: true,
      forceExtractionMode: "demo"
    });
  });

  it("defaults the Google Sheets runtime config to an unconfigured gateway instead of localhost", () => {
    const code = loadCodeModule();

    expect(code.getRuntimeConfig()).toEqual({
      gatewayBaseUrl: "",
      clientVersion: "google-sheets-addon-live-demo",
      reviewerSafeMode: false,
      forceExtractionMode: null
    });
  });

  it("loads the Google Sheets sidebar from the staged html/Sidebar template path", () => {
    const code = loadCodeModule();

    code.showHermesSidebar();

    expect(code.__htmlServiceMocks.createTemplateFromFile).toHaveBeenCalledWith("html/Sidebar");
    expect(code.__htmlServiceMocks.showSidebar).toHaveBeenCalledTimes(1);
  });

  it("includes the Hermes session id when Google Sheets uploads an image to a public gateway", () => {
    const fetch = vi.fn(() => ({
      getResponseCode() {
        return 201;
      },
      getContentText() {
        return JSON.stringify({
          attachment: {
            id: "att_001",
            type: "image",
            mimeType: "image/png",
            fileName: "table.png",
            size: 5,
            source: "upload",
            previewUrl: "https://gateway.test/api/uploads/att_001/content?uploadToken=upl_001",
            uploadToken: "upl_001",
            storageRef: "blob://att_001"
          }
        });
      }
    }));
    const code = loadCodeModule({
      scriptProperties: {
        HERMES_GATEWAY_URL: "https://gateway.test"
      },
      urlFetchApp: { fetch },
      utilities: {
        base64Decode(value: string) {
          return Buffer.from(value, "base64");
        },
        newBlob(bytes: unknown, mimeType: string, fileName: string) {
          return { bytes, mimeType, fileName };
        }
      }
    });

    const uploaded = code.uploadGatewayImageAttachment({
      fileName: "table.png",
      mimeType: "image/png",
      base64Data: "aGVsbG8=",
      source: "upload",
      sessionId: "sess_upload_001"
    });

    expect(uploaded).toMatchObject({
      attachment: {
        id: "att_001",
        uploadToken: "upl_001"
      }
    });
    expect(fetch).toHaveBeenCalledWith(
      "https://gateway.test/api/uploads/image",
      expect.objectContaining({
        method: "post",
        muteHttpExceptions: true,
        payload: expect.objectContaining({
          source: "upload",
          sessionId: "sess_upload_001"
        })
      })
    );
  });

  it("truncates oversized prompt and conversation content before building the Google Sheets request envelope", () => {
    const sidebar = loadSidebarContext();
    const oversized = "x".repeat(16_050);
    sidebar.__sidebarTestHooks.state.runtimeConfig = {
      clientVersion: "google-sheets-addon-live-demo",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    sidebar.__sidebarTestHooks.state.sessionId = "sess_test";

    const request = sidebar.__sidebarTestHooks.buildRequestEnvelope({
      userMessage: oversized,
      conversation: [
        { role: "assistant", content: oversized },
        { role: "user", content: "short" }
      ],
      snapshot: {
        host: {
          platform: "google_sheets",
          workbookTitle: "Revenue Demo",
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

  it("truncates oversized Google Sheets spreadsheet-context strings before building the host snapshot", () => {
    const longHeader = "H".repeat(300);
    const longValue = "V".repeat(4500);
    const longFormula = `=${"A".repeat(17000)}`;
    const longNote = "N".repeat(4500);
    const selectionRange = createRangeStub({
      a1Notation: "A1:B2",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        [longHeader, "Revenue"],
        [longValue, 123]
      ]
    });
    selectionRange.getFormulas = vi.fn(() => [
      ["", ""],
      [longFormula, ""]
    ]);

    const currentRegion = createRangeStub({
      a1Notation: "A1:B2",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        [longHeader, "Revenue"],
        [longValue, 123]
      ]
    });
    currentRegion.getFormulas = vi.fn(() => [
      ["", ""],
      [longFormula, ""]
    ]);
    currentRegion.offset = vi.fn(() => createRangeStub({
      a1Notation: "A1:B1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 2,
      values: [[longHeader, "Revenue"]]
    }));

    const activeCell = {
      getA1Notation: vi.fn(() => "A2"),
      getDisplayValue: vi.fn(() => longValue),
      getValue: vi.fn(() => longValue),
      getFormula: vi.fn(() => longFormula),
      getNote: vi.fn(() => longNote),
      getDataRegion: vi.fn(() => currentRegion)
    };
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      getRange: vi.fn()
    };
    const spreadsheet = {
      getActiveSheet: vi.fn(() => sheet),
      getActiveRange: vi.fn(() => selectionRange),
      getCurrentCell: vi.fn(() => activeCell),
      getName: vi.fn(() => "Revenue Demo"),
      getId: vi.fn(() => "spreadsheet_123")
    };

    const code = loadCodeModule({ spreadsheet });
    const snapshot = code.getSpreadsheetSnapshot("Explain the current selection");

    expect(snapshot.context.selection.headers[0]).toHaveLength(256);
    expect(snapshot.context.selection.headers[0].endsWith("…")).toBe(true);
    expect(String(snapshot.context.selection.values[1][0])).toHaveLength(4000);
    expect(String(snapshot.context.selection.values[1][0]).endsWith("…")).toBe(true);
    expect(snapshot.context.selection.formulas[1][0]).toHaveLength(16000);
    expect(snapshot.context.selection.formulas[1][0].endsWith("…")).toBe(true);
    expect(String(snapshot.context.activeCell.displayValue)).toHaveLength(4000);
    expect(String(snapshot.context.activeCell.displayValue).endsWith("…")).toBe(true);
    expect(snapshot.context.activeCell.formula).toHaveLength(16000);
    expect(snapshot.context.activeCell.formula.endsWith("…")).toBe(true);
    expect(snapshot.context.activeCell.note).toHaveLength(4000);
    expect(snapshot.context.activeCell.note.endsWith("…")).toBe(true);
  });

  it("does not overlap Google Sheets poll requests when a previous poll is still in flight", async () => {
    vi.useFakeTimers();
    const sidebar = loadSidebarContext();
    sidebar.fetch = vi.fn(() => new Promise(() => {}));

    void sidebar.__sidebarTestHooks.pollRun({
      runId: "run_poll_001",
      requestId: "req_poll_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    });

    await vi.advanceTimersByTimeAsync(900);
    expect(sidebar.fetch).toHaveBeenCalledTimes(1);

    await vi.advanceTimersByTimeAsync(5000);
    expect(sidebar.fetch).toHaveBeenCalledTimes(1);
  });

  it("loads the Google Sheets sidebar without crypto.randomUUID and still creates request ids", () => {
    const sidebar = loadSidebarContext({ disableRandomUUID: true });
    const hooks = (sidebar as any).__sidebarTestHooks;

    hooks.state.runtimeConfig = {
      gatewayBaseUrl: "https://example.test/hermes-gateway",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    hooks.state.sessionId = "sess_test";

    const request = hooks.buildRequestEnvelope({
      userMessage: "Explain this selection",
      conversation: [{ role: "user", content: "Explain this selection" }],
      snapshot: {
        source: {
          channel: "google_sheets",
          clientVersion: "test-client",
          sessionId: "sess_test"
        },
        host: {
          platform: "google_sheets",
          workbookTitle: "Sheet Demo",
          activeSheet: "Sheet1"
        },
        context: {}
      },
      attachments: []
    });

    expect(request.requestId).toMatch(/^req_/);
  });

  it("loads the Google Sheets sidebar when localStorage access is blocked", () => {
    const sidebar = loadSidebarContext({ throwOnStorageAccess: true });
    const hooks = (sidebar as any).__sidebarTestHooks;

    const request = hooks.buildRequestEnvelope({
      userMessage: "Explain this selection",
      conversation: [{ role: "user", content: "Explain this selection" }],
      snapshot: {
        source: {
          channel: "google_sheets",
          clientVersion: "test-client",
          sessionId: "sess_test"
        },
        host: {
          platform: "google_sheets",
          workbookTitle: "Sheet Demo",
          activeSheet: "Sheet1"
        },
        context: {}
      },
      attachments: []
    });

    expect(request.requestId).toMatch(/^req_/);
    expect(request.source.sessionId).toMatch(/^sess_/);
  });

  it("routes natural-language undo prompts to execution control instead of sending them through the model", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;
    const workbookSessionKey = "google_sheets::sheet-123";
    const snapshotStoreKey = `Hermes.ReversibleExecutions.v1::${workbookSessionKey}`;

    sidebar.window.localStorage.getItem = vi.fn((key: string) => key === snapshotStoreKey
      ? JSON.stringify({
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
        })
      : null);
    sidebar.window.localStorage.setItem = vi.fn();

    let successHandler: ((value: unknown) => unknown) | null = null;
    let failureHandler: ((error: unknown) => unknown) | null = null;
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler(handler: (error: unknown) => unknown) {
        failureHandler = handler;
        void failureHandler;
        return this;
      },
      getWorkbookSessionKey() {
        successHandler?.(workbookSessionKey);
      },
      getRuntimeConfig() {
        successHandler?.({
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        });
      },
      validateExecutionCellSnapshot(payload: unknown) {
        successHandler?.(payload);
      },
      applyExecutionCellSnapshot() {
        successHandler?.(null);
      },
      getSpreadsheetSnapshot() {
        throw new Error("sendPrompt should not request a fresh snapshot for undo.");
      }
    };

    const fetchMock = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      json: async () => {
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

        if (url.endsWith("/api/execution/undo")) {
          expect(JSON.parse(String(init?.body))).toMatchObject({
            executionId: "exec_001",
            workbookSessionKey
          });
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "google_sheets",
            executionId: "exec_undo_001",
            stepResults: [],
            summary: "Undid Sheet8!A1."
          };
        }

        throw new Error(`Unexpected fetch URL: ${url}`);
      }
    }));
    sidebar.fetch = fetchMock;

    hooks.elements.prompt.value = "undo";
    await hooks.sendPrompt();

    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls[0][0]).toBe(
      "http://127.0.0.1:8787/api/execution/history?workbookSessionKey=google_sheets%3A%3Asheet-123&limit=20"
    );
    expect(fetchMock.mock.calls[1][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(hooks.elements.messages.innerHTML).toContain("Undid Sheet8!A1.");
  });

  it("handles bare affirmations locally instead of sending an under-specified follow-up back through the model", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;
    const fetchMock = vi.fn();
    sidebar.fetch = fetchMock;

    hooks.state.messages = [
      {
        role: "assistant",
        content: "If you want, ask me to restore a specific range by describing what should be put back in Sheet8!A1:K33."
      }
    ];
    hooks.elements.prompt.value = "yep";

    await hooks.sendPrompt();

    expect(fetchMock).not.toHaveBeenCalled();
    expect(hooks.elements.messages.innerHTML).toContain(
      "I still need the exact range, cell, sheet, or action you want me to apply."
    );
    expect(hooks.elements.messages.innerHTML).toContain("Need more detail");
  });

  it("renders message status lines so confirm-path host errors are visible in the sidebar", () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    hooks.state.messages = [
      {
        role: "assistant",
        content: "Prepared a chart preview for Sheet1!A78.",
        statusLine: "Write-back failed."
      }
    ];
    hooks.renderMessages();

    expect(hooks.elements.messages.innerHTML).toContain("Write-back failed.");
  });

  it("keeps the Google Sheets sidebar pinned to the latest message after deferred layout growth", () => {
    vi.useFakeTimers();
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;
    const scrollIntoView = vi.fn();
    const messagesElement = hooks.elements.messages as {
      scrollTop: number;
      scrollHeight: number;
      lastElementChild?: { scrollIntoView?: ReturnType<typeof vi.fn> };
    };

    hooks.state.messages = [
      { role: "user", content: "Summarize this sheet." },
      { role: "assistant", content: "Thinking..." }
    ];
    messagesElement.scrollTop = 0;
    messagesElement.scrollHeight = 120;
    messagesElement.lastElementChild = { scrollIntoView };

    hooks.renderMessages();

    expect(messagesElement.scrollTop).toBe(120);

    messagesElement.scrollHeight = 420;
    vi.advanceTimersByTime(280);

    expect(messagesElement.scrollTop).toBe(420);
    expect(scrollIntoView).toHaveBeenCalled();
  });

  it("loads runtime config before sending a prompt when initialize has not finished yet", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;
    hooks.elements.prompt.value = "Summarize the current sheet";

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/requests") {
        return {
          ok: true,
          json: async () => ({ runId: "run_send_001" })
        };
      }

      if (url.startsWith("http://gateway.test/api/trace/run_send_001?")) {
        return {
          ok: true,
          json: async () => ({ events: [], nextIndex: 0 })
        };
      }

      if (url.startsWith("http://gateway.test/api/requests/run_send_001?")) {
        expect(url).toContain("includeTrace=0");
        return {
          ok: true,
          json: async () => ({
            status: "completed",
            response: {
              type: "chat",
              data: {
                message: "Done."
              },
              trace: []
            }
          })
        };
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });
    sidebar.fetch = fetchMock;

    let successHandler: ((value: unknown) => unknown) | null = null;
    let failureHandler: ((error: unknown) => unknown) | null = null;
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler(handler: (error: unknown) => unknown) {
        failureHandler = handler;
        return this;
      },
      getRuntimeConfig() {
        successHandler?.({
          gatewayBaseUrl: "http://gateway.test",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        });
      },
      getSpreadsheetSnapshot() {
        successHandler?.({
          host: {
            platform: "google_sheets",
            workbookTitle: "Sheet Demo",
            activeSheet: "Sheet1"
          },
          context: {}
        });
      },
      consumePrefillPrompt() {
        successHandler?.(null);
      }
    };
    void failureHandler;

    await hooks.sendPrompt();

    expect(fetchMock.mock.calls[0]?.[0]).toBe("http://gateway.test/api/requests");
    expect(hooks.state.runtimeConfig).toMatchObject({
      gatewayBaseUrl: "http://gateway.test",
      clientVersion: "google-sheets-addon-dev"
    });
  });

  it("does not fail open to localhost when Google Sheets runtime config loading fails", async () => {
    const sidebar = loadSidebarContext();
    sidebar.fetch = vi.fn(async () => {
      throw new Error("Unexpected fetch");
    });

    let successHandler: ((value: unknown) => unknown) | null = null;
    let failureHandler: ((error: unknown) => unknown) | null = null;
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler(handler: (error: unknown) => unknown) {
        failureHandler = handler;
        return this;
      },
      getRuntimeConfig() {
        failureHandler?.(new Error("runtime config unavailable"));
      },
      proxyGatewayJson(payload: unknown) {
        successHandler?.({
          proxied: true,
          path: (payload as { path: string }).path
        });
      }
    };

    const response = await sidebar.callGatewayJson("/api/requests", {
      method: "post",
      body: { requestId: "req_proxy_config_failure_001" }
    });

    expect(response).toEqual({
      proxied: true,
      path: "/api/requests"
    });
    expect(sidebar.fetch).not.toHaveBeenCalled();
    expect(sidebar.__sidebarTestHooks.state.runtimeConfig).toMatchObject({
      gatewayBaseUrl: "",
      clientVersion: "google-sheets-addon-dev"
    });
  });

  it("keeps Google Sheets snapshot session identity aligned with the sidebar session id", () => {
    const selectionRange = createRangeStub({
      a1Notation: "A1:B2",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        ["EEID", "Name"],
        ["E001", "Ada"]
      ]
    });
    const currentRegion = createRangeStub({
      a1Notation: "A1:B2",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        ["EEID", "Name"],
        ["E001", "Ada"]
      ]
    });
    currentRegion.offset = vi.fn(() => createRangeStub({
      a1Notation: "A1:B1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 2,
      values: [["EEID", "Name"]]
    }));

    const activeCell = {
      getA1Notation: vi.fn(() => "A2"),
      getDisplayValue: vi.fn(() => "E001"),
      getValue: vi.fn(() => "E001"),
      getFormula: vi.fn(() => ""),
      getNote: vi.fn(() => ""),
      getDataRegion: vi.fn(() => currentRegion)
    };
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      getRange: vi.fn()
    };
    const spreadsheet = {
      getActiveSheet: vi.fn(() => sheet),
      getActiveRange: vi.fn(() => selectionRange),
      getCurrentCell: vi.fn(() => activeCell),
      getName: vi.fn(() => "Employee Demo"),
      getId: vi.fn(() => "spreadsheet_123")
    };

    const code = loadCodeModule({ spreadsheet });
    const snapshot = code.getSpreadsheetSnapshot("Explain the current selection", "sess_sidebar_123");
    const followUpSnapshot = code.getSpreadsheetSnapshot("Explain the current selection");

    expect(snapshot.source.sessionId).toBe("sess_sidebar_123");
    expect(followUpSnapshot.source.sessionId).toBe("sess_sidebar_123");
  });

  it("polls only run status and does not request live trace events in Google Sheets", async () => {
    vi.useFakeTimers();
    const sidebar = loadSidebarContext();
    sidebar.fetch = vi.fn()
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

    const message = {
      runId: "run_poll_trace_gone_001",
      requestId: "req_poll_trace_gone_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    };

    void sidebar.__sidebarTestHooks.pollRun(message);

    await vi.advanceTimersByTimeAsync(900);

    expect(sidebar.fetch).toHaveBeenCalledTimes(1);
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("/api/requests/run_poll_trace_gone_001");
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("includeTrace=0");
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).not.toContain("/api/trace/");
    expect(message.content).toBe("Done.");
    expect(message.statusLine).not.toContain("Request failed");
    expect(message.tracePollingDisabled).toBe(true);
    expect(message.response.type).toBe("chat");
  });

  it("continues polling run status in Google Sheets without any trace call even across repeated attempts", async () => {
    vi.useFakeTimers();
    const sidebar = loadSidebarContext();
    sidebar.fetch = vi.fn()
      .mockResolvedValueOnce(
        new Response(
          JSON.stringify({
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
          }),
          {
            status: 200,
            headers: { "content-type": "application/json" }
          }
        )
      );

    const message = {
      runId: "run_poll_quota_001",
      requestId: "req_poll_quota_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    };

    void sidebar.__sidebarTestHooks.pollRun(message);

    await vi.advanceTimersByTimeAsync(900);

    expect(sidebar.fetch).toHaveBeenCalledTimes(1);
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("/api/requests/run_poll_quota_001");
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("includeTrace=0");
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).not.toContain("/api/trace/");
    expect(message.content).toBe("Done.");
    expect(message.statusLine).not.toContain("Request failed");
    expect(message.tracePollingDisabled).toBe(true);
    expect(message.response.type).toBe("chat");
  });

  it("does not proxy Google Sheets run-status polling when direct fetch fails", async () => {
    vi.useFakeTimers();
    const sidebar = loadSidebarContext();
    sidebar.__sidebarTestHooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://127.0.0.1:8787",
      clientVersion: "google-sheets-addon-dev",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    sidebar.fetch = vi.fn(async () => {
      throw new TypeError("Failed to fetch");
    });
    sidebar.callServer = vi.fn(async (functionName: string, payload: unknown) => {
      if (functionName === "proxyGatewayJson") {
        throw new Error(`Unexpected server call: ${functionName} ${JSON.stringify(payload)}`);
      }

      return null;
    });

    const message = {
      runId: "run_poll_no_proxy_001",
      requestId: "req_poll_no_proxy_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    };

    void sidebar.__sidebarTestHooks.pollRun(message);

    await vi.advanceTimersByTimeAsync(900);

    expect(sidebar.fetch).toHaveBeenCalledTimes(1);
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("/api/requests/run_poll_no_proxy_001");
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("includeTrace=0");
    expect(
      sidebar.callServer.mock.calls.filter((call: unknown[]) => call[0] === "proxyGatewayJson")
    ).toHaveLength(0);
    expect(message.content).toContain("The Hermes service could not be reached.");
    expect(message.statusLine).toBe("Request failed");
  });

  it("caps stored Google Sheets messages and per-message trace history to keep long sessions responsive", () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    const messages = Array.from({ length: 130 }, (_, index) => ({
      role: index % 2 === 0 ? "user" : "assistant",
      content: `message_${index}`
    }));
    const traces = Array.from({ length: 260 }, (_, index) => ({
      event: "result_generated",
      timestamp: `2026-04-22T00:00:${String(index % 60).padStart(2, "0")}.000Z`,
      label: `trace_${index}`
    }));

    const trimmedMessages = hooks.pruneStoredMessages(messages);
    const trimmedTrace = hooks.trimMessageTraceEvents(traces);

    expect(trimmedMessages).toHaveLength(100);
    expect(trimmedMessages[0].content).toBe("message_30");
    expect(trimmedMessages.at(-1)?.content).toBe("message_129");
    expect(trimmedTrace).toHaveLength(200);
    expect(trimmedTrace[0].label).toBe("trace_60");
    expect(trimmedTrace.at(-1)?.label).toBe("trace_259");
  });

  it("does not fetch the full currentRegion matrix when the current table exceeds the inline threshold", () => {
    const selectionRange = createRangeStub({
      a1Notation: "J6",
      row: 6,
      column: 10,
      numRows: 1,
      numColumns: 1,
      values: [[123]],
      displayValues: [["123"]]
    });
    const currentRegion = createRangeStub({
      a1Notation: "A1:T500",
      row: 1,
      column: 1,
      numRows: 500,
      numColumns: 20,
      values: [[123]],
      displayValues: [["123"]]
    });
    const headerRange = createRangeStub({
      a1Notation: "A1:T1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 20,
      values: [[
        "Date", "Category", "Product", "Region", "Units",
        "Revenue", "Rep", "Channel", "Segment", "Discount",
        "COGS", "Margin", "City", "State", "Country",
        "Quarter", "Month", "Year", "Customer", "Order ID"
      ]]
    });
    currentRegion.offset = vi.fn(() => headerRange);

    const activeCell = {
      getA1Notation: vi.fn(() => "J6"),
      getDisplayValue: vi.fn(() => "123"),
      getValue: vi.fn(() => 123),
      getFormula: vi.fn(() => ""),
      getNote: vi.fn(() => ""),
      getDataRegion: vi.fn(() => currentRegion)
    };
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      getRange: vi.fn()
    };
    const spreadsheet = {
      getActiveSheet: vi.fn(() => sheet),
      getActiveRange: vi.fn(() => selectionRange),
      getCurrentCell: vi.fn(() => activeCell),
      getName: vi.fn(() => "Revenue Demo"),
      getId: vi.fn(() => "spreadsheet_123")
    };

    const code = loadCodeModule({ spreadsheet });
    const snapshot = code.getSpreadsheetSnapshot("Explain the current selection");

    expect(currentRegion.offset).toHaveBeenCalledWith(0, 0, 1, 20);
    expect(currentRegion.getValues).not.toHaveBeenCalled();
    expect(currentRegion.getDisplayValues).not.toHaveBeenCalled();
    expect(currentRegion.getFormulas).not.toHaveBeenCalled();
    expect(snapshot.context.currentRegion).toMatchObject({
      range: "A1:T500",
      headers: headerRange.getValues()[0]
    });
    expect(snapshot.context.currentRegion.values).toBeUndefined();
    expect(snapshot.context.currentRegion.formulas).toBeUndefined();
    expect(snapshot.context.currentRegionArtifactTarget).toBe("A502");
    expect(snapshot.context.currentRegionAppendTarget).toBe("A501:T501");
  });

  it("does not fetch the full selected range matrix when the selected range exceeds the inline threshold", () => {
    const selectionRange = createRangeStub({
      a1Notation: "A1:T500",
      row: 1,
      column: 1,
      numRows: 500,
      numColumns: 20,
      values: [[123]],
      displayValues: [["123"]]
    });
    const selectionHeaderRange = createRangeStub({
      a1Notation: "A1:T1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 20,
      values: [[
        "Date", "Category", "Product", "Region", "Units",
        "Revenue", "Rep", "Channel", "Segment", "Discount",
        "COGS", "Margin", "City", "State", "Country",
        "Quarter", "Month", "Year", "Customer", "Order ID"
      ]]
    });
    selectionRange.offset = vi.fn(() => selectionHeaderRange);

    const currentRegion = createRangeStub({
      a1Notation: "A1:F10",
      row: 1,
      column: 1,
      numRows: 10,
      numColumns: 6,
      values: [[
        "Date", "Category", "Product", "Region", "Units", "Revenue"
      ]]
    });
    currentRegion.offset = vi.fn(() => createRangeStub({
      a1Notation: "A1:F1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 6,
      values: [[
        "Date", "Category", "Product", "Region", "Units", "Revenue"
      ]]
    }));

    const activeCell = {
      getA1Notation: vi.fn(() => "J6"),
      getDisplayValue: vi.fn(() => "123"),
      getValue: vi.fn(() => 123),
      getFormula: vi.fn(() => ""),
      getNote: vi.fn(() => ""),
      getDataRegion: vi.fn(() => currentRegion)
    };
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      getRange: vi.fn()
    };
    const spreadsheet = {
      getActiveSheet: vi.fn(() => sheet),
      getActiveRange: vi.fn(() => selectionRange),
      getCurrentCell: vi.fn(() => activeCell),
      getName: vi.fn(() => "Revenue Demo"),
      getId: vi.fn(() => "spreadsheet_123")
    };

    const code = loadCodeModule({ spreadsheet });
    const snapshot = code.getSpreadsheetSnapshot("Explain the current selection");

    expect(selectionRange.offset).toHaveBeenCalledWith(0, 0, 1, 20);
    expect(selectionRange.getValues).not.toHaveBeenCalled();
    expect(selectionRange.getDisplayValues).not.toHaveBeenCalled();
    expect(selectionRange.getFormulas).not.toHaveBeenCalled();
    expect(snapshot.context.selection).toMatchObject({
      range: "A1:T500",
      headers: selectionHeaderRange.getValues()[0]
    });
    expect(snapshot.context.selection.values).toBeUndefined();
    expect(snapshot.context.selection.formulas).toBeUndefined();
  });

  it("renders advisory formula-debug previews with intent metadata without treating them as write plans", () => {
    const sidebar = loadSidebarContext();

    const response = {
      type: "formula",
      data: {
        intent: "explain",
        targetCell: "F6",
        formula: "=SUMIF(B:B,,F:F)",
        formulaLanguage: "google_sheets",
        explanation: "The criteria argument is blank, so the formula cannot match the intended rows.",
        confidence: 0.93
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(false);
    expect(sidebar.getStructuredPreview(response)).toMatchObject({
      kind: "formula",
      intent: "explain",
      targetCell: "F6"
    });

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_formula_debug_001",
      requestId: "req_formula_debug_001"
    });

    expect(html).toContain("Formula");
    expect(html).toContain("explain");
    expect(html).toContain("F6");
    expect(html).toContain("=SUMIF(B:B,,F:F)");
  });

  it("keeps external data imports as first-class confirmable previews in Google Sheets", () => {
    const sidebar = loadSidebarContext();

    const response = {
      type: "external_data_plan",
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
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.getStructuredPreview(response)).toMatchObject({
      kind: "external_data_plan",
      sourceType: "market_data",
      provider: "googlefinance",
      targetSheet: "Market Data",
      targetRange: "B2",
      query: {
        symbol: "CURRENCY:BTCUSD",
        attribute: "price"
      }
    });

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_external_data_preview_001",
      requestId: "req_external_data_preview_001"
    });

    expect(html).toContain("Confirm External Data");
    expect(html).toContain("GOOGLEFINANCE");
    expect(html).toContain("symbol CURRENCY:BTCUSD");
  });

  it("renders a composite preview with dry-run and destructive flags and requires destructive confirmation for destructive child steps", () => {
    const sidebar = loadSidebarContext();
    const confirmMock = vi.fn(() => true);
    sidebar.window.confirm = confirmMock;

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
            stepId: "step_delete_rows",
            dependsOn: ["step_sort"],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              operation: "delete_rows",
              startIndex: 4,
              count: 2,
              explanation: "Delete stale rows after sorting.",
              confidence: 0.82,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A5:F6"],
              overwriteRisk: "high",
              confirmationLevel: "destructive"
            }
          }
        ],
        explanation: "Sort the table and then delete stale rows.",
        confidence: 0.88,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales!A5:F6"],
        overwriteRisk: "high",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: true
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.getRequiresConfirmation(response)).toBe(true);
    expect(sidebar.getStructuredPreview(response)).toMatchObject({
      kind: "composite_plan",
      stepCount: 2,
      dryRunRequired: true,
      steps: [
        { stepId: "step_sort", reversible: false },
        { stepId: "step_delete_rows", destructive: true, reversible: false }
      ]
    });

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_composite_preview_001",
      requestId: "req_composite_preview_001"
    });
    expect(html).toContain("Confirm Workflow");
    expect(html).toContain("dry run required");
    expect(html).toContain("step_delete_rows");

    expect(sidebar.buildWriteApprovalRequest({
      requestId: "req_composite_preview_001",
      runId: "run_composite_preview_001",
      workbookSessionKey: "google_sheets::sheet-123",
      plan: response.data
    })).toMatchObject({
      workbookSessionKey: "google_sheets::sheet-123",
      destructiveConfirmation: {
        confirmed: true
      }
    });
    expect(confirmMock).toHaveBeenCalledTimes(1);
  });

  it("keeps a demo-safe pivot-then-chart composite confirmable in the sidebar", () => {
    const sidebar = loadSidebarContext();

    const response = {
      type: "composite_plan",
      data: {
        steps: [
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
              explanation: "Build a pivot first.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "step_chart",
            dependsOn: ["step_pivot"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales Pivot",
              sourceRange: "A1:B2",
              targetSheet: "Sales Chart",
              targetRange: "A1",
              chartType: "line",
              categoryField: "Region",
              series: [{ field: "Revenue", label: "Revenue" }],
              explanation: "Chart the pivot output.",
              confidence: 0.88,
              requiresConfirmation: true,
              affectedRanges: ["Sales Pivot!A1:B2", "Sales Chart!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Build a pivot and then chart it.",
        confidence: 0.86,
        requiresConfirmation: true,
        affectedRanges: ["Sales Pivot!A1", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.getStructuredPreview(response)).toMatchObject({
      kind: "composite_plan",
      stepCount: 2,
      dryRunRequired: false,
      steps: [
        { stepId: "step_pivot", destructive: false, reversible: false },
        { stepId: "step_chart", destructive: false, reversible: false }
      ]
    });

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_composite_supported_001",
      requestId: "req_composite_supported_001"
    });
    expect(html).toContain("Confirm Workflow");
    expect(html).toContain("step_pivot");
    expect(html).toContain("step_chart");
    expect(html).not.toContain("does not support exact-safe pivot table creation yet");
  });

  it("flags unsupported composite steps before the user confirms the workflow", () => {
    const sidebar = loadSidebarContext();

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
              namedRangeName: "StatusOptions",
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

    expect(sidebar.isWritePlanResponse(response)).toBe(false);
    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_composite_unsupported_google_001",
      requestId: "req_composite_unsupported_google_001"
    });
    expect(html).toContain("Some workflow steps can't run in this Google Sheets runtime yet.");
    expect(html).toContain("This Google Sheets flow only supports exact year-first date text patterns");
    expect(html).not.toContain("Confirm Workflow");
  });

  it("flags unsupported filter child steps inside composite previews", () => {
    const sidebar = loadSidebarContext();

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

    expect(sidebar.isWritePlanResponse(response)).toBe(false);
    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_composite_filter_unsupported_google_001",
      requestId: "req_composite_filter_unsupported_google_001"
    });
    expect(html).toContain("Some workflow steps can't run in this Google Sheets runtime yet.");
    expect(html).toContain("can't combine multiple filter conditions with OR exactly");
    expect(html).not.toContain("Confirm Workflow");
  });

  it("routes dry-run, history, undo, and redo through the gateway using the workbook session key from Apps Script", async () => {
    const sidebar = loadSidebarContext();
    sidebar.__sidebarTestHooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://127.0.0.1:8787",
      clientVersion: "google-sheets-addon-dev",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    const localStorageBacking = new Map<string, string>();
    sidebar.window.localStorage.getItem = (key: string) => localStorageBacking.get(key) ?? null;
    sidebar.window.localStorage.setItem = (key: string, value: string) => {
      localStorageBacking.set(key, value);
    };
    localStorageBacking.set(
      "Hermes.ReversibleExecutions.v1::google_sheets::sheet-123",
      JSON.stringify({
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
      })
    );
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "validateExecutionCellSnapshot") {
        return payload;
      }

      if (functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });

    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        if (url.endsWith("/api/execution/dry-run")) {
          return {
            planDigest: "digest_001",
            workbookSessionKey: "google_sheets::sheet-123",
            simulated: true,
            predictedAffectedRanges: ["Sales!A1:F50"],
            predictedSummaries: ["Would sort and then delete stale rows."],
            overwriteRisk: "high",
            reversible: false,
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
            hostPlatform: "google_sheets",
            executionId: "exec_undo_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }

        if (url.endsWith("/api/execution/redo")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "google_sheets",
            executionId: "exec_redo_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }

        return {
          kind: "composite_update",
          operation: "composite_update",
          hostPlatform: "google_sheets",
          executionId: "exec_001",
          stepResults: [],
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.dryRunCompositePlan({
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

    await sidebar.listExecutionHistory({ limit: 5 });
    await sidebar.undoExecution("exec_001");
    await sidebar.redoExecution("exec_undo_001");

    expect(sidebar.fetch).toHaveBeenCalledTimes(4);
    expect(sidebar.fetch.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/dry-run");
    expect(JSON.parse(String(sidebar.fetch.mock.calls[0][1]?.body))).toMatchObject({
      workbookSessionKey: "google_sheets::sheet-123"
    });
    expect(sidebar.fetch.mock.calls[1][0]).toBe(
      "http://127.0.0.1:8787/api/execution/history?workbookSessionKey=google_sheets%3A%3Asheet-123&limit=5"
    );
    expect(sidebar.fetch.mock.calls[2][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(JSON.parse(String(sidebar.fetch.mock.calls[2][1]?.body))).toMatchObject({
      executionId: "exec_001",
      workbookSessionKey: "google_sheets::sheet-123"
    });
    expect(sidebar.fetch.mock.calls[3][0]).toBe("http://127.0.0.1:8787/api/execution/redo");
    expect(sidebar.callServer).toHaveBeenCalledWith("applyExecutionCellSnapshot", expect.objectContaining({
      targetSheet: "Sales",
      targetRange: "A1"
    }));
  });

  it("fails undo and redo before calling the gateway when the local snapshot validation fails", async () => {
    const sidebar = loadSidebarContext();

    sidebar.window.localStorage.getItem = vi.fn(() => JSON.stringify({
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
    }));
    sidebar.callServer = vi.fn(async (functionName: string) => {
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "validateExecutionCellSnapshot") {
        throw new Error("The saved undo snapshot no longer matches the current range shape.");
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn();

    await expect(sidebar.undoExecution("exec_001")).rejects.toThrow(
      "The saved undo snapshot no longer matches the current range shape."
    );
    await expect(sidebar.redoExecution("exec_undo_001")).rejects.toThrow(
      "The saved undo snapshot no longer matches the current range shape."
    );
    expect(sidebar.fetch).not.toHaveBeenCalled();
  });

  it("fails undo and redo before calling the gateway when the local snapshot store cannot persist redo lineage", async () => {
    const sidebar = loadSidebarContext();

    const backingStore = new Map<string, string>();
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
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
    }));
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) ?? null);
    sidebar.window.localStorage.setItem = vi.fn(() => {
      throw new Error("quota exceeded");
    });
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "validateExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn();

    await expect(sidebar.undoExecution("exec_001")).rejects.toThrow(
      "That history entry is no longer available for exact undo or redo in this sheet session."
    );
    await expect(sidebar.redoExecution("exec_undo_001")).rejects.toThrow(
      "That history entry is no longer available for exact undo or redo in this sheet session."
    );
    expect(sidebar.fetch).not.toHaveBeenCalled();
  });

  it("applies external data plans by anchoring a first-class formula into a single Google Sheets cell", () => {
    const targetRange = createRangeStub({
      a1Notation: "B2",
      row: 2,
      column: 2,
      numRows: 1,
      numColumns: 1
    });
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2");
        return targetRange;
      })
    };
    const spreadsheet = {
      getId() {
        return "sheet-123";
      },
      getSheetByName(name: string) {
        if (name === "Market Data") {
          return sheet;
        }
        return null;
      }
    };

    const code = loadCodeModule({ spreadsheet });
    const result = code.applyWritePlan({
      requestId: "req_external_data_apply_001",
      runId: "run_external_data_apply_001",
      approvalToken: "token",
      plan: {
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
    });

    expect(result).toMatchObject({
      kind: "external_data_update",
      hostPlatform: "google_sheets",
      sourceType: "market_data",
      provider: "googlefinance",
      targetSheet: "Market Data",
      targetRange: "B2",
      formula: '=GOOGLEFINANCE("CURRENCY:BTCUSD","price")'
    });
    expect(targetRange.setFormula).toHaveBeenCalledWith('=GOOGLEFINANCE("CURRENCY:BTCUSD","price")');
    expect(targetRange.getFormulas()).toEqual([['=GOOGLEFINANCE("CURRENCY:BTCUSD","price")']]);
    expect(code.flush).toHaveBeenCalled();
  });

  it("validates Google Sheets local execution snapshot shape before applying it", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 1
    });
    const spreadsheet = {
      getSheetByName(name: string) {
        if (name !== "Sales") {
          return null;
        }

        return {
          getRange(rangeA1: string) {
            expect(rangeA1).toBe("A1");
            return targetRange;
          }
        };
      }
    };
    const code = loadCodeModule({ spreadsheet });

    expect(() => code.validateExecutionCellSnapshot({
      targetSheet: "Sales",
      targetRange: "A1",
      cells: [[{ kind: "value", value: { type: "string", value: "before" } }]]
    })).toThrow("The saved undo snapshot no longer matches the current range shape.");
  });

  it("falls back to the Apps Script proxy when browser fetch throws before a response exists", async () => {
    const sidebar = loadSidebarContext();
    sidebar.__sidebarTestHooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://127.0.0.1:8787",
      clientVersion: "google-sheets-addon-dev",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };

    sidebar.fetch = vi.fn(async () => {
      throw new TypeError("Failed to fetch");
    });
    sidebar.callServer = vi.fn(async (functionName: string, payload: unknown) => {
      if (functionName !== "proxyGatewayJson") {
        throw new Error(`Unexpected server call: ${functionName}`);
      }

      return {
        proxied: true,
        path: (payload as { path: string }).path
      };
    });

    const response = await sidebar.callGatewayJson("/api/requests", {
      method: "post",
      body: { requestId: "req_proxy_001" }
    });

    expect(response).toEqual({
      proxied: true,
      path: "/api/requests"
    });
    expect(sidebar.fetch).toHaveBeenCalledTimes(1);
    expect(sidebar.callServer).toHaveBeenCalledWith("proxyGatewayJson", {
      path: "/api/requests",
      method: "post",
      headers: {},
      body: { requestId: "req_proxy_001" }
    });
  });

  it("preserves userAction guidance when Apps Script formats gateway errors", () => {
    const code = loadCodeModule();

    expect(code.extractGatewayErrorMessage(
      400,
      JSON.stringify({
        error: {
          code: "ATTACHMENT_UNAVAILABLE",
          message: "I can't access that uploaded file anymore.",
          userAction: "Reattach the file and tell me where to paste it."
        }
      })
    )).toBe(
      [
        "I can't access that uploaded file anymore.",
        "Reattach the file and tell me where to paste it."
      ].join("\n\n")
    );
  });

  it("translates raw 404 gateway text into an actionable sidebar error", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 404,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return "The requested resource doesn't exist.";
      }
    })).rejects.toThrow(
      "The Hermes request was sent to a service path that does not exist (https://gateway.test/api/requests, HTTP 404)."
    );
  });

  it("sanitizes raw requested-resource host errors in the Google Sheets sidebar", () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    expect(hooks.sanitizeHostExecutionError(
      new Error("The requested resource doesn't exist."),
      "Failed to contact Hermes."
    )).toBe(
      [
        "Hermes could not read the current sheet or selection.",
        "Select a normal sheet range, reload the sidebar, and retry. If it keeps happening, reopen the spreadsheet and try again."
      ].join("\n\n")
    );
  });

  it("returns a valid composite_update result from Code.gs when a demo-safe pivot step feeds a supported chart step", () => {
    const pivotSourceRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6,
      displayValues: [
        ["Region", "Rep", "Quarter", "Revenue", "Status", "Deals"]
      ]
    });
    const pivotOutputRange = createRangeStub({
      a1Notation: "A1:B2",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 2,
      displayValues: [
        ["Region", "Revenue"],
        ["West", "1200"]
      ]
    });
    const categoryRange = createRangeStub({
      a1Notation: "A1:A2",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 1
    });
    const revenueRange = createRangeStub({
      a1Notation: "B1:B2",
      row: 1,
      column: 2,
      numRows: 2,
      numColumns: 1
    });
    const pivotAnchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1
    });
    const chartAnchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1
    });
    const pivotTable = {
      addRowGroup: vi.fn(),
      addColumnGroup: vi.fn(),
      addPivotValue: vi.fn(),
      addFilter: vi.fn()
    };
    pivotAnchorRange.createPivotTable = vi.fn(() => pivotTable);
    const builtChart = { id: "chart-1" };
    const chartBuilder = {
      addRange: vi.fn().mockReturnThis(),
      setChartType: vi.fn().mockReturnThis(),
      setPosition: vi.fn().mockReturnThis(),
      setOption: vi.fn().mockReturnThis(),
      build: vi.fn(() => builtChart)
    };
    const chartSheet = {
      getRange: vi.fn(() => chartAnchorRange),
      newChart: vi.fn(() => chartBuilder),
      insertChart: vi.fn()
    };
    const spreadsheet = {
      getId() {
        return "sheet-123";
      },
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "Sales") {
          return {
            getRange: vi.fn((...args: unknown[]) => {
              if (args.length === 1 && args[0] === "A1:F50") {
                return pivotSourceRange;
              }
              return null;
            })
          };
        }
        if (sheetName === "Sales Pivot") {
          return {
            getRange: vi.fn((...args: unknown[]) => {
              if (args.length === 1 && args[0] === "A1") {
                return pivotAnchorRange;
              }
              if (args.length === 1 && typeof args[0] === "string" && args[0] !== "A1") {
                return pivotOutputRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 1 && args[2] === 2 && args[3] === 1) {
                return categoryRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 2 && args[2] === 2 && args[3] === 1) {
                return revenueRange;
              }
              return null;
            })
          };
        }
        if (sheetName === "Sales Chart") {
          return chartSheet;
        }
        return null;
      })
    };

    const code = loadCodeModule({ spreadsheet });

    expect(code.getWorkbookSessionKey()).toBe("google_sheets::sheet-123");
    expect(code.applyWritePlan({
      requestId: "req_composite_apply_001",
      runId: "run_composite_apply_001",
      approvalToken: "token",
      executionId: "exec_composite_apply_001",
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
              targetRange: "A1",
              rowGroups: ["Region"],
              valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
              explanation: "Build a pivot first.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "step_chart",
            dependsOn: ["step_pivot"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales Pivot",
              sourceRange: "A1:B2",
              targetSheet: "Sales Chart",
              targetRange: "A1",
              chartType: "line",
              categoryField: "Region",
              series: [{ field: "Revenue", label: "Revenue" }],
              explanation: "Chart the pivot output.",
              confidence: 0.88,
              requiresConfirmation: true,
              affectedRanges: ["Sales Pivot!A1:B2", "Sales Chart!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Build a pivot and then chart it.",
        confidence: 0.86,
        requiresConfirmation: true,
        affectedRanges: ["Sales Pivot!A1", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    })).toMatchObject({
      kind: "composite_update",
      operation: "composite_update",
      executionId: "exec_composite_apply_001",
      summary:
        "Workflow finished: 2 steps • 2 completed. " +
        "Completed: Created pivot table on Sales Pivot!A1; Created line chart on Sales Chart!A1.",
      stepResults: [
        {
          stepId: "step_pivot",
          status: "completed",
          summary: "Created pivot table on Sales Pivot!A1.",
          result: {
            kind: "pivot_table_update",
            operation: "pivot_table_update",
            hostPlatform: "google_sheets",
            sourceSheet: "Sales",
            sourceRange: "A1:F50",
            targetSheet: "Sales Pivot",
            targetRange: "A1",
            rowGroups: ["Region"],
            valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
            summary: "Created pivot table on Sales Pivot!A1."
          }
        },
        {
          stepId: "step_chart",
          status: "completed",
          summary: "Created line chart on Sales Chart!A1.",
          result: {
            kind: "chart_update",
            operation: "chart_update",
            hostPlatform: "google_sheets",
            sourceSheet: "Sales Pivot",
            sourceRange: "A1:B2",
            targetSheet: "Sales Chart",
            targetRange: "A1",
            chartType: "line",
            categoryField: "Region",
            series: [{ field: "Revenue", label: "Revenue" }],
            summary: "Created line chart on Sales Chart!A1."
          }
        }
      ]
    });

    expect(pivotAnchorRange.createPivotTable).toHaveBeenCalledWith(pivotSourceRange);
    expect(pivotTable.addPivotValue).toHaveBeenCalledWith(4, "SUM");
    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(1, categoryRange);
    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(2, revenueRange);
    expect(chartBuilder.setChartType).toHaveBeenCalledWith("LINE");
    expect(chartSheet.insertChart).toHaveBeenCalledWith(builtChart);
  });

  it("fails closed when a composite step appears before an unsatisfied dependency in Code.gs execution order", () => {
    const code = loadCodeModule({
      spreadsheet: {
        getId() {
          return "sheet-123";
        }
      }
    });

    expect(code.applyWritePlan({
      requestId: "req_composite_order_001",
      runId: "run_composite_order_001",
      approvalToken: "token",
      executionId: "exec_composite_order_001",
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
      }
    })).toMatchObject({
      kind: "composite_update",
      operation: "composite_update",
      executionId: "exec_composite_order_001",
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

  it("executes the latest pending write plan when the user types an explicit confirmation instead of sending a new Hermes request", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    hooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://gateway.test",
      clientVersion: "google-sheets-addon-dev",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    hooks.state.messages = [
      {
        role: "assistant",
        runId: "run_confirm_sheet_001",
        requestId: "req_confirm_sheet_001",
        response: {
          type: "workbook_structure_update",
          data: {
            operation: "create_sheet",
            sheetName: "dieu",
            explanation: "Create a new sheet named dieu.",
            confidence: 0.99,
            requiresConfirmation: true
          }
        },
        content: "Prepared a workbook update to create sheet dieu.",
        statusLine: ""
      }
    ];
    hooks.elements.prompt.value = 'Confirm create sheet "dieu"';

    const fetchMock = vi.fn(async (url: string, options: { body?: string }) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return {
          ok: true,
          json: async () => ({
            approvalToken: "approval-token",
            planDigest: "plan-digest"
          })
        };
      }

      if (url === "http://gateway.test/api/writeback/complete") {
        return {
          ok: true,
          json: async () => ({ ok: true })
        };
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });
    sidebar.fetch = fetchMock;

    let successHandler: ((value: unknown) => unknown) | null = null;
    let failureHandler: ((error: unknown) => unknown) | null = null;
    const getSpreadsheetSnapshotSpy = vi.fn(() => {
      throw new Error("sendPrompt should not request a fresh snapshot for typed confirmation.");
    });
    const applyWritePlanSpy = vi.fn((payload: Record<string, unknown>) => {
      successHandler?.({
        kind: "workbook_structure_update",
        hostPlatform: "google_sheets",
        sheetName: payload.plan && (payload.plan as any).sheetName,
        operation: payload.plan && (payload.plan as any).operation,
        summary: `Created sheet ${(payload.plan as any).sheetName}.`
      });
    });
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler(handler: (error: unknown) => unknown) {
        failureHandler = handler;
        void failureHandler;
        return this;
      },
      applyWritePlan: applyWritePlanSpy,
      getWorkbookSessionKey() {
        successHandler?.("google_sheets::sheet-123");
      },
      getSpreadsheetSnapshot: getSpreadsheetSnapshotSpy
    };

    await hooks.sendPrompt();

    expect(getSpreadsheetSnapshotSpy).not.toHaveBeenCalled();
    expect(fetchMock).toHaveBeenCalledTimes(2);
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      "http://gateway.test/api/writeback/approve",
      "http://gateway.test/api/writeback/complete"
    ]);
    expect(JSON.parse(String(fetchMock.mock.calls[0][1]?.body))).toMatchObject({
      requestId: "req_confirm_sheet_001",
      runId: "run_confirm_sheet_001",
      workbookSessionKey: "google_sheets::sheet-123"
    });
    expect(applyWritePlanSpy).toHaveBeenCalledWith(expect.objectContaining({
      requestId: "req_confirm_sheet_001",
      runId: "run_confirm_sheet_001",
      approvalToken: "approval-token",
      plan: expect.objectContaining({
        operation: "create_sheet",
        sheetName: "dieu"
      })
    }));
    expect(hooks.state.messages.map((message: { role: string; content: string }) => ({
      role: message.role,
      content: message.content
    }))).toContainEqual({
      role: "user",
      content: 'Confirm create sheet "dieu"'
    });
    expect(hooks.state.messages[0]).toMatchObject({
      content: "Created sheet dieu.",
      response: null,
      statusLine: ""
    });
    expect(hooks.elements.prompt.value).toBe("");
  });

  it("disambiguates typed workbook confirmations by sheet name instead of always executing the newest pending plan", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    hooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://gateway.test",
      clientVersion: "google-sheets-addon-dev",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    hooks.state.messages = [
      {
        role: "assistant",
        runId: "run_delete_foo_001",
        requestId: "req_delete_foo_001",
        response: {
          type: "workbook_structure_update",
          data: {
            operation: "delete_sheet",
            sheetName: "Foo",
            explanation: "Delete sheet Foo.",
            confidence: 0.99,
            requiresConfirmation: true
          }
        },
        content: "Prepared a workbook update to delete sheet Foo.",
        statusLine: ""
      },
      {
        role: "assistant",
        runId: "run_delete_bar_001",
        requestId: "req_delete_bar_001",
        response: {
          type: "workbook_structure_update",
          data: {
            operation: "delete_sheet",
            sheetName: "Bar",
            explanation: "Delete sheet Bar.",
            confidence: 0.99,
            requiresConfirmation: true
          }
        },
        content: "Prepared a workbook update to delete sheet Bar.",
        statusLine: ""
      }
    ];
    hooks.elements.prompt.value = "Confirm delete sheet Foo";

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return {
          ok: true,
          json: async () => ({
            approvalToken: "approval-token",
            planDigest: "plan-digest"
          })
        };
      }

      if (url === "http://gateway.test/api/writeback/complete") {
        return {
          ok: true,
          json: async () => ({ ok: true })
        };
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });
    sidebar.fetch = fetchMock;

    let successHandler: ((value: unknown) => unknown) | null = null;
    const applyWritePlanSpy = vi.fn((payload: Record<string, unknown>) => {
      successHandler?.({
        kind: "workbook_structure_update",
        hostPlatform: "google_sheets",
        sheetName: payload.plan && (payload.plan as any).sheetName,
        operation: payload.plan && (payload.plan as any).operation,
        summary: `Deleted sheet ${(payload.plan as any).sheetName}.`
      });
    });
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler() {
        return this;
      },
      applyWritePlan: applyWritePlanSpy,
      getWorkbookSessionKey() {
        successHandler?.("google_sheets::sheet-123");
      },
      getSpreadsheetSnapshot: vi.fn(() => {
        throw new Error("sendPrompt should not request a fresh snapshot for typed confirmation.");
      })
    };

    await hooks.sendPrompt();

    expect(applyWritePlanSpy).toHaveBeenCalledWith(expect.objectContaining({
      requestId: "req_delete_foo_001",
      runId: "run_delete_foo_001",
      plan: expect.objectContaining({
        operation: "delete_sheet",
        sheetName: "Foo"
      })
    }));
    expect(applyWritePlanSpy).not.toHaveBeenCalledWith(expect.objectContaining({
      requestId: "req_delete_bar_001",
      runId: "run_delete_bar_001",
      plan: expect.objectContaining({
        operation: "delete_sheet",
        sheetName: "Bar"
      })
    }));
  });

  it("retries Google Sheets writeback completion without reapplying the local change", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    hooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://gateway.test",
      clientVersion: "google-sheets-addon-dev",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    hooks.state.messages = [
      {
        role: "assistant",
        runId: "run_confirm_retry_001",
        requestId: "req_confirm_retry_001",
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
        statusLine: ""
      }
    ];

    let completeAttempts = 0;
    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return {
          ok: true,
          json: async () => ({
            approvalToken: "approval-token",
            planDigest: "plan-digest",
            executionId: "exec_retry_001"
          })
        };
      }

      if (url === "http://gateway.test/api/writeback/complete") {
        completeAttempts += 1;
        if (completeAttempts === 1) {
          return {
            ok: false,
            status: 500,
            text: async () => JSON.stringify({
              error: {
                message: "Temporary completion failure."
              }
            })
          };
        }

        return {
          ok: true,
          json: async () => ({ ok: true })
        };
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });
    sidebar.fetch = fetchMock;

    let successHandler: ((value: unknown) => unknown) | null = null;
    const applyWritePlanSpy = vi.fn((payload: Record<string, unknown>) => {
      successHandler?.({
        kind: "workbook_structure_update",
        hostPlatform: "google_sheets",
        sheetName: payload.plan && (payload.plan as any).sheetName,
        operation: payload.plan && (payload.plan as any).operation,
        summary: `Created sheet ${(payload.plan as any).sheetName}.`
      });
    });
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler() {
        return this;
      },
      applyWritePlan: applyWritePlanSpy,
      getWorkbookSessionKey() {
        successHandler?.("google_sheets::sheet-123");
      },
      getSpreadsheetSnapshot: vi.fn(() => {
        throw new Error("sendPrompt should not request a fresh snapshot for typed confirmation.");
      })
    };

    hooks.elements.prompt.value = 'Confirm create sheet "Retry Demo"';
    await hooks.sendPrompt();

    expect(applyWritePlanSpy).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      "http://gateway.test/api/writeback/approve",
      "http://gateway.test/api/writeback/complete"
    ]);
    expect(hooks.state.messages[0].statusLine).toBe(
      "Applied locally. Retry confirm to finish syncing Hermes history."
    );
    expect(hooks.state.messages[0].pendingCompletion).toMatchObject({
      requestId: "req_confirm_retry_001",
      runId: "run_confirm_retry_001",
      workbookSessionKey: "google_sheets::sheet-123",
      approvalToken: "approval-token",
      planDigest: "plan-digest"
    });

    hooks.elements.prompt.value = 'Confirm create sheet "Retry Demo"';
    await hooks.sendPrompt();

    expect(applyWritePlanSpy).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      "http://gateway.test/api/writeback/approve",
      "http://gateway.test/api/writeback/complete",
      "http://gateway.test/api/writeback/complete"
    ]);
    expect(hooks.state.messages[0]).toMatchObject({
      content: "Created sheet Retry Demo.",
      response: null,
      statusLine: ""
    });
    expect(hooks.state.messages[0].pendingCompletion).toBeUndefined();
  });

  it("disambiguates typed range confirmations by quoted sheet-range target instead of always executing the newest pending plan", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    hooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://gateway.test",
      clientVersion: "google-sheets-addon-dev",
      reviewerSafeMode: false,
      forceExtractionMode: null
    };
    hooks.state.messages = [
      {
        role: "assistant",
        runId: "run_update_sales_001",
        requestId: "req_update_sales_001",
        response: {
          type: "sheet_update",
          data: {
            targetSheet: "Sales",
            targetRange: "B2:C3",
            operation: "replace_range",
            values: [["North", 12], ["South", 20]],
            explanation: "Replace Sales!B2:C3.",
            confidence: 0.98,
            requiresConfirmation: true,
            shape: {
              rows: 2,
              columns: 2
            }
          }
        },
        content: "Prepared an update for Sales!B2:C3.",
        statusLine: ""
      },
      {
        role: "assistant",
        runId: "run_update_summary_001",
        requestId: "req_update_summary_001",
        response: {
          type: "sheet_update",
          data: {
            targetSheet: "Summary",
            targetRange: "D4:E5",
            operation: "replace_range",
            values: [["East", 30], ["West", 44]],
            explanation: "Replace Summary!D4:E5.",
            confidence: 0.98,
            requiresConfirmation: true,
            shape: {
              rows: 2,
              columns: 2
            }
          }
        },
        content: "Prepared an update for Summary!D4:E5.",
        statusLine: ""
      }
    ];
    hooks.elements.prompt.value = 'Confirm "Sales!B2:C3"';

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return {
          ok: true,
          json: async () => ({
            approvalToken: "approval-token",
            planDigest: "plan-digest"
          })
        };
      }

      if (url === "http://gateway.test/api/writeback/complete") {
        return {
          ok: true,
          json: async () => ({ ok: true })
        };
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });
    sidebar.fetch = fetchMock;

    let successHandler: ((value: unknown) => unknown) | null = null;
    const applyWritePlanSpy = vi.fn((payload: Record<string, unknown>) => {
      successHandler?.({
        kind: "range_write",
        hostPlatform: "google_sheets",
        ...((payload.plan as any) || {}),
        writtenRows: 2,
        writtenColumns: 2
      });
    });
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler() {
        return this;
      },
      applyWritePlan: applyWritePlanSpy,
      getWorkbookSessionKey() {
        successHandler?.("google_sheets::sheet-123");
      },
      getSpreadsheetSnapshot: vi.fn(() => {
        throw new Error("sendPrompt should not request a fresh snapshot for typed confirmation.");
      })
    };

    await hooks.sendPrompt();

    expect(applyWritePlanSpy).toHaveBeenCalledWith(expect.objectContaining({
      requestId: "req_update_sales_001",
      runId: "run_update_sales_001",
      plan: expect.objectContaining({
        targetSheet: "Sales",
        targetRange: "B2:C3"
      })
    }));
    expect(applyWritePlanSpy).not.toHaveBeenCalledWith(expect.objectContaining({
      requestId: "req_update_summary_001",
      runId: "run_update_summary_001",
      plan: expect.objectContaining({
        targetSheet: "Summary",
        targetRange: "D4:E5"
      })
    }));
  });

  it("does not treat unsupported Google Sheets previews as typed-confirmable write plans", () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    hooks.state.messages = [
      {
        role: "assistant",
        runId: "run_validation_unsupported_001",
        requestId: "req_validation_unsupported_001",
        response: {
          type: "data_validation_plan",
          data: {
            targetSheet: "Sheet1",
            targetRange: "B2:B20",
            ruleType: "list",
            namedRangeName: "StatusOptions",
            showDropdown: true,
            allowBlank: true,
            invalidDataBehavior: "warn",
            explanation: "Unsupported validation plan for Google Sheets.",
            confidence: 0.5,
            requiresConfirmation: true
          }
        },
        content: "Prepared a validation plan for Sheet1!B2:B20.",
        statusLine: ""
      },
      {
        role: "assistant",
        runId: "run_cleanup_unsupported_001",
        requestId: "req_cleanup_unsupported_001",
        response: {
          type: "data_cleanup_plan",
          data: {
            targetSheet: "Contacts",
            targetRange: "A2:B20",
            operation: "standardize_format",
            formatType: "date_text",
            formatPattern: "locale-sensitive-fuzzy",
            explanation: "Unsupported fuzzy cleanup plan for Google Sheets.",
            confidence: 0.5,
            requiresConfirmation: true
          }
        },
        content: "Prepared a cleanup plan for Contacts!A2:B20.",
        statusLine: ""
      }
    ];

    expect(sidebar.isWritePlanResponse(hooks.state.messages[0].response)).toBe(false);
    expect(sidebar.isWritePlanResponse(hooks.state.messages[1].response)).toBe(false);
    expect(sidebar.getLatestPendingWritePlanMessage("ok")).toBeNull();
    expect(sidebar.getLatestPendingWritePlanMessage('Confirm "Sheet1!B2:B20"')).toBeNull();
    expect(sidebar.getLatestPendingWritePlanMessage('Confirm "Contacts!A2:B20"')).toBeNull();
  });
});
