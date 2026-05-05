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
  spreadsheetAppOverrides?: Record<string, unknown>;
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
        let rawCriteriaType: unknown = null;
        let rawCriteriaValues: unknown[] = [];
        let hiddenValues: string[] | null = null;
        let visibleValues: string[] | null = null;
        return {
          whenTextEqualTo: vi.fn(function() {
            matchedValue = arguments[0] == null ? null : String(arguments[0]);
            return this;
          }),
          withCriteria: vi.fn(function(type: unknown, values: unknown[]) {
            rawCriteriaType = type;
            rawCriteriaValues = Array.isArray(values) ? values : [];
            return this;
          }),
          setHiddenValues: vi.fn(function(values: string[]) {
            hiddenValues = Array.isArray(values) ? values : [];
            return this;
          }),
          setVisibleValues: vi.fn(function(values: string[]) {
            visibleValues = Array.isArray(values) ? values : [];
            return this;
          }),
          build: vi.fn(function() {
            if (hiddenValues) {
              return { type: "hidden_values", values: hiddenValues };
            }
            if (visibleValues) {
              return { type: "visible_values", values: visibleValues };
            }
            if (rawCriteriaType) {
              return { type: rawCriteriaType, values: rawCriteriaValues };
            }
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
      },
      DataValidationCriteria: {
        VALUE_IN_LIST: "VALUE_IN_LIST",
        NUMBER_BETWEEN: "NUMBER_BETWEEN"
      },
      newDataValidation() {
        let criteriaType: unknown = null;
        let criteriaValues: unknown[] = [];
        let allowInvalid: boolean | undefined;
        let helpText: string | null | undefined;
        return {
          withCriteria(type: unknown, values: unknown[]) {
            criteriaType = type;
            criteriaValues = Array.isArray(values) ? values : [];
            return this;
          },
          setAllowInvalid(value: boolean) {
            allowInvalid = value;
            return this;
          },
          setHelpText(value: string | null) {
            helpText = value;
            return this;
          },
          build() {
            return {
              criteriaType,
              criteriaValues,
              allowInvalid,
              helpText
            };
          }
        };
      },
      ...(options.spreadsheetAppOverrides || {})
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
    proxyGatewayJson: context.proxyGatewayJson,
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
    "this.__sidebarTestHooks = { state, elements, renderMessages, sendPrompt, pollRun, sanitizeConversation, sanitizeHostExecutionError, getResponseMetaLine, buildRequestEnvelope, pruneStoredMessages, trimMessageTraceEvents, ensureRuntimeConfig, parseGatewayJsonResponse, initialize };",
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
    expect(code.sanitizeHostExecutionError(
      new Error("Google Sheets host requires append targetRange to match the full destination rectangle.")
    )).toBe(expectedAppendMessage);
    expect(sidebar.sanitizeHostExecutionError(
      new Error("Google Sheets host requires append targetRange to match the full destination rectangle.")
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

    expect(sidebar.sanitizeHostExecutionError(
      new Error("Unhandled failure at /srv/hermes/apps/google-sheets-addon/src/Code.gs:42 refresh_token=refresh_123"),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error("Unhandled failure at path=/srv/hermes/apps/google-sheets-addon/src/Code.gs:42"),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at path=C:\Users\runner\work\Sidebar.js.html:42`),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at \\runner\share\Sidebar.js.html:42`),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at source=\\runner\share\Sidebar.js.html:42`),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at ("C:\Users\runner\work\Sidebar.js.html:42")`),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at ("\\runner\share\Sidebar.js.html:42")`),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(code.sanitizeHostExecutionError(
      new Error("Request failed at https://internal.example/api with HERMES_API_SERVER_KEY=secret_123")
    )).toBe("Write-back failed.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error("Request failed at http://[::ffff:7f00:1]/latest/meta-data"),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(code.sanitizeHostExecutionError(
      new Error("Request failed at http://[::ffff:7f00:1]/latest/meta-data")
    )).toBe("Write-back failed.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error("Request failed at http://2130706433/latest/meta-data"),
      "Failed to contact Hermes."
    )).toBe("Failed to contact Hermes.");

    expect(code.sanitizeHostExecutionError(
      new Error("Request failed at http://2130706433/latest/meta-data")
    )).toBe("Write-back failed.");

    expect(code.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at path=C:\Users\runner\work\Code.gs:42`)
    )).toBe("Write-back failed.");

    expect(code.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at \\runner\share\Code.gs:42`)
    )).toBe("Write-back failed.");

    expect(code.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at source=\\runner\share\Code.gs:42`)
    )).toBe("Write-back failed.");

    expect(code.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at ("C:\Users\runner\work\Code.gs:42")`)
    )).toBe("Write-back failed.");

    expect(code.sanitizeHostExecutionError(
      new Error(String.raw`Unhandled writeback failure at ("\\runner\share\Code.gs:42")`)
    )).toBe("Write-back failed.");

    expect(sidebar.sanitizeHostExecutionError(
      new Error("Writeback failed for qa_HERMES_API_SERVER_KEY")
    )).toBe("Write-back failed.");
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

  it("fails fast when Google Sheets image upload is pointed at private gateway host aliases", () => {
    for (const gatewayUrl of [
      "https://localhost.:8787",
      "https://app.localhost:8787",
      "https://gateway.local:8787",
      "https://gateway.internal:8787",
      "https://127.0.0.2:8787",
      "https://169.254.1.10:8787",
      "https://2130706433:8787",
      "https://0177.0.0.1:8787",
      "https://0x7f.0.0.1:8787",
      "https://2852039166:8787"
    ]) {
      const fetch = vi.fn();
      const code = loadCodeModule({
        scriptProperties: {
          HERMES_GATEWAY_URL: gatewayUrl
        },
        urlFetchApp: { fetch }
      });

      expect(() => code.uploadGatewayImageAttachment({
        fileName: "table.png",
        mimeType: "image/png",
        base64Data: "aGVsbG8=",
        source: "upload",
        sessionId: "sess_upload_private_alias_001"
      })).toThrow(
        "Google Sheets image upload requires a public Hermes gateway URL.\n\n" +
        "Set HERMES_GATEWAY_URL to a reachable HTTPS or public address, then retry the upload."
      );
      expect(fetch).not.toHaveBeenCalled();
    }
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

  it("does not expose private gateway aliases through Google Sheets runtime config", () => {
    for (const gatewayUrl of [
      "https://localhost.:8787",
      "https://app.localhost:8787",
      "https://gateway.local:8787",
      "https://gateway.internal:8787",
      "https://127.0.0.2:8787",
      "https://169.254.1.10:8787"
    ]) {
      const code = loadCodeModule({
        scriptProperties: {
          HERMES_GATEWAY_URL: gatewayUrl
        }
      });

      expect(code.getRuntimeConfig()).toMatchObject({
        gatewayBaseUrl: ""
      });
    }
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
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("sessionId=sess_");

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

        if (url.endsWith("/api/execution/undo/prepare")) {
          expect(JSON.parse(String(init?.body))).toMatchObject({
            executionId: "exec_001",
            workbookSessionKey
          });
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "google_sheets",
            executionId: "exec_undo_preview_001",
            stepResults: [],
            summary: "Prepared Sheet8!A1 undo."
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

    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(String(fetchMock.mock.calls[0][0])).toContain(
      "workbookSessionKey=google_sheets%3A%3Asheet-123"
    );
    expect(String(fetchMock.mock.calls[0][0])).toContain("sessionId=sess_");
    expect(String(fetchMock.mock.calls[0][0])).toContain("limit=20");
    expect(fetchMock.mock.calls[1][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(fetchMock.mock.calls[2][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
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

  it("redacts unsafe proof metadata in Google Sheets meta lines", () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    const metaLine = hooks.getResponseMetaLine({
      type: "chat",
      skillsUsed: [
        "SelectionExplainerSkill",
        "/srv/hermes/private-tool.ts",
        "C:\\Users\\runner\\work\\private-tool.ts",
        "\\\\server\\share\\private-tool.ts",
        "https://169.254.169.254/latest/meta-data",
        "https://[fd00::1]/latest/meta-data",
        "https://[::ffff:7f00:1]/latest/meta-data",
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
    expect(metaLine).not.toContain("C:\\Users");
    expect(metaLine).not.toContain("\\\\server");
    expect(metaLine).not.toContain("169.254.169.254");
    expect(metaLine).not.toContain("fd00");
    expect(metaLine).not.toContain("::ffff");
    expect(metaLine).not.toContain("7f00");
    expect(metaLine).not.toContain("internal.example");
    expect(metaLine).not.toContain("provider https://internal");

    const embeddedMetaLine = hooks.getResponseMetaLine({
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

    const wrappedMetaLine = hooks.getResponseMetaLine({
      type: "chat",
      skillsUsed: [
        "SelectionExplainerSkill",
        String.raw`("/srv/hermes/private-tool.ts")`,
        String.raw`("C:\Users\runner\work\private-tool.ts")`,
        String.raw`("\\server\share\private-tool.ts")`
      ],
      downstreamProvider: {
        label: String.raw`("\\server\share\provider")`,
        model: "gpt-5"
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

    expect(wrappedMetaLine).toContain("skills SelectionExplainerSkill");
    expect(wrappedMetaLine).not.toContain("/srv/hermes");
    expect(wrappedMetaLine).not.toContain("C:\\Users");
    expect(wrappedMetaLine).not.toContain("\\\\server");
    expect(wrappedMetaLine).not.toContain("private-tool");
    expect(wrappedMetaLine).not.toContain("provider");
  });

  it("escapes quotes in Google Sheets preview action attributes", () => {
    const sidebar = loadSidebarContext();

    const html = sidebar.renderStructuredPreview({
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
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("sessionId=sess_");
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).not.toContain("/api/trace/");
    expect(message.content).toBe("Done.");
    expect(message.statusLine).not.toContain("Request failed");
    expect(message.tracePollingDisabled).toBe(true);
    expect(message.response.type).toBe("chat");
  });

  it("encodes Google Sheets run identifiers in request polling paths", async () => {
    vi.useFakeTimers();
    const sidebar = loadSidebarContext();
    sidebar.fetch = vi.fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({
        runId: "run/../unsafe?x=1",
        requestId: "req_poll_path_001",
        status: "processing"
      }), {
        status: 200,
        headers: { "content-type": "application/json" }
      }));

    const message = {
      runId: "run/../unsafe?x=1",
      requestId: "req_poll_path_001",
      traceIndex: 0,
      trace: [],
      statusLine: "Thinking...",
      content: "Thinking..."
    };

    void sidebar.__sidebarTestHooks.pollRun(message);

    await vi.advanceTimersByTimeAsync(900);

    expect(sidebar.fetch).toHaveBeenCalledTimes(1);
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain(
      "/api/requests/run%2F..%2Funsafe%3Fx%3D1?"
    );
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
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("sessionId=sess_");
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
    expect(String(sidebar.fetch.mock.calls[0]?.[0] || "")).toContain("sessionId=sess_");
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

  it("fails closed when a Google Sheets market data formula does not match the query", () => {
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
        formula: '=GOOGLEFINANCE("CURRENCY:ETHUSD","price")',
        explanation: "Anchor the latest BTC price in B2.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Market Data!B2"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(false);

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_external_data_preview_market_query_mismatch",
      requestId: "req_external_data_preview_market_query_mismatch"
    });

    expect(html).toContain("Google Sheets market data formula must match query.symbol.");
    expect(html).not.toContain("Confirm External Data");
  });

  it("fails closed when a Google Sheets market data formula uses single-quoted query arguments", () => {
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
        formula: "=GOOGLEFINANCE('CURRENCY:BTCUSD','price')",
        explanation: "Anchor the latest BTC price in B2.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Market Data!B2"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(false);

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_external_data_preview_market_single_quotes",
      requestId: "req_external_data_preview_market_single_quotes"
    });

    expect(html).toContain("Google Sheets market data formula must match query.symbol.");
    expect(html).not.toContain("Confirm External Data");
  });

  it("fails closed when a Google Sheets external data formula references a different source URL", () => {
    const sidebar = loadSidebarContext();

    const response = {
      type: "external_data_plan",
      data: {
        sourceType: "web_table_import",
        provider: "importdata",
        sourceUrl: "https://example.com/data.csv",
        selectorType: "direct",
        targetSheet: "Imports",
        targetRange: "A1",
        formula: '=IMPORTDATA("https://other.example/data.csv")',
        explanation: "Import a public CSV into the sheet.",
        confidence: 0.86,
        requiresConfirmation: true,
        affectedRanges: ["Imports!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(false);

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_external_data_preview_source_mismatch",
      requestId: "req_external_data_preview_source_mismatch"
    });

    expect(html).toContain("Google Sheets external data formulas must reference sourceUrl.");
    expect(html).not.toContain("Confirm External Data");
  });

  it("fails closed when Google Sheets external data source URLs use private aliases", () => {
    const sidebar = loadSidebarContext();
    const sourceUrls = [
      "http://[fc00::1]/data.csv",
      "http://[fe80::1]/data.csv",
      "http://[::ffff:127.0.0.1]/data.csv",
      "http://2130706433/data.csv",
      "http://0177.0.0.1/data.csv",
      "http://0x7f.0.0.1/data.csv",
      "http://2852039166/data.csv"
    ];

    sourceUrls.forEach((sourceUrl, index) => {
      const response = {
        type: "external_data_plan",
        data: {
          sourceType: "web_table_import",
          provider: "importdata",
          sourceUrl,
          selectorType: "direct",
          targetSheet: "Imports",
          targetRange: "A1",
          formula: `=IMPORTDATA("${sourceUrl}")`,
          explanation: "Import a public CSV into the sheet.",
          confidence: 0.86,
          requiresConfirmation: true,
          affectedRanges: ["Imports!A1"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      };

      expect(sidebar.isWritePlanResponse(response)).toBe(false);

      const html = sidebar.renderStructuredPreview(response, {
        runId: `run_external_data_preview_private_alias_${index}`,
        requestId: `req_external_data_preview_private_alias_${index}`
      });

      expect(html).toContain("Choose a public HTTP(S) source URL for the web import first.");
      expect(html).not.toContain("Confirm External Data");
    });
  });

  it("renders exact-safe Google Sheets table-like previews and rejects unsupported totals rows", () => {
    const sidebar = loadSidebarContext();
    const response = {
      type: "table_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        showBandedRows: true,
        showBandedColumns: false,
        showFilterButton: true,
        showTotalsRow: false,
        explanation: "Format the sales range as a filterable table.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };
    const unsupportedTotals = {
      type: "table_plan",
      data: {
        ...response.data,
        showTotalsRow: true
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.renderStructuredPreview(response, {
      runId: "run_table_preview_google_001",
      requestId: "req_table_preview_google_001"
    })).toContain("Confirm Table");
    expect(sidebar.isWritePlanResponse(unsupportedTotals)).toBe(false);
    expect(sidebar.renderStructuredPreview(unsupportedTotals, {
      runId: "run_table_preview_google_unsupported_001",
      requestId: "req_table_preview_google_unsupported_001"
    })).toContain("Google Sheets can format this range as a table-like range, but it cannot create an exact native totals row here.");
  });

  it("preserves contract metadata across Google Sheets structured previews", () => {
    const sidebar = loadSidebarContext();

    const shared = {
      explanation: "Preview the planned spreadsheet change.",
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
      expect(sidebar.getStructuredPreview(testCase.response), testCase.name).toMatchObject(testCase.expected);
    }
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
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.getStructuredPreview(response)).toMatchObject({
      kind: "composite_plan",
      stepCount: 2,
      dryRunRequired: false,
      reversible: true,
      steps: [
        { stepId: "step_pivot", destructive: false, reversible: true },
        { stepId: "step_chart", destructive: false, reversible: true }
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

  it("keeps snapshot-backed analytic child steps non-reversible when the Google Sheets composite is non-reversible", () => {
    const sidebar = loadSidebarContext();

    const preview = sidebar.getStructuredPreview({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_pivot",
            dependsOn: [],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:C20",
              targetSheet: "Sales Pivot",
              targetRange: "A1",
              rowGroups: ["Region"],
              valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
              explanation: "Create a pivot table.",
              confidence: 0.92,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:C20", "Sales Pivot!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Create a pivot table.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });

    expect(preview).toMatchObject({
      kind: "composite_plan",
      reversible: false,
      steps: [
        { stepId: "step_pivot", reversible: false }
      ]
    });
  });

  it("marks snapshot-backed table child steps as reversible in Google Sheets composite previews", () => {
    const sidebar = loadSidebarContext();

    const preview = sidebar.getStructuredPreview({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_table",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:F20",
              name: "SalesTable",
              hasHeaders: true,
              showBandedRows: true,
              showBandedColumns: false,
              showFilterButton: true,
              showTotalsRow: false,
              explanation: "Format the sales range as a table.",
              confidence: 0.92,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F20"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Format the sales range as a table.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F20"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });

    expect(preview).toMatchObject({
      kind: "composite_plan",
      reversible: true,
      steps: [
        { stepId: "step_table", reversible: true }
      ]
    });
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

        if (url.endsWith("/api/execution/undo/prepare")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "google_sheets",
            executionId: "exec_undo_preview_001",
            stepResults: [],
            summary: init?.body || ""
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

        if (url.endsWith("/api/execution/redo/prepare")) {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "google_sheets",
            executionId: "exec_redo_preview_001",
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

    expect(sidebar.fetch).toHaveBeenCalledTimes(6);
    expect(sidebar.fetch.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/dry-run");
    expect(JSON.parse(String(sidebar.fetch.mock.calls[0][1]?.body))).toMatchObject({
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey: "google_sheets::sheet-123"
    });
    expect(String(sidebar.fetch.mock.calls[1][0])).toContain(
      "workbookSessionKey=google_sheets%3A%3Asheet-123"
    );
    expect(String(sidebar.fetch.mock.calls[1][0])).toContain("sessionId=sess_");
    expect(String(sidebar.fetch.mock.calls[1][0])).toContain("limit=5");
    expect(sidebar.fetch.mock.calls[2][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(JSON.parse(String(sidebar.fetch.mock.calls[2][1]?.body))).toMatchObject({
      executionId: "exec_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey: "google_sheets::sheet-123"
    });
    expect(sidebar.fetch.mock.calls[3][0]).toBe("http://127.0.0.1:8787/api/execution/undo");
    expect(JSON.parse(String(sidebar.fetch.mock.calls[3][1]?.body))).toMatchObject({
      executionId: "exec_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey: "google_sheets::sheet-123"
    });
    expect(sidebar.fetch.mock.calls[4][0]).toBe("http://127.0.0.1:8787/api/execution/redo/prepare");
    expect(JSON.parse(String(sidebar.fetch.mock.calls[4][1]?.body))).toMatchObject({
      executionId: "exec_undo_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey: "google_sheets::sheet-123"
    });
    expect(sidebar.fetch.mock.calls[5][0]).toBe("http://127.0.0.1:8787/api/execution/redo");
    expect(JSON.parse(String(sidebar.fetch.mock.calls[5][1]?.body))).toMatchObject({
      executionId: "exec_undo_001",
      sessionId: expect.stringMatching(/^sess_/),
      workbookSessionKey: "google_sheets::sheet-123"
    });
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

  it("prepares undo and redo but does not commit the gateway when local snapshot apply fails", async () => {
    function loadFailingApplySidebar() {
      const sidebar = loadSidebarContext();
      sidebar.__sidebarTestHooks.state.runtimeConfig = {
        gatewayBaseUrl: "http://127.0.0.1:8787",
        clientVersion: "google-sheets-addon-dev",
        reviewerSafeMode: false,
        forceExtractionMode: null
      };
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
      sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
        backingStore.set(key, value);
      });
      sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
        if (functionName === "getWorkbookSessionKey") {
          return "google_sheets::sheet-123";
        }

        if (functionName === "validateExecutionCellSnapshot") {
          return payload;
        }

        if (functionName === "applyExecutionCellSnapshot") {
          throw new Error("Apps Script flush failed.");
        }

        throw new Error(`Unexpected server call: ${functionName}`);
      });
      sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
        ok: true,
        async json() {
          return {
            kind: "composite_update",
            operation: "composite_update",
            hostPlatform: "google_sheets",
            executionId: String(url).includes("/redo/") ? "exec_redo_preview_001" : "exec_undo_preview_001",
            stepResults: [],
            summary: init?.body || ""
          };
        }
      }));

      return sidebar;
    }

    const undo = loadFailingApplySidebar();
    await expect(undo.undoExecution("exec_001")).rejects.toThrow("Apps Script flush failed.");
    expect(undo.fetch).toHaveBeenCalledTimes(1);
    expect(undo.fetch.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/undo/prepare");
    expect(undo.fetch.mock.calls.some(([url]) => String(url).endsWith("/api/execution/undo"))).toBe(false);

    const redo = loadFailingApplySidebar();
    await expect(redo.redoExecution("exec_undo_001")).rejects.toThrow("Apps Script flush failed.");
    expect(redo.fetch).toHaveBeenCalledTimes(1);
    expect(redo.fetch.mock.calls[0][0]).toBe("http://127.0.0.1:8787/api/execution/redo/prepare");
    expect(redo.fetch.mock.calls.some(([url]) => String(url).endsWith("/api/execution/redo"))).toBe(false);
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

  it("attaches local undo snapshots for Google Sheets range format writes", () => {
    let backgrounds = [["#ffffff"]];
    let fontWeights = [["normal"]];
    let numberFormats = [["General"]];
    const cloneMatrix = (matrix: string[][]) => matrix.map((row) => [...row]);
    const targetRange = {
      getNumRows: vi.fn(() => 1),
      getNumColumns: vi.fn(() => 1),
      getRow: vi.fn(() => 2),
      getColumn: vi.fn(() => 2),
      getBackgrounds: vi.fn(() => cloneMatrix(backgrounds)),
      setBackground: vi.fn((value: string) => {
        backgrounds = [[value]];
      }),
      getFontWeights: vi.fn(() => cloneMatrix(fontWeights)),
      setFontWeight: vi.fn((value: string) => {
        fontWeights = [[value]];
      }),
      getNumberFormats: vi.fn(() => cloneMatrix(numberFormats)),
      setNumberFormat: vi.fn((value: string) => {
        numberFormats = [[value]];
      })
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_range_format_snapshot_sheets_001",
      runId: "run_range_format_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_range_format_snapshot_sheets_001",
      plan: {
        targetSheet: "Sales",
        targetRange: "B2",
        format: {
          backgroundColor: "#fff2cc",
          bold: true,
          numberFormat: "$#,##0"
        },
        explanation: "Format revenue.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    });

    expect(result).toMatchObject({
      kind: "range_format_update",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_range_format_snapshot_sheets_001",
        kind: "range_format",
        targetSheet: "Sales",
        targetRange: "B2",
        shape: {
          rows: 1,
          columns: 1
        },
        beforeFormat: {
          backgrounds: [["#ffffff"]],
          fontWeights: [["normal"]],
          numberFormats: [["General"]]
        },
        afterFormat: {
          backgrounds: [["#fff2cc"]],
          fontWeights: [["bold"]],
          numberFormats: [["$#,##0"]]
        }
      }
    });
  });

  it("fails closed for Google Sheets border formatting when exact border styles are unavailable", () => {
    const targetRange = {
      setBorder: vi.fn()
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:C3");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(() => code.applyWritePlan({
      requestId: "req_range_format_border_styles_sheets_001",
      runId: "run_range_format_border_styles_sheets_001",
      approvalToken: "token",
      plan: {
        targetSheet: "Sales",
        targetRange: "B2:C3",
        format: {
          border: {
            outer: {
              style: "solid",
              color: "#1f2937"
            }
          }
        },
        explanation: "Add an outside border.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host does not support exact border style solid.");

    expect(targetRange.setBorder).not.toHaveBeenCalled();
  });

  it("fails closed for Google Sheets border formatting when the range border API is unavailable", () => {
    const targetRange = {};
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:C3");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(() => code.applyWritePlan({
      requestId: "req_range_format_border_api_sheets_001",
      runId: "run_range_format_border_api_sheets_001",
      approvalToken: "token",
      plan: {
        targetSheet: "Sales",
        targetRange: "B2:C3",
        format: {
          border: {
            outer: {
              style: "none"
            }
          }
        },
        explanation: "Remove an outside border.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host does not support exact range borders on this range.");
  });

  it("attaches local undo snapshots for Google Sheets autofit row writes", () => {
    const rowHeights = new Map<number, number>([
      [2, 24],
      [3, 28]
    ]);
    const targetRange = {
      getNumRows: vi.fn(() => 2),
      getNumColumns: vi.fn(() => 3),
      getRow: vi.fn(() => 2),
      getColumn: vi.fn(() => 4)
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("D2:F3");
        return targetRange;
      }),
      getRowHeight: vi.fn((rowIndex: number) => rowHeights.get(rowIndex)),
      autoResizeRows: vi.fn((startRow: number, rowCount: number) => {
        for (let offset = 0; offset < rowCount; offset += 1) {
          rowHeights.set(startRow + offset, 18 + offset);
        }
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_autofit_rows_snapshot_sheets_001",
      runId: "run_autofit_rows_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_autofit_rows_snapshot_sheets_001",
      plan: {
        targetSheet: "Sales",
        targetRange: "D2:F3",
        operation: "autofit_rows",
        explanation: "Autofit row heights.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(sheet.autoResizeRows).toHaveBeenCalledWith(2, 2);
    expect(result).toMatchObject({
      kind: "sheet_structure_update",
      operation: "autofit_rows",
      targetSheet: "Sales",
      targetRange: "D2:F3",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_autofit_rows_snapshot_sheets_001",
        kind: "range_format",
        targetSheet: "Sales",
        targetRange: "D2:F3",
        shape: {
          rows: 2,
          columns: 3
        },
        beforeFormat: {
          rowHeights: [24, 28]
        },
        afterFormat: {
          rowHeights: [18, 19]
        }
      }
    });
  });

  it("attaches local undo snapshots for Google Sheets merge cell writes", () => {
    let merged = false;
    let values: unknown[][] = [
      ["Region", ""],
      ["West", "East"]
    ];
    let formulas = [
      ["", ""],
      ["", ""]
    ];
    const targetRange = {
      getA1Notation: vi.fn(() => "B2:C3"),
      getNumRows: vi.fn(() => 2),
      getNumColumns: vi.fn(() => 2),
      getValues: vi.fn(() => values.map((row) => [...row])),
      getFormulas: vi.fn(() => formulas.map((row) => [...row])),
      getMergedRanges: vi.fn(() => merged ? [targetRange] : []),
      merge: vi.fn(() => {
        merged = true;
        values = [
          ["Region", ""],
          ["", ""]
        ];
        formulas = [
          ["", ""],
          ["", ""]
        ];
      }),
      breakApart: vi.fn(() => {
        merged = false;
      })
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:C3");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_merge_cells_snapshot_sheets_001",
      runId: "run_merge_cells_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_merge_cells_snapshot_sheets_001",
      plan: {
        targetSheet: "Sales",
        targetRange: "B2:C3",
        operation: "merge_cells",
        explanation: "Merge the regional header.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(targetRange.merge).toHaveBeenCalledTimes(1);
    expect(result).toMatchObject({
      kind: "sheet_structure_update",
      operation: "merge_cells",
      targetSheet: "Sales",
      targetRange: "B2:C3",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_merge_cells_snapshot_sheets_001",
        kind: "range_merge",
        targetSheet: "Sales",
        targetRange: "B2:C3",
        before: {
          merged: false,
          cells: [
            [{ kind: "value", value: { type: "string", value: "Region" } }, { kind: "value", value: { type: "string", value: "" } }],
            [{ kind: "value", value: { type: "string", value: "West" } }, { kind: "value", value: { type: "string", value: "East" } }]
          ]
        },
        after: {
          merged: true,
          cells: [
            [{ kind: "value", value: { type: "string", value: "Region" } }, { kind: "value", value: { type: "string", value: "" } }],
            [{ kind: "value", value: { type: "string", value: "" } }, { kind: "value", value: { type: "string", value: "" } }]
          ]
        }
      }
    });
  });

  it("passes Google Sheets merge snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const before = {
      merged: false,
      cells: [[{ kind: "value", value: { type: "string", value: "Region" } }]]
    };
    const after = {
      merged: true,
      cells: [[{ kind: "value", value: { type: "string", value: "Region" } }]]
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_merge_cells_001"],
      executions: {
        exec_merge_cells_001: {
          baseExecutionId: "exec_merge_cells_001"
        }
      },
      bases: {
        exec_merge_cells_001: {
          baseExecutionId: "exec_merge_cells_001",
          kind: "range_merge",
          targetSheet: "Sales",
          targetRange: "B2:C3",
          before,
          after
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "sheet_structure_update",
          operation: "merge_cells",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_merge_cells_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "range_merge",
          targetSheet: "Sales",
          targetRange: "B2:C3",
          from: after,
          to: before
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "range_merge",
          targetSheet: "Sales",
          targetRange: "B2:C3",
          from: after,
          to: before
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets merge snapshots on the server", () => {
    let merged = true;
    const values: unknown[][] = [
      ["Region", ""],
      ["", ""]
    ];
    const formulas = [
      ["", ""],
      ["", ""]
    ];
    const cellValues = new Map<string, unknown>();
    const targetRange = {
      getNumRows: vi.fn(() => 2),
      getNumColumns: vi.fn(() => 2),
      getMergedRanges: vi.fn(() => merged ? [targetRange] : []),
      getA1Notation: vi.fn(() => "B2:C3"),
      getValues: vi.fn(() => values.map((row) => [...row])),
      getFormulas: vi.fn(() => formulas.map((row) => [...row])),
      getNotes: vi.fn(() => [
        ["", ""],
        ["", ""]
      ]),
      merge: vi.fn(() => {
        merged = true;
      }),
      breakApart: vi.fn(() => {
        merged = false;
      }),
      getCell: vi.fn((row: number, column: number) => ({
        setValue: vi.fn((value: unknown) => {
          cellValues.set(`${row},${column}`, value);
        }),
        setFormula: vi.fn((formula: string) => {
          cellValues.set(`${row},${column}`, formula);
        })
      }))
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:C3");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "range_merge",
      targetSheet: "Sales",
      targetRange: "B2:C3",
      from: {
        merged: true,
        cells: [
          [{ kind: "value", value: { type: "string", value: "Region" }, note: "" }, { kind: "value", value: { type: "string", value: "" }, note: "" }],
          [{ kind: "value", value: { type: "string", value: "" }, note: "" }, { kind: "value", value: { type: "string", value: "" }, note: "" }]
        ]
      },
      to: {
        merged: false,
        cells: [
          [{ kind: "value", value: { type: "string", value: "Region" } }, { kind: "value", value: { type: "string", value: "" } }],
          [{ kind: "value", value: { type: "string", value: "West" } }, { kind: "value", value: { type: "string", value: "East" } }]
        ]
      }
    })).toMatchObject({
      ok: true,
      targetSheet: "Sales",
      targetRange: "B2:C3"
    });

    expect(targetRange.breakApart).toHaveBeenCalledTimes(1);
    expect(targetRange.merge).not.toHaveBeenCalled();
    expect(merged).toBe(false);
    expect(cellValues.get("1,1")).toBe("Region");
    expect(cellValues.get("2,1")).toBe("West");
    expect(cellValues.get("2,2")).toBe("East");
    expect(code.flush).toHaveBeenCalled();
  });

  it("fails closed when Google Sheets merge snapshots no longer match the current merge state", () => {
    const values: unknown[][] = [
      ["Region", ""],
      ["West", "East"]
    ];
    const formulas = [
      ["", ""],
      ["", ""]
    ];
    const targetRange = {
      getNumRows: vi.fn(() => 2),
      getNumColumns: vi.fn(() => 2),
      getMergedRanges: vi.fn(() => []),
      getA1Notation: vi.fn(() => "B2:C3"),
      getValues: vi.fn(() => values.map((row) => [...row])),
      getFormulas: vi.fn(() => formulas.map((row) => [...row])),
      getNotes: vi.fn(() => [
        ["", ""],
        ["", ""]
      ]),
      breakApart: vi.fn(),
      merge: vi.fn()
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:C3");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(() => code.validateExecutionCellSnapshot({
      kind: "range_merge",
      targetSheet: "Sales",
      targetRange: "B2:C3",
      from: {
        merged: true,
        cells: [
          [{ kind: "value", value: { type: "string", value: "Region" } }, { kind: "value", value: { type: "string", value: "" } }],
          [{ kind: "value", value: { type: "string", value: "" } }, { kind: "value", value: { type: "string", value: "" } }]
        ]
      },
      to: {
        merged: false,
        cells: [
          [{ kind: "value", value: { type: "string", value: "Region" } }, { kind: "value", value: { type: "string", value: "" } }],
          [{ kind: "value", value: { type: "string", value: "West" } }, { kind: "value", value: { type: "string", value: "East" } }]
        ]
      }
    })).toThrow("Range merge state changed since this history entry was captured.");
    expect(targetRange.breakApart).not.toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Google Sheets autofit column writes", () => {
    const columnWidths = new Map<number, number>([
      [4, 64],
      [5, 72],
      [6, 80]
    ]);
    const targetRange = {
      getNumRows: vi.fn(() => 2),
      getNumColumns: vi.fn(() => 3),
      getRow: vi.fn(() => 2),
      getColumn: vi.fn(() => 4)
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("D2:F3");
        return targetRange;
      }),
      getColumnWidth: vi.fn((columnIndex: number) => columnWidths.get(columnIndex)),
      autoResizeColumns: vi.fn((startColumn: number, columnCount: number) => {
        for (let offset = 0; offset < columnCount; offset += 1) {
          columnWidths.set(startColumn + offset, 90 + offset);
        }
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_autofit_columns_snapshot_sheets_001",
      runId: "run_autofit_columns_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_autofit_columns_snapshot_sheets_001",
      plan: {
        targetSheet: "Sales",
        targetRange: "D2:F3",
        operation: "autofit_columns",
        explanation: "Autofit column widths.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(sheet.autoResizeColumns).toHaveBeenCalledWith(4, 3);
    expect(result).toMatchObject({
      kind: "sheet_structure_update",
      operation: "autofit_columns",
      targetSheet: "Sales",
      targetRange: "D2:F3",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_autofit_columns_snapshot_sheets_001",
        kind: "range_format",
        targetSheet: "Sales",
        targetRange: "D2:F3",
        shape: {
          rows: 2,
          columns: 3
        },
        beforeFormat: {
          columnWidths: [64, 72, 80]
        },
        afterFormat: {
          columnWidths: [90, 91, 92]
        }
      }
    });
  });

  it("passes Google Sheets range format snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
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
          beforeFormat: {
            backgrounds: [["#ffffff"]],
            fontWeights: [["normal"]],
            numberFormats: [["General"]]
          },
          afterFormat: {
            backgrounds: [["#fff2cc"]],
            fontWeights: [["bold"]],
            numberFormats: [["$#,##0"]]
          }
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "range_format_update",
          operation: "range_format_update",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_range_format_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "range_format",
          targetSheet: "Sales",
          targetRange: "B2",
          shape: {
            rows: 1,
            columns: 1
          },
          format: {
            backgrounds: [["#ffffff"]],
            fontWeights: [["normal"]],
            numberFormats: [["General"]]
          }
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "range_format",
          targetSheet: "Sales",
          targetRange: "B2",
          shape: {
            rows: 1,
            columns: 1
          },
          format: {
            backgrounds: [["#ffffff"]],
            fontWeights: [["normal"]],
            numberFormats: [["General"]]
          }
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("passes Google Sheets conditional format snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const before = {
      rules: [
        {
          ranges: ["B2:B20"],
          rule: {
            kind: "text_contains",
            text: "old"
          },
          format: {
            backgroundColor: "#eeeeee"
          }
        }
      ]
    };
    const after = {
      rules: [
        {
          ranges: ["B2:B20"],
          rule: {
            kind: "text_contains",
            text: "overdue"
          },
          format: {
            backgroundColor: "#ffcccc"
          }
        }
      ]
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_conditional_format_001"],
      executions: {
        exec_conditional_format_001: {
          baseExecutionId: "exec_conditional_format_001"
        }
      },
      bases: {
        exec_conditional_format_001: {
          baseExecutionId: "exec_conditional_format_001",
          kind: "conditional_format",
          targetSheet: "Sheet1",
          targetRange: "B2:B20",
          before,
          after
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "conditional_format_update",
          operation: "conditional_format_update",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_conditional_format_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "conditional_format",
          targetSheet: "Sheet1",
          targetRange: "B2:B20",
          from: after,
          to: before
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "conditional_format",
          targetSheet: "Sheet1",
          targetRange: "B2:B20",
          from: after,
          to: before
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets range format snapshots on the server", () => {
    let backgrounds = [["#fff2cc"]];
    let fontWeights = [["bold"]];
    let numberFormats = [["$#,##0"]];
    const cloneMatrix = (matrix: string[][]) => matrix.map((row) => [...row]);
    const targetRange = {
      getNumRows: vi.fn(() => 1),
      getNumColumns: vi.fn(() => 1),
      setBackgrounds: vi.fn((nextBackgrounds: string[][]) => {
        backgrounds = cloneMatrix(nextBackgrounds);
      }),
      setFontWeights: vi.fn((nextFontWeights: string[][]) => {
        fontWeights = cloneMatrix(nextFontWeights);
      }),
      setNumberFormats: vi.fn((nextNumberFormats: string[][]) => {
        numberFormats = cloneMatrix(nextNumberFormats);
      })
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "range_format",
      targetSheet: "Sales",
      targetRange: "B2",
      shape: {
        rows: 1,
        columns: 1
      },
      format: {
        backgrounds: [["#ffffff"]],
        fontWeights: [["normal"]],
        numberFormats: [["General"]]
      }
    })).toMatchObject({
      ok: true,
      targetSheet: "Sales",
      targetRange: "B2"
    });

    expect(backgrounds).toEqual([["#ffffff"]]);
    expect(fontWeights).toEqual([["normal"]]);
    expect(numberFormats).toEqual([["General"]]);
    expect(code.flush).toHaveBeenCalled();
  });

  it("passes Google Sheets data validation snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
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
          beforeValidation: {
            rule: null
          },
          afterValidation: {
            rule: {
              criteriaType: "NUMBER_BETWEEN",
              criteriaValues: [
                { kind: "scalar", value: 1 },
                { kind: "scalar", value: 10 }
              ],
              allowInvalid: false,
              helpText: "Enter a number from 1 to 10."
            }
          }
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "data_validation_update",
          operation: "data_validation_update",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_validation_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "data_validation",
          targetSheet: "Sales",
          targetRange: "B2:B20",
          shape: {
            rows: 19,
            columns: 1
          },
          validation: {
            rule: null
          }
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "data_validation",
          targetSheet: "Sales",
          targetRange: "B2:B20",
          shape: {
            rows: 19,
            columns: 1
          },
          validation: {
            rule: null
          }
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets data validation snapshots on the server", () => {
    let validationRule: unknown = {
      criteriaType: "NUMBER_BETWEEN",
      criteriaValues: [1, 10],
      allowInvalid: false,
      helpText: "Enter a number from 1 to 10."
    };
    const setDataValidation = vi.fn((nextRule: unknown) => {
      validationRule = nextRule;
    });
    const targetRange = {
      getNumRows: vi.fn(() => 19),
      getNumColumns: vi.fn(() => 1),
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:B20");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "data_validation",
      targetSheet: "Sales",
      targetRange: "B2:B20",
      shape: {
        rows: 19,
        columns: 1
      },
      validation: {
        rule: null
      }
    })).toMatchObject({
      ok: true,
      targetSheet: "Sales",
      targetRange: "B2:B20"
    });

    expect(setDataValidation).toHaveBeenCalledWith(null);
    expect(validationRule).toBeNull();

    expect(code.applyExecutionCellSnapshot({
      kind: "data_validation",
      targetSheet: "Sales",
      targetRange: "B2:B20",
      shape: {
        rows: 19,
        columns: 1
      },
      validation: {
        rule: {
          criteriaType: "NUMBER_BETWEEN",
          criteriaValues: [
            { kind: "scalar", value: 1 },
            { kind: "scalar", value: 10 }
          ],
          allowInvalid: false,
          helpText: "Enter a number from 1 to 10."
        }
      }
    })).toMatchObject({
      ok: true,
      targetSheet: "Sales",
      targetRange: "B2:B20"
    });

    expect(validationRule).toEqual({
      criteriaType: "NUMBER_BETWEEN",
      criteriaValues: [1, 10],
      allowInvalid: false,
      helpText: "Enter a number from 1 to 10."
    });
    expect(code.flush).toHaveBeenCalled();
  });

  it("fails closed before restoring Google Sheets validation snapshots when builder options are unavailable", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      getNumRows: vi.fn(() => 19),
      getNumColumns: vi.fn(() => 1),
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:B20");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({
      spreadsheet,
      spreadsheetAppOverrides: {
        newDataValidation() {
          return {
            withCriteria: vi.fn(function() {
              return this;
            }),
            build: vi.fn(() => ({ kind: "validation" }))
          };
        }
      }
    });

    expect(() => code.applyExecutionCellSnapshot({
      kind: "data_validation",
      targetSheet: "Sales",
      targetRange: "B2:B20",
      shape: {
        rows: 19,
        columns: 1
      },
      validation: {
        rule: {
          criteriaType: "NUMBER_BETWEEN",
          criteriaValues: [
            { kind: "scalar", value: 1 },
            { kind: "scalar", value: 10 }
          ],
          allowInvalid: false
        }
      }
    })).toThrow("Google Sheets host does not support exact validation invalid-entry behavior.");

    expect(setDataValidation).not.toHaveBeenCalled();
  });

  it("fails closed before restoring Google Sheets validation snapshot help text when builder support is unavailable", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      getNumRows: vi.fn(() => 19),
      getNumColumns: vi.fn(() => 1),
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:B20");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({
      spreadsheet,
      spreadsheetAppOverrides: {
        newDataValidation() {
          return {
            withCriteria: vi.fn(function() {
              return this;
            }),
            setAllowInvalid: vi.fn(function() {
              return this;
            }),
            build: vi.fn(() => ({ kind: "validation" }))
          };
        }
      }
    });

    expect(() => code.applyExecutionCellSnapshot({
      kind: "data_validation",
      targetSheet: "Sales",
      targetRange: "B2:B20",
      shape: {
        rows: 19,
        columns: 1
      },
      validation: {
        rule: {
          criteriaType: "NUMBER_BETWEEN",
          criteriaValues: [
            { kind: "scalar", value: 1 },
            { kind: "scalar", value: 10 }
          ],
          allowInvalid: false,
          helpText: "Enter a number from 1 to 10."
        }
      }
    })).toThrow("Google Sheets host does not support exact validation help text.");

    expect(setDataValidation).not.toHaveBeenCalled();
  });

  it("fails closed before applying Google Sheets validation when builder options are unavailable", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:B20");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({
      spreadsheet,
      spreadsheetAppOverrides: {
        newDataValidation() {
          return {
            requireValueInList: vi.fn(function() {
              return this;
            }),
            build: vi.fn(() => ({ kind: "validation" }))
          };
        }
      }
    });

    expect(() => code.applyWritePlan({
      requestId: "req_validation_builder_api_sheets_001",
      runId: "run_validation_builder_api_sheets_001",
      approvalToken: "token",
      plan: {
        type: "data_validation_plan",
        targetSheet: "Sales",
        targetRange: "B2:B20",
        ruleType: "list",
        values: ["Open", "Closed"],
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Add status dropdown validation.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host does not support exact validation invalid-entry behavior.");

    expect(setDataValidation).not.toHaveBeenCalled();
  });

  it("fails closed before applying Google Sheets validation help text when builder support is unavailable", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("B2:B20");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({
      spreadsheet,
      spreadsheetAppOverrides: {
        newDataValidation() {
          return {
            setAllowInvalid: vi.fn(function() {
              return this;
            }),
            requireValueInList: vi.fn(function() {
              return this;
            }),
            build: vi.fn(() => ({ kind: "validation" }))
          };
        }
      }
    });

    expect(() => code.applyWritePlan({
      requestId: "req_validation_builder_help_text_sheets_001",
      runId: "run_validation_builder_help_text_sheets_001",
      approvalToken: "token",
      plan: {
        type: "data_validation_plan",
        targetSheet: "Sales",
        targetRange: "B2:B20",
        ruleType: "list",
        values: ["Open", "Closed"],
        allowBlank: false,
        invalidDataBehavior: "reject",
        helpText: "Choose one of the approved statuses.",
        explanation: "Add status dropdown validation.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host does not support exact validation help text.");

    expect(setDataValidation).not.toHaveBeenCalled();
  });

  it("passes Google Sheets named range snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
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
            targetSheet: "Sales",
            targetRange: "A1:A10"
          },
          after: {
            exists: true,
            name: "NewRange",
            targetSheet: "Sales",
            targetRange: "A1:A10"
          }
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "named_range_update",
          operation: "rename",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_named_range_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "named_range",
          scope: "workbook",
          from: {
            exists: true,
            name: "NewRange",
            targetSheet: "Sales",
            targetRange: "A1:A10"
          },
          to: {
            exists: true,
            name: "OldRange",
            targetSheet: "Sales",
            targetRange: "A1:A10"
          }
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "named_range",
          scope: "workbook",
          from: {
            exists: true,
            name: "NewRange",
            targetSheet: "Sales",
            targetRange: "A1:A10"
          },
          to: {
            exists: true,
            name: "OldRange",
            targetSheet: "Sales",
            targetRange: "A1:A10"
          }
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets named range snapshots on the server", () => {
    const targetRange = {
      getA1Notation() {
        return "A1:A10";
      }
    };
    const namedRange = {
      getName: vi.fn(() => "NewRange"),
      setName: vi.fn(),
      setRange: vi.fn(),
      getRange: vi.fn(() => targetRange)
    };
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:A10");
        return targetRange;
      })
    };
    const spreadsheet = {
      getNamedRanges: vi.fn(() => [namedRange]),
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      }),
      setNamedRange: vi.fn(),
      removeNamedRange: vi.fn()
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "named_range",
      scope: "workbook",
      from: {
        exists: true,
        name: "NewRange",
        targetSheet: "Sales",
        targetRange: "A1:A10"
      },
      to: {
        exists: true,
        name: "OldRange",
        targetSheet: "Sales",
        targetRange: "A1:A10"
      }
    })).toMatchObject({
      ok: true,
      scope: "workbook"
    });

    expect(namedRange.setName).toHaveBeenCalledWith("OldRange");
    expect(namedRange.setRange).not.toHaveBeenCalled();
    expect(spreadsheet.setNamedRange).not.toHaveBeenCalled();
    expect(spreadsheet.removeNamedRange).not.toHaveBeenCalled();
    expect(code.flush).toHaveBeenCalled();
  });

  it("passes Google Sheets range filter snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const beforeFilter = {
      exists: true,
      targetRange: "A1:F25",
      criteria: [
        {
          criteriaType: "TEXT_EQUAL_TO",
          criteriaValues: [{ kind: "scalar", value: "Closed" }]
        },
        null,
        null,
        null,
        null,
        null
      ]
    };
    const afterFilter = {
      exists: true,
      targetRange: "A1:F25",
      criteria: [
        {
          criteriaType: "TEXT_EQUAL_TO",
          criteriaValues: [{ kind: "scalar", value: "Open" }]
        },
        null,
        null,
        null,
        null,
        null
      ]
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_filter_001"],
      executions: {
        exec_filter_001: {
          baseExecutionId: "exec_filter_001"
        }
      },
      bases: {
        exec_filter_001: {
          baseExecutionId: "exec_filter_001",
          kind: "range_filter",
          targetSheet: "Sales",
          targetRange: "A1:F25",
          beforeFilter,
          afterFilter
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "range_filter",
          operation: "range_filter",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_filter_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "range_filter",
          targetSheet: "Sales",
          targetRange: "A1:F25",
          filter: beforeFilter
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "range_filter",
          targetSheet: "Sales",
          targetRange: "A1:F25",
          filter: beforeFilter
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets range filter snapshots on the server", () => {
    const filterCriteriaByColumn = new Map<number, unknown>([
      [1, { type: "TEXT_EQUAL_TO", values: ["Open"] }]
    ]);
    const filter = {
      getRange() {
        return {
          getA1Notation() {
            return "A1:F25";
          }
        };
      },
      removeColumnFilterCriteria: vi.fn((columnPosition: number) => {
        filterCriteriaByColumn.delete(columnPosition);
      }),
      setColumnFilterCriteria: vi.fn((columnPosition: number, criteria: unknown) => {
        filterCriteriaByColumn.set(columnPosition, criteria);
      })
    };
    const targetRange = {
      getA1Notation() {
        return "A1:F25";
      },
      getNumColumns() {
        return 6;
      },
      getNumRows() {
        return 25;
      },
      createFilter: vi.fn(() => filter)
    };
    const sheet = {
      getFilter: vi.fn(() => filter),
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:F25");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "range_filter",
      targetSheet: "Sales",
      targetRange: "A1:F25",
      filter: {
        exists: true,
        targetRange: "A1:F25",
        criteria: [
          {
            criteriaType: "TEXT_EQUAL_TO",
            criteriaValues: [{ kind: "scalar", value: "Closed" }]
          },
          null,
          null,
          null,
          null,
          null
        ]
      }
    })).toMatchObject({
      ok: true,
      targetSheet: "Sales",
      targetRange: "A1:F25"
    });

    expect(filter.removeColumnFilterCriteria).toHaveBeenCalledTimes(6);
    expect(filter.setColumnFilterCriteria).toHaveBeenCalledWith(1, {
      type: "TEXT_EQUAL_TO",
      values: ["Closed"]
    });
    expect(filterCriteriaByColumn.get(1)).toEqual({
      type: "TEXT_EQUAL_TO",
      values: ["Closed"]
    });
    expect(code.flush).toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Google Sheets sheet rename writes", () => {
    let sheetName = "OldName";
    const sheet = {
      getName: vi.fn(() => sheetName),
      setName: vi.fn((nextName: string) => {
        sheetName = nextName;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("OldName");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_rename_sheet_snapshot_sheets_001",
      runId: "run_rename_sheet_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_rename_sheet_snapshot_sheets_001",
      plan: {
        operation: "rename_sheet",
        sheetName: "OldName",
        newSheetName: "NewName",
        explanation: "Rename the staging sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(sheetName).toBe("NewName");
    expect(result).toMatchObject({
      kind: "workbook_structure_update",
      operation: "rename_sheet",
      sheetName: "OldName",
      newSheetName: "NewName",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_rename_sheet_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "rename_sheet",
        before: {
          exists: true,
          name: "OldName"
        },
        after: {
          exists: true,
          name: "NewName"
        }
      }
    });
  });

  it("passes Google Sheets sheet rename snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const before = {
      exists: true,
      name: "OldName"
    };
    const after = {
      exists: true,
      name: "NewName"
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_rename_sheet_001"],
      executions: {
        exec_rename_sheet_001: {
          baseExecutionId: "exec_rename_sheet_001"
        }
      },
      bases: {
        exec_rename_sheet_001: {
          baseExecutionId: "exec_rename_sheet_001",
          kind: "workbook_structure",
          operation: "rename_sheet",
          before,
          after
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "workbook_structure_update",
          operation: "rename_sheet",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_rename_sheet_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "workbook_structure",
          operation: "rename_sheet",
          from: after,
          to: before
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "workbook_structure",
          operation: "rename_sheet",
          from: after,
          to: before
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets sheet rename snapshots on the server", () => {
    let sheetName = "NewName";
    const sheet = {
      getName: vi.fn(() => sheetName),
      setName: vi.fn((nextName: string) => {
        sheetName = nextName;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("NewName");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "rename_sheet",
      from: {
        exists: true,
        name: "NewName"
      },
      to: {
        exists: true,
        name: "OldName"
      }
    })).toMatchObject({
      ok: true,
      operation: "rename_sheet"
    });

    expect(sheetName).toBe("OldName");
    expect(code.flush).toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Google Sheets sheet visibility writes", () => {
    let hidden = false;
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      isSheetHidden: vi.fn(() => hidden),
      hideSheet: vi.fn(() => {
        hidden = true;
      }),
      showSheet: vi.fn(() => {
        hidden = false;
      })
    };
    const otherSheet = {
      getName: vi.fn(() => "Sheet2"),
      isSheetHidden: vi.fn(() => false)
    };
    const spreadsheet = {
      getSheets: vi.fn(() => [sheet, otherSheet]),
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_hide_sheet_snapshot_sheets_001",
      runId: "run_hide_sheet_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_hide_sheet_snapshot_sheets_001",
      plan: {
        operation: "hide_sheet",
        sheetName: "Sheet1",
        explanation: "Hide the staging sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(hidden).toBe(true);
    expect(result).toMatchObject({
      kind: "workbook_structure_update",
      operation: "hide_sheet",
      sheetName: "Sheet1",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_hide_sheet_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "sheet_visibility",
        before: {
          exists: true,
          name: "Sheet1",
          visibility: "visible"
        },
        after: {
          exists: true,
          name: "Sheet1",
          visibility: "hidden"
        }
      }
    });
  });

  it("passes Google Sheets sheet visibility snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const before = {
      exists: true,
      name: "Sheet1",
      visibility: "visible"
    };
    const after = {
      exists: true,
      name: "Sheet1",
      visibility: "hidden"
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_hide_sheet_001"],
      executions: {
        exec_hide_sheet_001: {
          baseExecutionId: "exec_hide_sheet_001"
        }
      },
      bases: {
        exec_hide_sheet_001: {
          baseExecutionId: "exec_hide_sheet_001",
          kind: "workbook_structure",
          operation: "sheet_visibility",
          before,
          after
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "workbook_structure_update",
          operation: "hide_sheet",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_hide_sheet_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "workbook_structure",
          operation: "sheet_visibility",
          from: after,
          to: before
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "workbook_structure",
          operation: "sheet_visibility",
          from: after,
          to: before
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets sheet visibility snapshots on the server", () => {
    let hidden = true;
    const sheet = {
      isSheetHidden: vi.fn(() => hidden),
      hideSheet: vi.fn(() => {
        hidden = true;
      }),
      showSheet: vi.fn(() => {
        hidden = false;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "sheet_visibility",
      from: {
        exists: true,
        name: "Sheet1",
        visibility: "hidden"
      },
      to: {
        exists: true,
        name: "Sheet1",
        visibility: "visible"
      }
    })).toMatchObject({
      ok: true,
      operation: "sheet_visibility"
    });

    expect(hidden).toBe(false);
    expect(sheet.showSheet).toHaveBeenCalledTimes(1);
    expect(sheet.hideSheet).not.toHaveBeenCalled();
    expect(code.flush).toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Google Sheets sheet tab color writes", () => {
    let tabColor = "#ffcccc";
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      getTabColor: vi.fn(() => tabColor),
      setTabColor: vi.fn((nextColor: string | null) => {
        tabColor = nextColor || "";
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_tab_color_snapshot_sheets_001",
      runId: "run_tab_color_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_tab_color_snapshot_sheets_001",
      plan: {
        operation: "set_sheet_tab_color",
        targetSheet: "Sheet1",
        color: "#00ff00",
        explanation: "Color the working sheet tab.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(tabColor).toBe("#00ff00");
    expect(result).toMatchObject({
      kind: "sheet_structure_update",
      operation: "set_sheet_tab_color",
      targetSheet: "Sheet1",
      color: "#00ff00",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_tab_color_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "sheet_tab_color",
        before: {
          exists: true,
          name: "Sheet1",
          color: "#ffcccc"
        },
        after: {
          exists: true,
          name: "Sheet1",
          color: "#00ff00"
        }
      }
    });
  });

  it("applies Google Sheets sheet tab color snapshots on the server", () => {
    let tabColor = "#00ff00";
    const sheet = {
      getTabColor: vi.fn(() => tabColor),
      setTabColor: vi.fn((nextColor: string | null) => {
        tabColor = nextColor || "";
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "sheet_tab_color",
      from: {
        exists: true,
        name: "Sheet1",
        color: "#00ff00"
      },
      to: {
        exists: true,
        name: "Sheet1",
        color: "#ffcccc"
      }
    })).toMatchObject({
      ok: true,
      operation: "sheet_tab_color"
    });

    expect(tabColor).toBe("#ffcccc");
    expect(sheet.setTabColor).toHaveBeenCalledWith("#ffcccc");
    expect(code.flush).toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Google Sheets freeze pane writes", () => {
    let frozenRows = 1;
    let frozenColumns = 0;
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      getFrozenRows: vi.fn(() => frozenRows),
      getFrozenColumns: vi.fn(() => frozenColumns),
      setFrozenRows: vi.fn((nextRows: number) => {
        frozenRows = nextRows;
      }),
      setFrozenColumns: vi.fn((nextColumns: number) => {
        frozenColumns = nextColumns;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_freeze_pane_snapshot_sheets_001",
      runId: "run_freeze_pane_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_freeze_pane_snapshot_sheets_001",
      plan: {
        operation: "freeze_panes",
        targetSheet: "Sheet1",
        frozenRows: 2,
        frozenColumns: 3,
        explanation: "Freeze the header rows and leading columns.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(frozenRows).toBe(2);
    expect(frozenColumns).toBe(3);
    expect(result).toMatchObject({
      kind: "sheet_structure_update",
      operation: "freeze_panes",
      targetSheet: "Sheet1",
      frozenRows: 2,
      frozenColumns: 3,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_freeze_pane_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "sheet_freeze_panes",
        before: {
          exists: true,
          name: "Sheet1",
          frozenRows: 1,
          frozenColumns: 0
        },
        after: {
          exists: true,
          name: "Sheet1",
          frozenRows: 2,
          frozenColumns: 3
        }
      }
    });
  });

  it("attaches local undo snapshots for Google Sheets row visibility writes", () => {
    const rowHidden = new Map([
      [3, false],
      [4, true],
      [5, false]
    ]);
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      isRowHiddenByUser: vi.fn((rowIndex: number) => rowHidden.get(rowIndex)),
      hideRows: vi.fn((startRow: number, rowCount: number) => {
        for (let offset = 0; offset < rowCount; offset += 1) {
          rowHidden.set(startRow + offset, true);
        }
      }),
      showRows: vi.fn((startRow: number, rowCount: number) => {
        for (let offset = 0; offset < rowCount; offset += 1) {
          rowHidden.set(startRow + offset, false);
        }
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_hide_rows_snapshot_sheets_001",
      runId: "run_hide_rows_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_hide_rows_snapshot_sheets_001",
      plan: {
        operation: "hide_rows",
        targetSheet: "Sheet1",
        startIndex: 2,
        count: 3,
        explanation: "Hide subtotal rows.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(rowHidden.get(3)).toBe(true);
    expect(rowHidden.get(4)).toBe(true);
    expect(rowHidden.get(5)).toBe(true);
    expect(result).toMatchObject({
      kind: "sheet_structure_update",
      operation: "hide_rows",
      targetSheet: "Sheet1",
      startIndex: 2,
      count: 3,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_hide_rows_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "row_column_visibility",
        before: {
          exists: true,
          name: "Sheet1",
          dimension: "rows",
          startIndex: 2,
          count: 3,
          hiddenStates: [false, true, false]
        },
        after: {
          exists: true,
          name: "Sheet1",
          dimension: "rows",
          startIndex: 2,
          count: 3,
          hiddenStates: [true, true, true]
        }
      }
    });
  });

  it("attaches local undo snapshots for Google Sheets column visibility writes", () => {
    const columnHidden = new Map([
      [2, true],
      [3, false]
    ]);
    const sheet = {
      getName: vi.fn(() => "Sheet1"),
      isColumnHiddenByUser: vi.fn((columnIndex: number) => columnHidden.get(columnIndex)),
      hideColumns: vi.fn((startColumn: number, columnCount: number) => {
        for (let offset = 0; offset < columnCount; offset += 1) {
          columnHidden.set(startColumn + offset, true);
        }
      }),
      showColumns: vi.fn((startColumn: number, columnCount: number) => {
        for (let offset = 0; offset < columnCount; offset += 1) {
          columnHidden.set(startColumn + offset, false);
        }
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_unhide_columns_snapshot_sheets_001",
      runId: "run_unhide_columns_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_unhide_columns_snapshot_sheets_001",
      plan: {
        operation: "unhide_columns",
        targetSheet: "Sheet1",
        startIndex: 1,
        count: 2,
        explanation: "Unhide working columns.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(columnHidden.get(2)).toBe(false);
    expect(columnHidden.get(3)).toBe(false);
    expect(result).toMatchObject({
      kind: "sheet_structure_update",
      operation: "unhide_columns",
      targetSheet: "Sheet1",
      startIndex: 1,
      count: 2,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_unhide_columns_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "row_column_visibility",
        before: {
          exists: true,
          name: "Sheet1",
          dimension: "columns",
          startIndex: 1,
          count: 2,
          hiddenStates: [true, false]
        },
        after: {
          exists: true,
          name: "Sheet1",
          dimension: "columns",
          startIndex: 1,
          count: 2,
          hiddenStates: [false, false]
        }
      }
    });
  });

  it("applies Google Sheets freeze pane snapshots on the server", () => {
    let frozenRows = 2;
    let frozenColumns = 3;
    const sheet = {
      getFrozenRows: vi.fn(() => frozenRows),
      getFrozenColumns: vi.fn(() => frozenColumns),
      setFrozenRows: vi.fn((nextRows: number) => {
        frozenRows = nextRows;
      }),
      setFrozenColumns: vi.fn((nextColumns: number) => {
        frozenColumns = nextColumns;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "sheet_freeze_panes",
      from: {
        exists: true,
        name: "Sheet1",
        frozenRows: 2,
        frozenColumns: 3
      },
      to: {
        exists: true,
        name: "Sheet1",
        frozenRows: 0,
        frozenColumns: 1
      }
    })).toMatchObject({
      ok: true,
      operation: "sheet_freeze_panes"
    });

    expect(frozenRows).toBe(0);
    expect(frozenColumns).toBe(1);
    expect(sheet.setFrozenRows).toHaveBeenCalledWith(0);
    expect(sheet.setFrozenColumns).toHaveBeenCalledWith(1);
    expect(code.flush).toHaveBeenCalled();
  });

  it("applies Google Sheets row visibility snapshots on the server", () => {
    const rowHidden = new Map([
      [3, true],
      [4, true],
      [5, true]
    ]);
    const sheet = {
      isRowHiddenByUser: vi.fn((rowIndex: number) => rowHidden.get(rowIndex)),
      hideRows: vi.fn((startRow: number, rowCount: number) => {
        for (let offset = 0; offset < rowCount; offset += 1) {
          rowHidden.set(startRow + offset, true);
        }
      }),
      showRows: vi.fn((startRow: number, rowCount: number) => {
        for (let offset = 0; offset < rowCount; offset += 1) {
          rowHidden.set(startRow + offset, false);
        }
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "row_column_visibility",
      from: {
        exists: true,
        name: "Sheet1",
        dimension: "rows",
        startIndex: 2,
        count: 3,
        hiddenStates: [true, true, true]
      },
      to: {
        exists: true,
        name: "Sheet1",
        dimension: "rows",
        startIndex: 2,
        count: 3,
        hiddenStates: [false, true, false]
      }
    })).toMatchObject({
      ok: true,
      operation: "row_column_visibility"
    });

    expect(rowHidden.get(3)).toBe(false);
    expect(rowHidden.get(4)).toBe(true);
    expect(rowHidden.get(5)).toBe(false);
    expect(code.flush).toHaveBeenCalled();
  });

  it("applies Google Sheets column visibility snapshots on the server", () => {
    const columnHidden = new Map([
      [2, false],
      [3, false]
    ]);
    const sheet = {
      isColumnHiddenByUser: vi.fn((columnIndex: number) => columnHidden.get(columnIndex)),
      hideColumns: vi.fn((startColumn: number, columnCount: number) => {
        for (let offset = 0; offset < columnCount; offset += 1) {
          columnHidden.set(startColumn + offset, true);
        }
      }),
      showColumns: vi.fn((startColumn: number, columnCount: number) => {
        for (let offset = 0; offset < columnCount; offset += 1) {
          columnHidden.set(startColumn + offset, false);
        }
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet1");
        return sheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "row_column_visibility",
      from: {
        exists: true,
        name: "Sheet1",
        dimension: "columns",
        startIndex: 1,
        count: 2,
        hiddenStates: [false, false]
      },
      to: {
        exists: true,
        name: "Sheet1",
        dimension: "columns",
        startIndex: 1,
        count: 2,
        hiddenStates: [true, false]
      }
    })).toMatchObject({
      ok: true,
      operation: "row_column_visibility"
    });

    expect(columnHidden.get(2)).toBe(true);
    expect(columnHidden.get(3)).toBe(false);
    expect(code.flush).toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Google Sheets sheet move writes", () => {
    const firstSheet = {
      getName: vi.fn(() => "Sheet1"),
      getIndex: vi.fn(() => sheets.indexOf(firstSheet) + 1)
    };
    const secondSheet = {
      getName: vi.fn(() => "Sheet2"),
      getIndex: vi.fn(() => sheets.indexOf(secondSheet) + 1)
    };
    const targetSheet = {
      getName: vi.fn(() => "Sheet3"),
      getIndex: vi.fn(() => sheets.indexOf(targetSheet) + 1)
    };
    const sheets = [firstSheet, secondSheet, targetSheet];
    let activeSheet: typeof targetSheet | null = null;
    const spreadsheet = {
      getSheets: vi.fn(() => sheets),
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet3");
        return targetSheet;
      }),
      setActiveSheet: vi.fn((sheet: typeof targetSheet) => {
        activeSheet = sheet;
      }),
      moveActiveSheet: vi.fn((position: number) => {
        if (!activeSheet) {
          throw new Error("No active sheet");
        }

        const currentIndex = sheets.indexOf(activeSheet);
        sheets.splice(currentIndex, 1);
        sheets.splice(position - 1, 0, activeSheet);
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_move_sheet_snapshot_sheets_001",
      runId: "run_move_sheet_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_move_sheet_snapshot_sheets_001",
      plan: {
        operation: "move_sheet",
        sheetName: "Sheet3",
        position: 0,
        explanation: "Move the staging sheet to the front.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(sheets.map((sheet) => sheet.getName())).toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    expect(spreadsheet.moveActiveSheet).toHaveBeenCalledWith(1);
    expect(result).toMatchObject({
      kind: "workbook_structure_update",
      operation: "move_sheet",
      sheetName: "Sheet3",
      positionResolved: 0,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_move_sheet_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "move_sheet",
        before: {
          exists: true,
          name: "Sheet3",
          position: 2
        },
        after: {
          exists: true,
          name: "Sheet3",
          position: 0
        }
      }
    });
  });

  it("applies Google Sheets sheet move snapshots on the server", () => {
    const firstSheet = {
      getName: vi.fn(() => "Sheet1"),
      getIndex: vi.fn(() => sheets.indexOf(firstSheet) + 1)
    };
    const targetSheet = {
      getName: vi.fn(() => "Sheet3"),
      getIndex: vi.fn(() => sheets.indexOf(targetSheet) + 1)
    };
    const lastSheet = {
      getName: vi.fn(() => "Sheet4"),
      getIndex: vi.fn(() => sheets.indexOf(lastSheet) + 1)
    };
    const sheets = [targetSheet, firstSheet, lastSheet];
    let activeSheet: typeof targetSheet | null = null;
    const spreadsheet = {
      getSheets: vi.fn(() => sheets),
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sheet3");
        return targetSheet;
      }),
      setActiveSheet: vi.fn((sheet: typeof targetSheet) => {
        activeSheet = sheet;
      }),
      moveActiveSheet: vi.fn((position: number) => {
        if (!activeSheet) {
          throw new Error("No active sheet");
        }

        const currentIndex = sheets.indexOf(activeSheet);
        sheets.splice(currentIndex, 1);
        sheets.splice(position - 1, 0, activeSheet);
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "move_sheet",
      from: {
        exists: true,
        name: "Sheet3",
        position: 0
      },
      to: {
        exists: true,
        name: "Sheet3",
        position: 2
      }
    })).toMatchObject({
      ok: true,
      operation: "move_sheet"
    });

    expect(sheets.map((sheet) => sheet.getName())).toEqual(["Sheet1", "Sheet4", "Sheet3"]);
    expect(spreadsheet.moveActiveSheet).toHaveBeenCalledWith(3);
    expect(code.flush).toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Google Sheets sheet create writes", () => {
    const existingSheet = {
      getName: vi.fn(() => "Sheet1"),
      getIndex: vi.fn(() => sheets.indexOf(existingSheet) + 1)
    };
    const createdSheet = {
      getName: vi.fn(() => "Staging"),
      getIndex: vi.fn(() => sheets.indexOf(createdSheet) + 1)
    };
    const sheets = [existingSheet];
    const spreadsheet = {
      getSheets: vi.fn(() => sheets),
      insertSheet: vi.fn((name: string, position: number) => {
        expect(name).toBe("Staging");
        sheets.splice(position, 0, createdSheet);
        return createdSheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_create_sheet_snapshot_sheets_001",
      runId: "run_create_sheet_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_create_sheet_snapshot_sheets_001",
      plan: {
        operation: "create_sheet",
        sheetName: "Staging",
        position: 0,
        explanation: "Create a staging sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(sheets.map((sheet) => sheet.getName())).toEqual(["Staging", "Sheet1"]);
    expect(spreadsheet.insertSheet).toHaveBeenCalledWith("Staging", 0);
    expect(result).toMatchObject({
      kind: "workbook_structure_update",
      operation: "create_sheet",
      sheetName: "Staging",
      positionResolved: 0,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_create_sheet_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "create_sheet",
        before: {
          exists: false,
          name: "Staging"
        },
        after: {
          exists: true,
          name: "Staging",
          position: 0
        }
      }
    });
  });

  it("attaches local undo snapshots for Google Sheets duplicate sheet writes", () => {
    let copiedSheetName = "Copy of Template";
    const sourceSheet = {
      getName: vi.fn(() => "Template"),
      getIndex: vi.fn(() => sheets.indexOf(sourceSheet) + 1),
      copyTo: vi.fn(() => {
        sheets.push(copiedSheet);
        return copiedSheet;
      })
    };
    const otherSheet = {
      getName: vi.fn(() => "Summary"),
      getIndex: vi.fn(() => sheets.indexOf(otherSheet) + 1)
    };
    const copiedSheet = {
      getName: vi.fn(() => copiedSheetName),
      getIndex: vi.fn(() => sheets.indexOf(copiedSheet) + 1),
      setName: vi.fn((nextName: string) => {
        copiedSheetName = nextName;
      })
    };
    const sheets = [sourceSheet, otherSheet];
    let activeSheet: typeof copiedSheet | null = null;
    const spreadsheet = {
      getSheets: vi.fn(() => sheets),
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Template");
        return sourceSheet;
      }),
      setActiveSheet: vi.fn((sheet: typeof copiedSheet) => {
        activeSheet = sheet;
      }),
      moveActiveSheet: vi.fn((position: number) => {
        if (!activeSheet) {
          throw new Error("No active sheet");
        }

        const currentIndex = sheets.indexOf(activeSheet);
        sheets.splice(currentIndex, 1);
        sheets.splice(position - 1, 0, activeSheet);
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_duplicate_sheet_snapshot_sheets_001",
      runId: "run_duplicate_sheet_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_duplicate_sheet_snapshot_sheets_001",
      plan: {
        operation: "duplicate_sheet",
        sheetName: "Template",
        newSheetName: "Template Copy",
        position: "end",
        explanation: "Duplicate the template sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(sheets.map((sheet) => sheet.getName())).toEqual(["Template", "Summary", "Template Copy"]);
    expect(spreadsheet.moveActiveSheet).toHaveBeenCalledWith(3);
    expect(result).toMatchObject({
      kind: "workbook_structure_update",
      operation: "duplicate_sheet",
      sheetName: "Template",
      newSheetName: "Template Copy",
      positionResolved: 2,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_duplicate_sheet_snapshot_sheets_001",
        kind: "workbook_structure",
        operation: "duplicate_sheet",
        source: {
          exists: true,
          name: "Template",
          position: 0
        },
        before: {
          exists: false,
          name: "Template Copy"
        },
        after: {
          exists: true,
          name: "Template Copy",
          position: 2
        }
      }
    });
  });

  it("passes Google Sheets duplicate sheet snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const source = {
      exists: true,
      name: "Template",
      position: 0
    };
    const before = {
      exists: false,
      name: "Template Copy"
    };
    const after = {
      exists: true,
      name: "Template Copy",
      position: 1
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_duplicate_sheet_001"],
      executions: {
        exec_duplicate_sheet_001: {
          baseExecutionId: "exec_duplicate_sheet_001"
        }
      },
      bases: {
        exec_duplicate_sheet_001: {
          baseExecutionId: "exec_duplicate_sheet_001",
          kind: "workbook_structure",
          operation: "duplicate_sheet",
          source,
          before,
          after
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "workbook_structure_update",
          operation: "duplicate_sheet",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_duplicate_sheet_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "workbook_structure",
          operation: "duplicate_sheet",
          source,
          from: after,
          to: before
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "workbook_structure",
          operation: "duplicate_sheet",
          source,
          from: after,
          to: before
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("passes Google Sheets chart snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:C20",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line",
      categoryField: "Month",
      series: [{ field: "Revenue", label: "Revenue" }],
      title: "Revenue",
      legendPosition: "none",
      explanation: "Chart revenue.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    };
    const before = {
      exists: false,
      targetRange: "A1"
    };
    const after = {
      exists: true,
      targetRange: "A1",
      chartId: 202
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_chart_001"],
      executions: {
        exec_chart_001: {
          baseExecutionId: "exec_chart_001"
        }
      },
      bases: {
        exec_chart_001: {
          baseExecutionId: "exec_chart_001",
          kind: "chart",
          targetSheet: "Sales Chart",
          targetRange: "A1",
          chartId: 202,
          before,
          after,
          plan
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "chart_update",
          operation: "chart_update",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_chart_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "chart",
          targetSheet: "Sales Chart",
          targetRange: "A1",
          chartId: 202,
          from: after,
          to: before,
          plan
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "chart",
          targetSheet: "Sales Chart",
          targetRange: "A1",
          chartId: 202,
          from: after,
          to: before,
          plan
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("passes Google Sheets pivot table snapshots through sidebar undo before committing", async () => {
    const backingStore = new Map<string, string>();
    const plan = {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region"],
      columnGroups: [],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      filters: [],
      explanation: "Build a pivot table by region.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    };
    const before = {
      exists: false,
      targetRange: "A1"
    };
    const after = {
      exists: true,
      targetRange: "A1"
    };
    backingStore.set("Hermes.ReversibleExecutions.v1::google_sheets::sheet-123", JSON.stringify({
      version: 1,
      order: ["exec_pivot_001"],
      executions: {
        exec_pivot_001: {
          baseExecutionId: "exec_pivot_001"
        }
      },
      bases: {
        exec_pivot_001: {
          baseExecutionId: "exec_pivot_001",
          kind: "pivot_table",
          targetSheet: "Sales Pivot",
          targetRange: "A1",
          before,
          after,
          plan
        }
      }
    }));
    const sidebar = loadSidebarContext();
    sidebar.window.localStorage.getItem = vi.fn((key: string) => backingStore.get(key) || null);
    sidebar.window.localStorage.setItem = vi.fn((key: string, value: string) => {
      backingStore.set(key, value);
    });
    const serverCalls: Array<{ functionName: string; payload?: Record<string, unknown> }> = [];
    sidebar.callServer = vi.fn(async (functionName: string, payload?: Record<string, unknown>) => {
      serverCalls.push({ functionName, payload });
      if (functionName === "getWorkbookSessionKey") {
        return "google_sheets::sheet-123";
      }

      if (functionName === "getRuntimeConfig") {
        return {
          gatewayBaseUrl: "http://127.0.0.1:8787",
          clientVersion: "google-sheets-addon-dev",
          reviewerSafeMode: false,
          forceExtractionMode: null
        };
      }

      if (functionName === "validateExecutionCellSnapshot" || functionName === "applyExecutionCellSnapshot") {
        return payload;
      }

      throw new Error(`Unexpected server call: ${functionName}`);
    });
    sidebar.fetch = vi.fn(async (url: string, init?: { body?: string }) => ({
      ok: true,
      async json() {
        return {
          kind: "pivot_table_update",
          operation: "pivot_table_update",
          hostPlatform: "google_sheets",
          executionId: String(url).endsWith("/api/execution/undo") ? "exec_undo_001" : "exec_undo_preview_001",
          summary: init?.body || ""
        };
      }
    }));

    await sidebar.undoExecution("exec_pivot_001");

    const snapshotServerCalls = serverCalls.filter((call) => call.functionName !== "getRuntimeConfig");
    expect(snapshotServerCalls).toEqual([
      { functionName: "getWorkbookSessionKey", payload: undefined },
      {
        functionName: "validateExecutionCellSnapshot",
        payload: {
          kind: "pivot_table",
          targetSheet: "Sales Pivot",
          targetRange: "A1",
          from: after,
          to: before,
          plan
        }
      },
      {
        functionName: "applyExecutionCellSnapshot",
        payload: {
          kind: "pivot_table",
          targetSheet: "Sales Pivot",
          targetRange: "A1",
          from: after,
          to: before,
          plan
        }
      }
    ]);
    expect(sidebar.fetch).toHaveBeenCalledTimes(2);
  });

  it("applies Google Sheets sheet create snapshots on the server", () => {
    const createdSheet = {
      getName: vi.fn(() => "Staging"),
      getIndex: vi.fn(() => sheets.indexOf(createdSheet) + 1)
    };
    const otherSheet = {
      getName: vi.fn(() => "Sheet1"),
      getIndex: vi.fn(() => sheets.indexOf(otherSheet) + 1)
    };
    const sheets = [createdSheet, otherSheet];
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Staging");
        return createdSheet;
      }),
      deleteSheet: vi.fn((sheet: typeof createdSheet) => {
        const index = sheets.indexOf(sheet);
        sheets.splice(index, 1);
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "create_sheet",
      from: {
        exists: true,
        name: "Staging",
        position: 0
      },
      to: {
        exists: false,
        name: "Staging"
      }
    })).toMatchObject({
      ok: true,
      operation: "create_sheet"
    });

    expect(sheets.map((sheet) => sheet.getName())).toEqual(["Sheet1"]);
    expect(spreadsheet.deleteSheet).toHaveBeenCalledWith(createdSheet);
    expect(code.flush).toHaveBeenCalled();
  });

  it("applies Google Sheets duplicate sheet snapshots on the server", () => {
    const sourceSheet = {
      getName: vi.fn(() => "Template"),
      getIndex: vi.fn(() => sheets.indexOf(sourceSheet) + 1)
    };
    const copiedSheet = {
      getName: vi.fn(() => "Template Copy"),
      getIndex: vi.fn(() => sheets.indexOf(copiedSheet) + 1)
    };
    const otherSheet = {
      getName: vi.fn(() => "Summary"),
      getIndex: vi.fn(() => sheets.indexOf(otherSheet) + 1)
    };
    const sheets = [sourceSheet, copiedSheet, otherSheet];
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Template Copy");
        return copiedSheet;
      }),
      deleteSheet: vi.fn((sheet: typeof copiedSheet) => {
        const index = sheets.indexOf(sheet);
        sheets.splice(index, 1);
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyExecutionCellSnapshot({
      kind: "workbook_structure",
      operation: "duplicate_sheet",
      source: {
        exists: true,
        name: "Template",
        position: 0
      },
      from: {
        exists: true,
        name: "Template Copy",
        position: 1
      },
      to: {
        exists: false,
        name: "Template Copy"
      }
    })).toMatchObject({
      ok: true,
      operation: "duplicate_sheet"
    });

    expect(sheets.map((sheet) => sheet.getName())).toEqual(["Template", "Summary"]);
    expect(spreadsheet.deleteSheet).toHaveBeenCalledWith(copiedSheet);
    expect(code.flush).toHaveBeenCalled();
  });

  it("applies Google Sheets chart snapshots on the server", () => {
    const chart = {
      getChartId: vi.fn(() => 202)
    };
    const charts = [chart];
    const chartSheet = {
      getCharts: vi.fn(() => [...charts]),
      removeChart: vi.fn((removedChart: typeof chart) => {
        const index = charts.indexOf(removedChart);
        if (index >= 0) {
          charts.splice(index, 1);
        }
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((name: string) => {
        expect(name).toBe("Sales Chart");
        return chartSheet;
      })
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.validateExecutionCellSnapshot({
      kind: "chart",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartId: 202,
      from: {
        exists: true,
        targetRange: "A1",
        chartId: 202
      },
      to: {
        exists: false,
        targetRange: "A1"
      },
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line"
      }
    })).toMatchObject({
      ok: true,
      targetSheet: "Sales Chart",
      targetRange: "A1"
    });

    expect(code.applyExecutionCellSnapshot({
      kind: "chart",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartId: 202,
      from: {
        exists: true,
        targetRange: "A1",
        chartId: 202
      },
      to: {
        exists: false,
        targetRange: "A1"
      },
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line"
      }
    })).toMatchObject({
      ok: true,
      targetSheet: "Sales Chart",
      targetRange: "A1"
    });

    expect(chartSheet.removeChart).toHaveBeenCalledWith(chart);
    expect(charts).toEqual([]);
    expect(code.flush).toHaveBeenCalled();
  });

  it("applies external data plans by anchoring a first-class formula into a single Google Sheets cell", () => {
    const targetRange = createRangeStub({
      a1Notation: "B2",
      row: 2,
      column: 2,
      numRows: 1,
      numColumns: 1,
      values: [["previous price"]],
      formulas: [[""]]
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
      executionId: "exec_external_data_apply_001",
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
      formula: '=GOOGLEFINANCE("CURRENCY:BTCUSD","price")',
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_external_data_apply_001",
        targetSheet: "Market Data",
        targetRange: "B2",
        beforeCells: [[
          {
            kind: "value",
            value: {
              type: "string",
              value: "previous price"
            }
          }
        ]],
        afterCells: [[
          {
            kind: "formula",
            formula: '=GOOGLEFINANCE("CURRENCY:BTCUSD","price")'
          }
        ]]
      }
    });
    expect(targetRange.setFormula).toHaveBeenCalledWith('=GOOGLEFINANCE("CURRENCY:BTCUSD","price")');
    expect(targetRange.getFormulas()).toEqual([['=GOOGLEFINANCE("CURRENCY:BTCUSD","price")']]);
    expect(code.flush).toHaveBeenCalled();
  });

  it("fails closed when an external data formula does not persist on the target cell", () => {
    const targetRange = createRangeStub({
      a1Notation: "B2",
      row: 2,
      column: 2,
      numRows: 1,
      numColumns: 1
    });
    targetRange.setFormula = vi.fn();
    targetRange.getFormulas = vi.fn(() => [[""]]);
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

    expect(() => code.applyWritePlan({
      requestId: "req_external_data_apply_mismatch",
      runId: "run_external_data_apply_mismatch",
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
    })).toThrow("Google Sheets host could not verify the applied external data formula.");
    expect(targetRange.setFormula).toHaveBeenCalledWith('=GOOGLEFINANCE("CURRENCY:BTCUSD","price")');
    expect(code.flush).toHaveBeenCalled();
  });

  it("fails closed before applying market data formulas that do not match the query", () => {
    const targetRange = createRangeStub({
      a1Notation: "B2",
      row: 2,
      column: 2,
      numRows: 1,
      numColumns: 1,
      values: [["previous price"]],
      formulas: [[""]]
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

    expect(() => code.applyWritePlan({
      requestId: "req_external_data_apply_market_query_mismatch",
      runId: "run_external_data_apply_market_query_mismatch",
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
        formula: '=GOOGLEFINANCE("CURRENCY:ETHUSD","price")',
        explanation: "Anchor the latest BTC price in B2.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Market Data!B2"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toThrow("Google Sheets market data formula must match query.symbol.");

    expect(targetRange.setFormula).not.toHaveBeenCalled();
  });

  it("fails closed before applying market data formulas with single-quoted query arguments", () => {
    const targetRange = createRangeStub({
      a1Notation: "B2",
      row: 2,
      column: 2,
      numRows: 1,
      numColumns: 1,
      values: [["previous price"]],
      formulas: [[""]]
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

    expect(() => code.applyWritePlan({
      requestId: "req_external_data_apply_market_single_quotes",
      runId: "run_external_data_apply_market_single_quotes",
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
        formula: "=GOOGLEFINANCE('CURRENCY:BTCUSD','price')",
        explanation: "Anchor the latest BTC price in B2.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Market Data!B2"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toThrow("Google Sheets market data formula must match query.symbol.");

    expect(targetRange.setFormula).not.toHaveBeenCalled();
  });

  it("fails closed before applying an external data formula with a mismatched source URL", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1
    });
    targetRange.setFormula = vi.fn();
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1");
        return targetRange;
      })
    };
    const spreadsheet = {
      getId() {
        return "sheet-123";
      },
      getSheetByName(name: string) {
        if (name === "Imports") {
          return sheet;
        }
        return null;
      }
    };

    const code = loadCodeModule({ spreadsheet });

    expect(() => code.applyWritePlan({
      requestId: "req_external_data_apply_source_mismatch",
      runId: "run_external_data_apply_source_mismatch",
      approvalToken: "token",
      plan: {
        sourceType: "web_table_import",
        provider: "importdata",
        sourceUrl: "https://example.com/data.csv",
        selectorType: "direct",
        targetSheet: "Imports",
        targetRange: "A1",
        formula: '=IMPORTDATA("https://other.example/data.csv")',
        explanation: "Import a public CSV into the sheet.",
        confidence: 0.86,
        requiresConfirmation: true,
        affectedRanges: ["Imports!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toThrow("Google Sheets external data formulas must reference sourceUrl.");

    expect(targetRange.setFormula).not.toHaveBeenCalled();
  });

  it("fails closed before applying external data formulas with private URL aliases", () => {
    const sourceUrls = [
      "http://[fc00::1]/data.csv",
      "http://[fe80::1]/data.csv",
      "http://[::ffff:127.0.0.1]/data.csv",
      "http://2130706433/data.csv",
      "http://0177.0.0.1/data.csv",
      "http://0x7f.0.0.1/data.csv",
      "http://2852039166/data.csv"
    ];

    sourceUrls.forEach((sourceUrl) => {
      const targetRange = createRangeStub({
        a1Notation: "A1",
        row: 1,
        column: 1,
        numRows: 1,
        numColumns: 1
      });
      targetRange.setFormula = vi.fn();
      const sheet = {
        getRange: vi.fn((rangeA1: string) => {
          expect(rangeA1).toBe("A1");
          return targetRange;
        })
      };
      const spreadsheet = {
        getId() {
          return "sheet-123";
        },
        getSheetByName(name: string) {
          if (name === "Imports") {
            return sheet;
          }
          return null;
        }
      };

      const code = loadCodeModule({ spreadsheet });

      expect(() => code.applyWritePlan({
        requestId: "req_external_data_apply_private_alias",
        runId: "run_external_data_apply_private_alias",
        approvalToken: "token",
        plan: {
          sourceType: "web_table_import",
          provider: "importdata",
          sourceUrl,
          selectorType: "direct",
          targetSheet: "Imports",
          targetRange: "A1",
          formula: `=IMPORTDATA("${sourceUrl}")`,
          explanation: "Import a public CSV into the sheet.",
          confidence: 0.86,
          requiresConfirmation: true,
          affectedRanges: ["Imports!A1"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      })).toThrow("Google Sheets external data sourceUrl must be a public HTTP(S) URL.");

      expect(targetRange.setFormula).not.toHaveBeenCalled();
    });
  });

  it("applies Google Sheets table-like plans with banding and filters", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6
    });
    const bandings: Array<Record<string, unknown>> = [];
    const rowBanding = {
      getRange: vi.fn(() => targetRange),
      remove: vi.fn(() => {
        bandings.splice(0, bandings.length);
      })
    };
    targetRange.getBandings = vi.fn(() => [...bandings]);
    targetRange.applyRowBanding = vi.fn(() => {
      bandings.push(rowBanding);
      return rowBanding;
    });
    let currentFilter: Record<string, unknown> | null = null;
    const filter = {
      getRange: vi.fn(() => targetRange),
      getColumnFilterCriteria: vi.fn(() => null)
    };
    targetRange.createFilter = vi.fn(() => {
      currentFilter = filter;
      return filter;
    });
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:F50");
        return targetRange;
      }),
      getFilter: vi.fn(() => currentFilter)
    };
    const spreadsheet = {
      getNamedRanges: vi.fn(() => []),
      setNamedRange: vi.fn(),
      getId() {
        return "sheet-123";
      },
      getSheetByName(name: string) {
        if (name === "Sales") {
          return sheet;
        }
        return null;
      }
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_table_apply_google_001",
      runId: "run_table_apply_google_001",
      approvalToken: "token",
      executionId: "exec_table_apply_google_001",
      plan: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        showBandedRows: true,
        showBandedColumns: false,
        showFilterButton: true,
        showTotalsRow: false,
        explanation: "Format the sales range as a filterable table.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toMatchObject({
      kind: "table_update",
      operation: "table_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      name: "SalesTable",
      hasHeaders: true,
      summary: "Formatted table-like range SalesTable on Sales!A1:F50."
    });
    expect(result).toMatchObject({
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_table_apply_google_001",
        entries: [
          {
            baseExecutionId: "exec_table_apply_google_001",
            kind: "named_range",
            scope: "workbook",
            before: {
              exists: false,
              name: "SalesTable"
            },
            after: {
              exists: true,
              name: "SalesTable",
              targetSheet: "Sales",
              targetRange: "A1:F50"
            }
          },
          {
            baseExecutionId: "exec_table_apply_google_001",
            kind: "range_filter",
            targetSheet: "Sales",
            targetRange: "A1:F50",
            beforeFilter: {
              exists: false
            },
            afterFilter: {
              exists: true,
              targetRange: "A1:F50",
              criteria: [null, null, null, null, null, null]
            }
          },
          {
            baseExecutionId: "exec_table_apply_google_001",
            kind: "range_banding",
            targetSheet: "Sales",
            targetRange: "A1:F50",
            beforeBanding: {
              bandings: []
            },
            afterBanding: {
              bandings: [
                {
                  axis: "rows",
                  targetRange: "A1:F50"
                }
              ]
            }
          }
        ]
      }
    });
    expect(targetRange.applyRowBanding).toHaveBeenCalledTimes(1);
    expect(targetRange.createFilter).toHaveBeenCalledTimes(1);
    expect(spreadsheet.setNamedRange).toHaveBeenCalledWith("SalesTable", targetRange);
    expect(code.flush).toHaveBeenCalled();
  });

  it("suppresses Google Sheets table-like local snapshots when banding already exists", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6
    });
    const bandings: Array<Record<string, unknown>> = [
      {
        getRange: vi.fn(() => targetRange),
        remove: vi.fn()
      }
    ];
    const appliedBanding = {
      getRange: vi.fn(() => targetRange),
      remove: vi.fn()
    };
    targetRange.getBandings = vi.fn(() => [...bandings]);
    targetRange.applyRowBanding = vi.fn(() => {
      bandings.push(appliedBanding);
      return appliedBanding;
    });
    let currentFilter: Record<string, unknown> | null = null;
    const filter = {
      getRange: vi.fn(() => targetRange),
      getColumnFilterCriteria: vi.fn(() => null)
    };
    targetRange.createFilter = vi.fn(() => {
      currentFilter = filter;
      return filter;
    });
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:F50");
        return targetRange;
      }),
      getFilter: vi.fn(() => currentFilter)
    };
    const spreadsheet = {
      getNamedRanges: vi.fn(() => []),
      setNamedRange: vi.fn(),
      getId() {
        return "sheet-123";
      },
      getSheetByName(name: string) {
        if (name === "Sales") {
          return sheet;
        }
        return null;
      }
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_table_apply_google_existing_banding_001",
      runId: "run_table_apply_google_existing_banding_001",
      approvalToken: "token",
      executionId: "exec_table_apply_google_existing_banding_001",
      plan: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        showBandedRows: true,
        showFilterButton: true,
        explanation: "Format the sales range as a filterable table.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toMatchObject({
      kind: "table_update",
      operation: "table_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      name: "SalesTable",
      hasHeaders: true
    });
    expect(result.__hermesLocalExecutionSnapshot).toBeUndefined();
    expect(targetRange.applyRowBanding).toHaveBeenCalledTimes(1);
    expect(targetRange.createFilter).toHaveBeenCalledTimes(1);
    expect(spreadsheet.setNamedRange).toHaveBeenCalledWith("SalesTable", targetRange);
    expect(code.flush).toHaveBeenCalled();
  });

  it("applies Google Sheets table banding snapshots on the server", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6
    });
    const bandings: Array<Record<string, unknown>> = [];
    const createBanding = (axis: "rows" | "columns") => {
      const banding = {
        __hermesAxis: axis,
        getRange: vi.fn(() => targetRange),
        remove: vi.fn(() => {
          const index = bandings.indexOf(banding);
          if (index >= 0) {
            bandings.splice(index, 1);
          }
        })
      };
      return banding;
    };
    const initialBanding = createBanding("rows");
    bandings.push(initialBanding);
    targetRange.getBandings = vi.fn(() => [...bandings]);
    targetRange.applyRowBanding = vi.fn(() => {
      const banding = createBanding("rows");
      bandings.push(banding);
      return banding;
    });
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:F50");
        return targetRange;
      })
    };
    const spreadsheet = {
      getId() {
        return "sheet-123";
      },
      getSheetByName(name: string) {
        if (name === "Sales") {
          return sheet;
        }
        return null;
      }
    };
    const code = loadCodeModule({ spreadsheet });
    const rowBandingState = {
      bandings: [
        {
          axis: "rows",
          targetRange: "A1:F50"
        }
      ]
    };
    const emptyBandingState = {
      bandings: []
    };

    expect(code.validateExecutionCellSnapshot({
      kind: "range_banding",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      from: rowBandingState,
      to: emptyBandingState
    })).toEqual({
      ok: true,
      targetSheet: "Sales",
      targetRange: "A1:F50"
    });

    expect(code.applyExecutionCellSnapshot({
      kind: "range_banding",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      from: rowBandingState,
      to: emptyBandingState
    })).toEqual({
      ok: true,
      targetSheet: "Sales",
      targetRange: "A1:F50"
    });
    expect(initialBanding.remove).toHaveBeenCalledTimes(1);
    expect(bandings).toHaveLength(0);

    expect(code.applyExecutionCellSnapshot({
      kind: "range_banding",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      from: emptyBandingState,
      to: rowBandingState
    })).toEqual({
      ok: true,
      targetSheet: "Sales",
      targetRange: "A1:F50"
    });
    expect(targetRange.applyRowBanding).toHaveBeenCalledTimes(1);
    expect(bandings).toHaveLength(1);
  });

  it("leaves Google Sheets table-like optional formatting off when flags are omitted", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6
    });
    targetRange.applyRowBanding = undefined;
    targetRange.applyColumnBanding = undefined;
    targetRange.createFilter = undefined;
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:F50");
        return targetRange;
      })
    };
    const spreadsheet = {
      getId() {
        return "sheet-123";
      },
      setNamedRange: vi.fn(),
      getSheetByName(name: string) {
        if (name === "Sales") {
          return sheet;
        }
        return null;
      }
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyWritePlan({
      requestId: "req_table_apply_google_002",
      runId: "run_table_apply_google_002",
      approvalToken: "token",
      executionId: "exec_table_apply_google_002",
      plan: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        explanation: "Create a table-like range without additional formatting.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toMatchObject({
      kind: "table_update",
      operation: "table_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      name: "SalesTable",
      hasHeaders: true,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_table_apply_google_002",
        kind: "named_range",
        scope: "workbook",
        before: {
          exists: false,
          name: "SalesTable"
        },
        after: {
          exists: true,
          name: "SalesTable",
          targetSheet: "Sales",
          targetRange: "A1:F50"
        }
      }
    });
    expect(spreadsheet.setNamedRange).toHaveBeenCalledWith("SalesTable", targetRange);
    expect(code.flush).toHaveBeenCalled();
  });

  it("attaches composite undo snapshots for Google Sheets table-like range names and filters", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6
    });
    let currentFilter: Record<string, unknown> | null = null;
    const filter = {
      getRange: vi.fn(() => targetRange),
      getColumnFilterCriteria: vi.fn(() => null)
    };
    targetRange.createFilter = vi.fn(() => {
      currentFilter = filter;
      return filter;
    });
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:F50");
        return targetRange;
      }),
      getFilter: vi.fn(() => currentFilter)
    };
    const spreadsheet = {
      getNamedRanges: vi.fn(() => []),
      getId() {
        return "sheet-123";
      },
      setNamedRange: vi.fn(),
      getSheetByName(name: string) {
        if (name === "Sales") {
          return sheet;
        }
        return null;
      }
    };
    const code = loadCodeModule({ spreadsheet });

    expect(code.applyWritePlan({
      requestId: "req_table_apply_google_filter_snapshot",
      runId: "run_table_apply_google_filter_snapshot",
      approvalToken: "token",
      executionId: "exec_table_apply_google_filter_snapshot",
      plan: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        showFilterButton: true,
        explanation: "Create a named table-like range with a header filter.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toMatchObject({
      kind: "table_update",
      operation: "table_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales",
      targetRange: "A1:F50",
      name: "SalesTable",
      hasHeaders: true,
      showFilterButton: true,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_table_apply_google_filter_snapshot",
        entries: [
          {
            baseExecutionId: "exec_table_apply_google_filter_snapshot",
            kind: "named_range",
            scope: "workbook",
            before: {
              exists: false,
              name: "SalesTable"
            },
            after: {
              exists: true,
              name: "SalesTable",
              targetSheet: "Sales",
              targetRange: "A1:F50"
            }
          },
          {
            baseExecutionId: "exec_table_apply_google_filter_snapshot",
            kind: "range_filter",
            targetSheet: "Sales",
            targetRange: "A1:F50",
            beforeFilter: {
              exists: false
            },
            afterFilter: {
              exists: true,
              targetRange: "A1:F50",
              criteria: [null, null, null, null, null, null]
            }
          }
        ]
      }
    });
    expect(targetRange.createFilter).toHaveBeenCalledTimes(1);
    expect(spreadsheet.setNamedRange).toHaveBeenCalledWith("SalesTable", targetRange);
    expect(code.flush).toHaveBeenCalled();
  });

  it("does not name Google Sheets table-like ranges when required formatting is unsupported", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6
    });
    targetRange.applyRowBanding = undefined;
    const sheet = {
      getRange: vi.fn((rangeA1: string) => {
        expect(rangeA1).toBe("A1:F50");
        return targetRange;
      })
    };
    const spreadsheet = {
      setNamedRange: vi.fn(),
      getSheetByName(name: string) {
        if (name === "Sales") {
          return sheet;
        }
        return null;
      }
    };
    const code = loadCodeModule({ spreadsheet });

    expect(() => code.applyWritePlan({
      requestId: "req_table_apply_google_unsupported_name_001",
      runId: "run_table_apply_google_unsupported_name_001",
      approvalToken: "token",
      executionId: "exec_table_apply_google_unsupported_name_001",
      plan: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        showBandedRows: true,
        explanation: "Format the sales range as a filterable table.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toThrow("Google Sheets host does not support exact table-like row banding.");

    expect(spreadsheet.setNamedRange).not.toHaveBeenCalled();
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

  it("rejects non-api paths and unsafe methods in the Apps Script proxy", () => {
    const fetch = vi.fn(() => {
      throw new Error("Unexpected proxy fetch");
    });
    const code = loadCodeModule({
      scriptProperties: {
        HERMES_GATEWAY_URL: "https://gateway.test"
      },
      urlFetchApp: { fetch }
    });

    expect(() => code.proxyGatewayJson({
      path: "/health"
    })).toThrow("Hermes gateway proxy only supports /api/ routes.");

    expect(() => code.proxyGatewayJson({
      path: "/api/requests",
      method: "delete"
    })).toThrow("Hermes gateway proxy only supports GET and POST.");

    expect(fetch).not.toHaveBeenCalled();
  });

  it("fails closed before proxying when the Apps Script gateway URL is private", () => {
    const fetch = vi.fn(() => {
      throw new Error("Unexpected proxy fetch");
    });
    const code = loadCodeModule({
      scriptProperties: {
        HERMES_GATEWAY_URL: "https://gateway.local:8787"
      },
      urlFetchApp: { fetch }
    });

    expect(() => code.proxyGatewayJson({
      path: "/api/requests",
      method: "post",
      body: { requestId: "req_private_proxy_001" }
    })).toThrow("Hermes gateway URL is not configured.");

    expect(fetch).not.toHaveBeenCalled();
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

  it("sanitizes raw text gateway failures before display in Google Sheets", async () => {
    const code = loadCodeModule();
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    expect(code.extractGatewayErrorMessage(
      500,
      "ReferenceError at /srv/hermes/services/gateway/src/app.ts:99 HERMES_API_SERVER_KEY=secret_123"
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      "Gateway failed for qa_HERMES_API_SERVER_KEY"
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      String.raw`Gateway failed at \\runner\share\server.ts:42`
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      String.raw`Gateway failed path=\\runner\share\server.ts:42`
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      String.raw`Gateway failed at ("C:\Users\runner\work\server.ts:42")`
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      String.raw`Gateway failed at ("\\runner\share\server.ts:42")`
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      String.raw`Gateway failed at ("/srv/hermes/server.ts:42")`
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      "Gateway failed while fetching http://169.254.169.254/latest/meta-data"
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      "Gateway failed while fetching http://[::ffff:7f00:1]/latest/meta-data"
    )).toBe("Hermes gateway request failed with 500.");

    expect(code.extractGatewayErrorMessage(
      500,
      "Gateway failed while fetching http://[::1]/debug"
    )).toBe("Hermes gateway request failed with 500.");

    await expect(hooks.parseGatewayJsonResponse({
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

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return "Gateway failed while fetching http://[::1]/debug";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return "Gateway failed while fetching http://169.254.169.254/latest/meta-data";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return "Gateway failed while fetching http://[::ffff:7f00:1]/latest/meta-data";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
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

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return String.raw`Gateway failed at \\runner\share\server.ts:42`;
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return String.raw`Gateway failed path=\\runner\share\server.ts:42`;
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return String.raw`Gateway failed at ("C:\Users\runner\work\server.ts:42")`;
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return String.raw`Gateway failed at ("\\runner\share\server.ts:42")`;
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        throw new Error("not json");
      },
      async text() {
        return String.raw`Gateway failed at ("/srv/hermes/server.ts:42")`;
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    let htmlError: Error | undefined;
    try {
      await hooks.parseGatewayJsonResponse({
        ok: false,
        status: 502,
        url: "http://127.0.0.1:8787/internal/debug",
        async json() {
          throw new Error("not json");
        },
        async text() {
          return "<html><body>not gateway</body></html>";
        }
      });
    } catch (error) {
      htmlError = error as Error;
    }
    expect(htmlError?.message).toContain("the configured gateway, HTTP 502");
    expect(htmlError?.message).not.toContain("127.0.0.1");
    expect(htmlError?.message).not.toContain("/internal/debug");
  });

  it("sanitizes JSON gateway failures before display in Google Sheets", async () => {
    const code = loadCodeModule();
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;
    const bodyText = JSON.stringify({
      error: {
        message: "ReferenceError at /srv/hermes/services/gateway/src/app.ts:99 HERMES_API_SERVER_KEY=secret_123",
        userAction: "Inspect https://internal.example/debug for stack trace details."
      }
    });

    expect(code.extractGatewayErrorMessage(500, bodyText)).toBe(
      "Hermes gateway request failed with 500."
    );

    const uncBodyText = JSON.stringify({
      error: {
        message: "Gateway failed while formatting diagnostics.",
        userAction: String.raw`Inspect \\runner\share\debug.log before retrying.`
      }
    });

    expect(code.extractGatewayErrorMessage(500, uncBodyText)).toBe(
      "Hermes gateway request failed with 500."
    );

    const labeledUncBodyText = JSON.stringify({
      error: {
        message: "Gateway failed while formatting diagnostics.",
        userAction: String.raw`Inspect source=\\runner\share\debug.log before retrying.`
      }
    });

    expect(code.extractGatewayErrorMessage(500, labeledUncBodyText)).toBe(
      "Hermes gateway request failed with 500."
    );

    const wrappedPosixBodyText = JSON.stringify({
      error: {
        message: "Gateway failed while formatting diagnostics.",
        userAction: String.raw`Inspect ("/srv/hermes/debug.log") before retrying.`
      }
    });

    expect(code.extractGatewayErrorMessage(500, wrappedPosixBodyText)).toBe(
      "Hermes gateway request failed with 500."
    );

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        return JSON.parse(bodyText);
      },
      async text() {
        return "";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        return JSON.parse(wrappedPosixBodyText);
      },
      async text() {
        return "";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        return JSON.parse(uncBodyText);
      },
      async text() {
        return "";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");

    await expect(hooks.parseGatewayJsonResponse({
      ok: false,
      status: 500,
      url: "https://gateway.test/api/requests",
      async json() {
        return JSON.parse(labeledUncBodyText);
      },
      async text() {
        return "";
      }
    })).rejects.toThrow("Hermes gateway request failed with HTTP 500.");
  });

  it("sanitizes malformed successful gateway JSON before display in Google Sheets", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;

    await expect(hooks.parseGatewayJsonResponse({
      ok: true,
      status: 200,
      async json() {
        throw new Error("Unexpected token at /srv/hermes/services/gateway/src/app.ts:99");
      }
    })).rejects.toThrow(
      "The Hermes service returned a response Google Sheets could not use.\n\n" +
      "Retry the request, then reload the sidebar if it keeps happening."
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
      getAnchorCell: vi.fn(() => pivotAnchorRange),
      remove: vi.fn(),
      addRowGroup: vi.fn(),
      addColumnGroup: vi.fn(),
      addPivotValue: vi.fn(),
      addFilter: vi.fn()
    };
    const pivotTables: Array<Record<string, unknown>> = [];
    pivotAnchorRange.createPivotTable = vi.fn(() => {
      pivotTables.push(pivotTable);
      return pivotTable;
    });
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
            getPivotTables: vi.fn(() => pivotTables),
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

  it("attaches a composite undo snapshot when every Google Sheets step captures a local snapshot", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:B1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 2,
      values: [["", ""]],
      formulas: [["", ""]]
    });
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("A1:B1");
            return targetRange;
          })
        };
      })
    };
    const code = loadCodeModule({ spreadsheet });

    const result = code.applyWritePlan({
      requestId: "req_composite_snapshot_sheets_001",
      runId: "run_composite_snapshot_sheets_001",
      approvalToken: "token",
      executionId: "exec_composite_snapshot_sheets_001",
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
      }
    });

    expect(result).toMatchObject({
      kind: "composite_update",
      undoReady: true,
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_composite_snapshot_sheets_001",
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
            planDigest: "plan-digest",
            executionId: "exec_confirm_sheet_001"
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
            planDigest: "plan-digest",
            executionId: "exec_delete_foo_001"
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

  it("keeps Google Sheets writebacks pending when completion returns malformed success JSON", async () => {
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
        runId: "run_malformed_ack_001",
        requestId: "req_malformed_ack_001",
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
        statusLine: ""
      }
    ];

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return {
          ok: true,
          json: async () => ({
            approvalToken: "approval-token",
            planDigest: "plan-digest",
            executionId: "exec_malformed_ack_001"
          })
        };
      }

      if (url === "http://gateway.test/api/writeback/complete") {
        return {
          ok: true,
          json: async () => ({ ok: false })
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

    hooks.elements.prompt.value = 'Confirm create sheet "Malformed Ack"';
    await hooks.sendPrompt();

    expect(applyWritePlanSpy).toHaveBeenCalledTimes(1);
    expect(hooks.state.messages[0].statusLine).toBe(
      "Applied locally. Retry confirm to finish syncing Hermes history."
    );
    expect(hooks.state.messages[0].pendingCompletion).toMatchObject({
      requestId: "req_malformed_ack_001",
      runId: "run_malformed_ack_001",
      workbookSessionKey: "google_sheets::sheet-123",
      approvalToken: "approval-token",
      planDigest: "plan-digest"
    });
    expect(hooks.state.messages[0].response).toBeTruthy();
  });

  it("does not execute Google Sheets writes when approval responses omit required tokens", async () => {
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
        runId: "run_missing_token_001",
        requestId: "req_missing_token_001",
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
        statusLine: ""
      }
    ];

    const fetchMock = vi.fn(async (url: string) => {
      if (url === "http://gateway.test/api/writeback/approve") {
        return {
          ok: true,
          json: async () => ({
            planDigest: "plan-digest",
            executionId: "exec_missing_token_001"
          })
        };
      }

      throw new Error(`Unexpected fetch URL: ${url}`);
    });
    sidebar.fetch = fetchMock;

    let successHandler: ((value: unknown) => unknown) | null = null;
    let failureHandler: ((value: unknown) => unknown) | null = null;
    const applyWritePlanSpy = vi.fn(() => {
      failureHandler?.(new Error("applyWritePlan should not be called"));
    });
    sidebar.google.script.run = {
      withSuccessHandler(handler: (value: unknown) => unknown) {
        successHandler = handler;
        return this;
      },
      withFailureHandler(handler: (value: unknown) => unknown) {
        failureHandler = handler;
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

    hooks.elements.prompt.value = 'Confirm create sheet "Missing Token"';
    await hooks.sendPrompt();

    expect(applyWritePlanSpy).not.toHaveBeenCalled();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(hooks.state.messages[0].pendingCompletion).toBeUndefined();
    expect(hooks.state.messages[0].statusLine).toContain("writeback approval response");
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
            planDigest: "plan-digest",
            executionId: "exec_update_sales_001"
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
      },
      {
        role: "assistant",
        runId: "run_format_unsupported_001",
        requestId: "req_format_unsupported_001",
        response: {
          type: "range_format_update",
          data: {
            targetSheet: "Sheet1",
            targetRange: "B2:B20",
            format: {
              underline: true,
              strikethrough: true
            },
            explanation: "Apply text line styling.",
            confidence: 0.75,
            requiresConfirmation: true
          }
        },
        content: "Prepared a formatting update for Sheet1!B2:B20.",
        statusLine: ""
      }
    ];

    expect(sidebar.isWritePlanResponse(hooks.state.messages[0].response)).toBe(false);
    expect(sidebar.isWritePlanResponse(hooks.state.messages[1].response)).toBe(false);
    expect(sidebar.isWritePlanResponse(hooks.state.messages[2].response)).toBe(false);
    expect(sidebar.renderStructuredPreview(hooks.state.messages[2].response, {
      runId: "run_format_unsupported_001",
      requestId: "req_format_unsupported_001"
    })).toContain("Google Sheets cannot apply underline and strikethrough together as exact static formatting.");
    expect(sidebar.getLatestPendingWritePlanMessage("ok")).toBeNull();
    expect(sidebar.getLatestPendingWritePlanMessage('Confirm "Sheet1!B2:B20"')).toBeNull();
    expect(sidebar.getLatestPendingWritePlanMessage('Confirm "Contacts!A2:B20"')).toBeNull();
  });
});
