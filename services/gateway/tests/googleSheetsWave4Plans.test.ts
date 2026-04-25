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
} = {}) {
  const flush = options.flush || vi.fn();
  const context = {
    console,
    module: { exports: {} },
    exports: {},
    SpreadsheetApp: {
      getActive() {
        return options.spreadsheet;
      },
      CopyPasteType: {
        PASTE_FORMULA: "PASTE_FORMULA",
        PASTE_FORMAT: "PASTE_FORMAT"
      },
      flush,
      newConditionalFormatRule() {
        throw new Error("not needed in wave 4 tests");
      },
      newDataValidation() {
        throw new Error("not needed in wave 4 tests");
      },
      WrapStrategy: {
        WRAP: "WRAP",
        CLIP: "CLIP",
        OVERFLOW: "OVERFLOW"
      }
    },
    ...require(path.join(REPO_ROOT, "apps/google-sheets-addon/src/Wave1Plans.js"))
  };

  vm.runInNewContext(codeScript, context, { filename: CODE_PATH });

  return {
    ...context.module.exports,
    applyWritePlan: context.applyWritePlan,
    flush
  };
}

function createSidebarElementStub() {
  return {
    innerHTML: "",
    value: "",
    scrollTop: 0,
    scrollHeight: 0,
    addEventListener() {},
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

function loadSidebarContext(options: {
  confirm?: (message: string) => boolean;
} = {}) {
  const scriptWithoutBootstrap = sidebarScript.replace(/\n\s*initialize\(\);\s*$/, "\n");
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
          return null;
        },
        setItem() {},
        removeItem() {}
      },
      confirm: options.confirm || (() => true),
      setInterval,
      clearInterval,
      addEventListener() {}
    },
    crypto: {
      randomUUID() {
        return "test-uuid";
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
          withSuccessHandler() {
            return this;
          },
          withFailureHandler() {
            return this;
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

  vm.runInNewContext(scriptWithoutBootstrap, context, { filename: SIDEBAR_PATH });
  return context;
}

function createRangeStub(options: {
  a1Notation: string;
  row: number;
  column: number;
  numRows: number;
  numColumns: number;
  values?: unknown[][];
  formulas?: (string | null)[][];
}) {
  let currentValues = (options.values || []).map((row) => [...row]);
  let currentFormulas = options.formulas
    ? options.formulas.map((row) => [...row])
    : currentValues.map((row) =>
        row.map((value) => typeof value === "string" && value.startsWith("=") ? value : "")
      );

  const range = {
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
    getValues: vi.fn(() => currentValues.map((row) => [...row])),
    getDisplayValues: vi.fn(() =>
      currentValues.map((row) => row.map((value) => value == null ? "" : String(value)))
    ),
    getFormulas: vi.fn(() =>
      currentFormulas.map((row) => row.map((value) => (value == null ? "" : value)))
    ),
    setValues: vi.fn((nextValues: unknown[][]) => {
      currentValues = nextValues.map((row) => [...row]);
      currentFormulas = nextValues.map((row) => row.map(() => ""));
    }),
    setFormulas: vi.fn((nextFormulas: (string | null)[][]) => {
      currentFormulas = nextFormulas.map((row) => [...row]);
      currentValues = nextFormulas.map((row) => row.map((value) => value == null ? "" : value));
    }),
    clearContent: vi.fn(() => {
      currentValues = Array.from({ length: options.numRows }, () =>
        Array.from({ length: options.numColumns }, () => "")
      );
      currentFormulas = Array.from({ length: options.numRows }, () =>
        Array.from({ length: options.numColumns }, () => "")
      );
    }),
    clearFormat: vi.fn(),
    getCell(rowIndex: number, columnIndex: number) {
      return {
        setValue(value: unknown) {
          currentValues[rowIndex - 1][columnIndex - 1] = value;
          currentFormulas[rowIndex - 1][columnIndex - 1] = "";
        }
      };
    },
    getResizedRange: vi.fn((rowDelta: number, columnDelta: number) =>
      createRangeStub({
        a1Notation: options.a1Notation,
        row: options.row,
        column: options.column,
        numRows: rowDelta + 1,
        numColumns: columnDelta + 1,
        values: Array.from({ length: rowDelta + 1 }, () =>
          Array.from({ length: columnDelta + 1 }, () => "")
        )
      })
    )
  };

  Object.defineProperty(range, "values", {
    configurable: true,
    get() {
      return currentValues;
    }
  });

  Object.defineProperty(range, "formulas", {
    configurable: true,
    get() {
      return currentFormulas;
    }
  });

  return range;
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("Google Sheets wave 4 transfer and cleanup plans", () => {
  it("renders future-tense transfer and cleanup previews", () => {
    const sidebar = loadSidebarContext();

    const transferHtml = sidebar.renderStructuredPreview({
      type: "range_transfer_plan",
      data: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Archive",
        targetRange: "D5:E6",
        operation: "copy",
        pasteMode: "values",
        transpose: false,
        explanation: "Copy finalized rows into the archive block.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Archive!D5:E6"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    }, {
      runId: "run_transfer_preview",
      requestId: "req_transfer_preview"
    });

    const cleanupHtml = sidebar.renderStructuredPreview({
      type: "data_cleanup_plan",
      data: {
        targetSheet: "Contacts",
        targetRange: "A2:B4",
        operation: "trim_whitespace",
        explanation: "Trim imported contact fields in place.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:B4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    }, {
      runId: "run_cleanup_preview",
      requestId: "req_cleanup_preview"
    });

    expect(transferHtml).toContain("Will copy RawData!A2:B3 to Archive!D5:E6.");
    expect(transferHtml).toContain("Confirm Transfer");
    expect(cleanupHtml).toContain("Will trim whitespace in Contacts!A2:B4.");
    expect(cleanupHtml).toContain("Confirm Cleanup");

    const formatTransferHtml = sidebar.renderStructuredPreview({
      type: "range_transfer_plan",
      data: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Archive",
        targetRange: "D5:E6",
        operation: "copy",
        pasteMode: "formats",
        transpose: false,
        explanation: "Copy the source formatting into the archive block.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Archive!D5:E6"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    }, {
      runId: "run_transfer_formats_preview",
      requestId: "req_transfer_formats_preview"
    });
    const titleCleanupHtml = sidebar.renderStructuredPreview({
      type: "data_cleanup_plan",
      data: {
        targetSheet: "Contacts",
        targetRange: "A2:A4",
        operation: "normalize_case",
        mode: "title",
        explanation: "Normalize names into title case.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:A4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    }, {
      runId: "run_cleanup_title_preview",
      requestId: "req_cleanup_title_preview"
    });

    expect(formatTransferHtml).toContain("Confirm Transfer");
    expect(titleCleanupHtml).toContain("Confirm Cleanup");
  });

  it("builds destructive transfer approvals with a second confirmation payload", () => {
    const confirm = vi.fn(() => true);
    const sidebar = loadSidebarContext({ confirm });

    const approvalRequest = sidebar.buildWriteApprovalRequest({
      requestId: "req_transfer_destructive",
      runId: "run_transfer_destructive",
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B20",
        targetSheet: "Archive",
        targetRange: "D5:E23",
        operation: "move",
        pasteMode: "values",
        transpose: false,
        explanation: "Move finalized rows into the archive block.",
        confidence: 0.95,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B20", "Archive!D5:E23"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive"
      }
    });

    expect(confirm).toHaveBeenCalledTimes(1);
    expect(confirm.mock.calls[0]?.[0]).toContain("destructive");
    expect(approvalRequest).toMatchObject({
      requestId: "req_transfer_destructive",
      runId: "run_transfer_destructive",
      destructiveConfirmation: {
        confirmed: true
      },
      plan: {
        operation: "move",
        confirmationLevel: "destructive"
      }
    });
  });

  it("builds destructive cleanup approvals with a second confirmation payload", () => {
    const confirm = vi.fn(() => true);
    const sidebar = loadSidebarContext({ confirm });

    const approvalRequest = sidebar.buildWriteApprovalRequest({
      requestId: "req_cleanup_destructive",
      runId: "run_cleanup_destructive",
      plan: {
        targetSheet: "Contacts",
        targetRange: "A2:F100",
        operation: "remove_blank_rows",
        keyColumns: ["A"],
        explanation: "Remove blank rows before export.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:F100"],
        overwriteRisk: "high",
        confirmationLevel: "destructive"
      }
    });

    expect(confirm).toHaveBeenCalledTimes(1);
    expect(confirm.mock.calls[0]?.[0]).toContain("destructive");
    expect(approvalRequest).toMatchObject({
      requestId: "req_cleanup_destructive",
      runId: "run_cleanup_destructive",
      destructiveConfirmation: {
        confirmed: true
      },
      plan: {
        operation: "remove_blank_rows",
        confirmationLevel: "destructive"
      }
    });
  });

  it("rejects Google Sheets import writes when the destination contains a blank-rendering formula", () => {
    const targetRange = createRangeStub({
      a1Notation: "A1:B2",
      row: 1,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        ["", ""],
        ["", ""]
      ],
      formulas: [
        ['=IF(TRUE,"","")', ""],
        ["", ""]
      ]
    });
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Imported");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("A1:B2");
            return targetRange;
          })
        };
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Imported",
        targetRange: "A1:B2",
        shape: { rows: 2, columns: 2 },
        headers: ["Name", "Amount"],
        values: [["Ada", 42]],
        explanation: "Import a table into the destination.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Imported!A1:B2"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_import_formula_guard_google_001",
      runId: "run_import_formula_guard_google_001",
      approvalToken: "token"
    })).toThrow("Target range already contains content.");
    expect(targetRange.setValues).not.toHaveBeenCalled();
    expect(flush).not.toHaveBeenCalled();
  });

  it("replaces a transfer plan message with the completed transfer summary", () => {
    const sidebar = loadSidebarContext();
    const message = {
      role: "assistant",
      content: "Prepared a transfer plan from RawData!A2:B3 to Archive!D5:E6.",
      response: {
        type: "range_transfer_plan",
        data: {
          sourceSheet: "RawData",
          sourceRange: "A2:B3",
          targetSheet: "Archive",
          targetRange: "D5:E6",
          operation: "copy",
          pasteMode: "values",
          transpose: false,
          explanation: "Copy finalized rows into the archive block.",
          confidence: 0.94,
          requiresConfirmation: true,
          affectedRanges: ["RawData!A2:B3", "Archive!D5:E6"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      }
    };

    sidebar.applyWritebackResultToMessage(message, {
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      hostPlatform: "google_sheets",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Archive",
      targetRange: "D5:E6",
      transferOperation: "copy",
      pasteMode: "values",
      transpose: false,
      summary: "Copied RawData!A2:B3 to Archive!D5:E6."
    });

    expect(message.content).toBe("Copied RawData!A2:B3 to Archive!D5:E6.");
    expect(message.response).toBeNull();
    expect(message.statusLine).toBe("");
  });

  it("replaces a cleanup plan message with the completed cleanup summary", () => {
    const sidebar = loadSidebarContext();
    const message = {
      role: "assistant",
      content: "Prepared a cleanup plan for Contacts!A2:B4.",
      response: {
        type: "data_cleanup_plan",
        data: {
          targetSheet: "Contacts",
          targetRange: "A2:B4",
          operation: "trim_whitespace",
          explanation: "Trim imported contact fields in place.",
          confidence: 0.92,
          requiresConfirmation: true,
          affectedRanges: ["Contacts!A2:B4"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      }
    };

    sidebar.applyWritebackResultToMessage(message, {
      kind: "data_cleanup_update",
      operation: "data_cleanup_update",
      hostPlatform: "google_sheets",
      targetSheet: "Contacts",
      targetRange: "A2:B4",
      cleanupOperation: "trim_whitespace",
      summary: "Trimmed whitespace in Contacts!A2:B4."
    });

    expect(message.content).toBe("Trimmed whitespace in Contacts!A2:B4.");
    expect(message.response).toBeNull();
    expect(message.statusLine).toBe("");
  });

  it("supports exact-safe standardize_format previews and rejects fuzzy patterns", () => {
    const sidebar = loadSidebarContext();

    const supportedHtml = sidebar.renderStructuredPreview({
      type: "data_cleanup_plan",
      data: {
        targetSheet: "Contacts",
        targetRange: "B2:B5",
        operation: "standardize_format",
        formatType: "date_text",
        formatPattern: "YYYY-MM-DD",
        explanation: "Normalize date strings into ISO format.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!B2:B5"],
        overwriteRisk: "medium",
        confirmationLevel: "standard"
      }
    }, {
      runId: "run_cleanup_unsupported_preview",
      requestId: "req_cleanup_unsupported_preview"
    });

    expect(supportedHtml).toContain("Confirm Cleanup");
    expect(supportedHtml).toContain("Will standardize format in Contacts!B2:B5.");

    const unsupportedHtml = sidebar.renderStructuredPreview({
      type: "data_cleanup_plan",
      data: {
        targetSheet: "Contacts",
        targetRange: "B2:B5",
        operation: "standardize_format",
        formatType: "date_text",
        formatPattern: "locale-sensitive-fuzzy",
        explanation: "Normalize date strings using a fuzzy locale format.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!B2:B5"],
        overwriteRisk: "medium",
        confirmationLevel: "standard"
      }
    }, {
      runId: "run_cleanup_unsupported_preview_fuzzy",
      requestId: "req_cleanup_unsupported_preview_fuzzy"
    });

    expect(unsupportedHtml).toContain(
      "This Google Sheets flow only supports exact year-first date text patterns"
    );
    expect(unsupportedHtml).not.toContain("Confirm Cleanup");
  });

  it("applies a Google Sheets range transfer plan and returns a typed result", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A2:B3",
      row: 2,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"]
      ]
    });
    const targetRange = createRangeStub({
      a1Notation: "D5:E6",
      row: 5,
      column: 4,
      numRows: 2,
      numColumns: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });

    const sourceSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:B3");
        return sourceRange;
      })
    };
    const targetSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("D5:E6");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return sourceSheet;
        }

        expect(sheetName).toBe("Archive");
        return targetSheet;
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Archive",
        targetRange: "D5:E6",
        operation: "copy",
        pasteMode: "values",
        transpose: false,
        explanation: "Copy finalized rows into the archive block.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Archive!D5:E6"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toEqual({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      hostPlatform: "google_sheets",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Archive",
      targetRange: "D5:E6",
      transferOperation: "copy",
      pasteMode: "values",
      transpose: false,
      summary: "Copied RawData!A2:B3 to Archive!D5:E6."
    });
    expect(targetRange.setValues).toHaveBeenCalledWith([
      ["Ada", "Lovelace"],
      ["Grace", "Hopper"]
    ]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies append transfers at the first trailing blank block in Google Sheets", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A2:B3",
      row: 2,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"]
      ]
    });
    const targetRange = createRangeStub({
      a1Notation: "D5:E9",
      row: 5,
      column: 4,
      numRows: 5,
      numColumns: 2,
      values: [
        ["Existing", "One"],
        ["Existing", "Two"],
        ["", ""],
        ["", ""],
        ["", ""]
      ]
    });
    const appendWriteRange = createRangeStub({
      a1Notation: "D7:E8",
      row: 7,
      column: 4,
      numRows: 2,
      numColumns: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    const targetSheet = {
      getRange: vi.fn((firstArg: string | number, column?: number, numRows?: number, numColumns?: number) => {
        if (typeof firstArg === "string") {
          expect(firstArg).toBe("D5:E9");
          return targetRange;
        }

        expect(firstArg).toBe(7);
        expect(column).toBe(4);
        expect(numRows).toBe(2);
        expect(numColumns).toBe(2);
        return appendWriteRange;
      })
    };
    const sourceSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:B3");
        return sourceRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return sourceSheet;
        }

        expect(sheetName).toBe("Archive");
        return targetSheet;
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Archive",
        targetRange: "D5:E9",
        operation: "append",
        pasteMode: "values",
        transpose: false,
        explanation: "Append finalized rows after the existing archive block.",
        confidence: 0.95,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Archive!D5:E9"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toEqual({
      kind: "range_transfer_update",
      hostPlatform: "google_sheets",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Archive",
      targetRange: "D7:E8",
      transferOperation: "append",
      pasteMode: "values",
      transpose: false,
      summary: "Appended RawData!A2:B3 into Archive!D7:E8."
    });
    expect(appendWriteRange.setValues).toHaveBeenCalledWith([
      ["Ada", "Lovelace"],
      ["Grace", "Hopper"]
    ]);
    expect(targetRange.setValues).not.toHaveBeenCalled();
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("uses Google Sheets copy semantics for formula-mode transfers so relative references can rebase", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A2:C3",
      row: 2,
      column: 1,
      numRows: 2,
      numColumns: 3,
      values: [
        [10, true, "label"],
        [20, false, "next"]
      ],
      formulas: [
        ["=A1+1", 42, "plain text"],
        [true, "=SUM(B1:B2)", "done"]
      ]
    });
    const targetRange = createRangeStub({
      a1Notation: "F2:H3",
      row: 2,
      column: 6,
      numRows: 2,
      numColumns: 3,
      values: [
        ["", "", ""],
        ["", "", ""]
      ],
      formulas: [
        ["", "", ""],
        ["", "", ""]
      ]
    });
    sourceRange.copyTo = vi.fn((destination: typeof targetRange, pasteType: string, transpose: boolean) => {
      expect(destination).toBe(targetRange);
      expect(pasteType).toBe("PASTE_FORMULA");
      expect(transpose).toBe(false);
      destination.__appliedFormulas = [
        ["=F1+1", 42, "plain text"],
        [true, "=SUM(G1:G2)", "done"]
      ];
    });

    const sourceSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:C3");
        return sourceRange;
      })
    };
    const targetSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("F2:H3");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return sourceSheet;
        }

        expect(sheetName).toBe("Report");
        return targetSheet;
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:C3",
        targetSheet: "Report",
        targetRange: "F2:H3",
        operation: "copy",
        pasteMode: "formulas",
        transpose: false,
        explanation: "Copy formulas and constants into the report block.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:C3", "Report!F2:H3"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toEqual({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      hostPlatform: "google_sheets",
      sourceSheet: "RawData",
      sourceRange: "A2:C3",
      targetSheet: "Report",
      targetRange: "F2:H3",
      transferOperation: "copy",
      pasteMode: "formulas",
      transpose: false,
      summary: "Copied RawData!A2:C3 to Report!F2:H3."
    });
    expect(sourceRange.copyTo).toHaveBeenCalledTimes(1);
    expect(targetRange.setFormulas).not.toHaveBeenCalled();
    expect(targetRange.__appliedFormulas).toEqual([
      ["=F1+1", 42, "plain text"],
      [true, "=SUM(G1:G2)", "done"]
    ]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies format-only move transfers in Google Sheets and clears the source formatting after success", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A2:B3",
      row: 2,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"]
      ]
    });
    const targetRange = createRangeStub({
      a1Notation: "F2:G3",
      row: 2,
      column: 6,
      numRows: 2,
      numColumns: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    sourceRange.copyTo = vi.fn((destination: typeof targetRange, pasteType: string, transpose: boolean) => {
      expect(destination).toBe(targetRange);
      expect(pasteType).toBe("PASTE_FORMAT");
      expect(transpose).toBe(false);
    });

    const sourceSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:B3");
        return sourceRange;
      })
    };
    const targetSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("F2:G3");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return sourceSheet;
        }

        expect(sheetName).toBe("Report");
        return targetSheet;
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Report",
        targetRange: "F2:G3",
        operation: "move",
        pasteMode: "formats",
        transpose: false,
        explanation: "Move the source formatting into the report block.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Report!F2:G3"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive"
      }
    });

    expect(result).toEqual({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      hostPlatform: "google_sheets",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Report",
      targetRange: "F2:G3",
      transferOperation: "move",
      pasteMode: "formats",
      transpose: false,
      summary: "Moved RawData!A2:B3 to Report!F2:G3."
    });
    expect(sourceRange.copyTo).toHaveBeenCalledTimes(1);
    expect(sourceRange.clearFormat).toHaveBeenCalledTimes(1);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies format-only append transfers in Google Sheets", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A2:B3",
      row: 2,
      column: 1,
      numRows: 2,
      numColumns: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"]
      ]
    });
    const targetRange = createRangeStub({
      a1Notation: "D5:E9",
      row: 5,
      column: 4,
      numRows: 5,
      numColumns: 2,
      values: [
        ["Existing", "One"],
        ["Existing", "Two"],
        ["", ""],
        ["", ""],
        ["", ""]
      ]
    });
    const appendWriteRange = createRangeStub({
      a1Notation: "D7:E8",
      row: 7,
      column: 4,
      numRows: 2,
      numColumns: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    sourceRange.copyTo = vi.fn((destination: typeof appendWriteRange, pasteType: string, transpose: boolean) => {
      expect(destination).toBe(appendWriteRange);
      expect(pasteType).toBe("PASTE_FORMAT");
      expect(transpose).toBe(false);
    });

    const targetSheet = {
      getRange: vi.fn((firstArg: string | number, column?: number, numRows?: number, numColumns?: number) => {
        if (typeof firstArg === "string") {
          expect(firstArg).toBe("D5:E9");
          return targetRange;
        }

        expect(firstArg).toBe(7);
        expect(column).toBe(4);
        expect(numRows).toBe(2);
        expect(numColumns).toBe(2);
        return appendWriteRange;
      })
    };
    const sourceSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:B3");
        return sourceRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return sourceSheet;
        }

        expect(sheetName).toBe("Archive");
        return targetSheet;
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Archive",
        targetRange: "D5:E9",
        operation: "append",
        pasteMode: "formats",
        transpose: false,
        explanation: "Append the source formatting after the existing archive block.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Archive!D5:E9"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toEqual({
      kind: "range_transfer_update",
      hostPlatform: "google_sheets",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Archive",
      targetRange: "D7:E8",
      transferOperation: "append",
      pasteMode: "formats",
      transpose: false,
      summary: "Appended RawData!A2:B3 into Archive!D7:E8."
    });
    expect(sourceRange.copyTo).toHaveBeenCalledTimes(1);
    expect(appendWriteRange.setValues).not.toHaveBeenCalled();
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("fails closed when a fully blank multi-row append target is too small in Google Sheets", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A2:B4",
      row: 2,
      column: 1,
      numRows: 3,
      numColumns: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"],
        ["Katherine", "Johnson"]
      ]
    });
    const targetRange = createRangeStub({
      a1Notation: "D5:E6",
      row: 5,
      column: 4,
      numRows: 2,
      numColumns: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    const targetSheet = {
      getRange: vi.fn((firstArg: string | number) => {
        if (typeof firstArg === "string") {
          expect(firstArg).toBe("D5:E6");
          return targetRange;
        }

        throw new Error("append target should not expand a multi-row approved window");
      })
    };
    const sourceSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:B4");
        return sourceRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return sourceSheet;
        }

        expect(sheetName).toBe("Archive");
        return targetSheet;
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B4",
        targetSheet: "Archive",
        targetRange: "D5:E6",
        operation: "append",
        pasteMode: "values",
        transpose: false,
        explanation: "Append finalized rows after the existing archive block.",
        confidence: 0.95,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B4", "Archive!D5:E6"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toThrow("Google Sheets host cannot append exactly within the approved target range.");

    expect(targetRange.setValues).not.toHaveBeenCalled();
    expect(flush).not.toHaveBeenCalled();
  });

  it("applies trim_whitespace cleanup in Google Sheets and returns a typed result", () => {
    const targetRange = createRangeStub({
      a1Notation: "A2:B4",
      row: 2,
      column: 1,
      numRows: 3,
      numColumns: 2,
      values: [
        ["  Ada  ", " Lovelace "],
        ["Grace ", " Hopper"],
        [" Alan", "Turing  "]
      ]
    });
    const sheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:B4");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn(() => sheet)
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Contacts",
        targetRange: "A2:B4",
        operation: "trim_whitespace",
        explanation: "Trim imported contact fields in place.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:B4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toMatchObject({
      kind: "data_cleanup_update",
      operation: "trim_whitespace",
      targetSheet: "Contacts",
      targetRange: "A2:B4"
    });
    expect(targetRange.values).toEqual([
      ["Ada", "Lovelace"],
      ["Grace", "Hopper"],
      ["Alan", "Turing"]
    ]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies trim_whitespace cleanup with formula-aware wrappers in Google Sheets", () => {
    const targetRange = createRangeStub({
      a1Notation: "A2:A4",
      row: 2,
      column: 1,
      numRows: 3,
      numColumns: 1,
      values: [
        [" Ada "],
        [" Grace "],
        [" Alan "]
      ],
      formulas: [
        ['=" Ada "'],
        ['=" Grace "'],
        [""]
      ]
    });
    const sheet = {
      getRange: vi.fn(() => targetRange)
    };
    const spreadsheet = {
      getSheetByName: vi.fn(() => sheet)
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Contacts",
        targetRange: "A2:A4",
        operation: "trim_whitespace",
        explanation: "Trim formula-driven contact names in place.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:A4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toMatchObject({
      kind: "data_cleanup_update",
      operation: "trim_whitespace",
      targetSheet: "Contacts",
      targetRange: "A2:A4"
    });
    expect(targetRange.setValues).toHaveBeenCalledWith([
      ['=LET(_hermes_value, " Ada ", IF(ISTEXT(_hermes_value), TRIM(_hermes_value), _hermes_value))'],
      ['=LET(_hermes_value, " Grace ", IF(ISTEXT(_hermes_value), TRIM(_hermes_value), _hermes_value))'],
      ["Alan"]
    ]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies normalize_case title cleanup in Google Sheets", () => {
    const targetRange = createRangeStub({
      a1Notation: "A2:A4",
      row: 2,
      column: 1,
      numRows: 3,
      numColumns: 1,
      values: [
        ["ada lovelace"],
        ["GRACE HOPPER"],
        ["alan turing"]
      ]
    });
    const sheet = {
      getRange: vi.fn(() => targetRange)
    };
    const spreadsheet = {
      getSheetByName: vi.fn(() => sheet)
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Contacts",
        targetRange: "A2:A4",
        operation: "normalize_case",
        mode: "title",
        explanation: "Normalize contact names into title case.",
        confidence: 0.86,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:A4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toMatchObject({
      kind: "data_cleanup_update",
      operation: "normalize_case",
      mode: "title",
      targetSheet: "Contacts",
      targetRange: "A2:A4"
    });
    expect(targetRange.values).toEqual([
      ["Ada Lovelace"],
      ["Grace Hopper"],
      ["Alan Turing"]
    ]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("fails closed for non-formula-aware cleanup plans when the target range contains formulas in Google Sheets", () => {
    const targetRange = createRangeStub({
      a1Notation: "A2:A4",
      row: 2,
      column: 1,
      numRows: 3,
      numColumns: 1,
      values: [
        ["Ada"],
        ["Ada"],
        ["Grace"]
      ],
      formulas: [
        ['="Ada"'],
        [""],
        [""]
      ]
    });
    const sheet = {
      getRange: vi.fn(() => targetRange)
    };
    const spreadsheet = {
      getSheetByName: vi.fn(() => sheet)
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Contacts",
        targetRange: "A2:A4",
        operation: "remove_duplicate_rows",
        explanation: "Remove duplicate contacts in place.",
        confidence: 0.82,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:A4"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive"
      }
    })).toThrow("Google Sheets host cannot apply cleanup plans exactly when the target range contains formulas.");

    expect(flush).not.toHaveBeenCalled();
  });

  it("fails closed for unsupported cleanup semantics in Google Sheets apply", () => {
    const targetRange = createRangeStub({
      a1Notation: "B2:B5",
      row: 2,
      column: 2,
      numRows: 4,
      numColumns: 1,
      values: [
        ["2026/04/20"],
        ["2026/04/21"],
        ["2026/04/22"],
        ["2026/04/23"]
      ]
    });
    const sheet = {
      getRange: vi.fn(() => targetRange)
    };
    const spreadsheet = {
      getSheetByName: vi.fn(() => sheet)
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(applyWritePlan({
      plan: {
        targetSheet: "Contacts",
        targetRange: "B2:B5",
        operation: "standardize_format",
        formatType: "date_text",
        formatPattern: "YYYY-MM-DD",
        explanation: "Normalize date strings into ISO format.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!B2:B5"],
        overwriteRisk: "medium",
        confirmationLevel: "standard"
      }
    })).toMatchObject({
      kind: "data_cleanup_update",
      hostPlatform: "google_sheets",
      targetSheet: "Contacts",
      targetRange: "B2:B5",
      operation: "standardize_format",
      formatType: "date_text",
      formatPattern: "YYYY-MM-DD",
      summary: "Standardized format in Contacts!B2:B5."
    });

    expect(targetRange.values).toEqual([
      ["2026-04-20"],
      ["2026-04-21"],
      ["2026-04-22"],
      ["2026-04-23"]
    ]);
    expect(targetRange.setValues).toHaveBeenCalledWith([
      ["2026-04-20"],
      ["2026-04-21"],
      ["2026-04-22"],
      ["2026-04-23"]
    ]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("fails closed for overlapping move transfers in Google Sheets", () => {
    const sheet = {
      getRange: vi.fn((rangeName: string) => {
        if (rangeName === "A1:B2") {
          return createRangeStub({
            a1Notation: "A1:B2",
            row: 1,
            column: 1,
            numRows: 2,
            numColumns: 2,
            values: [
              [1, 2],
              [3, 4]
            ]
          });
        }

        expect(rangeName).toBe("B2:C3");
        return createRangeStub({
          a1Notation: "B2:C3",
          row: 2,
          column: 2,
          numRows: 2,
          numColumns: 2,
          values: [
            ["", ""],
            ["", ""]
          ]
        });
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn(() => sheet)
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        sourceSheet: "Sheet1",
        sourceRange: "A1:B2",
        targetSheet: "Sheet1",
        targetRange: "B2:C3",
        operation: "move",
        pasteMode: "values",
        transpose: false,
        explanation: "Move the block down by one row and one column.",
        confidence: 0.82,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!A1:B2", "Sheet1!B2:C3"],
        overwriteRisk: "high",
        confirmationLevel: "destructive"
      }
    })).toThrow("Google Sheets host cannot apply an overlapping move transfer exactly.");

    expect(flush).not.toHaveBeenCalled();
  });
});
