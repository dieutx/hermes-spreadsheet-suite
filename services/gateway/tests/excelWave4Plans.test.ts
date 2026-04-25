import { afterEach, describe, expect, it, vi } from "vitest";

const TASKPANE_MODULE_URL = new URL(
  "../../../apps/excel-addin/src/taskpane/taskpane.js",
  import.meta.url
).href;

function createElementStub() {
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

async function loadTaskpaneModule(
  excelContext: Record<string, unknown>,
  options: {
    confirm?: (message: string) => boolean;
  } = {}
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
  const localStorage = {
    getItem(key: string) {
      return storage.get(key) ?? null;
    },
    setItem(key: string, value: string) {
      storage.set(key, value);
    },
    removeItem(key: string) {
      storage.delete(key);
    }
  };

  vi.stubGlobal("window", {
    location: { search: "" },
    localStorage,
    confirm: options.confirm || (() => true),
    addEventListener() {},
    setInterval,
    clearInterval
  });
  vi.stubGlobal("document", {
    getElementById(id: string) {
      return elements.get(id) ?? createElementStub();
    }
  });
  vi.stubGlobal("fetch", vi.fn());
  vi.stubGlobal("Office", {
    PlatformType: { Mac: "Mac" },
    context: {
      platform: "PC",
      diagnostics: { version: "test-client" },
      document: {
        url: "",
        settings: {
          get() {
            return null;
          },
          remove() {},
          saveAsync() {}
        }
      },
      displayLanguage: "en-US"
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

function createRangeStub(options: {
  address: string;
  rowCount: number;
  columnCount: number;
  values?: unknown[][];
  formulas?: (string | null)[][];
}) {
  let currentValues = options.values ? options.values.map((row) => [...row]) : [];
  let currentFormulas = options.formulas
    ? options.formulas.map((row) => [...row])
    : currentValues.map((row) => row.map((value) => value == null ? "" : String(value)));

  const range = {
    address: options.address,
    rowCount: options.rowCount,
    columnCount: options.columnCount,
    values: currentValues,
    formulas: currentFormulas,
    load: vi.fn(),
    clear: vi.fn(),
    getResizedRange: vi.fn((rowDelta: number, columnDelta: number) => createRangeStub({
      address: options.address,
      rowCount: rowDelta + 1,
      columnCount: columnDelta + 1,
      values: Array.from({ length: rowDelta + 1 }, () => Array.from({ length: columnDelta + 1 }, () => "")),
      formulas: Array.from({ length: rowDelta + 1 }, () => Array.from({ length: columnDelta + 1 }, () => ""))
    }))
  };

  Object.defineProperty(range, "values", {
    configurable: true,
    get() {
      return currentValues;
    },
    set(nextValues) {
      currentValues = nextValues;
    }
  });

  Object.defineProperty(range, "formulas", {
    configurable: true,
    get() {
      return currentFormulas;
    },
    set(nextFormulas) {
      currentFormulas = nextFormulas;
    }
  });

  return range;
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("Excel wave 4 transfer and cleanup plans", () => {
  it("recognizes range transfer and cleanup previews as write plans", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const transferResponse = {
      type: "range_transfer_plan",
      data: {
        sourceSheet: "RawData",
        sourceRange: "A2:C4",
        targetSheet: "Report",
        targetRange: "B5:D7",
        operation: "copy",
        pasteMode: "values",
        transpose: false,
        explanation: "Copy the cleaned rows into the report sheet.",
        confidence: 0.96,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:C4", "Report!B5:D7"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };
    const cleanupResponse = {
      type: "data_cleanup_plan",
      data: {
        targetSheet: "Contacts",
        targetRange: "A2:B4",
        operation: "trim_whitespace",
        explanation: "Trim extra whitespace in the imported contacts.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:B4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(taskpane.isWritePlanResponse(transferResponse)).toBe(true);
    expect(taskpane.isWritePlanResponse(cleanupResponse)).toBe(true);
    expect(taskpane.getRequiresConfirmation(transferResponse)).toBe(true);
    expect(taskpane.getRequiresConfirmation(cleanupResponse)).toBe(true);

    expect(taskpane.getStructuredPreview(transferResponse)).toMatchObject({
      kind: "range_transfer_plan",
      sourceSheet: "RawData",
      targetSheet: "Report",
      operation: "copy",
      pasteMode: "values"
    });
    expect(taskpane.getStructuredPreview(cleanupResponse)).toMatchObject({
      kind: "data_cleanup_plan",
      targetSheet: "Contacts",
      targetRange: "A2:B4",
      operation: "trim_whitespace"
    });

    expect(taskpane.renderStructuredPreview(transferResponse, {
      runId: "run_transfer_preview",
      requestId: "req_transfer_preview"
    })).toContain("Confirm Transfer");
    expect(taskpane.renderStructuredPreview(cleanupResponse, {
      runId: "run_cleanup_preview",
      requestId: "req_cleanup_preview"
    })).toContain("Confirm Cleanup");

    const formatTransferPreview = taskpane.renderStructuredPreview({
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
    const titleCleanupPreview = taskpane.renderStructuredPreview({
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

    expect(formatTransferPreview).toContain("Confirm Transfer");
    expect(titleCleanupPreview).toContain("Confirm Cleanup");
  });

  it("builds destructive transfer approvals with a second confirmation payload", async () => {
    const confirm = vi.fn(() => true);
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, { confirm });

    const approvalRequest = taskpane.buildWriteApprovalRequest({
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
        explanation: "Move the finalized rows into the archive block.",
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

  it("builds destructive cleanup approvals with a second confirmation payload", async () => {
    const confirm = vi.fn(() => true);
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    }, { confirm });

    const approvalRequest = taskpane.buildWriteApprovalRequest({
      requestId: "req_cleanup_destructive",
      runId: "run_cleanup_destructive",
      plan: {
        targetSheet: "Contacts",
        targetRange: "A2:F100",
        operation: "remove_duplicate_rows",
        keyColumns: ["A", "C"],
        explanation: "Remove duplicate contact rows before export.",
        confidence: 0.93,
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
        operation: "remove_duplicate_rows",
        confirmationLevel: "destructive"
      }
    });
  });

  it("rejects Excel import writes when the destination contains a blank-rendering formula", async () => {
    const targetRange = createRangeStub({
      address: "Imported!A1:B2",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["", ""],
        ["", ""]
      ],
      formulas: [
        ['=IF(TRUE,"","")', ""],
        ["", ""]
      ]
    });
    const worksheets = {
      getItem: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Imported");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("A1:B2");
            return targetRange;
          })
        };
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    });

    await expect(taskpane.applyWritePlan({
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
      requestId: "req_import_formula_guard_excel_001",
      runId: "run_import_formula_guard_excel_001",
      approvalToken: "token"
    })).rejects.toThrow("Target range already contains content.");
    expect(targetRange.values).toEqual([
      ["", ""],
      ["", ""]
    ]);
  });

  it("replaces a transfer plan message with the completed transfer summary", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });
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

    taskpane.applyWritebackResultToMessage(message, {
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      hostPlatform: "excel",
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

  it("replaces a cleanup plan message with the completed cleanup summary", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });
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

    taskpane.applyWritebackResultToMessage(message, {
      kind: "data_cleanup_update",
      operation: "data_cleanup_update",
      hostPlatform: "excel",
      targetSheet: "Contacts",
      targetRange: "A2:B4",
      cleanupOperation: "trim_whitespace",
      summary: "Trimmed whitespace in Contacts!A2:B4."
    });

    expect(message.content).toBe("Trimmed whitespace in Contacts!A2:B4.");
    expect(message.response).toBeNull();
    expect(message.statusLine).toBe("");
  });

  it("applies a move range transfer in Excel and clears the source after the target write succeeds", async () => {
    const operationLog: string[] = [];
    const sourceRange = createRangeStub({
      address: "RawData!A2:B3",
      rowCount: 2,
      columnCount: 2,
      values: [
        [" Alpha ", "Beta"],
        ["Gamma", "Delta"]
      ],
      formulas: [
        [null, null],
        [null, null]
      ]
    });
    sourceRange.clear = vi.fn(() => {
      operationLog.push("source.clear");
    });

    const targetRange = createRangeStub({
      address: "Archive!D5:E6",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["", ""],
        ["", ""]
      ],
      formulas: [
        [null, null],
        [null, null]
      ]
    });
    Object.defineProperty(targetRange, "values", {
      configurable: true,
      get() {
        return [
          ["", ""],
          ["", ""]
        ];
      },
      set(nextValues) {
        operationLog.push("target.values");
        targetRange.__appliedValues = nextValues;
      }
    });

    const worksheets = {
      getItem: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return {
            getRange: vi.fn((rangeName: string) => {
              expect(rangeName).toBe("A2:B3");
              return sourceRange;
            })
          };
        }

        expect(sheetName).toBe("Archive");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("D5:E6");
            return targetRange;
          })
        };
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Archive",
        targetRange: "D5:E6",
        operation: "move",
        pasteMode: "values",
        transpose: false,
        explanation: "Move finalized rows into the archive sheet.",
        confidence: 0.95,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Archive!D5:E6"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive"
      },
      requestId: "req_range_transfer_move_excel_001",
      runId: "run_range_transfer_move_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Archive",
      targetRange: "D5:E6",
      transferOperation: "move",
      pasteMode: "values",
      transpose: false
    });

    expect(operationLog).toEqual(["target.values", "source.clear"]);
    expect(sourceRange.clear).toHaveBeenCalledTimes(1);
    expect(targetRange.__appliedValues).toEqual([
      [" Alpha ", "Beta"],
      ["Gamma", "Delta"]
    ]);
  });

  it("uses Excel copy semantics for formula-mode transfers so relative references can rebase", async () => {
    const sourceRange = createRangeStub({
      address: "RawData!A2:C3",
      rowCount: 2,
      columnCount: 3,
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
      address: "Report!F2:H3",
      rowCount: 2,
      columnCount: 3,
      values: [
        ["", "", ""],
        ["", "", ""]
      ],
      formulas: [
        ["", "", ""],
        ["", "", ""]
      ]
    });
    targetRange.copyFrom = vi.fn(
      (
        copiedRange: typeof sourceRange,
        copyType: string,
        skipBlanks: boolean,
        transpose: boolean
      ) => {
        expect(copiedRange).toBe(sourceRange);
        expect(copyType).toBe("Formulas");
        expect(skipBlanks).toBe(false);
        expect(transpose).toBe(false);
        targetRange.__appliedFormulas = [
          ["=F1+1", 42, "plain text"],
          [true, "=SUM(G1:G2)", "done"]
        ];
      }
    );
    Object.defineProperty(targetRange, "formulas", {
      configurable: true,
      get() {
        return [
          ["", "", ""],
          ["", "", ""]
        ];
      },
      set(nextFormulas) {
        targetRange.__literalFormulaAssignment = nextFormulas;
      }
    });

    const worksheets = {
      getItem: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return {
            getRange: vi.fn((rangeName: string) => {
              expect(rangeName).toBe("A2:C3");
              return sourceRange;
            })
          };
        }

        expect(sheetName).toBe("Report");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("F2:H3");
            return targetRange;
          })
        };
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    });

    await expect(taskpane.applyWritePlan({
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
      },
      requestId: "req_range_transfer_formulas_excel_001",
      runId: "run_range_transfer_formulas_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:C3",
      targetSheet: "Report",
      targetRange: "F2:H3",
      transferOperation: "copy",
      pasteMode: "formulas",
      transpose: false
    });

    expect(targetRange.copyFrom).toHaveBeenCalledTimes(1);
    expect(targetRange.__literalFormulaAssignment).toBeUndefined();
    expect(targetRange.__appliedFormulas).toEqual([
      ["=F1+1", 42, "plain text"],
      [true, "=SUM(G1:G2)", "done"]
    ]);
  });

  it("applies an append transfer in Excel when targetRange is an exact-safe append anchor", async () => {
    const sourceRange = createRangeStub({
      address: "RawData!A2:B3",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"]
      ]
    });
    const expandedWriteRange = createRangeStub({
      address: "Report!D5:E6",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    Object.defineProperty(expandedWriteRange, "values", {
      configurable: true,
      get() {
        return [
          ["", ""],
          ["", ""]
        ];
      },
      set(nextValues) {
        expandedWriteRange.__appliedValues = nextValues;
      }
    });

    const anchorRange = createRangeStub({
      address: "Report!D5:E5",
      rowCount: 1,
      columnCount: 2,
      values: [["", ""]]
    });
    anchorRange.getResizedRange = vi.fn((rowDelta: number, columnDelta: number) => {
      expect(rowDelta).toBe(1);
      expect(columnDelta).toBe(0);
      return expandedWriteRange;
    });

    const worksheets = {
      getItem: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return {
            getRange: vi.fn((rangeName: string) => {
              expect(rangeName).toBe("A2:B3");
              return sourceRange;
            })
          };
        }

        expect(sheetName).toBe("Report");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("D5:E5");
            return anchorRange;
          })
        };
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Report",
        targetRange: "D5:E5",
        operation: "append",
        pasteMode: "values",
        transpose: false,
        explanation: "Append the next batch at the approved start row.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Report!D5:E5"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_range_transfer_append_anchor_excel_001",
      runId: "run_range_transfer_append_anchor_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Report",
      targetRange: "D5:E6",
      transferOperation: "append",
      pasteMode: "values",
      transpose: false,
      summary: "Appended RawData!A2:B3 into Report!D5:E6."
    });

    expect(anchorRange.getResizedRange).toHaveBeenCalledTimes(1);
    expect(expandedWriteRange.__appliedValues).toEqual([
      ["Ada", "Lovelace"],
      ["Grace", "Hopper"]
    ]);
  });

  it("applies a format-only append transfer in Excel when targetRange is an exact-safe append anchor", async () => {
    const sourceRange = createRangeStub({
      address: "RawData!A2:B3",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"]
      ]
    });
    const expandedWriteRange = createRangeStub({
      address: "Report!D5:E6",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    expandedWriteRange.copyFrom = vi.fn((
      copiedRange: typeof sourceRange,
      copyType: string,
      skipBlanks: boolean,
      transpose: boolean
    ) => {
      expect(copiedRange).toBe(sourceRange);
      expect(copyType).toBe("Formats");
      expect(skipBlanks).toBe(false);
      expect(transpose).toBe(false);
    });

    const anchorRange = createRangeStub({
      address: "Report!D5:E5",
      rowCount: 1,
      columnCount: 2,
      values: [["", ""]]
    });
    anchorRange.getResizedRange = vi.fn((rowDelta: number, columnDelta: number) => {
      expect(rowDelta).toBe(1);
      expect(columnDelta).toBe(0);
      return expandedWriteRange;
    });

    const worksheets = {
      getItem: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return {
            getRange: vi.fn((rangeName: string) => {
              expect(rangeName).toBe("A2:B3");
              return sourceRange;
            })
          };
        }

        expect(sheetName).toBe("Report");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("D5:E5");
            return anchorRange;
          })
        };
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Report",
        targetRange: "D5:E5",
        operation: "append",
        pasteMode: "formats",
        transpose: false,
        explanation: "Append the source formatting at the approved start row.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Report!D5:E5"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_range_transfer_append_formats_excel_001",
      runId: "run_range_transfer_append_formats_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Report",
      targetRange: "D5:E6",
      transferOperation: "append",
      pasteMode: "formats",
      transpose: false,
      summary: "Appended RawData!A2:B3 into Report!D5:E6."
    });

    expect(anchorRange.getResizedRange).toHaveBeenCalledTimes(1);
    expect(expandedWriteRange.copyFrom).toHaveBeenCalledTimes(1);
  });

  it("reports the expanded target range for anchor-based non-append transfers in Excel", async () => {
    const sourceRange = createRangeStub({
      address: "RawData!A2:B3",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["Ada", "Lovelace"],
        ["Grace", "Hopper"]
      ]
    });
    const expandedWriteRange = createRangeStub({
      address: "Report!F2:G3",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    Object.defineProperty(expandedWriteRange, "values", {
      configurable: true,
      get() {
        return [
          ["", ""],
          ["", ""]
        ];
      },
      set(nextValues) {
        expandedWriteRange.__appliedValues = nextValues;
      }
    });

    const anchorRange = createRangeStub({
      address: "Report!F2",
      rowCount: 1,
      columnCount: 1,
      values: [[""]]
    });
    anchorRange.getResizedRange = vi.fn((rowDelta: number, columnDelta: number) => {
      expect(rowDelta).toBe(1);
      expect(columnDelta).toBe(1);
      return expandedWriteRange;
    });

    const worksheets = {
      getItem: vi.fn((sheetName: string) => {
        if (sheetName === "RawData") {
          return {
            getRange: vi.fn((rangeName: string) => {
              expect(rangeName).toBe("A2:B3");
              return sourceRange;
            })
          };
        }

        expect(sheetName).toBe("Report");
        return {
          getRange: vi.fn((rangeName: string) => {
            expect(rangeName).toBe("F2");
            return anchorRange;
          })
        };
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: { worksheets }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        sourceSheet: "RawData",
        sourceRange: "A2:B3",
        targetSheet: "Report",
        targetRange: "F2",
        operation: "copy",
        pasteMode: "values",
        transpose: false,
        explanation: "Copy finalized rows into the report block.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["RawData!A2:B3", "Report!F2"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_range_transfer_anchor_excel_001",
      runId: "run_range_transfer_anchor_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_transfer_update",
      operation: "range_transfer_update",
      sourceSheet: "RawData",
      sourceRange: "A2:B3",
      targetSheet: "Report",
      targetRange: "F2:G3",
      transferOperation: "copy",
      pasteMode: "values",
      transpose: false,
      summary: "Copied RawData!A2:B3 to Report!F2:G3."
    });

    expect(anchorRange.getResizedRange).toHaveBeenCalledTimes(1);
    expect(expandedWriteRange.__appliedValues).toEqual([
      ["Ada", "Lovelace"],
      ["Grace", "Hopper"]
    ]);
  });

  it("applies trim_whitespace cleanup in Excel", async () => {
    const targetRange = createRangeStub({
      address: "Contacts!A2:B4",
      rowCount: 3,
      columnCount: 2,
      values: [
        ["  Ada  ", " Lovelace "],
        ["Grace ", " Hopper"],
        [" Alan", "Turing  "]
      ],
      formulas: [
        ["  Ada  ", " Lovelace "],
        ["Grace ", " Hopper"],
        [" Alan", "Turing  "]
      ]
    });
    const worksheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A2:B4");
        return targetRange;
      })
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
        targetSheet: "Contacts",
        targetRange: "A2:B4",
        operation: "trim_whitespace",
        explanation: "Trim extra whitespace in imported contacts.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:B4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_cleanup_trim_excel_001",
      runId: "run_cleanup_trim_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
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
  });

  it("wraps existing formulas for exact-safe cleanup plans in Excel", async () => {
    const targetRange = createRangeStub({
      address: "Contacts!A2:A4",
      rowCount: 3,
      columnCount: 1,
      values: [
        [" Ada "],
        [" Grace "],
        [" Alan "]
      ],
      formulas: [
        [""],
        ["=UPPER(A2)"],
        [""]
      ]
    });
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
        targetSheet: "Contacts",
        targetRange: "A2:A4",
        operation: "trim_whitespace",
        explanation: "Trim extra whitespace in imported contacts.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:A4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_cleanup_formulas_excel_001",
      runId: "run_cleanup_formulas_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "data_cleanup_update",
      hostPlatform: "excel_windows",
      targetSheet: "Contacts",
      targetRange: "A2:A4",
      operation: "trim_whitespace",
      summary: "Trimmed whitespace in Contacts!A2:A4."
    });

    expect(targetRange.formulas).toEqual([
      ["Ada"],
      ["=LET(_hermes_value, UPPER(A2), IF(ISTEXT(_hermes_value), TRIM(_hermes_value), _hermes_value))"],
      ["Alan"]
    ]);
  });

  it("fails closed for unsupported normalize_case modes in Excel", async () => {
    const targetRange = createRangeStub({
      address: "Contacts!A2:A4",
      rowCount: 3,
      columnCount: 1,
      values: [
        ["Ada Lovelace"],
        ["Grace Hopper"],
        ["Alan Turing"]
      ]
    });
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
        targetSheet: "Contacts",
        targetRange: "A2:A4",
        operation: "normalize_case",
        mode: "sentence",
        explanation: "Normalize names into sentence case.",
        confidence: 0.81,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:A4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_cleanup_normalize_case_invalid_excel_001",
      runId: "run_cleanup_normalize_case_invalid_excel_001",
      approvalToken: "token"
    })).rejects.toThrow(
      "Excel host does not support exact-safe cleanup semantics for normalize_case mode sentence."
    );
  });

  it("applies normalize_case title cleanup in Excel", async () => {
    const targetRange = createRangeStub({
      address: "Contacts!A2:A4",
      rowCount: 3,
      columnCount: 1,
      values: [
        ["ada lovelace"],
        ["GRACE HOPPER"],
        ["alan turing"]
      ]
    });
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
        targetSheet: "Contacts",
        targetRange: "A2:A4",
        operation: "normalize_case",
        mode: "title",
        explanation: "Normalize names into title case.",
        confidence: 0.84,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!A2:A4"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_cleanup_normalize_case_title_excel_001",
      runId: "run_cleanup_normalize_case_title_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
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
  });

  it("fails closed for overlapping move transfers in Excel", async () => {
    const sourceRange = createRangeStub({
      address: "Sheet1!A1:B2",
      rowCount: 2,
      columnCount: 2,
      values: [
        [1, 2],
        [3, 4]
      ]
    });
    const targetRange = createRangeStub({
      address: "Sheet1!B2:C3",
      rowCount: 2,
      columnCount: 2,
      values: [
        ["", ""],
        ["", ""]
      ]
    });
    const worksheet = {
      getRange: vi.fn((rangeName: string) => rangeName === "A1:B2" ? sourceRange : targetRange)
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
      },
      requestId: "req_range_transfer_overlap_excel_001",
      runId: "run_range_transfer_overlap_excel_001",
      approvalToken: "token"
    })).rejects.toThrow("Excel host cannot apply an overlapping move transfer exactly.");
  });

  it("applies exact-safe standardize_format cleanup semantics in Excel", async () => {
    const targetRange = createRangeStub({
      address: "Contacts!B2:B5",
      rowCount: 4,
      columnCount: 1,
      values: [
        ["2026/04/20"],
        ["2026/04/21"],
        ["2026/04/22"],
        ["2026/04/23"]
      ]
    });
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
        targetSheet: "Contacts",
        targetRange: "B2:B5",
        operation: "standardize_format",
        formatType: "date_text",
        formatPattern: "YYYY-MM-DD",
        explanation: "Normalize date strings into ISO format.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Contacts!B2:B5"],
        overwriteRisk: "medium",
        confirmationLevel: "standard"
      },
      requestId: "req_cleanup_standardize_excel_001",
      runId: "run_cleanup_standardize_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "data_cleanup_update",
      hostPlatform: "excel_windows",
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
  });
});
