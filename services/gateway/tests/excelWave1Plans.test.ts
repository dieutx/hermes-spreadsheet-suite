import { afterEach, describe, expect, it, vi } from "vitest";
import {
  getSheetStructureStatusSummary,
  isSheetStructurePlan
} from "../../../apps/excel-addin/src/taskpane/structurePlan.js";
import {
  buildExcelSortFields,
  getRangeFilterStatusSummary,
  getRangeSortStatusSummary,
  isRangeFilterPlan,
  isRangeSortPlan
} from "../../../apps/excel-addin/src/taskpane/sortFilterPlan.js";

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

async function loadTaskpaneModule(excelContext: Record<string, unknown>) {
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
      displayLanguage: "en-US",
      requirements: {
        isSetSupported(setName: string, minVersion: string) {
          return setName === "ExcelApi" && minVersion === "1.18";
        }
      }
    },
    onReady() {}
  });
  vi.stubGlobal("Excel", {
    run: async (callback: (context: Record<string, unknown>) => unknown) => callback(excelContext),
    InsertShiftDirection: {
      down: "down",
      right: "right"
    },
    DeleteShiftDirection: {
      up: "up",
      left: "left"
    },
    GroupOption: {
      byRows: "byRows",
      byColumns: "byColumns"
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

function normalizeNoteKey(address: string) {
  return address
    .slice(address.lastIndexOf("!") + 1)
    .replace(/\$/g, "");
}

function createNoteCollectionMock(initialNotes: Record<string, string> = {}) {
  const notes = new Map(Object.entries(initialNotes).map(([key, value]) => [
    normalizeNoteKey(key),
    value
  ]));

  function createNoteObject(key: string) {
    return {
      get isNullObject() {
        return !notes.has(key);
      },
      get content() {
        return notes.get(key) ?? "";
      },
      set content(value: unknown) {
        notes.set(key, String(value));
      },
      load: vi.fn(),
      delete: vi.fn(() => {
        notes.delete(key);
      })
    };
  }

  return {
    __notes: notes,
    add: vi.fn((cellOrAddress: { address?: string } | string, content: unknown) => {
      const address = typeof cellOrAddress === "string"
        ? cellOrAddress
        : cellOrAddress.address || "";
      const key = normalizeNoteKey(address);
      notes.set(key, String(content));
      return createNoteObject(key);
    }),
    getItemOrNullObject: vi.fn((address: string) => createNoteObject(normalizeNoteKey(address)))
  };
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("Excel wave 1 plan helpers", () => {
  it("detects structure, sort, and filter plans", () => {
    expect(isSheetStructurePlan({
      operation: "insert_rows",
      targetSheet: "Sheet1"
    })).toBe(true);

    expect(isRangeSortPlan({
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      keys: [
        { columnRef: "Status", direction: "asc" }
      ]
    })).toBe(true);

    expect(isRangeFilterPlan({
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      conditions: [
        { columnRef: "Status", operator: "equals", value: "Open" }
      ],
      combiner: "and"
    })).toBe(true);
  });

  it("builds readable Excel sort/filter helper outputs", () => {
    expect(buildExcelSortFields({
      keys: [
        { columnRef: "Status", direction: "asc" },
        { columnRef: 3, direction: "desc" }
      ]
    })).toEqual([
      { key: "Status", ascending: true },
      { key: 3, ascending: false }
    ]);

    expect(getSheetStructureStatusSummary({
      operation: "insert_rows",
      targetSheet: "Sheet1",
      startIndex: 7,
      count: 3
    })).toBe("Inserted 3 rows at Sheet1 row 8.");

    expect(getRangeSortStatusSummary({
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      keys: [
        { columnRef: "Status", direction: "asc" },
        { columnRef: "Due Date", direction: "desc" }
      ]
    })).toBe("Sorted Sheet1!A1:F25 by Status (ascending), Due Date (descending).");

    expect(getRangeFilterStatusSummary({
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      conditions: [
        { columnRef: "Status", operator: "equals", value: "Open" }
      ],
      combiner: "and",
      clearExistingFilters: true
    })).toBe("Applied filter to Sheet1!A1:F25.");
  });

  it("builds wave 1 body text and structured previews through taskpane integration", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    expect(taskpane.getResponseBodyText({
      type: "sheet_structure_update",
      data: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 7,
        count: 3,
        explanation: "Insert rows above the totals block.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    })).toContain("Prepared a sheet structure update");

    expect(taskpane.getStructuredPreview({
      type: "sheet_structure_update",
      data: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 7,
        count: 3,
        explanation: "Insert rows above the totals block.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    })).toMatchObject({
      kind: "sheet_structure_update",
      targetSheet: "Sheet1",
      operation: "insert_rows"
    });

    expect(taskpane.getResponseBodyText({
      type: "range_sort_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Status", direction: "asc" },
          { columnRef: "Due Date", direction: "desc" }
        ],
        explanation: "Sort by status then due date.",
        confidence: 0.93,
        requiresConfirmation: true
      }
    })).toContain("Prepared a sort plan");

    expect(taskpane.getStructuredPreview({
      type: "range_sort_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Status", direction: "asc" },
          { columnRef: "Due Date", direction: "desc" }
        ],
        explanation: "Sort by status then due date.",
        confidence: 0.93,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "range_sort_plan",
      targetSheet: "Sheet1",
      targetRange: "A1:F25"
    });

    expect(taskpane.getResponseBodyText({
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep open rows only.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    })).toContain("Prepared a filter plan");

    expect(taskpane.getStructuredPreview({
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep open rows only.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "range_filter_plan",
      targetSheet: "Sheet1",
      targetRange: "A1:F25"
    });
  });

  it("returns typed writeback results for applied wave 1 plans", async () => {
    const insertedRows = { insert: vi.fn() };
    const sortApi = { apply: vi.fn() };
    const filterApi = { apply: vi.fn(), clearCriteria: vi.fn() };
    const targetRange = {
      load: vi.fn(),
      rowCount: 25,
      columnCount: 6,
      values: [
        ["Status", "Amount"],
        ["Open", 10]
      ],
      getSort: () => sortApi
    };
    const worksheet = {
      getRange: vi.fn((address?: string) => address === "5:6" ? insertedRows : targetRange),
      getRangeByIndexes: vi.fn(() => insertedRows),
      autoFilter: filterApi
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 4,
        count: 2,
        explanation: "Insert two rows.",
        confidence: 0.88,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      },
      requestId: "req_struct_001",
      runId: "run_struct_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "sheet_structure_update",
      targetSheet: "Sheet1",
      operation: "insert_rows"
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Status", direction: "asc" }
        ],
        explanation: "Sort by the Status column.",
        confidence: 0.91,
        requiresConfirmation: true
      },
      requestId: "req_sort_001",
      runId: "run_sort_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_sort",
      targetSheet: "Sheet1",
      targetRange: "A1:F25"
    });
    expect(sortApi.apply).toHaveBeenCalledWith(
      [{ key: 0, ascending: true }],
      false,
      true
    );

    await expect(taskpane.applyWritePlan({
      plan: {
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
        requiresConfirmation: true
      },
      requestId: "req_filter_001",
      runId: "run_filter_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_filter",
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      conditions: [
        { columnRef: "Status", operator: "equals", value: "Open" }
      ],
      combiner: "and",
      clearExistingFilters: true
    });
    expect(filterApi.clearCriteria).toHaveBeenCalledTimes(1);
    expect(filterApi.apply).toHaveBeenCalledWith(
      targetRange,
      0,
      { filterOn: "custom", criterion1: "=Open" }
    );
  });

  it("attaches local undo snapshots for Excel range filter writes", async () => {
    let criteria = [{ filterOn: "custom", criterion1: "=Closed" }];
    const filterApi = {
      load: vi.fn(),
      clearCriteria: vi.fn(() => {
        criteria = [];
      }),
      apply: vi.fn((_target: unknown, columnIndex: number, nextCriteria: Record<string, unknown>) => {
        criteria[columnIndex] = nextCriteria;
      }),
      get criteria() {
        return criteria;
      }
    };
    const targetRange = {
      load: vi.fn(),
      values: [
        ["Status", "Amount"],
        ["Open", 10]
      ]
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange),
      autoFilter: filterApi
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
        requiresConfirmation: true
      },
      requestId: "req_filter_snapshot_excel_001",
      runId: "run_filter_snapshot_excel_001",
      approvalToken: "token",
      executionId: "exec_filter_snapshot_excel_001"
    });

    expect(result).toMatchObject({
      kind: "range_filter",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_filter_snapshot_excel_001",
        kind: "range_filter",
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        beforeCriteria: [
          { filterOn: "custom", criterion1: "=Closed" }
        ],
        afterCriteria: [
          { filterOn: "custom", criterion1: "=Open" }
        ]
      }
    });
  });

  it("maps Excel sort and filter column refs to zero-based offsets within the target range", async () => {
    const sortApi = { apply: vi.fn() };
    const filterApi = { apply: vi.fn(), clearCriteria: vi.fn() };
    const targetRange = {
      load: vi.fn(),
      rowCount: 9,
      columnCount: 2,
      values: [
        ["Status", "Amount"],
        ["Open", 10]
      ],
      getSort: () => sortApi
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange),
      autoFilter: filterApi
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:C10",
        hasHeader: true,
        keys: [
          { columnRef: "C", direction: "desc" }
        ],
        explanation: "Sort by the Amount column.",
        confidence: 0.91,
        requiresConfirmation: true
      },
      requestId: "req_sort_offset_001",
      runId: "run_sort_offset_001",
      approvalToken: "token"
    });

    expect(sortApi.apply).toHaveBeenCalledWith(
      [{ key: 1, ascending: false }],
      false,
      true
    );

    await taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:C10",
        hasHeader: true,
        conditions: [
          { columnRef: "Amount", operator: "greaterThan", value: 5 }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Filter by the Amount column.",
        confidence: 0.89,
        requiresConfirmation: true
      },
      requestId: "req_filter_offset_001",
      runId: "run_filter_offset_001",
      approvalToken: "token"
    });

    expect(filterApi.apply).toHaveBeenCalledWith(
      targetRange,
      1,
      { filterOn: "custom", criterion1: ">5" }
    );
  });

  it("uses full row and column address ranges for Excel sheet structure mutations", async () => {
    const rowRange = { insert: vi.fn() };
    const columnRange = { delete: vi.fn() };
    const indexedRange = {
      insert: vi.fn(),
      delete: vi.fn()
    };
    const worksheet = {
      getRange: vi.fn((address: string) => {
        if (address === "5:6") {
          return rowRange;
        }

        if (address === "B:D") {
          return columnRange;
        }

        throw new Error(`Unexpected Excel range address: ${address}`);
      }),
      getRangeByIndexes: vi.fn(() => indexedRange)
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 4,
        count: 2,
        explanation: "Insert two full worksheet rows.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      },
      requestId: "req_insert_full_rows_excel_001",
      runId: "run_insert_full_rows_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "sheet_structure_update",
      operation: "insert_rows",
      startIndex: 4,
      count: 2
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        operation: "delete_columns",
        startIndex: 1,
        count: 3,
        explanation: "Delete three full worksheet columns.",
        confidence: 0.9,
        requiresConfirmation: true,
        confirmationLevel: "destructive"
      },
      requestId: "req_delete_full_columns_excel_001",
      runId: "run_delete_full_columns_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "sheet_structure_update",
      operation: "delete_columns",
      startIndex: 1,
      count: 3
    });

    expect(worksheet.getRange).toHaveBeenCalledWith("5:6");
    expect(worksheet.getRange).toHaveBeenCalledWith("B:D");
    expect(rowRange.insert).toHaveBeenCalledWith("down");
    expect(columnRange.delete).toHaveBeenCalledWith("left");
    expect(worksheet.getRangeByIndexes).not.toHaveBeenCalled();
    expect(indexedRange.insert).not.toHaveBeenCalled();
    expect(indexedRange.delete).not.toHaveBeenCalled();
  });

  it("attaches local undo snapshots for Excel sheet rename writes", async () => {
    const worksheet = {
      name: "OldName",
      position: 0,
      visibility: "visible"
    };
    const worksheets = {
      items: [worksheet],
      load: vi.fn()
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets
      }
    });

    const result = await taskpane.applyWritePlan({
      plan: {
        operation: "rename_sheet",
        sheetName: "OldName",
        newSheetName: "NewName",
        explanation: "Rename the staging sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      },
      requestId: "req_rename_sheet_snapshot_excel_001",
      runId: "run_rename_sheet_snapshot_excel_001",
      approvalToken: "token",
      executionId: "exec_rename_sheet_snapshot_excel_001"
    });

    expect(worksheet.name).toBe("NewName");
    expect(result).toMatchObject({
      kind: "workbook_structure_update",
      operation: "rename_sheet",
      sheetName: "OldName",
      newSheetName: "NewName",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_rename_sheet_snapshot_excel_001",
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

  it("fails closed when Excel cannot expose freeze pane operations", async () => {
    const worksheet = {
      getRangeByIndexes: vi.fn(() => ({ address: "Sheet1!B2" })),
      freezePanes: {}
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        operation: "freeze_panes",
        frozenRows: 1,
        frozenColumns: 1,
        explanation: "Freeze the header row and first column.",
        confidence: 0.9,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      },
      requestId: "req_freeze_no_api_excel_001",
      runId: "run_freeze_no_api_excel_001",
      approvalToken: "token"
    })).rejects.toThrow("Excel host does not support freezing panes on this sheet.");

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        operation: "unfreeze_panes",
        frozenRows: 0,
        frozenColumns: 0,
        explanation: "Unfreeze panes.",
        confidence: 0.9,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      },
      requestId: "req_unfreeze_no_api_excel_001",
      runId: "run_unfreeze_no_api_excel_001",
      approvalToken: "token"
    })).rejects.toThrow("Excel host does not support unfreezing panes on this sheet.");
  });

  it("applies note-only sheet updates through the Excel host without changing cell values", async () => {
    let assignedValues: unknown[][] | null = null;
    let assignedFormulas: unknown[][] | null = null;
    const cells = [
      { address: "Sheet1!B2", load: vi.fn() },
      { address: "Sheet1!C2", load: vi.fn() }
    ];
    const targetRange = {
      load: vi.fn(),
      rowCount: 1,
      columnCount: 2,
      values: [["Open", "Closed"]],
      formulas: [["", ""]],
      getCell: vi.fn((rowIndex: number, columnIndex: number) => {
        expect(rowIndex).toBe(0);
        return cells[columnIndex];
      })
    };
    Object.defineProperty(targetRange, "values", {
      configurable: true,
      get() {
        return [["Open", "Closed"]];
      },
      set(nextValues) {
        assignedValues = nextValues;
      }
    });
    Object.defineProperty(targetRange, "formulas", {
      configurable: true,
      get() {
        return [["", ""]];
      },
      set(nextFormulas) {
        assignedFormulas = nextFormulas;
      }
    });
    const noteCollection = createNoteCollectionMock({
      C2: "Old closeout note"
    });
    const worksheet = {
      notes: noteCollection,
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
        targetSheet: "Sheet1",
        targetRange: "B2:C2",
        operation: "set_notes",
        notes: [["Needs review", ""]],
        explanation: "Attach review notes without changing the cell values.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low",
        shape: { rows: 1, columns: 2 }
      },
      requestId: "req_note_only_excel_001",
      runId: "run_note_only_excel_001",
      approvalToken: "token",
      executionId: "exec_note_only_excel_001"
    });

    expect(result).toMatchObject({
      kind: "range_write",
      hostPlatform: "excel_windows",
      targetSheet: "Sheet1",
      targetRange: "B2:C2",
      operation: "set_notes",
      writtenRows: 1,
      writtenColumns: 2
    });
    expect(noteCollection.add).toHaveBeenCalledWith(cells[0], "Needs review");
    expect(noteCollection.__notes.get("B2")).toBe("Needs review");
    expect(noteCollection.__notes.has("C2")).toBe(false);
    expect(assignedValues).toBeNull();
    expect(assignedFormulas).toBeNull();
    expect(result.__hermesLocalExecutionSnapshot.afterCells[0][0]).toMatchObject({
      kind: "value",
      note: "Needs review"
    });
    expect(result.__hermesLocalExecutionSnapshot.afterCells[0][1]).toMatchObject({
      kind: "value",
      note: ""
    });
  });

  it("applies notes included in mixed Excel sheet updates", async () => {
    const appliedCells: Array<{ address: string; kind: string; value: unknown }> = [];
    const cells = [
      { address: "Sheet1!D4", load: vi.fn() },
      { address: "Sheet1!E4", load: vi.fn() },
      { address: "Sheet1!D5", load: vi.fn() },
      { address: "Sheet1!E5", load: vi.fn() }
    ].map((cell) => ({
      ...cell,
      set values(value: unknown) {
        appliedCells.push({ address: cell.address, kind: "value", value });
      },
      set formulas(value: unknown) {
        appliedCells.push({ address: cell.address, kind: "formula", value });
      }
    }));
    const targetRange = {
      load: vi.fn(),
      rowCount: 2,
      columnCount: 2,
      values: [
        ["", ""],
        ["", ""]
      ],
      formulas: [
        ["", ""],
        ["", ""]
      ],
      getCell: vi.fn((rowIndex: number, columnIndex: number) => cells[(rowIndex * 2) + columnIndex])
    };
    const noteCollection = createNoteCollectionMock({
      E4: "Leave this existing note alone"
    });
    const worksheet = {
      notes: noteCollection,
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
        targetSheet: "Sheet1",
        targetRange: "D4:E5",
        operation: "mixed_update",
        values: [
          ["North", null],
          [null, "Total"]
        ],
        formulas: [
          [null, "=SUM(D2:D3)"],
          [null, null]
        ],
        notes: [
          ["Review regional input", ""],
          [null, "Manual total override"]
        ],
        explanation: "Write labels, formulas, and review notes together.",
        confidence: 0.9,
        requiresConfirmation: true,
        overwriteRisk: "medium",
        shape: { rows: 2, columns: 2 }
      },
      requestId: "req_mixed_notes_excel_001",
      runId: "run_mixed_notes_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_write",
      hostPlatform: "excel_windows",
      operation: "mixed_update",
      writtenRows: 2,
      writtenColumns: 2
    });

    expect(appliedCells).toEqual([
      { address: "Sheet1!D4", kind: "value", value: [["North"]] },
      { address: "Sheet1!E4", kind: "formula", value: [["=SUM(D2:D3)"]] },
      { address: "Sheet1!D5", kind: "value", value: [[null]] },
      { address: "Sheet1!E5", kind: "value", value: [["Total"]] }
    ]);
    expect(noteCollection.__notes.get("D4")).toBe("Review regional input");
    expect(noteCollection.__notes.get("E4")).toBe("Leave this existing note alone");
    expect(noteCollection.__notes.get("E5")).toBe("Manual total override");
  });

  it("advertises Excel note-write support in request capabilities", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const request = taskpane.buildRequestEnvelope({
      userMessage: "Add notes to the review cells",
      conversation: [{ role: "user", content: "Add notes to the review cells" }],
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

    expect(request.capabilities.supportsNoteWrites).toBe(true);
  });

  it("fails closed in preview for unsupported Excel filter combinations", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const unsupportedCombinerResponse = {
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "Amount", operator: "greaterThan", value: 10 }
        ],
        combiner: "or",
        clearExistingFilters: true,
        explanation: "Keep open rows or large amounts.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    };
    const duplicateColumnResponse = {
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Amount", operator: "greaterThan", value: 10 },
          { columnRef: "amount", operator: "lessThan", value: 100 }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep amounts in range.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    };
    const unsupportedDuplicateColumnResponse = {
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "status", operator: "notEquals", value: "Closed" },
          { columnRef: "STATUS", operator: "contains", value: "Active" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep only active status rows.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    };

    expect(taskpane.isWritePlanResponse(unsupportedCombinerResponse)).toBe(false);
    expect(taskpane.renderStructuredPreview(unsupportedCombinerResponse, {
      runId: "run_filter_preview_unsupported_excel_combiner",
      requestId: "req_filter_preview_unsupported_excel_combiner"
    })).toContain("can't combine those filter conditions exactly");

    expect(taskpane.isWritePlanResponse(duplicateColumnResponse)).toBe(true);

    const supportedDuplicateHtml = taskpane.renderStructuredPreview(duplicateColumnResponse, {
      runId: "run_filter_preview_supported_excel_duplicate",
      requestId: "req_filter_preview_supported_excel_duplicate"
    });

    expect(supportedDuplicateHtml).toContain("Confirm Filter");

    const duplicateHtml = taskpane.renderStructuredPreview(unsupportedDuplicateColumnResponse, {
      runId: "run_filter_preview_unsupported_excel_duplicate",
      requestId: "req_filter_preview_unsupported_excel_duplicate"
    });

    expect(taskpane.isWritePlanResponse(unsupportedDuplicateColumnResponse)).toBe(false);
    expect(duplicateHtml).toContain("same filter column");
    expect(duplicateHtml).not.toContain("Confirm Filter");
  });

  it("rejects unsupported Excel filter combiner and repeated same-column conditions", async () => {
    const filterApi = { apply: vi.fn(), clearCriteria: vi.fn() };
    const targetRange = {
      load: vi.fn(),
      values: [
        ["Status", "Amount"],
        ["Open", 10]
      ]
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange),
      autoFilter: filterApi
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "Amount", operator: "greaterThan", value: 5 }
        ],
        combiner: "or",
        clearExistingFilters: true,
        explanation: "Unsupported OR filter.",
        confidence: 0.89,
        requiresConfirmation: true
      },
      requestId: "req_filter_or_001",
      runId: "run_filter_or_001",
      approvalToken: "token"
    })).rejects.toThrow(/does not support filter combiners other than and/i);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Amount", operator: "greaterThan", value: 10 },
          { columnRef: "Amount", operator: "lessThan", value: 100 }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep amounts in range.",
        confidence: 0.89,
        requiresConfirmation: true
      },
      requestId: "req_filter_repeat_supported_001",
      runId: "run_filter_repeat_supported_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "range_filter",
      targetSheet: "Sheet1",
      targetRange: "A1:F25"
    });

    expect(filterApi.clearCriteria).toHaveBeenCalledTimes(1);
    expect(filterApi.apply).toHaveBeenCalledWith(
      targetRange,
      1,
      {
        filterOn: "custom",
        criterion1: ">10",
        criterion2: "<100",
        operator: "and"
      }
    );
    filterApi.clearCriteria.mockClear();
    filterApi.apply.mockClear();

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "Status", operator: "notEquals", value: "Closed" },
          { columnRef: "Status", operator: "contains", value: "Active" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Unsupported repeated column filter.",
        confidence: 0.89,
        requiresConfirmation: true
      },
      requestId: "req_filter_repeat_001",
      runId: "run_filter_repeat_001",
      approvalToken: "token"
    })).rejects.toThrow(/does not support multiple conditions for the same column/i);

    expect(filterApi.clearCriteria).not.toHaveBeenCalled();
    expect(filterApi.apply).not.toHaveBeenCalled();
  });
});
