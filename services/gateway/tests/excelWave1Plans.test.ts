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
      displayLanguage: "en-US"
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
      getRange: vi.fn(() => targetRange),
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
      [{ key: 1, ascending: true }],
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
      1,
      { filterOn: "custom", criterion1: "=Open" }
    );
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
      2,
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
