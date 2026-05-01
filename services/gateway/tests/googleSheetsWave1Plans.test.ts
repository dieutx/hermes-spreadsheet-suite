import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { describe, expect, it, vi } from "vitest";

const require = createRequire(import.meta.url);
const TEST_FILE_PATH = fileURLToPath(import.meta.url);
const TEST_DIR = path.dirname(TEST_FILE_PATH);
const REPO_ROOT = path.resolve(TEST_DIR, "../../..");
const CODE_PATH = path.join(REPO_ROOT, "apps/google-sheets-addon/src/Code.gs");
const SIDEBAR_PATH = path.join(REPO_ROOT, "apps/google-sheets-addon/html/Sidebar.js.html");
const FORBIDDEN_ABSOLUTE_PREFIX = ["", "root", "claude"].join("/");
const TEST_SOURCE = fs.readFileSync(TEST_FILE_PATH, "utf8");
const codeScript = fs.readFileSync(CODE_PATH, "utf8");
const sidebarHtml = fs.readFileSync(SIDEBAR_PATH, "utf8");
const sidebarScript = sidebarHtml.match(/<script>([\s\S]*)<\/script>/)?.[1] || "";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const {
  isSheetStructurePlan_,
  isRangeSortPlan_,
  isRangeFilterPlan_,
  buildSortSpec_,
  getSheetStructureStatusSummary_,
  getRangeSortStatusSummary_,
  getRangeFilterStatusSummary_
} = require(path.join(REPO_ROOT, "apps/google-sheets-addon/src/Wave1Plans.js"));

function createFilterCriteriaBuilder() {
  return {
    _built: null,
    build() {
      return this._built;
    },
    setHiddenValues(values: unknown[]) {
      this._built = { kind: "hiddenValues", values };
      return this;
    },
    whenTextEqualTo(value: string) {
      this._built = { kind: "textEquals", value };
      return this;
    },
    whenTextContains(value: string) {
      this._built = { kind: "textContains", value };
      return this;
    },
    whenTextStartsWith(value: string) {
      this._built = { kind: "textStartsWith", value };
      return this;
    },
    whenTextEndsWith(value: string) {
      this._built = { kind: "textEndsWith", value };
      return this;
    },
    whenNumberEqualTo(value: number) {
      this._built = { kind: "numberEquals", value };
      return this;
    },
    whenNumberNotEqualTo(value: number) {
      this._built = { kind: "numberNotEquals", value };
      return this;
    },
    whenNumberGreaterThan(value: number) {
      this._built = { kind: "numberGreaterThan", value };
      return this;
    },
    whenNumberGreaterThanOrEqualTo(value: number) {
      this._built = { kind: "numberGreaterThanOrEqual", value };
      return this;
    },
    whenNumberLessThan(value: number) {
      this._built = { kind: "numberLessThan", value };
      return this;
    },
    whenNumberLessThanOrEqualTo(value: number) {
      this._built = { kind: "numberLessThanOrEqual", value };
      return this;
    },
    whenCellEmpty() {
      this._built = { kind: "cellEmpty" };
      return this;
    },
    whenCellNotEmpty() {
      this._built = { kind: "cellNotEmpty" };
      return this;
    }
  };
}

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
      flush,
      newFilterCriteria() {
        return createFilterCriteriaBuilder();
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

function loadSidebarContext() {
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

describe("Google Sheets wave 1 helper compilation", () => {
  it("does not hard-code repo paths in this test file", () => {
    expect(TEST_SOURCE).not.toContain(FORBIDDEN_ABSOLUTE_PREFIX);
  });

  it("detects wave 1 structure, sort, and filter plans", () => {
    expect(
      isSheetStructurePlan_({
        operation: "merge_cells",
        targetSheet: "Sheet1",
        targetRange: "A1:C1"
      })
    ).toBe(true);

    expect(
      isRangeSortPlan_({
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        keys: [{ columnRef: "Status", direction: "asc" }]
      })
    ).toBe(true);

    expect(
      isRangeFilterPlan_({
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        conditions: [{ columnRef: "Status", operator: "equals", value: "Open" }],
        combiner: "and"
      })
    ).toBe(true);

    expect(isSheetStructurePlan_({ targetRange: "A1:C1" })).toBe(false);
    expect(isRangeSortPlan_({ targetSheet: "Sheet1", targetRange: "A1:F25", keys: [] })).toBe(false);
    expect(
      isRangeFilterPlan_({
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        conditions: []
      })
    ).toBe(false);
  });

  it("compiles Google Sheets sort specs only from positive numeric keys", () => {
    expect(
      buildSortSpec_({
        hasHeader: true,
        keys: [
          { columnRef: 2, direction: "desc" }
        ]
      })
    ).toEqual([
      { dimensionIndex: 1, sortOrder: "DESCENDING" }
    ]);
  });

  it("does not compile invalid or unresolved sort keys", () => {
    expect(
      buildSortSpec_({
        hasHeader: true,
        keys: [
          { columnRef: 0, direction: "asc" },
          { columnRef: -1, direction: "desc" },
          { columnRef: "", direction: "asc" },
          { columnRef: "Status", direction: "desc" },
          { columnRef: 3, direction: "asc" }
        ]
      })
    ).toEqual([
      { dimensionIndex: 2, sortOrder: "ASCENDING" }
    ]);
  });

  it("builds readable Google Sheets status summaries", () => {
    expect(
      getSheetStructureStatusSummary_({
        operation: "insert_rows",
        targetSheet: "Sheet1",
        startIndex: 7,
        count: 3
      })
    ).toBe("Inserted 3 rows at Sheet1 row 8.");

    expect(
      getRangeSortStatusSummary_({
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        keys: [
          { columnRef: "Status", direction: "asc" },
          { columnRef: "Due Date", direction: "desc" }
        ]
      })
    ).toBe("Sorted Sheet1!A1:F25 by Status asc, Due Date desc.");

    expect(
      getRangeFilterStatusSummary_({
        targetSheet: "Sheet1",
        targetRange: "A1:F25"
      })
    ).toBe("Applied filter to Sheet1!A1:F25");
  });

  it("supports exact text notEquals filters via hidden values", () => {
    const setColumnFilterCriteria = vi.fn();
    const createFilter = vi.fn(() => ({
      setColumnFilterCriteria,
      removeColumnFilterCriteria: vi.fn()
    }));
    const targetRange = {
      getNumColumns() {
        return 2;
      },
      getNumRows() {
        return 4;
      },
      getColumn() {
        return 3;
      },
      getRow() {
        return 1;
      },
      getA1Notation() {
        return "C1:D4";
      },
      getDisplayValues() {
        return [
          ["Status", "Amount"],
          ["Open", "10"],
          ["Closed", "12"],
          ["Open", "7"]
        ];
      },
      createFilter
    };
    const sheet = {
      getFilter() {
        return null;
      },
      getRange(rangeA1: string) {
        expect(rangeA1).toBe("C1:D4");
        return targetRange;
      }
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "C1:D4",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "notEquals", value: "Closed" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep rows whose status is not exactly Closed.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    });

    expect(result).toMatchObject({
      kind: "range_filter",
      targetSheet: "Sheet1",
      targetRange: "C1:D4",
      hasHeader: true,
      conditions: [
        { columnRef: "Status", operator: "notEquals", value: "Closed" }
      ],
      combiner: "and",
      clearExistingFilters: true
    });
    expect(createFilter).toHaveBeenCalledTimes(1);
    expect(setColumnFilterCriteria).toHaveBeenCalledWith(1, {
      kind: "hiddenValues",
      values: ["Closed"]
    });
  });

  it("applies exact top-N Google Sheets filters with hidden values", () => {
    const setColumnFilterCriteria = vi.fn();
    const createFilter = vi.fn(() => ({
      setColumnFilterCriteria,
      removeColumnFilterCriteria: vi.fn()
    }));
    const targetRange = {
      getNumColumns() {
        return 2;
      },
      getNumRows() {
        return 5;
      },
      getColumn() {
        return 3;
      },
      getRow() {
        return 1;
      },
      getA1Notation() {
        return "C1:D5";
      },
      getValues() {
        return [
          ["Status", "Amount"],
          ["Open", 10],
          ["Closed", 12],
          ["Open", 7],
          ["Pending", ""]
        ];
      },
      getDisplayValues() {
        return [
          ["Status", "Amount"],
          ["Open", "10"],
          ["Closed", "12"],
          ["Open", "7"],
          ["Pending", ""]
        ];
      },
      createFilter
    };
    const sheet = {
      getFilter() {
        return null;
      },
      getRange(rangeA1: string) {
        expect(rangeA1).toBe("C1:D5");
        return targetRange;
      }
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "C1:D5",
        hasHeader: true,
        conditions: [
          { columnRef: "Amount", operator: "topN", value: 2 }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep the top two amounts.",
        confidence: 0.88,
        requiresConfirmation: true
      }
    });

    expect(result).toMatchObject({
      kind: "range_filter",
      targetSheet: "Sheet1",
      targetRange: "C1:D5",
      conditions: [
        { columnRef: "Amount", operator: "topN", value: 2 }
      ]
    });
    expect(setColumnFilterCriteria).toHaveBeenCalledWith(2, {
      kind: "hiddenValues",
      values: ["7", ""]
    });
  });

  it("attaches local undo snapshots for Google Sheets range filter writes", () => {
    const criteriaByColumn = new Map<number, unknown>([
      [1, { kind: "textEquals", value: "Closed" }]
    ]);
    const existingFilter = {
      getRange() {
        return {
          getA1Notation() {
            return "C1:D4";
          }
        };
      },
      getColumnFilterCriteria: vi.fn((columnPosition: number) =>
        criteriaByColumn.get(columnPosition) || null
      ),
      removeColumnFilterCriteria: vi.fn((columnPosition: number) => {
        criteriaByColumn.delete(columnPosition);
      }),
      setColumnFilterCriteria: vi.fn((columnPosition: number, criteria: unknown) => {
        criteriaByColumn.set(columnPosition, criteria);
      })
    };
    const targetRange = {
      getNumColumns() {
        return 2;
      },
      getNumRows() {
        return 4;
      },
      getColumn() {
        return 3;
      },
      getRow() {
        return 1;
      },
      getA1Notation() {
        return "C1:D4";
      },
      getDisplayValues() {
        return [
          ["Status", "Amount"],
          ["Open", "10"],
          ["Closed", "12"],
          ["Open", "7"]
        ];
      },
      createFilter: vi.fn(() => existingFilter)
    };
    const sheet = {
      getFilter() {
        return existingFilter;
      },
      getRange(rangeA1: string) {
        expect(rangeA1).toBe("C1:D4");
        return targetRange;
      }
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      executionId: "exec_filter_snapshot_sheets_001",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "C1:D4",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep open work.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    });

    expect(result).toMatchObject({
      kind: "range_filter",
      targetSheet: "Sheet1",
      targetRange: "C1:D4",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_filter_snapshot_sheets_001",
        kind: "range_filter",
        targetSheet: "Sheet1",
        targetRange: "C1:D4",
        beforeFilter: {
          exists: true,
          targetRange: "C1:D4",
          criteria: [{ kind: "textEquals", value: "Closed" }, null]
        },
        afterFilter: {
          exists: true,
          targetRange: "C1:D4",
          criteria: [{ kind: "textEquals", value: "Open" }, null]
        }
      }
    });
    expect(existingFilter.removeColumnFilterCriteria).toHaveBeenCalledTimes(2);
    expect(existingFilter.setColumnFilterCriteria).toHaveBeenCalledWith(1, {
      kind: "textEquals",
      value: "Open"
    });
  });

  it("fails safely instead of discarding an existing different-range filter when clearExistingFilters is false", () => {
    const remove = vi.fn();
    const existingFilter = {
      getRange() {
        return {
          getA1Notation() {
            return "A1:B5";
          }
        };
      },
      remove,
      removeColumnFilterCriteria: vi.fn(),
      setColumnFilterCriteria: vi.fn()
    };
    const targetRange = {
      getNumColumns() {
        return 2;
      },
      getNumRows() {
        return 5;
      },
      getColumn() {
        return 3;
      },
      getRow() {
        return 1;
      },
      getA1Notation() {
        return "C1:D5";
      },
      getDisplayValues() {
        return [
          ["Status", "Amount"],
          ["Open", "10"],
          ["Closed", "12"],
          ["Open", "7"],
          ["Open", "5"]
        ];
      },
      createFilter: vi.fn()
    };
    const sheet = {
      getFilter() {
        return existingFilter;
      },
      getRange() {
        return targetRange;
      }
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "C1:D5",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" }
        ],
        combiner: "and",
        clearExistingFilters: false,
        explanation: "Filter open rows in the target block only.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    })).toThrow("Existing sheet filter applies to A1:B5. Rejecting this plan without clearExistingFilters=true.");

    expect(remove).not.toHaveBeenCalled();
  });

  it("fails closed instead of partially applying Google Sheets sort plans with unresolved keys", () => {
    const sort = vi.fn();
    const targetRange = {
      getNumColumns() {
        return 2;
      },
      getNumRows() {
        return 4;
      },
      getColumn() {
        return 1;
      },
      getRow() {
        return 1;
      },
      getA1Notation() {
        return "A1:B4";
      },
      getValues() {
        return [
          ["Status", "Amount"],
          ["Open", 10],
          ["Closed", 12],
          ["Open", 7]
        ];
      },
      getFormulas() {
        return [
          ["", ""],
          ["", ""],
          ["", ""],
          ["", ""]
        ];
      },
      getDisplayValues() {
        return [
          ["Status", "Amount"],
          ["Open", "10"],
          ["Closed", "12"],
          ["Open", "7"]
        ];
      }
    };
    const sheet = {
      getRange(...args: unknown[]) {
        if (args.length === 1) {
          expect(args[0]).toBe("A1:B4");
          return targetRange;
        }

        return {
          sort
        };
      }
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      executionId: "exec_sort_unresolved_001",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:B4",
        hasHeader: true,
        keys: [
          { columnRef: "Status", direction: "asc" },
          { columnRef: "Missing Header", direction: "desc" }
        ],
        explanation: "Sort by status and a missing field.",
        confidence: 0.88,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host could not resolve sort key Missing Header inside the target range.");

    expect(sort).not.toHaveBeenCalled();
  });

  it("fails safely for grid-filter semantics that cannot be represented exactly", () => {
    const targetRange = {
      getNumColumns() {
        return 2;
      },
      getNumRows() {
        return 4;
      },
      getColumn() {
        return 3;
      },
      getRow() {
        return 1;
      },
      getA1Notation() {
        return "C1:D4";
      },
      getDisplayValues() {
        return [
          ["Status", "Amount"],
          ["Closed", "12"],
          ["Open", "10"],
          ["Pending", "10"]
        ];
      },
      createFilter: vi.fn(() => ({
        setColumnFilterCriteria: vi.fn(),
        removeColumnFilterCriteria: vi.fn()
      }))
    };
    const sheet = {
      getFilter() {
        return null;
      },
      getRange() {
        return targetRange;
      }
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "C1:D4",
        hasHeader: true,
        conditions: [
          { columnRef: "Amount", operator: "topN", value: 2 }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep the top two amounts.",
        confidence: 0.88,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets grid filters cannot represent topN exactly when duplicate display values cross the cutoff.");

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "C1:D4",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "Amount", operator: "greaterThan", value: 10 }
        ],
        combiner: "or",
        clearExistingFilters: true,
        explanation: "Keep open rows or large amounts.",
        confidence: 0.87,
        requiresConfirmation: true
      }
    })).toThrow('Google Sheets grid filters cannot represent combiner "or" exactly for multiple conditions.');
  });

  it("fails closed for out-of-range numeric relative column refs on offset ranges", () => {
    const { resolveRelativeColumnRef_ } = loadCodeModule();
    const targetRange = {
      getNumColumns() {
        return 2;
      },
      getColumn() {
        return 3;
      },
      getDisplayValues() {
        return [["Status", "Amount"]];
      }
    };

    expect(resolveRelativeColumnRef_(2, targetRange, true)).toBe(2);
    expect(resolveRelativeColumnRef_(3, targetRange, true)).toBeNull();
    expect(resolveRelativeColumnRef_("3", targetRange, true)).toBeNull();
  });

  it("captures Google Sheets notes in local execution snapshots for undo and redo", () => {
    const values = [["A", "B"]];
    const formulas = [["", ""]];
    const notes = [["old note", ""]];
    const getCell = (row: number, column: number) => ({
      setFormula(value: string) {
        formulas[row - 1][column - 1] = value;
      },
      setValue(value: unknown) {
        values[row - 1][column - 1] = value;
        formulas[row - 1][column - 1] = "";
      },
      setNote(value: string) {
        notes[row - 1][column - 1] = value;
      }
    });
    const targetRange = {
      getNumRows() {
        return 1;
      },
      getNumColumns() {
        return 2;
      },
      getValues() {
        return values.map((row) => row.slice());
      },
      getFormulas() {
        return formulas.map((row) => row.slice());
      },
      getNotes() {
        return notes.map((row) => row.slice());
      },
      setNotes(nextNotes: string[][]) {
        nextNotes.forEach((row, rowIndex) => {
          row.forEach((note, columnIndex) => {
            notes[rowIndex][columnIndex] = note;
          });
        });
      },
      getCell
    };
    const sheet = {
      getRange(rangeA1: string) {
        expect(rangeA1).toBe("A1:B1");
        return targetRange;
      }
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      }
    };
    const { applyWritePlan, applyExecutionCellSnapshot } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      executionId: "exec_notes_sheets_001",
      plan: {
        targetSheet: "Sheet1",
        targetRange: "A1:B1",
        operation: "set_notes",
        shape: { rows: 1, columns: 2 },
        notes: [["new note", ""]],
        explanation: "Update one cell note and clear the adjacent note.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    });

    const snapshot = result.__hermesLocalExecutionSnapshot;
    expect(snapshot.beforeCells[0][0]).toMatchObject({ kind: "value", note: "old note" });
    expect(snapshot.beforeCells[0][1]).toMatchObject({ kind: "value", note: "" });
    expect(snapshot.afterCells[0][0]).toMatchObject({ kind: "value", note: "new note" });
    expect(snapshot.afterCells[0][1]).toMatchObject({ kind: "value", note: "" });

    applyExecutionCellSnapshot({
      targetSheet: "Sheet1",
      targetRange: "A1:B1",
      cells: snapshot.beforeCells
    });
    expect(notes).toEqual([["old note", ""]]);

    applyExecutionCellSnapshot({
      targetSheet: "Sheet1",
      targetRange: "A1:B1",
      cells: snapshot.afterCells
    });
    expect(notes).toEqual([["new note", ""]]);
  });

  it("renders detailed sort and filter previews including keys, conditions, and filter reset state", () => {
    const sidebar = loadSidebarContext();
    const sortPreview = sidebar.renderStructuredPreview({
      type: "range_sort_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        keys: [
          { columnRef: "Status", direction: "asc" },
          { columnRef: "Due Date", direction: "desc" }
        ],
        explanation: "Sort open items first, then latest due date.",
        confidence: 0.93,
        requiresConfirmation: true
      }
    }, {
      runId: "run_sort",
      requestId: "req_sort"
    });

    expect(sortPreview).toContain("Status asc");
    expect(sortPreview).toContain("Due Date desc");

    const filterPreview = sidebar.renderStructuredPreview({
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "notEquals", value: "Closed" },
          { columnRef: "Amount", operator: "greaterThan", value: 10 }
        ],
        combiner: "and",
        clearExistingFilters: false,
        explanation: "Keep rows that are not closed and above the threshold.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    }, {
      runId: "run_filter",
      requestId: "req_filter"
    });

    expect(filterPreview).toContain("Status notEquals Closed");
    expect(filterPreview).toContain("Amount greaterThan 10");
    expect(filterPreview).toContain("Preserve existing filters");
  });

  it("fails closed in preview for unsupported Google Sheets filter combinations", () => {
    const sidebar = loadSidebarContext();
    const unsupportedOrResponse = {
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
        confidence: 0.88,
        requiresConfirmation: true
      }
    };
    const topNResponse = {
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Amount", operator: "topN", value: 2 }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep the top two amounts.",
        confidence: 0.88,
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
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "status", operator: "notEquals", value: "Closed" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Keep only active status rows.",
        confidence: 0.88,
        requiresConfirmation: true
      }
    };

    expect(sidebar.isWritePlanResponse(unsupportedOrResponse)).toBe(false);
    expect(sidebar.renderStructuredPreview(unsupportedOrResponse, {
      runId: "run_filter_preview_unsupported_google_or",
      requestId: "req_filter_preview_unsupported_google_or"
    })).toContain("can't combine multiple filter conditions with OR exactly");

    expect(sidebar.isWritePlanResponse(topNResponse)).toBe(true);
    expect(sidebar.renderStructuredPreview(topNResponse, {
      runId: "run_filter_preview_unsupported_google_topn",
      requestId: "req_filter_preview_unsupported_google_topn"
    })).toContain("Confirm Filter");

    const duplicateHtml = sidebar.renderStructuredPreview(duplicateColumnResponse, {
      runId: "run_filter_preview_unsupported_google_duplicate",
      requestId: "req_filter_preview_unsupported_google_duplicate"
    });

    expect(sidebar.isWritePlanResponse(duplicateColumnResponse)).toBe(false);
    expect(duplicateHtml).toContain("same filter column");
    expect(duplicateHtml).not.toContain("Confirm Filter");
  });

  it("renders detailed sheet structure previews with operation-specific details", () => {
    const sidebar = loadSidebarContext();
    const insertPreview = sidebar.renderStructuredPreview({
      type: "sheet_structure_update",
      data: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 7,
        count: 3,
        explanation: "Insert spacing rows below the header.",
        confirmationLevel: "standard",
        confidence: 0.91,
        requiresConfirmation: true
      }
    }, {
      runId: "run_insert_rows",
      requestId: "req_insert_rows"
    });

    expect(insertPreview).toContain("Rows 8-10");

    const freezePreview = sidebar.renderStructuredPreview({
      type: "sheet_structure_update",
      data: {
        targetSheet: "Sheet1",
        operation: "freeze_panes",
        frozenRows: 1,
        frozenColumns: 2,
        explanation: "Keep headers pinned while scrolling.",
        confirmationLevel: "standard",
        confidence: 0.89,
        requiresConfirmation: true
      }
    }, {
      runId: "run_freeze_panes",
      requestId: "req_freeze_panes"
    });

    expect(freezePreview).toContain("Frozen rows 1");
    expect(freezePreview).toContain("Frozen columns 2");
  });

  it("routes wave 1 plans through Code.gs apply handling", () => {
    expect(codeScript).toContain("isSheetStructurePlan_(plan)");
    expect(codeScript).toContain("kind: 'sheet_structure_update'");
    expect(codeScript).toContain("getSheetStructureStatusSummary_(plan)");

    expect(codeScript).toContain("isRangeSortPlan_(plan)");
    expect(codeScript).toContain("kind: 'range_sort'");
    expect(codeScript).toContain("getRangeSortStatusSummary_(plan)");

    expect(codeScript).toContain("isRangeFilterPlan_(plan)");
    expect(codeScript).toContain("kind: 'range_filter'");
    expect(codeScript).toContain("getRangeFilterStatusSummary_(plan)");
  });
});
