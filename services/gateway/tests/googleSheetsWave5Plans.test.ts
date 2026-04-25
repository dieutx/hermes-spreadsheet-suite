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
      flush,
      newFilterCriteria() {
        let matchedValue: unknown = null;
        let matchedType: string | null = null;
        const bindBuilderMethod = (type: string) => vi.fn(function() {
          matchedType = type;
          matchedValue = arguments.length > 1
            ? Array.from(arguments)
            : arguments[0] == null
              ? null
              : arguments[0];
          return this;
        });
        return {
          whenTextEqualTo: bindBuilderMethod("text_equal_to"),
          whenTextNotEqualTo: bindBuilderMethod("text_not_equal_to"),
          whenNumberEqualTo: bindBuilderMethod("number_equal_to"),
          whenNumberNotEqualTo: bindBuilderMethod("number_not_equal_to"),
          whenNumberGreaterThan: bindBuilderMethod("number_greater_than"),
          whenNumberGreaterThanOrEqualTo: bindBuilderMethod("number_greater_than_or_equal_to"),
          whenNumberLessThan: bindBuilderMethod("number_less_than"),
          whenNumberLessThanOrEqualTo: bindBuilderMethod("number_less_than_or_equal_to"),
          whenNumberBetween: bindBuilderMethod("number_between"),
          whenNumberNotBetween: bindBuilderMethod("number_not_between"),
          build: vi.fn(function() {
            return { type: matchedType, value: matchedValue };
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
      WrapStrategy: {
        WRAP: "WRAP",
        CLIP: "CLIP",
        OVERFLOW: "OVERFLOW"
      }
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
  vm.runInNewContext(
    "this.__sidebarTestHooks = { state, elements, renderMessages };",
    context,
    { filename: `${SIDEBAR_PATH}#test-hooks` }
  );
  return context;
}

function createRangeStub(options: {
  a1Notation: string;
  row: number;
  column: number;
  numRows: number;
  numColumns: number;
  values?: unknown[][];
  displayValues?: string[][];
}) {
  let currentValues = (options.values || []).map((row) => [...row]);
  let currentDisplayValues = (options.displayValues || []).map((row) => [...row]);

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
    setValues: vi.fn((nextValues: unknown[][]) => {
      currentValues = nextValues.map((row) => [...row]);
      currentDisplayValues = nextValues.map((row) => row.map((value) => value == null ? "" : String(value)));
    }),
    getValues: vi.fn(() => currentValues.map((row) => [...row])),
    getDisplayValues: vi.fn(() =>
      (currentDisplayValues.length > 0 ? currentDisplayValues : currentValues.map((row) => row.map((value) => value == null ? "" : String(value))))
        .map((row) => [...row])
    ),
    getFormulas: vi.fn(() =>
      currentValues.map((row) => row.map(() => ""))
    ),
    getResizedRange: vi.fn()
  };

  Object.defineProperty(range, "values", {
    configurable: true,
    get() {
      return currentValues;
    }
  });

  return range;
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("Google Sheets wave 5 analysis, pivot, and chart plans", () => {
  it("treats chat-only analysis reports as non-write previews and resolves materialized report approvals to the real range", () => {
    const sidebar = loadSidebarContext();

    const chatOnlyResponse = {
      type: "analysis_report_plan",
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
        explanation: "Summarize the selected sales range.",
        confidence: 0.92,
        requiresConfirmation: false,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "none",
        confirmationLevel: "standard"
      }
    };
    const materializedResponse = {
      type: "analysis_report_plan",
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
          },
          {
            type: "group_breakdown",
            title: "By region",
            summary: "West leads closed-won revenue.",
            sourceRanges: ["Sales!A1:F50", "Sales!H1:J20"]
          }
        ],
        explanation: "Materialize the analysis report onto a report sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(chatOnlyResponse)).toBe(false);
    expect(sidebar.getRequiresConfirmation(chatOnlyResponse)).toBe(false);
    expect(sidebar.getResponseBodyText(chatOnlyResponse)).toBe(
      "Prepared a chat-only analysis report for Sales!A1:F50."
    );
    expect(sidebar.renderStructuredPreview(chatOnlyResponse, {
      runId: "run_analysis_chat_only",
      requestId: "req_analysis_chat_only"
    })).not.toContain("Confirm Analysis Report");

    expect(sidebar.isWritePlanResponse(materializedResponse)).toBe(true);
    expect(sidebar.getRequiresConfirmation(materializedResponse)).toBe(true);
    expect(sidebar.getResponseBodyText(materializedResponse)).toBe(
      "Prepared an analysis report preview for Sales Report!A1:D6."
    );
    expect(sidebar.getStructuredPreview(materializedResponse)).toMatchObject({
      kind: "analysis_report_plan",
      outputMode: "materialize_report",
      targetSheet: "Sales Report",
      targetRange: "A1:D6",
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"]
    });

    const analysisHtml = sidebar.renderStructuredPreview(materializedResponse, {
      runId: "run_analysis_report_preview",
      requestId: "req_analysis_report_preview"
    });
    expect(analysisHtml).toContain("Confirm Analysis Report");
    expect(analysisHtml).toContain("Sales Report!A1:D6");
    expect(analysisHtml).toContain("Revenue summary");
    expect(analysisHtml).toContain("West leads closed-won revenue.");
    expect(sidebar.buildWriteApprovalRequest({
      requestId: "req_analysis_report_preview",
      runId: "run_analysis_report_preview",
      plan: materializedResponse.data
    })).toMatchObject({
      requestId: "req_analysis_report_preview",
      runId: "run_analysis_report_preview",
      plan: {
        targetSheet: "Sales Report",
        targetRange: "A1:D6",
        affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"]
      }
    });
  });

  it("treats a demo-safe chart plan as a confirmable Google Sheets write preview", () => {
    const sidebar = loadSidebarContext();
    const response = {
      type: "chart_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Month",
        series: [
          { field: "Revenue", label: "Revenue" },
          { field: "Margin", label: "Margin" }
        ],
        title: "Revenue vs Margin",
        legendPosition: "hidden",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.getRequiresConfirmation(response)).toBe(true);

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_chart_preview",
      requestId: "req_chart_preview"
    });

    expect(html).toContain("Will create a line chart on Sales Chart!A1.");
    expect(html).toContain("Confirm Chart");
    expect(html).not.toContain("does not support exact-safe chart creation yet");
  });

  it("treats a demo-safe pivot plan as a confirmable Google Sheets write preview", () => {
    const sidebar = loadSidebarContext();
    const response = {
      type: "pivot_table_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        columnGroups: [],
        valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
        filters: [{ field: "Deals", operator: "greater_than", value: 5 }],
        sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.getRequiresConfirmation(response)).toBe(true);

    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_pivot_preview",
      requestId: "req_pivot_preview"
    });

    expect(html).toContain("Will create a pivot table on Sales Pivot!A1.");
    expect(html).toContain("Confirm Pivot Table");
    expect(html).not.toContain("does not support exact-safe pivot table creation yet");
  });

  it("fails closed in the sidebar for ambiguous pivot value sorting across both axes", () => {
    const sidebar = loadSidebarContext();

    const sortResponse = {
      type: "pivot_table_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        columnGroups: ["Quarter"],
        valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
        filters: [{ field: "Status", operator: "equal_to", value: "Closed Won" }],
        sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(sortResponse)).toBe(false);
    expect(sidebar.getRequiresConfirmation(sortResponse)).toBe(false);
    const sortHtml = sidebar.renderStructuredPreview(sortResponse, {
      runId: "run_pivot_sort_preview",
      requestId: "req_pivot_sort_preview"
    });
    expect(sortHtml).toContain("This Google Sheets flow can't sort pivot values when both row and column groups are present yet.");
    expect(sortHtml).not.toContain("Confirm Pivot Table");
  });

  it("applies a materialized analysis report in Google Sheets using the resolved range", () => {
    const targetAnchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1,
      values: [[""]]
    });
    const resolvedTargetRange = createRangeStub({
      a1Notation: "A1:D6",
      row: 1,
      column: 1,
      numRows: 6,
      numColumns: 4,
      values: Array.from({ length: 6 }, () => Array.from({ length: 4 }, () => ""))
    });
    targetAnchorRange.getResizedRange = vi.fn(() => resolvedTargetRange);

    const sheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A1");
        return targetAnchorRange;
      })
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sales Report");
        return sheet;
      })
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
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
          },
          {
            type: "group_breakdown",
            title: "By region",
            summary: "West leads closed-won revenue.",
            sourceRanges: ["Sales!A1:F50", "Sales!H1:J20"]
          }
        ],
        explanation: "Materialize the analysis report onto a report sheet.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(result).toMatchObject({
      kind: "analysis_report_update",
      hostPlatform: "google_sheets",
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report",
      targetSheet: "Sales Report",
      targetRange: "A1:D6",
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"],
      summary: "Created analysis report on Sales Report!A1:D6."
    });
    expect(resolvedTargetRange.setValues).toHaveBeenCalledWith([
      ["Analysis report", "", "", ""],
      ["Source sheet", "Sales", "", ""],
      ["Source range", "A1:F50", "", ""],
      ["Section", "Title", "Summary", "Source ranges"],
      ["summary_stats", "Revenue summary", "Average revenue is 12,500.", "Sales!A1:F50"],
      ["group_breakdown", "By region", "West leads closed-won revenue.", "Sales!A1:F50, Sales!H1:J20"]
    ]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("fails closed for unsupported Google Sheets apply branches", () => {
    const spreadsheet = {
      getSheetByName: vi.fn()
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
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
        explanation: "Summarize the selected sales range.",
        confidence: 0.92,
        requiresConfirmation: false,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "none",
        confirmationLevel: "standard"
      }
    })).toThrow("Chat-only analysis reports are not writeback eligible.");

    expect(() => applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        columnGroups: ["Quarter"],
        valueAggregations: [
          { field: "Revenue", aggregation: "sum" }
        ],
        sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toThrow("Google Sheets host cannot sort pivot values when both row and column groups are present exactly.");

    expect(() => applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region"],
        valueAggregations: [
          { field: "Revenue", aggregation: "sum" }
        ],
        filters: [
          { field: "Status", operator: "contains", value: "Closed" }
        ],
        explanation: "Build a pivot table by region.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })).toThrow("Unsupported pivot filter operator: contains");
  });

  it("applies a demo-safe chart plan through Code.gs", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A1:C20",
      row: 1,
      column: 1,
      numRows: 20,
      numColumns: 3,
      values: [
        ["Month", "Revenue", "Margin"],
        ...Array.from({ length: 19 }, () => ["Jan", 10, 5])
      ]
    });
    const categoryRange = createRangeStub({
      a1Notation: "A1:A20",
      row: 1,
      column: 1,
      numRows: 20,
      numColumns: 1,
      values: Array.from({ length: 20 }, (_, index) => [index === 0 ? "Month" : "Jan"])
    });
    const revenueRange = createRangeStub({
      a1Notation: "B1:B20",
      row: 1,
      column: 2,
      numRows: 20,
      numColumns: 1,
      values: Array.from({ length: 20 }, (_, index) => [index === 0 ? "Revenue" : 10])
    });
    const marginRange = createRangeStub({
      a1Notation: "C1:C20",
      row: 1,
      column: 3,
      numRows: 20,
      numColumns: 1,
      values: Array.from({ length: 20 }, (_, index) => [index === 0 ? "Margin" : 5])
    });
    const anchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1,
      values: [[""]]
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
      getRange: vi.fn(() => anchorRange),
      newChart: vi.fn(() => chartBuilder),
      insertChart: vi.fn()
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "Sales") {
          return {
            getRange: vi.fn((...args: unknown[]) => {
              if (args.length === 1 && args[0] === "A1:C20") {
                return sourceRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 1 && args[2] === 20 && args[3] === 1) {
                return categoryRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 2 && args[2] === 20 && args[3] === 1) {
                return revenueRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 3 && args[2] === 20 && args[3] === 1) {
                return marginRange;
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

    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });
    const result = applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Month",
        series: [
          { field: "Revenue", label: "Revenue" },
          { field: "Margin", label: "Margin" }
        ],
        title: "Revenue vs Margin",
        legendPosition: "hidden",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(1, categoryRange);
    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(2, revenueRange);
    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(3, marginRange);
    expect(chartBuilder.setChartType).toHaveBeenCalledWith("LINE");
    expect(chartBuilder.setPosition).toHaveBeenCalledWith(1, 1, 0, 0);
    expect(chartBuilder.setOption).toHaveBeenNthCalledWith(1, "title", "Revenue vs Margin");
    expect(chartBuilder.setOption).toHaveBeenNthCalledWith(2, "legend", { position: "none" });
    expect(chartBuilder.build).toHaveBeenCalledTimes(1);
    expect(chartSheet.newChart).toHaveBeenCalledTimes(1);
    expect(chartSheet.insertChart).toHaveBeenCalledWith(builtChart);
    expect(result).toMatchObject({
      kind: "chart_update",
      operation: "chart_update",
      hostPlatform: "google_sheets",
      sourceSheet: "Sales",
      sourceRange: "A1:C20",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line",
      categoryField: "Month",
      series: [
        { field: "Revenue", label: "Revenue" },
        { field: "Margin", label: "Margin" }
      ],
      title: "Revenue vs Margin",
      legendPosition: "hidden",
      summary: "Created line chart on Sales Chart!A1."
    });
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("accepts a demo-safe chart plan when a series label is omitted", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A1:C20",
      row: 1,
      column: 1,
      numRows: 20,
      numColumns: 3,
      values: [
        ["Month", "Revenue", "Margin"],
        ...Array.from({ length: 19 }, () => ["Jan", 10, 5])
      ]
    });
    const categoryRange = createRangeStub({
      a1Notation: "A1:A20",
      row: 1,
      column: 1,
      numRows: 20,
      numColumns: 1,
      values: Array.from({ length: 20 }, (_, index) => [index === 0 ? "Month" : "Jan"])
    });
    const revenueRange = createRangeStub({
      a1Notation: "B1:B20",
      row: 1,
      column: 2,
      numRows: 20,
      numColumns: 1,
      values: Array.from({ length: 20 }, (_, index) => [index === 0 ? "Revenue" : 10])
    });
    const marginRange = createRangeStub({
      a1Notation: "C1:C20",
      row: 1,
      column: 3,
      numRows: 20,
      numColumns: 1,
      values: Array.from({ length: 20 }, (_, index) => [index === 0 ? "Margin" : 5])
    });
    const anchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1,
      values: [[""]]
    });
    const builtChart = { id: "chart-2" };
    const chartBuilder = {
      addRange: vi.fn().mockReturnThis(),
      setChartType: vi.fn().mockReturnThis(),
      setPosition: vi.fn().mockReturnThis(),
      setOption: vi.fn().mockReturnThis(),
      build: vi.fn(() => builtChart)
    };
    const chartSheet = {
      getRange: vi.fn(() => anchorRange),
      newChart: vi.fn(() => chartBuilder),
      insertChart: vi.fn()
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "Sales") {
          return {
            getRange: vi.fn((...args: unknown[]) => {
              if (args.length === 1 && args[0] === "A1:C20") {
                return sourceRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 1 && args[2] === 20 && args[3] === 1) {
                return categoryRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 2 && args[2] === 20 && args[3] === 1) {
                return revenueRange;
              }
              if (args.length === 4 && args[0] === 1 && args[1] === 3 && args[2] === 20 && args[3] === 1) {
                return marginRange;
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

    const { applyWritePlan } = loadCodeModule({ spreadsheet });
    const result = applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Month",
        series: [
          { field: "Revenue" },
          { field: "Margin", label: "Margin" }
        ],
        title: "Revenue vs Margin",
        legendPosition: "hidden",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(1, categoryRange);
    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(2, revenueRange);
    expect(chartBuilder.addRange).toHaveBeenNthCalledWith(3, marginRange);
    expect(chartBuilder.setOption).toHaveBeenNthCalledWith(1, "title", "Revenue vs Margin");
    expect(chartBuilder.setOption).toHaveBeenNthCalledWith(2, "legend", { position: "none" });
    expect(chartSheet.insertChart).toHaveBeenCalledWith(builtChart);
    expect(result).toMatchObject({
      kind: "chart_update",
      operation: "chart_update",
      hostPlatform: "google_sheets",
      sourceSheet: "Sales",
      sourceRange: "A1:C20",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line",
      series: [
        { field: "Revenue" },
        { field: "Margin" }
      ],
      summary: "Created line chart on Sales Chart!A1."
    });
  });

  it("applies exact-safe chart series labels when the plan renames legend entries", () => {
    const sourceRange = createRangeStub({
      a1Notation: "A1:C20",
      row: 1,
      column: 1,
      numRows: 20,
      numColumns: 3,
      values: [
        ["Month", "Revenue", "Margin"],
        ...Array.from({ length: 19 }, () => ["Jan", 10, 5])
      ]
    });
    const anchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1,
      values: [[""]]
    });
    const chartBuilder = {
      addRange: vi.fn().mockReturnThis(),
      setChartType: vi.fn().mockReturnThis(),
      setPosition: vi.fn().mockReturnThis(),
      setOption: vi.fn().mockReturnThis(),
      build: vi.fn()
    };
    const chartSheet = {
      getRange: vi.fn(() => anchorRange),
      newChart: vi.fn(() => chartBuilder),
      insertChart: vi.fn()
    };
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "Sales") {
          return { getRange: vi.fn(() => sourceRange) };
        }
        if (sheetName === "Sales Chart") {
          return chartSheet;
        }
        return null;
      })
    };

    const { applyWritePlan } = loadCodeModule({ spreadsheet });
    const result = applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Month",
        series: [
          { field: "Revenue", label: "Gross Revenue" },
          { field: "Margin", label: "Margin" }
        ],
        title: "Revenue vs Margin",
        legendPosition: "hidden",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(chartSheet.newChart).toHaveBeenCalledTimes(1);
    expect(chartBuilder.setOption).toHaveBeenNthCalledWith(1, "title", "Revenue vs Margin");
    expect(chartBuilder.setOption).toHaveBeenNthCalledWith(2, "legend", { position: "none" });
    expect(chartBuilder.setOption).toHaveBeenNthCalledWith(3, "series", {
      0: { labelInLegend: "Gross Revenue" }
    });
    expect(result).toMatchObject({
      kind: "chart_update",
      chartType: "line",
      summary: "Created line chart on Sales Chart!A1."
    });
  });

  it("keeps chart previews confirmable when the plan renames series labels", () => {
    const sidebar = loadSidebarContext();
    const response = {
      type: "chart_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Month",
        series: [
          { field: "Revenue", label: "Gross Revenue" },
          { field: "Margin", label: "Margin" }
        ],
        title: "Revenue vs Margin",
        legendPosition: "hidden",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(sidebar.isWritePlanResponse(response)).toBe(true);
    expect(sidebar.getRequiresConfirmation(response)).toBe(true);
    const html = sidebar.renderStructuredPreview(response, {
      runId: "run_chart_label_preview",
      requestId: "req_chart_label_preview"
    });
    expect(html).toContain("Will create a line chart on Sales Chart!A1.");
    expect(html).toContain("Confirm Chart");
  });

  it("applies a demo-safe pivot plan through Code.gs", () => {
    const anchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1,
      values: [[""]]
    });
    const sourceRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6,
      displayValues: [
        ["Region", "Rep", "Quarter", "Revenue", "Status", "Deals"]
      ]
    });
    const regionGroup = {
      sortAscending: vi.fn(),
      sortDescending: vi.fn(),
      sortBy: vi.fn()
    };
    const repGroup = {
      sortAscending: vi.fn(),
      sortDescending: vi.fn(),
      sortBy: vi.fn()
    };
    const revenuePivotValue = { id: "revenue-value" };
    const pivotTable = {
      addRowGroup: vi.fn((columnIndex: number) => {
        if (columnIndex === 1) {
          return regionGroup;
        }
        if (columnIndex === 2) {
          return repGroup;
        }
        return null;
      }),
      addColumnGroup: vi.fn(),
      addPivotValue: vi.fn(() => revenuePivotValue),
      addFilter: vi.fn()
    };
    anchorRange.createPivotTable = vi.fn(() => pivotTable);

    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "Sales") {
          return { getRange: vi.fn(() => sourceRange) };
        }
        if (sheetName === "Sales Pivot") {
          return { getRange: vi.fn(() => anchorRange) };
        }
        return null;
      })
    };

    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });
    const result = applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        columnGroups: [],
        valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
        filters: [{ field: "Deals", operator: "greater_than", value: 5 }],
        sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(anchorRange.createPivotTable).toHaveBeenCalledWith(sourceRange);
    expect(pivotTable.addRowGroup).toHaveBeenCalledTimes(2);
    expect(pivotTable.addRowGroup).toHaveBeenNthCalledWith(1, 1);
    expect(pivotTable.addRowGroup).toHaveBeenNthCalledWith(2, 2);
    expect(pivotTable.addColumnGroup).not.toHaveBeenCalled();
    expect(pivotTable.addPivotValue).toHaveBeenCalledTimes(1);
    expect(pivotTable.addPivotValue).toHaveBeenCalledWith(4, "SUM");
    expect(pivotTable.addFilter).toHaveBeenCalledTimes(1);
    expect(pivotTable.addFilter).toHaveBeenCalledWith(6, {
      type: "number_greater_than",
      value: 5
    });
    expect(repGroup.sortBy).toHaveBeenCalledWith(revenuePivotValue, []);
    expect(repGroup.sortDescending).toHaveBeenCalledTimes(1);
    expect(result).toMatchObject({
      kind: "pivot_table_update",
      operation: "pivot_table_update",
      hostPlatform: "google_sheets",
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      columnGroups: [],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      filters: [{ field: "Deals", operator: "greater_than", value: 5 }],
      sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
      summary: "Created pivot table on Sales Pivot!A1."
    });
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies bounded numeric pivot filters through Code.gs", () => {
    const anchorRange = createRangeStub({
      a1Notation: "A1",
      row: 1,
      column: 1,
      numRows: 1,
      numColumns: 1,
      values: [[""]]
    });
    const sourceRange = createRangeStub({
      a1Notation: "A1:F50",
      row: 1,
      column: 1,
      numRows: 50,
      numColumns: 6,
      displayValues: [
        ["Region", "Rep", "Quarter", "Revenue", "Deals", "Discount"]
      ]
    });
    const pivotTable = {
      addRowGroup: vi.fn(),
      addColumnGroup: vi.fn(),
      addPivotValue: vi.fn(),
      addFilter: vi.fn()
    };
    anchorRange.createPivotTable = vi.fn(() => pivotTable);
    const spreadsheet = {
      getSheetByName: vi.fn((sheetName: string) => {
        if (sheetName === "Sales") {
          return { getRange: vi.fn(() => sourceRange) };
        }
        if (sheetName === "Sales Pivot") {
          return { getRange: vi.fn(() => anchorRange) };
        }
        return null;
      })
    };

    const { applyWritePlan } = loadCodeModule({ spreadsheet });
    const result = applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region"],
        columnGroups: [],
        valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
        filters: [
          { field: "Deals", operator: "between", value: 5, value2: "20" },
          { field: "Discount", operator: "not_between", value: "0.1", value2: 0.3 }
        ],
        explanation: "Build a pivot table with bounded filters.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(pivotTable.addFilter).toHaveBeenNthCalledWith(1, 5, {
      type: "number_between",
      value: [5, 20]
    });
    expect(pivotTable.addFilter).toHaveBeenNthCalledWith(2, 6, {
      type: "number_not_between",
      value: [0.1, 0.3]
    });
    expect(result).toMatchObject({
      kind: "pivot_table_update",
      operation: "pivot_table_update",
      filters: [
        { field: "Deals", operator: "between", value: 5, value2: "20" },
        { field: "Discount", operator: "not_between", value: "0.1", value2: 0.3 }
      ],
      summary: "Created pivot table on Sales Pivot!A1."
    });
  });

  it("uses Wave 5 update summaries for sidebar status and response text", () => {
    const sidebar = loadSidebarContext();
    const message = {
      role: "assistant",
      content: "Pending Wave 5 update",
      response: {
        type: "analysis_report_plan",
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
          explanation: "Materialize the analysis report onto a report sheet.",
          confidence: 0.91,
          requiresConfirmation: true,
          affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      },
      statusLine: "Working..."
    };

    expect(sidebar.getResponseBodyText({
      type: "analysis_report_update",
      data: {
        operation: "analysis_report_update",
        targetSheet: "Sales Report",
        targetRange: "A1:D6",
        summary: "Created analysis report on Sales Report!A1:D6."
      }
    })).toBe("Created analysis report on Sales Report!A1:D6.");
    expect(sidebar.getResponseBodyText({
      type: "pivot_table_update",
      data: {
        operation: "pivot_table_update",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        summary: "Created pivot table on Sales Pivot!A1."
      }
    })).toBe("Created pivot table on Sales Pivot!A1.");
    expect(sidebar.getResponseBodyText({
      type: "chart_update",
      data: {
        operation: "chart_update",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        summary: "Created line chart on Sales Chart!A1."
      }
    })).toBe("Created line chart on Sales Chart!A1.");

    expect(sidebar.getWritebackStatusLine({
      kind: "analysis_report_update",
      operation: "analysis_report_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales Report",
      targetRange: "A1:D6",
      summary: "Created analysis report on Sales Report!A1:D6."
    })).toBe("Created analysis report on Sales Report!A1:D6.");
    expect(sidebar.getWritebackStatusLine({
      kind: "pivot_table_update",
      operation: "pivot_table_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      summary: "Created pivot table on Sales Pivot!A1."
    })).toBe("Created pivot table on Sales Pivot!A1.");
    expect(sidebar.getWritebackStatusLine({
      kind: "chart_update",
      operation: "chart_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line",
      summary: "Created line chart on Sales Chart!A1."
    })).toBe("Created line chart on Sales Chart!A1.");

    sidebar.applyWritebackResultToMessage(message, {
      kind: "analysis_report_update",
      operation: "analysis_report_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales Report",
      targetRange: "A1:D6",
      summary: "Created analysis report on Sales Report!A1:D6."
    });

    expect(message.content).toBe("Created analysis report on Sales Report!A1:D6.");
    expect(message.response).toBeNull();
    expect(message.statusLine).toBe("");
  });

  it("sends the normalized resolved plan through the live confirm wiring to applyWritePlan", async () => {
    const sidebar = loadSidebarContext();
    const hooks = (sidebar as any).__sidebarTestHooks;
    const confirmButton = {
      ...createSidebarElementStub(),
      _attributes: {
        "data-confirm-run": "run_analysis_report_live_confirm",
        "data-request": "req_analysis_report_live_confirm"
      } as Record<string, string>,
      setAttribute(name: string, value: string) {
        this._attributes[name] = value;
      },
      getAttribute(name: string) {
        return this._attributes[name] ?? null;
      }
    };

    hooks.state.runtimeConfig = {
      gatewayBaseUrl: "http://gateway.test"
    };
    hooks.state.messages = [
      {
        role: "assistant",
        runId: "run_analysis_report_live_confirm",
        requestId: "req_analysis_report_live_confirm",
        response: {
          type: "analysis_report_plan",
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
              },
              {
                type: "group_breakdown",
                title: "By region",
                summary: "West leads closed-won revenue.",
                sourceRanges: ["Sales!A1:F50", "Sales!H1:J20"]
              }
            ],
            explanation: "Materialize the analysis report onto a report sheet.",
            confidence: 0.91,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
            overwriteRisk: "low",
            confirmationLevel: "standard"
          }
        },
        content: "",
        statusLine: ""
      }
    ];

    hooks.elements.messages.querySelectorAll = vi.fn((selector: string) => {
      if (selector === "[data-confirm-run]") {
        return [confirmButton];
      }
      return [];
    });

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
    const applyWritePlanSpy = vi.fn((payload: Record<string, unknown>) => {
      successHandler?.({
        kind: "analysis_report_update",
        operation: "analysis_report_update",
        hostPlatform: "google_sheets",
        targetSheet: payload.plan && (payload.plan as any).targetSheet,
        targetRange: payload.plan && (payload.plan as any).targetRange,
        summary: `Created analysis report on ${(payload.plan as any).targetSheet}!${(payload.plan as any).targetRange}.`
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
      }
    };

    hooks.renderMessages();
    await confirmButton.trigger("click");

    expect(fetchMock).toHaveBeenCalledTimes(2);
    const approveBody = JSON.parse(String(fetchMock.mock.calls[0]?.[1]?.body || "{}"));
    expect(approveBody.plan).toMatchObject({
      targetRange: "A1:D6",
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"]
    });
    expect(applyWritePlanSpy).toHaveBeenCalledWith(expect.objectContaining({
      requestId: "req_analysis_report_live_confirm",
      runId: "run_analysis_report_live_confirm",
      approvalToken: "approval-token",
      plan: expect.objectContaining({
        targetRange: "A1:D6",
        affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"]
      })
    }));
  });
});
