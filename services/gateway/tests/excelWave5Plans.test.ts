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
    ChartType: {
      barClustered: "BarClustered",
      columnClustered: "ColumnClustered",
      barStacked: "BarStacked",
      columnStacked: "ColumnStacked",
      line: "Line",
      area: "Area",
      pie: "Pie",
      xyScatter: "XYScatter"
    },
    ChartSeriesBy: {
      columns: "Columns"
    },
    ChartLegendPosition: {
      top: "Top",
      bottom: "Bottom",
      left: "Left",
      right: "Right"
    },
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

describe("Excel wave 5 analysis, pivot, and chart plans", () => {
  it("keeps structured preview helpers null-safe for messages without a response", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    expect(() => taskpane.getStructuredPreview(null)).not.toThrow();
    expect(() => taskpane.getStructuredPreview(undefined)).not.toThrow();
    expect(taskpane.getStructuredPreview(null)).toBe(null);
    expect(taskpane.getStructuredPreview(undefined)).toBe(null);

    expect(() => taskpane.renderStructuredPreview(null, {})).not.toThrow();
    expect(() => taskpane.renderStructuredPreview(undefined, {})).not.toThrow();
    expect(taskpane.renderStructuredPreview(null, {})).toBe("");
    expect(taskpane.renderStructuredPreview(undefined, {})).toBe("");
  });

  it("treats chat-only analysis reports as non-write plans, expands materialized analysis report previews, enables safe pivot previews, keeps unsupported pivot non-executable, and allows exact-safe chart previews", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

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
    const pivotResponse = {
      type: "pivot_table_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        columnGroups: [],
        valueAggregations: [
          { field: "Revenue", aggregation: "sum" }
        ],
        filters: [
          { field: "Status", operator: "equal_to", value: "Closed Won" }
        ],
        sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };
    const chartResponse = {
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
        legendPosition: "bottom",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };
    const unsupportedPivotSortResponse = {
      type: "pivot_table_plan",
      data: {
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
    };

    expect(taskpane.isWritePlanResponse(chatOnlyResponse)).toBe(false);
    expect(taskpane.getRequiresConfirmation(chatOnlyResponse)).toBe(false);
    expect(taskpane.renderStructuredPreview(chatOnlyResponse, {
      runId: "run_analysis_chat_only",
      requestId: "req_analysis_chat_only"
    })).not.toContain("Confirm Analysis Report");

    expect(taskpane.isWritePlanResponse(materializedResponse)).toBe(true);
    expect(taskpane.getRequiresConfirmation(materializedResponse)).toBe(true);
    expect(taskpane.getResponseBodyText(materializedResponse)).toBe(
      "Prepared an analysis report preview for Sales Report!A1:D6."
    );
    expect(taskpane.getStructuredPreview(materializedResponse)).toMatchObject({
      kind: "analysis_report_plan",
      outputMode: "materialize_report",
      targetSheet: "Sales Report",
      targetRange: "A1:D6"
    });
    const analysisHtml = taskpane.renderStructuredPreview(materializedResponse, {
      runId: "run_analysis_report_preview",
      requestId: "req_analysis_report_preview"
    });
    expect(analysisHtml).toContain("Confirm Analysis Report");
    expect(analysisHtml).toContain("Sales Report!A1:D6");
    expect(analysisHtml).toContain("Revenue summary");
    expect(analysisHtml).toContain("West leads closed-won revenue.");
    expect(taskpane.buildWriteApprovalRequest({
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

    expect(taskpane.isWritePlanResponse(pivotResponse)).toBe(true);
    expect(taskpane.getRequiresConfirmation(pivotResponse)).toBe(true);
    expect(taskpane.getResponseBodyText(pivotResponse)).toBe(
      "Prepared a pivot table preview for Sales Pivot!A1."
    );
    const pivotHtml = taskpane.renderStructuredPreview(pivotResponse, {
      runId: "run_pivot_preview",
      requestId: "req_pivot_preview"
    });
    expect(pivotHtml).toContain("Will create a pivot table on Sales Pivot!A1.");
    expect(pivotHtml).not.toContain("This Excel runtime can't create pivot tables safely yet.");
    expect(pivotHtml).toContain("Confirm Pivot Table");

    expect(taskpane.isWritePlanResponse(unsupportedPivotSortResponse)).toBe(false);
    const unsupportedPivotSortHtml = taskpane.renderStructuredPreview(unsupportedPivotSortResponse, {
      runId: "run_pivot_sort_preview",
      requestId: "req_pivot_sort_preview"
    });
    expect(unsupportedPivotSortHtml).toContain("This Excel runtime can't sort pivot values when both row and column groups are present yet.");
    expect(unsupportedPivotSortHtml).not.toContain("Confirm Pivot Table");

    expect(taskpane.isWritePlanResponse(chartResponse)).toBe(true);
    expect(taskpane.getRequiresConfirmation(chartResponse)).toBe(true);
    expect(taskpane.getResponseBodyText(chartResponse)).toBe(
      "Prepared a chart preview for Sales Chart!A1."
    );
    const chartHtml = taskpane.renderStructuredPreview(chartResponse, {
      runId: "run_chart_preview",
      requestId: "req_chart_preview"
    });
    expect(chartHtml).toContain("Will create a line chart on Sales Chart!A1.");
    expect(chartHtml).toContain("Confirm Chart");
    expect(chartHtml).not.toContain("can't create charts safely yet");
  });

  it("applies a materialized analysis report in Excel using the same resolved plan sent for approval", async () => {
    const expandedTargetRange = createRangeStub({
      address: "Sales Report!A1:D6",
      rowCount: 6,
      columnCount: 4,
      values: Array.from({ length: 6 }, () => Array.from({ length: 4 }, () => ""))
    });
    Object.defineProperty(expandedTargetRange, "values", {
      configurable: true,
      get() {
        return Array.from({ length: 6 }, () => Array.from({ length: 4 }, () => ""));
      },
      set(nextValues) {
        expandedTargetRange.__appliedValues = nextValues;
      }
    });

    const worksheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A1:D6");
        return expandedTargetRange;
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn((sheetName: string) => {
            expect(sheetName).toBe("Sales Report");
            return worksheet;
          })
        }
      }
    });

    const approvalRequest = taskpane.buildWriteApprovalRequest({
      requestId: "req_analysis_report_apply_excel_001",
      runId: "run_analysis_report_apply_excel_001",
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

    expect(approvalRequest).toMatchObject({
      requestId: "req_analysis_report_apply_excel_001",
      runId: "run_analysis_report_apply_excel_001",
      plan: {
        targetSheet: "Sales Report",
        targetRange: "A1:D6",
        affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"]
      }
    });

    await expect(taskpane.applyWritePlan({
      plan: approvalRequest.plan,
      requestId: "req_analysis_report_apply_excel_001",
      runId: "run_analysis_report_apply_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "analysis_report_update",
      targetSheet: "Sales Report",
      targetRange: "A1:D6",
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1:D6"],
      summary: "Created analysis report on Sales Report!A1:D6."
    });

    expect(worksheet.getRange).toHaveBeenCalledWith("A1:D6");
    expect(expandedTargetRange.__appliedValues).toEqual([
      ["Analysis report", "", "", ""],
      ["Source sheet", "Sales", "", ""],
      ["Source range", "A1:F50", "", ""],
      ["Section", "Title", "Summary", "Source ranges"],
      ["summary_stats", "Revenue summary", "Average revenue is 12,500.", "Sales!A1:F50"],
      ["group_breakdown", "By region", "West leads closed-won revenue.", "Sales!A1:F50, Sales!H1:J20"]
    ]);
  });

  it("applies a safe pivot table in Excel", async () => {
    const sourceRange = createRangeStub({
      address: "Sales!A1:F50",
      rowCount: 50,
      columnCount: 6,
      values: [
        ["Region", "Rep", "Quarter", "Revenue", "Status", "Month"],
        ["West", "Ada", "Q1", 100, "Closed Won", "Jan"]
      ]
    });
    const anchorRange = createRangeStub({
      address: "Sales Pivot!A1",
      rowCount: 1,
      columnCount: 1,
      values: [[""]],
      formulas: [[""]]
    });
    const createdDataHierarchies: Array<{ name: string; summarizeBy?: string }> = [];
    const pivotFieldMap = Object.fromEntries(
      ["Region", "Rep", "Revenue", "Status"].map((name) => [
        name,
        {
          name,
          applyFilter: vi.fn(),
          sortByLabels: vi.fn(),
          sortByValues: vi.fn()
        }
      ])
    );
    const hierarchyMap = Object.fromEntries(
      ["Region", "Rep", "Revenue", "Status"].map((name) => [
        name,
        {
          name,
          fields: {
            getItem: vi.fn(() => pivotFieldMap[name])
          }
        }
      ])
    );
    const pivotTable = {
      hierarchies: {
        getItem: vi.fn((name: string) => hierarchyMap[name])
      },
      rowHierarchies: {
        add: vi.fn()
      },
      columnHierarchies: {
        add: vi.fn()
      },
      filterHierarchies: {
        add: vi.fn()
      },
      dataHierarchies: {
        add: vi.fn((hierarchy: { name: string }) => {
          const dataHierarchy = { name: hierarchy.name, summarizeBy: undefined as string | undefined };
          createdDataHierarchies.push(dataHierarchy);
          return dataHierarchy;
        }),
        getItem: vi.fn((name: string) => createdDataHierarchies.find((hierarchy) => hierarchy.name === name))
      }
    };
    const sourceWorksheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A1:F50");
        return sourceRange;
      })
    };
    const targetWorksheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A1");
        return anchorRange;
      }),
      pivotTables: {
        add: vi.fn((name: string, source: unknown, destination: unknown) => {
          expect(name).toContain("HermesPivot_Sales_Pivot_A1_");
          expect(source).toBe(sourceRange);
          expect(destination).toBe(anchorRange);
          return pivotTable;
        })
      }
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn((sheetName: string) => {
            if (sheetName === "Sales") {
              return sourceWorksheet;
            }
            if (sheetName === "Sales Pivot") {
              return targetWorksheet;
            }
            throw new Error(`Unexpected sheet lookup: ${sheetName}`);
          })
        }
      }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region", "Rep"],
        columnGroups: [],
        valueAggregations: [
          { field: "Revenue", aggregation: "sum" }
        ],
        filters: [
          { field: "Status", operator: "equal_to", value: "Closed Won" }
        ],
        sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
        explanation: "Build a pivot table by region and rep.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_pivot_apply_excel_001",
      runId: "run_pivot_apply_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "pivot_table_update",
      operation: "pivot_table_update",
      hostPlatform: "excel_windows",
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      columnGroups: [],
      valueAggregations: [
        { field: "Revenue", aggregation: "sum" }
      ],
      filters: [
        { field: "Status", operator: "equal_to", value: "Closed Won" }
      ],
      sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
      summary: "Created pivot table on Sales Pivot!A1."
    });

    expect(pivotTable.rowHierarchies.add).toHaveBeenCalledTimes(2);
    expect(pivotTable.rowHierarchies.add).toHaveBeenNthCalledWith(1, hierarchyMap.Region);
    expect(pivotTable.rowHierarchies.add).toHaveBeenNthCalledWith(2, hierarchyMap.Rep);
    expect(pivotTable.columnHierarchies.add).not.toHaveBeenCalled();
    expect(pivotTable.filterHierarchies.add).toHaveBeenCalledWith(hierarchyMap.Status);
    expect(pivotTable.dataHierarchies.add).toHaveBeenCalledWith(hierarchyMap.Revenue);
    expect(createdDataHierarchies).toEqual([
      { name: "Revenue", summarizeBy: "Sum" }
    ]);
    expect(pivotFieldMap.Status.applyFilter).toHaveBeenCalledWith({
      labelFilter: {
        condition: "Equals",
        comparator: "Closed Won"
      }
    });
    expect(pivotFieldMap.Rep.sortByValues).toHaveBeenCalledWith("Descending", createdDataHierarchies[0], []);

  });

  it("applies bounded numeric pivot filters in Excel", async () => {
    const sourceRange = createRangeStub({
      address: "Sales!A1:F50",
      rowCount: 50,
      columnCount: 6,
      values: [
        ["Region", "Rep", "Quarter", "Revenue", "Deals", "Discount"],
        ...Array.from({ length: 49 }, () => ["West", "Avery", "Q1", 100, 12, 0.2])
      ]
    });
    const anchorRange = createRangeStub({
      address: "Sales Pivot!A1",
      rowCount: 1,
      columnCount: 1,
      values: [[""]],
      formulas: [[""]]
    });
    const pivotFieldMap = Object.fromEntries(
      ["Region", "Revenue", "Deals", "Discount"].map((name) => [
        name,
        {
          name,
          applyFilter: vi.fn()
        }
      ])
    );
    const hierarchyMap = Object.fromEntries(
      ["Region", "Revenue", "Deals", "Discount"].map((name) => [
        name,
        {
          name,
          fields: {
            getItem: vi.fn(() => pivotFieldMap[name])
          }
        }
      ])
    );
    const pivotTable = {
      hierarchies: {
        getItem: vi.fn((name: string) => hierarchyMap[name])
      },
      rowHierarchies: {
        add: vi.fn()
      },
      columnHierarchies: {
        add: vi.fn()
      },
      filterHierarchies: {
        add: vi.fn()
      },
      dataHierarchies: {
        add: vi.fn((hierarchy: { name: string }) => ({ name: hierarchy.name, summarizeBy: undefined }))
      }
    };
    const sourceWorksheet = {
      getRange: vi.fn(() => sourceRange)
    };
    const targetWorksheet = {
      getRange: vi.fn(() => anchorRange),
      pivotTables: {
        add: vi.fn(() => pivotTable)
      }
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn((sheetName: string) => {
            if (sheetName === "Sales") {
              return sourceWorksheet;
            }
            if (sheetName === "Sales Pivot") {
              return targetWorksheet;
            }
            throw new Error(`Unexpected sheet lookup: ${sheetName}`);
          })
        }
      }
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Sales Pivot",
        targetRange: "A1",
        rowGroups: ["Region"],
        columnGroups: [],
        valueAggregations: [
          { field: "Revenue", aggregation: "sum" }
        ],
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
      },
      requestId: "req_pivot_between_apply_excel_001",
      runId: "run_pivot_between_apply_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "pivot_table_update",
      operation: "pivot_table_update",
      filters: [
        { field: "Deals", operator: "between", value: 5, value2: "20" },
        { field: "Discount", operator: "not_between", value: "0.1", value2: 0.3 }
      ],
      summary: "Created pivot table on Sales Pivot!A1."
    });

    expect(pivotTable.filterHierarchies.add).toHaveBeenNthCalledWith(1, hierarchyMap.Deals);
    expect(pivotTable.filterHierarchies.add).toHaveBeenNthCalledWith(2, hierarchyMap.Discount);
    expect(pivotFieldMap.Deals.applyFilter).toHaveBeenCalledWith({
      labelFilter: {
        condition: "Between",
        lowerBound: "5",
        upperBound: "20"
      }
    });
    expect(pivotFieldMap.Discount.applyFilter).toHaveBeenCalledWith({
      labelFilter: {
        condition: "Between",
        lowerBound: "0.1",
        upperBound: "0.3",
        exclusive: true
      }
    });
  });

  it("applies an exact-safe chart plan through Excel Office.js", async () => {
    const sourceRange = createRangeStub({
      address: "Sales!A1:C20",
      rowCount: 20,
      columnCount: 3,
      values: [
        ["Month", "Revenue", "Margin"],
        ...Array.from({ length: 19 }, () => ["Jan", 10, 5])
      ]
    });
    const targetAnchorRange = createRangeStub({
      address: "Sales Chart!A1",
      rowCount: 1,
      columnCount: 1,
      values: [[""]]
    });
    const chart = {
      title: {
        text: "",
        visible: false
      },
      legend: {
        position: "",
        visible: true
      },
      setPosition: vi.fn()
    };
    const chartSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A1");
        return targetAnchorRange;
      }),
      charts: {
        add: vi.fn((chartType: string, range: unknown, seriesBy: string) => {
          expect(chartType).toBe("Line");
          expect(range).toBe(sourceRange);
          expect(seriesBy).toBe("Columns");
          return chart;
        })
      }
    };
    const sourceSheet = {
      getRange: vi.fn((rangeName: string) => {
        expect(rangeName).toBe("A1:C20");
        return sourceRange;
      })
    };
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn((sheetName: string) => {
            if (sheetName === "Sales") {
              return sourceSheet;
            }

            if (sheetName === "Sales Chart") {
              return chartSheet;
            }

            throw new Error(`Unexpected sheet lookup: ${sheetName}`);
          })
        }
      }
    });

    await expect(taskpane.applyWritePlan({
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
        legendPosition: "bottom",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      },
      requestId: "req_chart_apply_excel_001",
      runId: "run_chart_apply_excel_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "chart_update",
      operation: "chart_update",
      hostPlatform: "excel_windows",
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
      legendPosition: "bottom",
      summary: "Created line chart on Sales Chart!A1."
    });

    expect(sourceSheet.getRange).toHaveBeenCalledWith("A1:C20");
    expect(chartSheet.getRange).toHaveBeenCalledWith("A1");
    expect(chartSheet.charts.add).toHaveBeenCalledTimes(1);
    expect(chart.setPosition).toHaveBeenCalledWith(targetAnchorRange);
    expect(chart.title.visible).toBe(true);
    expect(chart.title.text).toBe("Revenue vs Margin");
    expect(chart.legend.visible).toBe(true);
    expect(chart.legend.position).toBe("Bottom");
  });

  it("fails closed for chart previews and apply paths when the plan would rename series labels", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn()
        }
      }
    });

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
        legendPosition: "bottom",
        explanation: "Chart monthly revenue and margin.",
        confidence: 0.93,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(false);
    expect(taskpane.getRequiresConfirmation(response)).toBe(true);
    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_chart_label_preview_excel_001",
      requestId: "req_chart_label_preview_excel_001"
    });
    expect(html).toContain("can't rename chart series labels during creation");
    expect(html).not.toContain("Confirm Chart");

    await expect(taskpane.applyWritePlan({
      plan: response.data,
      requestId: "req_chart_label_apply_excel_001",
      runId: "run_chart_label_apply_excel_001",
      approvalToken: "token"
    })).rejects.toThrow("Excel host can't rename chart series labels during creation.");
  });
});
