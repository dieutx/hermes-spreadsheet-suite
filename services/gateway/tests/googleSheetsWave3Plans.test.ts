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

function createConditionalFormatBuilder() {
  return {
    _rule: null as null | Record<string, unknown>,
    _ranges: [] as unknown[],
    _format: {} as Record<string, unknown>,
    _gradient: {} as Record<string, unknown>,
    whenTextContains(text: string) {
      this._rule = {
        kind: "text_contains",
        text
      };
      return this;
    },
    whenFormulaSatisfied(formula: string) {
      this._rule = {
        kind: "custom_formula",
        formula
      };
      return this;
    },
    whenNumberBetween(value: number, value2: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "between",
        value,
        value2
      };
      return this;
    },
    whenNumberNotBetween(value: number, value2: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "not_between",
        value,
        value2
      };
      return this;
    },
    whenNumberEqualTo(value: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "equal_to",
        value
      };
      return this;
    },
    whenNumberNotEqualTo(value: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "not_equal_to",
        value
      };
      return this;
    },
    whenNumberGreaterThan(value: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "greater_than",
        value
      };
      return this;
    },
    whenNumberGreaterThanOrEqualTo(value: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "greater_than_or_equal_to",
        value
      };
      return this;
    },
    whenNumberLessThan(value: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "less_than",
        value
      };
      return this;
    },
    whenNumberLessThanOrEqualTo(value: number) {
      this._rule = {
        kind: "number_compare",
        comparator: "less_than_or_equal_to",
        value
      };
      return this;
    },
    setBackground(color: string) {
      this._format.backgroundColor = color;
      return this;
    },
    setFontColor(color: string) {
      this._format.textColor = color;
      return this;
    },
    setBold(value: boolean) {
      this._format.bold = value;
      return this;
    },
    setItalic(value: boolean) {
      this._format.italic = value;
      return this;
    },
    setUnderline(value: boolean) {
      this._format.underline = value;
      return this;
    },
    setStrikethrough(value: boolean) {
      this._format.strikethrough = value;
      return this;
    },
    setGradientMinpoint(color: string) {
      this._gradient.min = { type: "min", color };
      return this;
    },
    setGradientMinpointWithValue(color: string, type: string, value: string) {
      this._gradient.min = { type, color, value };
      return this;
    },
    setGradientMidpoint(color: string) {
      this._gradient.mid = { type: "min", color };
      return this;
    },
    setGradientMidpointWithValue(color: string, type: string, value: string) {
      this._gradient.mid = { type, color, value };
      return this;
    },
    setGradientMaxpoint(color: string) {
      this._gradient.max = { type: "max", color };
      return this;
    },
    setGradientMaxpointWithValue(color: string, type: string, value: string) {
      this._gradient.max = { type, color, value };
      return this;
    },
    setRanges(ranges: unknown[]) {
      this._ranges = ranges.slice();
      return this;
    },
    build() {
      const gradientPoints = ["min", "mid", "max"]
        .map((key) => this._gradient[key])
        .filter(Boolean);
      const builtRule = {
        rule: gradientPoints.length > 0
          ? {
              kind: "color_scale",
              points: gradientPoints.map((point) => ({ ...point }))
            }
          : this._rule,
        format: { ...this._format },
        ranges: this._ranges.slice(),
        getRanges() {
          return this.ranges.slice();
        }
      };
      return builtRule;
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
      InterpolationType: {
        NUMBER: "NUMBER",
        PERCENT: "PERCENT",
        PERCENTILE: "PERCENTILE"
      },
      newConditionalFormatRule() {
        return createConditionalFormatBuilder();
      },
      newDataValidation() {
        return {
          build() {
            return {};
          }
        };
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

function createRangeStub(a1Notation: string, row: number, column: number, numRows: number, numColumns: number) {
  return {
    getA1Notation() {
      return a1Notation;
    },
    getRow() {
      return row;
    },
    getColumn() {
      return column;
    },
    getNumRows() {
      return numRows;
    },
    getNumColumns() {
      return numColumns;
    }
  };
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe("Google Sheets wave 3 conditional-format plans", () => {
  it("renders conditional_format_plan previews and suppresses confirmation for unsupported mappings", () => {
    const sidebar = loadSidebarContext();

    const supportedResponse = {
      type: "conditional_format_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "replace_all_on_target",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc",
          textColor: "#990000",
          bold: true,
          italic: false,
          underline: true,
          strikethrough: true
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: true
      }
    };

    expect(sidebar.isWritePlanResponse(supportedResponse)).toBe(true);
    expect(sidebar.getRequiresConfirmation(supportedResponse)).toBe(true);
    expect(sidebar.getResponseBodyText(supportedResponse)).toBe(
      "Prepared a conditional formatting plan for Sheet1!B2:B20."
    );
    expect(sidebar.getStructuredPreview(supportedResponse)).toMatchObject({
      kind: "conditional_format_plan",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target",
      ruleType: "text_contains",
      summary: "Will replace all conditional formatting on Sheet1!B2:B20."
    });

    const supportedHtml = sidebar.renderStructuredPreview(supportedResponse, {
      runId: "run_conditional_format_preview",
      requestId: "req_conditional_format_preview"
    });

    expect(supportedHtml).toContain("Confirm Conditional Formatting");
    expect(supportedHtml).toContain("replace_all_on_target");
    expect(supportedHtml).toContain("text_contains");
    expect(supportedHtml).toContain("Will replace all conditional formatting on Sheet1!B2:B20.");

    const unsupportedHtml = sidebar.renderStructuredPreview({
      type: "conditional_format_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "cell_empty",
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Highlight empty cells.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    }, {
      runId: "run_conditional_format_unsupported_preview",
      requestId: "req_conditional_format_unsupported_preview"
    });

    expect(unsupportedHtml).toContain(
      "This Google Sheets flow can't apply that conditional formatting rule exactly for cell_empty."
    );
    expect(unsupportedHtml).not.toContain("Confirm Conditional Formatting");
  });

  it("treats exact-safe non-text conditional formatting plans as confirmable in Google Sheets", () => {
    const sidebar = loadSidebarContext();
    const basePlan = {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      explanation: "Apply an exact-safe conditional formatting rule.",
      confidence: 0.82,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false
    };
    const responses = [
      {
        type: "conditional_format_plan",
        data: {
          ...basePlan,
          ruleType: "single_color",
          comparator: "equal_to",
          value: "overdue",
          style: {
            backgroundColor: "#ffcccc"
          }
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          ...basePlan,
          ruleType: "number_compare",
          comparator: "greater_than",
          value: 10,
          style: {
            backgroundColor: "#ffcccc"
          }
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          ...basePlan,
          ruleType: "custom_formula",
          formula: "=$C2=\"overdue\"",
          style: {
            backgroundColor: "#ffcccc"
          }
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          ...basePlan,
          ruleType: "duplicate_values",
          style: {
            backgroundColor: "#ffcccc"
          }
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          ...basePlan,
          ruleType: "top_n",
          rank: 3,
          direction: "top",
          style: {
            backgroundColor: "#ffcccc"
          }
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          ...basePlan,
          ruleType: "average_compare",
          direction: "above",
          style: {
            backgroundColor: "#ffcccc"
          }
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          ...basePlan,
          ruleType: "color_scale",
          points: [
            { type: "min", color: "#ff0000" },
            { type: "max", color: "#00ff00" }
          ]
        }
      }
    ];

    for (const response of responses) {
      expect(sidebar.isWritePlanResponse(response)).toBe(true);
    }
  });

  it("applies a Google Sheets conditional-format plan", () => {
    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const existingRule = {
      getRanges() {
        return [];
      }
    };
    const setConditionalFormatRules = vi.fn();
    const sheet = {
      getRange: vi.fn((a1Notation: string) => {
        expect(a1Notation).toBe("B2:B20");
        return targetRange;
      }),
      getConditionalFormatRules: vi.fn(() => [existingRule]),
      setConditionalFormatRules
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "replace_all_on_target",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc",
          textColor: "#990000",
          bold: true,
          italic: false,
          underline: true,
          strikethrough: true
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: true
      }
    });

    expect(result).toMatchObject({
      kind: "conditional_format_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target",
      ruleType: "text_contains",
      text: "overdue",
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: true,
      summary: "Replaced conditional formatting on Sheet1!B2:B20."
    });

    expect(setConditionalFormatRules).toHaveBeenCalledTimes(1);
    const updatedRules = setConditionalFormatRules.mock.calls[0][0];
    expect(updatedRules).toHaveLength(2);
    expect(updatedRules[0]).toBe(existingRule);
    expect(updatedRules[1]).toMatchObject({
      rule: {
        kind: "text_contains",
        text: "overdue"
      },
      format: {
        backgroundColor: "#ffcccc",
        textColor: "#990000",
        bold: true,
        italic: false,
        underline: true,
        strikethrough: true
      }
    });
    expect(updatedRules[1].getRanges()).toEqual([targetRange]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies Google Sheets number_compare conditional formatting", () => {
    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const setConditionalFormatRules = vi.fn();
    const sheet = {
      getRange: vi.fn(() => targetRange),
      getConditionalFormatRules: vi.fn(() => []),
      setConditionalFormatRules
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "number_compare",
        comparator: "greater_than",
        value: 10,
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Highlight values above 10.",
        confidence: 0.88,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    });

    expect(result).toMatchObject({
      kind: "conditional_format_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "number_compare",
      comparator: "greater_than",
      value: 10,
      summary: "Added conditional formatting on Sheet1!B2:B20."
    });

    expect(setConditionalFormatRules).toHaveBeenCalledTimes(1);
    expect(setConditionalFormatRules.mock.calls[0][0][0]).toMatchObject({
      rule: {
        kind: "number_compare",
        comparator: "greater_than",
        value: 10
      },
      format: {
        backgroundColor: "#ffcccc"
      }
    });
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies Google Sheets single_color equality rules through exact formulas", () => {
    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const setConditionalFormatRules = vi.fn();
    const sheet = {
      getRange: vi.fn(() => targetRange),
      getConditionalFormatRules: vi.fn(() => []),
      setConditionalFormatRules
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "single_color",
        comparator: "equal_to",
        value: "overdue",
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Highlight overdue cells.",
        confidence: 0.85,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    });

    expect(setConditionalFormatRules).toHaveBeenCalledTimes(1);
    expect(setConditionalFormatRules.mock.calls[0][0][0]).toMatchObject({
      rule: {
        kind: "custom_formula",
        formula: "=B2=\"overdue\""
      },
      format: {
        backgroundColor: "#ffcccc"
      }
    });
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies Google Sheets color-scale conditional formatting", () => {
    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const setConditionalFormatRules = vi.fn();
    const sheet = {
      getRange: vi.fn(() => targetRange),
      getConditionalFormatRules: vi.fn(() => []),
      setConditionalFormatRules
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "color_scale",
        points: [
          { type: "min", color: "#ff0000" },
          { type: "max", color: "#00ff00" }
        ],
        explanation: "Apply a two-color scale.",
        confidence: 0.84,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    });

    expect(setConditionalFormatRules).toHaveBeenCalledTimes(1);
    expect(setConditionalFormatRules.mock.calls[0][0][0]).toMatchObject({
      rule: {
        kind: "color_scale",
        points: [
          { type: "min", color: "#ff0000" },
          { type: "max", color: "#00ff00" }
        ]
      }
    });
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("applies a Google Sheets clear_on_target conditional-format plan", () => {
    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const exactMatchRule = {
      getRanges() {
        return [targetRange];
      }
    };
    const preservedRange = createRangeStub("D2:D20", 2, 4, 19, 1);
    const preservedRule = {
      getRanges() {
        return [preservedRange];
      }
    };
    const setConditionalFormatRules = vi.fn();
    const sheet = {
      getRange: vi.fn((a1Notation: string) => {
        expect(a1Notation).toBe("B2:B20");
        return targetRange;
      }),
      getConditionalFormatRules: vi.fn(() => [exactMatchRule, preservedRule]),
      setConditionalFormatRules
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    const result = applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "clear_on_target",
        explanation: "Clear existing rules on the target range.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: true
      }
    });

    expect(result).toMatchObject({
      kind: "conditional_format_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "clear_on_target",
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: true,
      summary: "Cleared conditional formatting on Sheet1!B2:B20."
    });

    expect(setConditionalFormatRules).toHaveBeenCalledTimes(1);
    expect(setConditionalFormatRules).toHaveBeenCalledWith([preservedRule]);
    expect(flush).toHaveBeenCalledTimes(1);
  });

  it("fails closed when Google Sheets cannot represent the requested conditional-format semantics exactly", () => {
    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const sheet = {
      getRange: vi.fn(() => targetRange),
      getConditionalFormatRules: vi.fn(() => []),
      setConditionalFormatRules: vi.fn()
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "cell_empty",
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Highlight empty cells.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    })).toThrow(
      "Google Sheets host does not support exact conditional-format mapping for ruleType cell_empty."
    );

    expect(sheet.setConditionalFormatRules).not.toHaveBeenCalled();
    expect(flush).not.toHaveBeenCalled();
  });

  it("fails closed when an existing conditional-format rule overlaps the target range without matching exactly", () => {
    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const overlappingRange = createRangeStub("B2:C20", 2, 2, 19, 2);
    const overlappingRule = {
      getRanges() {
        return [overlappingRange];
      }
    };
    const sheet = {
      getRange: vi.fn(() => targetRange),
      getConditionalFormatRules: vi.fn(() => [overlappingRule]),
      setConditionalFormatRules: vi.fn()
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "replace_all_on_target",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Replace target rules only.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: true
      }
    })).toThrow(
      "Google Sheets host cannot modify conditional formatting exactly when an existing rule overlaps the target range without matching it exactly."
    );

    expect(sheet.setConditionalFormatRules).not.toHaveBeenCalled();
    expect(flush).not.toHaveBeenCalled();
  });

  it("fails closed for unsupported conditional-format managementMode values", () => {
    const sidebar = loadSidebarContext();

    const unsupportedHtml = sidebar.renderStructuredPreview({
      type: "conditional_format_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "merge_with_existing",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Use an unsupported management mode.",
        confidence: 0.5,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    }, {
      runId: "run_conditional_format_unsupported_mode_preview",
      requestId: "req_conditional_format_unsupported_mode_preview"
    });

    expect(unsupportedHtml).toContain(
      "This Google Sheets flow can't manage conditional formatting with mode merge_with_existing."
    );
    expect(unsupportedHtml).not.toContain("Confirm Conditional Formatting");

    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const sheet = {
      getRange: vi.fn(() => targetRange),
      getConditionalFormatRules: vi.fn(() => []),
      setConditionalFormatRules: vi.fn()
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "merge_with_existing",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Use an unsupported management mode.",
        confidence: 0.5,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    })).toThrow(
      "Google Sheets host does not support exact conditional-format managementMode merge_with_existing."
    );

    expect(sheet.setConditionalFormatRules).not.toHaveBeenCalled();
    expect(flush).not.toHaveBeenCalled();
  });

  it("fails closed end to end for unsupported conditional-format style fields", () => {
    const sidebar = loadSidebarContext();

    const unsupportedHtml = sidebar.renderStructuredPreview({
      type: "conditional_format_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc",
          numberFormat: "0.00"
        },
        explanation: "Use an unsupported style field.",
        confidence: 0.5,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    }, {
      runId: "run_conditional_format_unsupported_style_preview",
      requestId: "req_conditional_format_unsupported_style_preview"
    });

    expect(unsupportedHtml).toContain(
      "This Google Sheets flow can't set these conditional-format style fields: numberFormat."
    );
    expect(unsupportedHtml).not.toContain("Confirm Conditional Formatting");

    const targetRange = createRangeStub("B2:B20", 2, 2, 19, 1);
    const sheet = {
      getRange: vi.fn(() => targetRange),
      getConditionalFormatRules: vi.fn(() => []),
      setConditionalFormatRules: vi.fn()
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      }
    };
    const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc",
          numberFormat: "0.00"
        },
        explanation: "Use an unsupported style field.",
        confidence: 0.5,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      }
    })).toThrow(
      "Google Sheets host does not support exact conditional-format style mapping for fields: numberFormat."
    );

    expect(sheet.setConditionalFormatRules).not.toHaveBeenCalled();
    expect(flush).not.toHaveBeenCalled();
  });
});
