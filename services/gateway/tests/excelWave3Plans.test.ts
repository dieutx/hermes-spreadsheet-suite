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
    ConditionalFormatType: {
      colorScale: "colorScale",
      containsText: "containsText"
    },
    ConditionalFormatColorCriterionType: {
      highestValue: "highestValue",
      lowestValue: "lowestValue",
      number: "number",
      percent: "percent",
      percentile: "percentile"
    }
  });

  return import(`${TASKPANE_MODULE_URL}?t=${Date.now()}_${Math.random()}`);
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("Excel wave 3 conditional-format plans", () => {
  it("recognizes conditional_format_plan in preview and confirmation helpers", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const response = {
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
          underline: true,
          strikethrough: true
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        confirmationLevel: "destructive",
        replacesExistingRules: true
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(true);
    expect(taskpane.getRequiresConfirmation(response)).toBe(true);
    expect(taskpane.getResponseBodyText(response)).toBe(
      "Prepared a conditional formatting plan for Sheet1!B2:B20."
    );
    expect(taskpane.getStructuredPreview(response)).toMatchObject({
      kind: "conditional_format_plan",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target",
      confirmationLevel: "destructive",
      ruleType: "text_contains"
    });

    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_conditional_format_preview",
      requestId: "req_conditional_format_preview"
    });

    expect(html).toContain("Confirm Conditional Formatting");
    expect(html).toContain("replace_all_on_target");
    expect(html).toContain("destructive");
    expect(html).toContain("text_contains");

    const confirm = vi.fn(() => true);
    vi.stubGlobal("confirm", confirm);
    expect(taskpane.buildWriteApprovalRequest({
      requestId: "req_conditional_format_preview",
      runId: "run_conditional_format_preview",
      plan: taskpane.getStructuredPreview(response)
    })).toMatchObject({
      requestId: "req_conditional_format_preview",
      runId: "run_conditional_format_preview",
      destructiveConfirmation: { confirmed: true },
      plan: {
        kind: "conditional_format_plan",
        confirmationLevel: "destructive"
      }
    });
    expect(confirm).toHaveBeenCalledTimes(1);
  });

  it("renders color scale conditional-format previews as confirmable in Excel", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    const response = {
      type: "conditional_format_plan",
      data: {
        targetSheet: "Summary",
        targetRange: "A2:D20",
        managementMode: "add",
        ruleType: "color_scale",
        points: [
          { type: "min", color: "#f8696b" },
          { type: "percentile", value: 50, color: "#ffeb84" },
          { type: "max", color: "#63be7b" }
        ],
        explanation: "Apply a red-yellow-green percentile scale.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Summary!A2:D20"],
        confirmationLevel: "standard",
        replacesExistingRules: false
      }
    };

    expect(taskpane.isWritePlanResponse(response)).toBe(true);
    const html = taskpane.renderStructuredPreview(response, {
      runId: "run_conditional_color_scale_preview",
      requestId: "req_conditional_color_scale_preview"
    });

    expect(html).toContain("Confirm Conditional Formatting");
    expect(html).toContain("color_scale");
    expect(html).not.toContain("can't apply that conditional formatting rule exactly");
  });

  it("applies a conditional formatting rule in Excel", async () => {
    const clearAll = vi.fn();
    const styleAssignments: Record<string, unknown> = {};
    const containsTextRule = { text: "" };
    const conditionalFormat = {
      containsText: {
        rule: containsTextRule,
        format: {
          fill: {},
          font: {}
        }
      }
    };

    Object.defineProperty(conditionalFormat.containsText.format.fill, "color", {
      configurable: true,
      set(value) {
        styleAssignments.backgroundColor = value;
      }
    });
    Object.defineProperty(conditionalFormat.containsText.format.font, "color", {
      configurable: true,
      set(value) {
        styleAssignments.textColor = value;
      }
    });
    Object.defineProperty(conditionalFormat.containsText.format.font, "bold", {
      configurable: true,
      set(value) {
        styleAssignments.bold = value;
      }
    });
    Object.defineProperty(conditionalFormat.containsText.format.font, "italic", {
      configurable: true,
      set(value) {
        styleAssignments.italic = value;
      }
    });
    Object.defineProperty(conditionalFormat.containsText.format.font, "underline", {
      configurable: true,
      set(value) {
        styleAssignments.underline = value;
      }
    });
    Object.defineProperty(conditionalFormat.containsText.format.font, "strikethrough", {
      configurable: true,
      set(value) {
        styleAssignments.strikethrough = value;
      }
    });

    const add = vi.fn(() => conditionalFormat);
    const targetRange = {
      conditionalFormats: {
        clearAll,
        add
      }
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange)
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
        confirmationLevel: "destructive",
        replacesExistingRules: true
      },
      requestId: "req_conditional_format_apply_001",
      runId: "run_conditional_format_apply_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "conditional_format_update",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target",
      ruleType: "text_contains",
      text: "overdue",
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: true,
      summary: "Replaced conditional formatting on Sheet1!B2:B20."
    });

    expect(clearAll).toHaveBeenCalledTimes(1);
    expect(add).toHaveBeenCalledWith("containsText");
    expect(containsTextRule.text).toBe("overdue");
    expect(styleAssignments).toEqual({
      backgroundColor: "#ffcccc",
      textColor: "#990000",
      bold: true,
      italic: false,
      underline: "Single",
      strikethrough: true
    });
  });

  it("attaches local undo snapshots for Excel conditional formatting writes", async () => {
    const clearAll = vi.fn();
    const add = vi.fn(() => ({
      containsText: {
        rule: {},
        format: {
          fill: {},
          font: {}
        }
      }
    }));
    const targetRange = {
      conditionalFormats: {
        items: [],
        clearAll,
        add
      }
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange)
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

    const result = await taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "replace_all_on_target",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc",
          textColor: "#990000",
          bold: true
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        confirmationLevel: "destructive",
        replacesExistingRules: true
      },
      requestId: "req_conditional_format_snapshot_excel_001",
      runId: "run_conditional_format_snapshot_excel_001",
      approvalToken: "token",
      executionId: "exec_conditional_format_snapshot_excel_001"
    });

    expect(result).toMatchObject({
      kind: "conditional_format_update",
      __hermesLocalExecutionSnapshot: {
        baseExecutionId: "exec_conditional_format_snapshot_excel_001",
        kind: "conditional_format",
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        before: {
          rules: []
        },
        after: {
          rules: [{
            ranges: ["B2:B20"],
            rule: {
              kind: "text_contains",
              text: "overdue"
            },
            format: {
              backgroundColor: "#ffcccc",
              textColor: "#990000",
              bold: true
            }
          }]
        }
      }
    });
  });

  it("does not block Excel conditional formatting when undo snapshot inspection is unavailable", async () => {
    const clearAll = vi.fn();
    const add = vi.fn(() => ({
      containsText: {
        rule: {},
        format: {
          fill: {},
          font: {}
        }
      }
    }));
    const targetRange = {
      conditionalFormats: {
        get items() {
          throw new Error("PropertyNotLoaded");
        },
        clearAll,
        add
      }
    };
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

    const result = await taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "replace_all_on_target",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc"
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.94,
        requiresConfirmation: true,
        confirmationLevel: "destructive"
      },
      requestId: "req_conditional_format_no_snapshot_excel_001",
      runId: "run_conditional_format_no_snapshot_excel_001",
      approvalToken: "token",
      executionId: "exec_conditional_format_no_snapshot_excel_001"
    });

    expect(result).toMatchObject({
      kind: "conditional_format_update",
      targetSheet: "Sheet1",
      targetRange: "B2:B20"
    });
    expect(result).not.toHaveProperty("__hermesLocalExecutionSnapshot");
    expect(clearAll).toHaveBeenCalledTimes(1);
    expect(add).toHaveBeenCalledWith("containsText");
  });

  it("applies a color scale conditional formatting rule in Excel", async () => {
    let assignedCriteria: unknown = null;
    const clearAll = vi.fn();
    const colorScale = {};
    Object.defineProperty(colorScale, "criteria", {
      configurable: true,
      set(value) {
        assignedCriteria = value;
      }
    });

    const add = vi.fn(() => ({
      colorScale
    }));
    const targetRange = {
      conditionalFormats: {
        clearAll,
        add
      }
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange)
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
        targetSheet: "Summary",
        targetRange: "A2:D20",
        managementMode: "add",
        ruleType: "color_scale",
        points: [
          { type: "min", color: "#f8696b" },
          { type: "percentile", value: 50, color: "#ffeb84" },
          { type: "max", color: "#63be7b" }
        ],
        explanation: "Apply a red-yellow-green percentile scale.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Summary!A2:D20"],
        confirmationLevel: "standard",
        replacesExistingRules: false
      },
      requestId: "req_conditional_color_scale_apply_001",
      runId: "run_conditional_color_scale_apply_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "conditional_format_update",
      targetSheet: "Summary",
      targetRange: "A2:D20",
      managementMode: "add",
      ruleType: "color_scale",
      affectedRanges: ["Summary!A2:D20"],
      replacesExistingRules: false,
      summary: "Added conditional formatting to Summary!A2:D20."
    });

    expect(clearAll).not.toHaveBeenCalled();
    expect(add).toHaveBeenCalledWith("colorScale");
    expect(assignedCriteria).toEqual({
      minimum: {
        formula: null,
        type: "lowestValue",
        color: "#f8696b"
      },
      midpoint: {
        formula: "50",
        type: "percentile",
        color: "#ffeb84"
      },
      maximum: {
        formula: null,
        type: "highestValue",
        color: "#63be7b"
      }
    });
  });

  it("fails closed when Excel cannot represent the requested conditional-format semantics exactly", async () => {
    const styleAssignments: Record<string, unknown> = {};
    const add = vi.fn(() => ({
      colorScale: {
        criteria: [],
        format: {
          fill: {},
          font: {}
        }
      }
    }));
    const addedFormat = add();
    Object.defineProperty(addedFormat.colorScale.format.fill, "color", {
      configurable: true,
      set(value) {
        styleAssignments.backgroundColor = value;
      }
    });
    Object.defineProperty(addedFormat.colorScale.format.font, "color", {
      configurable: true,
      set(value) {
        styleAssignments.textColor = value;
      }
    });
    add.mockReset();
    add.mockImplementation(() => addedFormat);

    const worksheet = {
      getRange: vi.fn(() => ({
        conditionalFormats: {
          clearAll: vi.fn(),
          add
        }
      }))
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
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "data_bar",
        style: {
          backgroundColor: "#ff0000"
        },
        explanation: "Apply a data bar.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      },
      requestId: "req_conditional_format_invalid_001",
      runId: "run_conditional_format_invalid_001",
      approvalToken: "token"
    })).rejects.toThrow(
      "Excel host does not support exact conditional-format mapping for ruleType data_bar."
    );

    expect(add).not.toHaveBeenCalled();
    expect(styleAssignments).toEqual({});
    expect(addedFormat.colorScale.criteria).toEqual([]);
  });

  it("fails closed when a conditional-format plan requests unsupported style fields", async () => {
    const styleAssignments: Record<string, unknown> = {};
    const containsTextRule = { text: "" };
    const add = vi.fn(() => ({
      containsText: {
        rule: containsTextRule,
        format: {
          fill: {},
          font: {}
        }
      }
    }));
    const addedFormat = add();
    Object.defineProperty(addedFormat.containsText.format.fill, "color", {
      configurable: true,
      set(value) {
        styleAssignments.backgroundColor = value;
      }
    });
    Object.defineProperty(addedFormat.containsText.format.font, "color", {
      configurable: true,
      set(value) {
        styleAssignments.textColor = value;
      }
    });
    Object.defineProperty(addedFormat.containsText.format.font, "bold", {
      configurable: true,
      set(value) {
        styleAssignments.bold = value;
      }
    });
    add.mockReset();
    add.mockImplementation(() => addedFormat);

    const worksheet = {
      getRange: vi.fn(() => ({
        conditionalFormats: {
          clearAll: vi.fn(),
          add
        }
      }))
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
        targetRange: "B2:B20",
        managementMode: "add",
        ruleType: "text_contains",
        text: "overdue",
        style: {
          backgroundColor: "#ffcccc",
          numberFormat: "0.00"
        },
        explanation: "Highlight overdue rows.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false
      },
      requestId: "req_conditional_format_style_invalid_001",
      runId: "run_conditional_format_style_invalid_001",
      approvalToken: "token"
    })).rejects.toThrow(
      "Excel host does not support exact conditional-format style mapping for fields: numberFormat."
    );

    expect(add).not.toHaveBeenCalled();
    expect(containsTextRule.text).toBe("");
    expect(styleAssignments).toEqual({});
  });
});
