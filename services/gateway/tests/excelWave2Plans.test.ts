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
    CellControlType: {
      checkbox: "Checkbox"
    },
    DataValidationOperator: {
      between: "Between",
      notBetween: "NotBetween",
      equalTo: "EqualTo",
      notEqualTo: "NotEqualTo",
      greaterThan: "GreaterThan",
      greaterThanOrEqualTo: "GreaterThanOrEqualTo",
      lessThan: "LessThan",
      lessThanOrEqualTo: "LessThanOrEqualTo"
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

describe("Excel wave 2 plan helpers", () => {
  it("renders validation and named-range previews and body text", async () => {
    const taskpane = await loadTaskpaneModule({
      sync: vi.fn(async () => {})
    });

    expect(taskpane.isWritePlanResponse({
      type: "data_validation_plan"
    })).toBe(true);
    expect(taskpane.isWritePlanResponse({
      type: "named_range_update"
    })).toBe(true);

    expect(taskpane.getRequiresConfirmation({
      type: "data_validation_plan",
      data: {
        requiresConfirmation: true
      }
    })).toBe(true);
    expect(taskpane.getRequiresConfirmation({
      type: "named_range_update",
      data: {
        requiresConfirmation: true
      }
    })).toBe(true);

    expect(taskpane.getResponseBodyText({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "whole_number",
        comparator: "between",
        value: 1,
        value2: 10,
        allowBlank: false,
        invalidDataBehavior: "reject",
        helpText: "Choose a valid number.",
        explanation: "Restrict the status column to approved values.",
        confidence: 0.95,
        requiresConfirmation: true
      }
    })).toBe("Prepared a validation plan for Sheet1!B2:B20.");

    expect(taskpane.getStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        namedRangeName: "StatusOptions",
        allowBlank: false,
        invalidDataBehavior: "reject",
        helpText: "Choose a valid number.",
        explanation: "Restrict the status column to approved values.",
        confidence: 0.95,
        requiresConfirmation: true,
        replacesExistingValidation: true
      }
    })).toMatchObject({
      kind: "data_validation_plan",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      ruleType: "list",
      namedRangeName: "StatusOptions",
      replacesExistingValidation: true
    });

    const validationHtml = taskpane.renderStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        namedRangeName: "StatusOptions",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict the status column to approved values.",
        confidence: 0.95,
        requiresConfirmation: true,
        replacesExistingValidation: true
      }
    }, {
      runId: "run_validation_preview",
      requestId: "req_validation_preview"
    });

    expect(validationHtml).toContain("source named range StatusOptions");
    expect(validationHtml).toContain("replaces existing validation");

    expect(taskpane.getStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "C2:C20",
        ruleType: "custom_formula",
        formula: "=COUNTIF($A$2:$A$20,C2)=0",
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "Prevent duplicate entries.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "data_validation_plan",
      ruleType: "custom_formula",
      formula: "=COUNTIF($A$2:$A$20,C2)=0"
    });

    const customFormulaHtml = taskpane.renderStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "C2:C20",
        ruleType: "custom_formula",
        formula: "=COUNTIF($A$2:$A$20,C2)=0",
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "Prevent duplicate entries.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    }, {
      runId: "run_validation_custom_formula",
      requestId: "req_validation_custom_formula"
    });

    expect(customFormulaHtml).toContain("=COUNTIF($A$2:$A$20,C2)=0");

    const checkboxHtml = taskpane.renderStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "D2:D20",
        ruleType: "checkbox",
        checkedValue: true,
        uncheckedValue: false,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Use checkboxes for completion state.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    }, {
      runId: "run_validation_checkbox",
      requestId: "req_validation_checkbox"
    });

    expect(checkboxHtml).toContain("checked true");
    expect(checkboxHtml).toContain("unchecked false");

    expect(taskpane.getResponseBodyText({
      type: "named_range_update",
      data: {
        operation: "retarget",
        scope: "sheet",
        name: "InputRange",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Retarget the named input block.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    })).toBe("Prepared a named range update for InputRange.");

    expect(taskpane.getStructuredPreview({
      type: "named_range_update",
      data: {
        operation: "retarget",
        scope: "sheet",
        name: "InputRange",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Retarget the named input block.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "named_range_update",
      operation: "retarget",
      name: "InputRange",
      scope: "sheet",
      targetSheet: "Sheet1",
      targetRange: "B2:D20"
    });

    expect(taskpane.getStructuredPreview({
      type: "named_range_update",
      data: {
        operation: "create",
        scope: "workbook",
        name: "SalesData",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Create a workbook-scoped name.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "named_range_update",
      operation: "create",
      summary: "Created named range SalesData at Sheet1!B2:D20."
    });
  });

  it("applies a whole-number validation rule in Excel", async () => {
    const validationRule = { set: vi.fn() };
    const ignoreBlanks = { set: vi.fn() };
    const prompt = { set: vi.fn() };
    const errorAlert = { set: vi.fn() };
    const dataValidation = {};
    Object.defineProperty(dataValidation, "rule", {
      configurable: true,
      set(rule) {
        validationRule.set(rule);
      }
    });
    Object.defineProperty(dataValidation, "ignoreBlanks", {
      configurable: true,
      set(value) {
        ignoreBlanks.set(value);
      }
    });
    Object.defineProperty(dataValidation, "prompt", {
      configurable: true,
      set(value) {
        prompt.set(value);
      }
    });
    Object.defineProperty(dataValidation, "errorAlert", {
      configurable: true,
      set(value) {
        errorAlert.set(value);
      }
    });

    const targetRange = {
      load: vi.fn(),
      address: "Sheet1!B2:B20",
      rowCount: 19,
      columnCount: 1,
      dataValidation
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange),
      names: {
        add: vi.fn()
      }
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: {
          add: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "whole_number",
        comparator: "between",
        value: 1,
        value2: 10,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict values to whole numbers between 1 and 10.",
        confidence: 0.96,
        requiresConfirmation: true
      },
      requestId: "req_validation_001",
      runId: "run_validation_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "data_validation_update",
      targetSheet: "Sheet1",
      targetRange: "B2:B20"
    });

    expect(validationRule.set).toHaveBeenCalledWith({
      wholeNumber: {
        operator: "Between",
        formula1: 1,
        formula2: 10
      }
    });
    expect(ignoreBlanks.set).toHaveBeenCalledWith(false);
    expect(prompt.set).toHaveBeenCalledWith({
      title: "Validation",
      message: "Restrict values to whole numbers between 1 and 10."
    });
    expect(errorAlert.set).toHaveBeenCalledWith({
      title: "Invalid data",
      message: "Values must match the approved validation rule.",
      style: "stop",
      showAlert: true
    });
  });

  it("maps Excel data validation comparators to Office.js operators", async () => {
    const validationRule = { set: vi.fn() };
    const ignoreBlanks = { set: vi.fn() };
    const prompt = { set: vi.fn() };
    const errorAlert = { set: vi.fn() };
    const dataValidation = {};
    Object.defineProperty(dataValidation, "rule", {
      configurable: true,
      set(rule) {
        validationRule.set(rule);
      }
    });
    Object.defineProperty(dataValidation, "ignoreBlanks", {
      configurable: true,
      set(value) {
        ignoreBlanks.set(value);
      }
    });
    Object.defineProperty(dataValidation, "prompt", {
      configurable: true,
      set(value) {
        prompt.set(value);
      }
    });
    Object.defineProperty(dataValidation, "errorAlert", {
      configurable: true,
      set(value) {
        errorAlert.set(value);
      }
    });
    const targetRange = {
      load: vi.fn(),
      address: "Sheet1!C2:C20",
      rowCount: 19,
      columnCount: 1,
      dataValidation
    };
    const worksheet = {
      getRange: vi.fn(() => targetRange),
      names: {
        add: vi.fn()
      }
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: {
          add: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "C2:C20",
        ruleType: "decimal",
        comparator: "greater_than_or_equal_to",
        value: 5,
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "Restrict values to at least five.",
        confidence: 0.95,
        requiresConfirmation: true
      },
      requestId: "req_validation_operator_001",
      runId: "run_validation_operator_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "data_validation_update",
      targetSheet: "Sheet1",
      targetRange: "C2:C20"
    });

    expect(validationRule.set).toHaveBeenCalledWith({
      decimal: {
        operator: "GreaterThanOrEqualTo",
        formula1: 5,
        formula2: undefined
      }
    });
  });

  it("applies a checkbox validation rule in Excel via cell control", async () => {
    const control = { set: vi.fn() };
    const targetRange = {
      load: vi.fn(),
      address: "Sheet1!D2:D20",
      rowCount: 19,
      columnCount: 1
    };
    Object.defineProperty(targetRange, "control", {
      configurable: true,
      set(value) {
        control.set(value);
      }
    });

    const worksheet = {
      getRange: vi.fn(() => targetRange),
      names: {
        add: vi.fn()
      }
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: {
          add: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "D2:D20",
        ruleType: "checkbox",
        checkedValue: true,
        uncheckedValue: false,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Use checkboxes for completion state.",
        confidence: 0.94,
        requiresConfirmation: true
      },
      requestId: "req_validation_checkbox_001",
      runId: "run_validation_checkbox_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "data_validation_update",
      targetSheet: "Sheet1",
      targetRange: "D2:D20",
      summary: "Applied validation to Sheet1!D2:D20."
    });

    expect(control.set).toHaveBeenCalledWith({
      type: "Checkbox"
    });
  });

  it("retargets a sheet-scoped named range in Excel", async () => {
    const namedRange = {
      reference: "Sheet1!A1:A2",
      name: "InputRange",
      delete: vi.fn()
    };
    const sheetNames = {
      getItem: vi.fn(() => namedRange),
      add: vi.fn()
    };
    const worksheet = {
      getRange: vi.fn(() => ({
        address: "Sheet1!B2:D20"
      })),
      names: sheetNames
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: {
          add: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        operation: "retarget",
        scope: "sheet",
        name: "InputRange",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Retarget the named input block.",
        confidence: 0.91,
        requiresConfirmation: true
      },
      requestId: "req_named_range_001",
      runId: "run_named_range_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "named_range_update",
      operation: "retarget",
      scope: "sheet",
      sheetName: "Sheet1",
      name: "InputRange",
      targetSheet: "Sheet1",
      targetRange: "B2:D20"
    });

    expect(sheetNames.getItem).toHaveBeenCalledWith("InputRange");
    expect(namedRange.reference).toBe("Sheet1!B2:D20");
  });

  it("creates, renames, and deletes named ranges in Excel", async () => {
    const workbookNamedRange = {
      reference: "Sheet1!A1:A2",
      name: "OldRange",
      delete: vi.fn()
    };
    const workbookNames = {
      add: vi.fn(),
      getItem: vi.fn(() => workbookNamedRange)
    };
    const worksheet = {
      getRange: vi.fn(() => ({
        address: "Sheet1!B2:D20"
      })),
      names: {
        add: vi.fn(),
        getItem: vi.fn(() => workbookNamedRange)
      }
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: workbookNames
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        operation: "create",
        scope: "workbook",
        name: "SalesData",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Create a workbook-scoped name.",
        confidence: 0.91,
        requiresConfirmation: true
      },
      requestId: "req_named_range_create_001",
      runId: "run_named_range_create_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "named_range_update",
      operation: "create",
      name: "SalesData",
      summary: "Created named range SalesData at Sheet1!B2:D20."
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        operation: "rename",
        scope: "workbook",
        name: "OldRange",
        newName: "NewRange",
        explanation: "Rename the workbook name.",
        confidence: 0.92,
        requiresConfirmation: true
      },
      requestId: "req_named_range_rename_001",
      runId: "run_named_range_rename_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "named_range_update",
      operation: "rename",
      name: "OldRange"
    });

    await expect(taskpane.applyWritePlan({
      plan: {
        operation: "delete",
        scope: "workbook",
        name: "OldRange",
        explanation: "Delete the workbook name.",
        confidence: 0.93,
        requiresConfirmation: true
      },
      requestId: "req_named_range_delete_001",
      runId: "run_named_range_delete_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "named_range_update",
      operation: "delete",
      name: "OldRange"
    });

    expect(workbookNames.add).toHaveBeenCalledWith("SalesData", expect.objectContaining({
      address: "Sheet1!B2:D20"
    }));
    expect(workbookNamedRange.name).toBe("NewRange");
    expect(workbookNamedRange.delete).toHaveBeenCalled();
  });

  it("uses sheetName to resolve sheet-scoped named range renames when targetSheet is absent", async () => {
    const sheetNamedRange = {
      reference: "Sheet Scope!A1:A2",
      name: "OldRange"
    };
    const sheetNames = {
      getItem: vi.fn(() => sheetNamedRange)
    };
    const worksheet = {
      names: sheetNames,
      getRange: vi.fn()
    };
    const worksheets = {
      getItem: vi.fn((sheetName: string) => {
        expect(sheetName).toBe("Sheet Scope");
        return worksheet;
      })
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets,
        names: {
          getItem: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        operation: "rename",
        scope: "sheet",
        sheetName: "Sheet Scope",
        name: "OldRange",
        newName: "RenamedRange",
        explanation: "Rename the sheet-scoped name.",
        confidence: 0.92,
        requiresConfirmation: true
      },
      requestId: "req_named_range_sheet_rename_001",
      runId: "run_named_range_sheet_rename_001",
      approvalToken: "token"
    })).resolves.toMatchObject({
      kind: "named_range_update",
      operation: "rename",
      scope: "sheet",
      sheetName: "Sheet Scope",
      name: "OldRange",
      newName: "RenamedRange"
    });

    expect(sheetNames.getItem).toHaveBeenCalledWith("OldRange");
    expect(sheetNamedRange.name).toBe("RenamedRange");
  });

  it("fails closed for unsupported validation rule types", async () => {
    const worksheet = {
      getRange: vi.fn(() => ({
        address: "Sheet1!B2:B20",
        rowCount: 19,
        columnCount: 1,
        dataValidation: {}
      })),
      names: {
        add: vi.fn()
      }
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: {
          add: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "unknown_rule",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Unsupported validation rule.",
        confidence: 0.5,
        requiresConfirmation: true
      },
      requestId: "req_validation_invalid_001",
      runId: "run_validation_invalid_001",
      approvalToken: "token"
    })).rejects.toThrow("Unsupported Excel data validation rule type.");
  });

  it("fails closed for Excel checkbox plans with non-boolean custom values", async () => {
    const targetRange = {
      load: vi.fn(),
      address: "Sheet1!D2:D20",
      rowCount: 19,
      columnCount: 1
    };
    Object.defineProperty(targetRange, "control", {
      configurable: true,
      set() {}
    });
    const worksheet = {
      getRange: vi.fn(() => targetRange),
      names: {
        add: vi.fn()
      }
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: {
          add: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "D2:D20",
        ruleType: "checkbox",
        checkedValue: "Y",
        uncheckedValue: "N",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Use checkboxes with custom values.",
        confidence: 0.9,
        requiresConfirmation: true
      },
      requestId: "req_validation_checkbox_invalid_001",
      runId: "run_validation_checkbox_invalid_001",
      approvalToken: "token"
    })).rejects.toThrow("Excel checkbox controls only support boolean true/false values.");
  });

  it("fails closed for partial named range plans without a target range", async () => {
    const worksheet = {
      getRange: vi.fn(() => ({
        address: "Sheet1!",
      })),
      names: {
        getItem: vi.fn()
      }
    };
    const context = {
      sync: vi.fn(async () => {}),
      workbook: {
        worksheets: {
          getItem: vi.fn(() => worksheet)
        },
        names: {
          add: vi.fn()
        }
      }
    };
    const taskpane = await loadTaskpaneModule(context);

    await expect(taskpane.applyWritePlan({
      plan: {
        operation: "retarget",
        scope: "sheet",
        name: "InputRange",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        explanation: "Malformed retarget request.",
        confidence: 0.4,
        requiresConfirmation: true
      },
      requestId: "req_named_range_invalid_001",
      runId: "run_named_range_invalid_001",
      approvalToken: "token"
    })).rejects.toThrow("Named range create and retarget require targetSheet and targetRange.");
  });
});
