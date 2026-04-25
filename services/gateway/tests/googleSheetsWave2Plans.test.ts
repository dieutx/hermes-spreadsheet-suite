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

function createDataValidationBuilder() {
  return {
    _rule: null as null | Record<string, unknown>,
    _allowInvalid: undefined as boolean | undefined,
    _helpText: undefined as string | undefined,
    requireValueInList(values: string[], showDropdown?: boolean) {
      this._rule = {
        kind: "list",
        values,
        showDropdown
      };
      return this;
    },
    requireValueInRange(range: Record<string, unknown>, showDropdown?: boolean) {
      this._rule = {
        kind: "list",
        sourceRange: range,
        showDropdown
      };
      return this;
    },
    requireCheckbox(checkedValue?: unknown, uncheckedValue?: unknown) {
      this._rule = {
        kind: "checkbox",
        checkedValue,
        uncheckedValue
      };
      return this;
    },
    requireNumberBetween(value: number, value2: number) {
      this._rule = { kind: "number_between", value, value2 };
      return this;
    },
    requireNumberNotBetween(value: number, value2: number) {
      this._rule = { kind: "number_not_between", value, value2 };
      return this;
    },
    requireNumberEqualTo(value: number) {
      this._rule = { kind: "number_equal_to", value };
      return this;
    },
    requireNumberNotEqualTo(value: number) {
      this._rule = { kind: "number_not_equal_to", value };
      return this;
    },
    requireNumberGreaterThan(value: number) {
      this._rule = { kind: "number_greater_than", value };
      return this;
    },
    requireNumberGreaterThanOrEqualTo(value: number) {
      this._rule = { kind: "number_greater_than_or_equal_to", value };
      return this;
    },
    requireNumberLessThan(value: number) {
      this._rule = { kind: "number_less_than", value };
      return this;
    },
    requireNumberLessThanOrEqualTo(value: number) {
      this._rule = { kind: "number_less_than_or_equal_to", value };
      return this;
    },
    requireDateBetween(value: Date, value2: Date) {
      this._rule = { kind: "date_between", value, value2 };
      return this;
    },
    requireDateNotBetween(value: Date, value2: Date) {
      this._rule = { kind: "date_not_between", value, value2 };
      return this;
    },
    requireDateEqualTo(value: Date) {
      this._rule = { kind: "date_equal_to", value };
      return this;
    },
    requireDateAfter(value: Date) {
      this._rule = { kind: "date_after", value };
      return this;
    },
    requireDateOnOrAfter(value: Date) {
      this._rule = { kind: "date_on_or_after", value };
      return this;
    },
    requireDateBefore(value: Date) {
      this._rule = { kind: "date_before", value };
      return this;
    },
    requireDateOnOrBefore(value: Date) {
      this._rule = { kind: "date_on_or_before", value };
      return this;
    },
    requireTextLengthBetween(value: number, value2: number) {
      this._rule = { kind: "text_length_between", value, value2 };
      return this;
    },
    requireTextLengthNotBetween(value: number, value2: number) {
      this._rule = { kind: "text_length_not_between", value, value2 };
      return this;
    },
    requireTextLengthEqualTo(value: number) {
      this._rule = { kind: "text_length_equal_to", value };
      return this;
    },
    requireTextLengthGreaterThan(value: number) {
      this._rule = { kind: "text_length_greater_than", value };
      return this;
    },
    requireTextLengthGreaterThanOrEqualTo(value: number) {
      this._rule = { kind: "text_length_greater_than_or_equal_to", value };
      return this;
    },
    requireTextLengthLessThan(value: number) {
      this._rule = { kind: "text_length_less_than", value };
      return this;
    },
    requireTextLengthLessThanOrEqualTo(value: number) {
      this._rule = { kind: "text_length_less_than_or_equal_to", value };
      return this;
    },
    requireFormulaSatisfied(formula: string) {
      this._rule = { kind: "custom_formula", formula };
      return this;
    },
    setAllowInvalid(value: boolean) {
      this._allowInvalid = value;
      return this;
    },
    setHelpText(value: string) {
      this._helpText = value;
      return this;
    },
    build() {
      return {
        ...this._rule,
        allowInvalid: this._allowInvalid,
        helpText: this._helpText
      };
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
      newDataValidation() {
        return createDataValidationBuilder();
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

afterEach(() => {
  vi.restoreAllMocks();
});

describe("Google Sheets wave 2 plans", () => {
  it("renders data validation and named range previews in the sidebar", () => {
    const sidebar = loadSidebarContext();

    expect(sidebar.isWritePlanResponse({
      type: "data_validation_plan"
    })).toBe(true);
    expect(sidebar.isWritePlanResponse({
      type: "named_range_update"
    })).toBe(true);

    expect(sidebar.getRequiresConfirmation({
      type: "data_validation_plan",
      data: {
        requiresConfirmation: true
      }
    })).toBe(true);
    expect(sidebar.getRequiresConfirmation({
      type: "named_range_update",
      data: {
        requiresConfirmation: true
      }
    })).toBe(true);

    expect(sidebar.getResponseBodyText({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        namedRangeName: "StatusOptions",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict values to the approved list.",
        confidence: 0.94,
        requiresConfirmation: true
      }
    })).toBe("Prepared a validation plan for Sheet1!B2:B20.");

    expect(sidebar.getStructuredPreview({
      type: "named_range_update",
      data: {
        operation: "retarget",
        scope: "workbook",
        name: "InputRange",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Retarget the input block.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "named_range_update",
      operation: "retarget",
      name: "InputRange",
      targetSheet: "Sheet1",
      targetRange: "B2:D20"
    });

    const validationHtml = sidebar.renderStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        namedRangeName: "StatusOptions",
        showDropdown: false,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict values to the approved list.",
        confidence: 0.94,
        requiresConfirmation: true,
        replacesExistingValidation: true
      }
    }, {
      runId: "run_validation_preview",
      requestId: "req_validation_preview"
    });

    expect(validationHtml).toContain("source named range StatusOptions");
    expect(validationHtml).toContain("replaces existing validation");
    expect(validationHtml).toContain("dropdown hidden");
    expect(validationHtml).toContain("Will validate Sheet1!B2:B20.");
    expect(validationHtml).toContain("Confirm Validation");

    const comparatorHtml = sidebar.renderStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "C2:C20",
        ruleType: "whole_number",
        comparator: "between",
        value: 1,
        value2: 10,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Restrict values to integers between 1 and 10.",
        confidence: 0.95,
        requiresConfirmation: true
      }
    }, {
      runId: "run_validation_comparator_preview",
      requestId: "req_validation_comparator_preview"
    });

    expect(comparatorHtml).toContain("value 1");
    expect(comparatorHtml).toContain("value2 10");

    const namedRangeHtml = sidebar.renderStructuredPreview({
      type: "named_range_update",
      data: {
        operation: "create",
        scope: "workbook",
        name: "SalesData",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Create a workbook name for the table.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    }, {
      runId: "run_named_range_preview",
      requestId: "req_named_range_preview"
    });

    expect(namedRangeHtml).toContain("Confirm Named Range Update");
    expect(namedRangeHtml).toContain("SalesData");
    expect(namedRangeHtml).toContain("Sheet1!B2:D20");
    expect(namedRangeHtml).toContain("Will create named range SalesData at Sheet1!B2:D20.");

    const unsupportedValidationHtml = sidebar.renderStructuredPreview({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        namedRangeName: "StatusOptions",
        showDropdown: true,
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "This plan should preview as unsupported for Google Sheets.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    }, {
      runId: "run_validation_unsupported_preview",
      requestId: "req_validation_unsupported_preview"
    });

    expect(unsupportedValidationHtml).toContain("This validation would behave differently in Google Sheets.");
    expect(unsupportedValidationHtml).not.toContain("Confirm Validation");

    const unsupportedNamedRangeHtml = sidebar.renderStructuredPreview({
      type: "named_range_update",
      data: {
        operation: "create",
        scope: "sheet",
        sheetName: "Sheet1",
        name: "InputRange",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "This plan should preview as unsupported for Google Sheets.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    }, {
      runId: "run_named_range_unsupported_preview",
      requestId: "req_named_range_unsupported_preview"
    });

    expect(unsupportedNamedRangeHtml).toContain("This Google Sheets flow only supports workbook-level named ranges.");
    expect(unsupportedNamedRangeHtml).not.toContain("Confirm Named Range Update");
  });

  it("applies Google Sheets validation plans across the supported Wave 2 families", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      getA1Notation() {
        return "B2:B20";
      },
      setDataValidation
    };
    const sourceRange = {
      getA1Notation() {
        return "Lookup!A1:A4";
      }
    };
    const sheet = {
      getRange: vi.fn((a1Notation: string) => {
        if (a1Notation === "B2:B20") {
          return targetRange;
        }

        if (a1Notation === "A1:A4") {
          return sourceRange;
        }

        throw new Error(`Unexpected range lookup: ${a1Notation}`);
      })
    };
    const namedRange = {
      getName() {
        return "StatusOptions";
      },
      getRange() {
        return sourceRange;
      }
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      },
      getNamedRanges() {
        return [namedRange];
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    const wholeNumberResult = applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "whole_number",
        comparator: "between",
        value: 1,
        value2: 10,
        allowBlank: false,
        invalidDataBehavior: "reject",
        helpText: "Enter an integer between 1 and 10.",
        explanation: "Restrict entries to integers between 1 and 10.",
        confidence: 0.95,
        requiresConfirmation: true
      }
    });

    expect(wholeNumberResult).toMatchObject({
      kind: "data_validation_update",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      summary: "Applied validation to Sheet1!B2:B20."
    });
    expect(setDataValidation).toHaveBeenLastCalledWith(expect.objectContaining({
      allowInvalid: false,
      helpText: "Enter an integer between 1 and 10."
    }));

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "whole_number",
        comparator: "between",
        value: 1,
        value2: 10,
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "Allow blanks but only keep valid integers when filled.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    });

    expect(setDataValidation).toHaveBeenLastCalledWith(expect.objectContaining({
      kind: "custom_formula",
      formula: "=OR(ISBLANK(B2),AND(ISNUMBER(B2),B2=INT(B2),AND(B2>=1,B2<=10)))",
      allowInvalid: true
    }));

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        namedRangeName: "StatusOptions",
        showDropdown: true,
        allowBlank: false,
        invalidDataBehavior: "warn",
        explanation: "Use the named range dropdown.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    });

    expect(setDataValidation).toHaveBeenLastCalledWith(expect.objectContaining({
      kind: "list",
      sourceRange,
      showDropdown: true,
      allowInvalid: true
    }));

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "checkbox",
        checkedValue: "Y",
        uncheckedValue: "N",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Use Y/N checkboxes.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(setDataValidation).toHaveBeenLastCalledWith(expect.objectContaining({
      kind: "checkbox",
      checkedValue: "Y",
      uncheckedValue: "N"
    }));

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "decimal",
        comparator: "greater_than",
        value: 0.5,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Require a decimal above the threshold.",
        confidence: 0.88,
        requiresConfirmation: true
      }
    });

    expect(setDataValidation).toHaveBeenLastCalledWith(expect.objectContaining({
      kind: "number_greater_than",
      value: 0.5
    }));

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "date",
        comparator: "less_than_or_equal_to",
        value: "2026-12-31",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Keep dates inside the reporting year.",
        confidence: 0.87,
        requiresConfirmation: true
      }
    });

    const dateRule = setDataValidation.mock.calls.at(-1)?.[0] as Record<string, unknown>;
    const dateValue = dateRule.value as Date;
    expect(dateRule.kind).toBe("date_on_or_before");
    expect(Object.prototype.toString.call(dateValue)).toBe("[object Date]");
    expect(dateValue.getFullYear()).toBe(2026);
    expect(dateValue.getMonth()).toBe(11);
    expect(dateValue.getDate()).toBe(31);

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "text_length",
        comparator: "less_than_or_equal_to",
        value: 12,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Keep codes short enough for downstream systems.",
        confidence: 0.86,
        requiresConfirmation: true
      }
    });

    expect(setDataValidation).toHaveBeenLastCalledWith(expect.objectContaining({
      kind: "text_length_less_than_or_equal_to",
      value: 12
    }));

    applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "custom_formula",
        formula: "=COUNTIF($B$2:$B$20,B2)=1",
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "Disallow duplicates.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(setDataValidation).toHaveBeenLastCalledWith(expect.objectContaining({
      kind: "custom_formula",
      formula: "=OR(ISBLANK(B2),(COUNTIF($B$2:$B$20,B2)=1))",
      allowInvalid: true
    }));
  });

  it("creates, renames, retargets, and deletes workbook named ranges in Google Sheets", () => {
    const targetRange = {
      getA1Notation() {
        return "B2:D20";
      }
    };
    const existingNamedRange = {
      getName: vi.fn(() => "SalesData"),
      setName: vi.fn(),
      setRange: vi.fn()
    };
    const setNamedRange = vi.fn();
    const removeNamedRange = vi.fn();
    const sheet = {
      getRange: vi.fn((a1Notation: string) => {
        expect(a1Notation).toBe("B2:D20");
        return targetRange;
      })
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      },
      getNamedRanges() {
        return [existingNamedRange];
      },
      setNamedRange,
      removeNamedRange
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(applyWritePlan({
      plan: {
        operation: "create",
        scope: "workbook",
        name: "NewRange",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Create a workbook name.",
        confidence: 0.93,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "named_range_update",
      operation: "create",
      scope: "workbook",
      name: "NewRange",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      summary: "Created named range NewRange at Sheet1!B2:D20."
    });

    expect(applyWritePlan({
      plan: {
        operation: "rename",
        scope: "workbook",
        name: "SalesData",
        newName: "SalesData2026",
        explanation: "Rename the workbook name.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "named_range_update",
      operation: "rename",
      name: "SalesData"
    });

    expect(applyWritePlan({
      plan: {
        operation: "retarget",
        scope: "workbook",
        name: "SalesData",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Retarget the workbook name.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "named_range_update",
      operation: "retarget",
      scope: "workbook",
      name: "SalesData",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      summary: "Retargeted SalesData to Sheet1!B2:D20."
    });

    expect(applyWritePlan({
      plan: {
        operation: "delete",
        scope: "workbook",
        name: "SalesData",
        explanation: "Delete the workbook name.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    })).toMatchObject({
      kind: "named_range_update",
      operation: "delete",
      name: "SalesData",
      summary: "Deleted named range SalesData."
    });

    expect(setNamedRange).toHaveBeenCalledWith("NewRange", targetRange);
    expect(existingNamedRange.setName).toHaveBeenCalledWith("SalesData2026");
    expect(existingNamedRange.setRange).toHaveBeenCalledWith(targetRange);
    expect(removeNamedRange).toHaveBeenCalledWith("SalesData");
  });

  it("fails closed when creating a Google Sheets named range that already exists", () => {
    const existingNamedRange = {
      getName: vi.fn(() => "SalesData")
    };
    const setNamedRange = vi.fn();
    const sheet = {
      getRange: vi.fn()
    };
    const spreadsheet = {
      getSheetByName(sheetName: string) {
        expect(sheetName).toBe("Sheet1");
        return sheet;
      },
      getNamedRanges() {
        return [existingNamedRange];
      },
      setNamedRange
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        operation: "create",
        scope: "workbook",
        name: "SalesData",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Create a workbook name.",
        confidence: 0.93,
        requiresConfirmation: true
      }
    })).toThrow("Named range already exists: SalesData");

    expect(setNamedRange).not.toHaveBeenCalled();
    expect(sheet.getRange).not.toHaveBeenCalled();
  });

  it("fails closed for unsupported Google Sheets named range scope", () => {
    const spreadsheet = {
      getSheetByName: vi.fn()
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        operation: "create",
        scope: "sheet",
        sheetName: "Sheet1",
        name: "InputRange",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        explanation: "Create a sheet-scoped named range.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host does not support sheet-scoped named ranges.");
  });

  it("fails closed for unsupported invalidDataBehavior values", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      getA1Notation() {
        return "B2:B20";
      },
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn(() => targetRange)
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      },
      getNamedRanges() {
        return [];
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        values: ["Open", "Closed"],
        allowBlank: false,
        invalidDataBehavior: "soft_warn",
        explanation: "Malformed policy should fail closed.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    })).toThrow("Unsupported invalidDataBehavior: soft_warn");
    expect(setDataValidation).not.toHaveBeenCalled();
  });

  it("fails closed for list validation plans that require allowBlank=true", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      getA1Notation() {
        return "B2:B20";
      },
      setDataValidation
    };
    const sourceRange = {
      getA1Notation() {
        return "Lookup!A1:A4";
      }
    };
    const sheet = {
      getRange: vi.fn((a1Notation: string) => {
        if (a1Notation === "B2:B20") {
          return targetRange;
        }

        if (a1Notation === "A1:A4") {
          return sourceRange;
        }

        throw new Error(`Unexpected range lookup: ${a1Notation}`);
      })
    };
    const namedRange = {
      getName() {
        return "StatusOptions";
      },
      getRange() {
        return sourceRange;
      }
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      },
      getNamedRanges() {
        return [namedRange];
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        namedRangeName: "StatusOptions",
        showDropdown: true,
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "This exact list+blank semantics should fail closed.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host cannot represent allowBlank=true exactly for list validation.");
    expect(setDataValidation).not.toHaveBeenCalled();
  });

  it("fails closed for single-value checkbox plans when allowBlank=false", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      getA1Notation() {
        return "B2:B20";
      },
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn(() => targetRange)
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      },
      getNamedRanges() {
        return [];
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "checkbox",
        checkedValue: "APPROVED",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Single-value checkbox cannot preserve no-blank semantics exactly.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    })).toThrow("Google Sheets host cannot represent allowBlank=false exactly for single-value checkbox validation.");
    expect(setDataValidation).not.toHaveBeenCalled();
  });

  it("fails closed for invalid date literals instead of rolling them over", () => {
    const setDataValidation = vi.fn();
    const targetRange = {
      getA1Notation() {
        return "B2:B20";
      },
      setDataValidation
    };
    const sheet = {
      getRange: vi.fn(() => targetRange)
    };
    const spreadsheet = {
      getSheetByName() {
        return sheet;
      },
      getNamedRanges() {
        return [];
      }
    };
    const { applyWritePlan } = loadCodeModule({ spreadsheet });

    expect(() => applyWritePlan({
      plan: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "date",
        comparator: "less_than_or_equal_to",
        value: "2026-13-40",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Malformed date should fail closed.",
        confidence: 0.5,
        requiresConfirmation: true
      }
    })).toThrow("Invalid date literal: 2026-13-40");
    expect(setDataValidation).not.toHaveBeenCalled();
  });
});
