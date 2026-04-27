import { describe, expect, it } from "vitest";
import {
  expandRangeBorderLines,
  getCompositeStepWritebackStatusLine,
  getConditionalFormatStatusSummary,
  getDataCleanupStatusSummary,
  getRangeTransferStatusSummary,
  getWorkbookStructureStatusSummary,
  getWritebackStatusLine,
  hasNonEmptyNoteValues,
  isRangeFormatPlan,
  isWorkbookStructurePlan,
  mapHorizontalAlignmentToExcel,
  mapVerticalAlignmentToExcel,
  mapWrapStrategyToExcel
} from "../../../apps/excel-addin/src/taskpane/writePlan.js";

describe("Excel write plan helpers", () => {
  it("detects when a plan contains non-empty note values", () => {
    expect(hasNonEmptyNoteValues({
      notes: [
        [null, ""],
        ["Needs review", null]
      ]
    })).toBe(true);
  });

  it("treats empty or missing note matrices as unsupported-noop", () => {
    expect(hasNonEmptyNoteValues({})).toBe(false);
    expect(hasNonEmptyNoteValues({
      notes: [
        [null, ""]
      ]
    })).toBe(false);
  });

  it("detects workbook structure and range format plans", () => {
    expect(isWorkbookStructurePlan({
      operation: "create_sheet",
      sheetName: "SheetX"
    })).toBe(true);
    expect(isWorkbookStructurePlan({
      operation: "set_formulas",
      targetSheet: "Sheet1"
    })).toBe(false);

    expect(isRangeFormatPlan({
      targetSheet: "Sheet1",
      targetRange: "A1:B2",
      format: {
        backgroundColor: "#ffffff"
      }
    })).toBe(true);
    expect(isRangeFormatPlan({
      targetSheet: "Sheet1",
      targetRange: "A1:B2"
    })).toBe(false);
  });

  it("expands range border groups with specific side overrides last", () => {
    expect(expandRangeBorderLines({
      outer: {
        style: "solid",
        color: "#1f1f1f"
      },
      inner: {
        style: "dotted",
        color: "#d9d9d9"
      },
      top: {
        style: "thick",
        color: "#000000"
      }
    })).toEqual([
      { side: "top", line: { style: "solid", color: "#1f1f1f" } },
      { side: "bottom", line: { style: "solid", color: "#1f1f1f" } },
      { side: "left", line: { style: "solid", color: "#1f1f1f" } },
      { side: "right", line: { style: "solid", color: "#1f1f1f" } },
      { side: "innerHorizontal", line: { style: "dotted", color: "#d9d9d9" } },
      { side: "innerVertical", line: { style: "dotted", color: "#d9d9d9" } },
      { side: "top", line: { style: "thick", color: "#000000" } }
    ]);
  });

  it("builds status summaries for workbook and range writes", () => {
    expect(getWorkbookStructureStatusSummary({
      operation: "rename_sheet",
      sheetName: "Old",
      newSheetName: "New"
    })).toBe("Renamed sheet Old to New.");

    expect(getWritebackStatusLine({
      kind: "workbook_structure_update",
      summary: "Created sheet New Sheet."
    })).toBe("Created sheet New Sheet.");

    expect(getWritebackStatusLine({
      kind: "range_write",
      targetSheet: "Sheet1",
      targetRange: "A1:B2"
    })).toBe("Write applied to Sheet1!A1:B2");

    expect(getCompositeStepWritebackStatusLine({
      targetSheet: "Lookup_Demo",
      targetRange: "C1",
      formulas: [["=A1"]]
    }, {
      kind: "range_write",
      targetSheet: "Lookup_Demo",
      targetRange: "C1",
      writtenRows: 1,
      writtenColumns: 1
    })).toBe("Set a formula in Lookup_Demo!C1.");

    expect(getWritebackStatusLine({
      kind: "sheet_structure_update",
      summary: "Inserted 2 rows at Sheet1 row 5."
    })).toBe("Inserted 2 rows at Sheet1 row 5.");

    expect(getWritebackStatusLine({
      kind: "range_sort",
      summary: "Sorted Sheet1!A1:F25 by Status (ascending)."
    })).toBe("Sorted Sheet1!A1:F25 by Status (ascending).");

    expect(getWritebackStatusLine({
      kind: "range_filter",
      summary: "Applied filter to Sheet1!A1:F25."
    })).toBe("Applied filter to Sheet1!A1:F25.");
  });

  it("builds status summaries for data validation and named range writes", () => {
    expect(getWritebackStatusLine({
      kind: "data_validation_update",
      summary: "Applied validation to Sheet1!B2:B20."
    })).toBe("Applied validation to Sheet1!B2:B20.");

    expect(getWritebackStatusLine({
      kind: "named_range_update",
      summary: "Retargeted InputRange to Sheet1!B2:D20."
    })).toBe("Retargeted InputRange to Sheet1!B2:D20.");
  });

  it("builds status summaries for wave 5 artifact writes", () => {
    expect(getWritebackStatusLine({
      kind: "analysis_report_update",
      summary: "Created analysis report on Sales Report!A1."
    })).toBe("Created analysis report on Sales Report!A1.");

    expect(getWritebackStatusLine({
      kind: "pivot_table_update",
      summary: "Created pivot table on Sales Pivot!A1."
    })).toBe("Created pivot table on Sales Pivot!A1.");

    expect(getWritebackStatusLine({
      kind: "chart_update",
      summary: "Created line chart on Sales Chart!A1."
    })).toBe("Created line chart on Sales Chart!A1.");
  });

  it("surfaces completed composite step actions instead of only step counts", () => {
    expect(getWritebackStatusLine({
      kind: "composite_update",
      executionId: "exec_001",
      stepResults: [
        { stepId: "create_sheet", status: "completed", summary: "Created sheet Lookup_Demo." },
        { stepId: "set_formula", status: "completed", summary: "Set a formula in Lookup_Demo!C1." }
      ],
      summary: "Workflow finished: 2 steps • 2 completed."
    })).toBe(
      "Workflow finished: 2 steps • 2 completed. Completed: Created sheet Lookup_Demo; Set a formula in Lookup_Demo!C1."
    );
  });

  it("builds status summaries for range transfer and cleanup writes", () => {
    expect(getRangeTransferStatusSummary({
      sourceSheet: "RawData",
      sourceRange: "A2:C10",
      targetSheet: "Report",
      targetRange: "B5:D13",
      operation: "copy"
    })).toBe("Copied RawData!A2:C10 to Report!B5:D13.");

    expect(getRangeTransferStatusSummary({
      sourceSheet: "RawData",
      sourceRange: "A2:C10",
      targetSheet: "Archive",
      targetRange: "A2:C10",
      operation: "move"
    })).toBe("Moved RawData!A2:C10 to Archive!A2:C10.");

    expect(getDataCleanupStatusSummary({
      targetSheet: "Contacts",
      targetRange: "A2:F100",
      operation: "trim_whitespace"
    })).toBe("Trimmed whitespace in Contacts!A2:F100.");

    expect(getDataCleanupStatusSummary({
      targetSheet: "Contacts",
      targetRange: "A2:F100",
      operation: "remove_duplicate_rows"
    })).toBe("Removed duplicate rows from Contacts!A2:F100.");

    expect(getWritebackStatusLine({
      kind: "range_transfer_update",
      summary: "Copied RawData!A2:C10 to Report!B5:D13."
    })).toBe("Copied RawData!A2:C10 to Report!B5:D13.");

    expect(getWritebackStatusLine({
      kind: "data_cleanup_update",
      summary: "Trimmed whitespace in Contacts!A2:F100."
    })).toBe("Trimmed whitespace in Contacts!A2:F100.");
  });

  it("builds status summaries for conditional-format writes", () => {
    expect(getConditionalFormatStatusSummary({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add"
    })).toBe("Added conditional formatting to Sheet1!B2:B20.");

    expect(getConditionalFormatStatusSummary({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target"
    })).toBe("Replaced conditional formatting on Sheet1!B2:B20.");

    expect(getConditionalFormatStatusSummary({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "clear_on_target"
    })).toBe("Cleared conditional formatting on Sheet1!B2:B20.");

    expect(getWritebackStatusLine({
      kind: "conditional_format_update",
      summary: "Added conditional formatting to Sheet1!B2:B20."
    })).toBe("Added conditional formatting to Sheet1!B2:B20.");
  });

  it("maps spreadsheet alignment and wrap semantics to Excel values", () => {
    expect(mapHorizontalAlignmentToExcel("center")).toBe("Center");
    expect(mapHorizontalAlignmentToExcel("general")).toBe("General");
    expect(mapVerticalAlignmentToExcel("middle")).toBe("Center");
    expect(mapVerticalAlignmentToExcel("bottom")).toBe("Bottom");
    expect(mapWrapStrategyToExcel("wrap")).toBe(true);
    expect(mapWrapStrategyToExcel("clip")).toBeUndefined();
    expect(mapWrapStrategyToExcel("overflow")).toBe(false);
  });
});
