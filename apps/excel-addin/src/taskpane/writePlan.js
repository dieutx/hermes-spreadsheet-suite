import {
  getAnalysisReportStatusSummary,
  getChartStatusSummary,
  getPivotTableStatusSummary
} from "./analysisArtifactsPlan.js?v=20260423t";
import { getCompositeStatusSummary } from "./compositePlan.js?v=20260423t";

export function hasNonEmptyNoteValues(plan) {
  return Array.isArray(plan?.notes) && plan.notes.some((row) =>
    Array.isArray(row) && row.some((value) => value != null && value !== "")
  );
}

const WORKBOOK_STRUCTURE_OPERATIONS = new Set([
  "create_sheet",
  "delete_sheet",
  "rename_sheet",
  "duplicate_sheet",
  "move_sheet",
  "hide_sheet",
  "unhide_sheet"
]);

export function isWorkbookStructurePlan(plan) {
  return WORKBOOK_STRUCTURE_OPERATIONS.has(plan?.operation);
}

export function isRangeFormatPlan(plan) {
  return Boolean(
    plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    typeof plan.format === "object" &&
    plan.format !== null
  );
}

export function isTablePlan(plan) {
  return Boolean(
    plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    typeof plan.hasHeaders === "boolean"
  );
}

export function getTablePreviewSummary(plan) {
  const label = typeof plan?.name === "string" && plan.name.trim().length > 0
    ? ` ${plan.name.trim()}`
    : "";
  return `Will format table${label} on ${plan.targetSheet}!${plan.targetRange}.`;
}

export function getTableStatusSummary(plan, options = {}) {
  const label = typeof plan?.name === "string" && plan.name.trim().length > 0
    ? ` ${plan.name.trim()}`
    : "";
  const verb = options.tableLike ? "Formatted table-like range" : "Created table";
  return `${verb}${label} on ${plan.targetSheet}!${plan.targetRange}.`;
}

export function getWorkbookStructureStatusSummary(plan) {
  switch (plan?.operation) {
    case "create_sheet":
      return `Created sheet ${plan.sheetName}.`;
    case "delete_sheet":
      return `Deleted sheet ${plan.sheetName}.`;
    case "rename_sheet":
      return `Renamed sheet ${plan.sheetName} to ${plan.newSheetName}.`;
    case "duplicate_sheet":
      return `Duplicated sheet ${plan.sheetName}${plan.newSheetName ? ` as ${plan.newSheetName}` : ""}.`;
    case "move_sheet":
      return `Moved sheet ${plan.sheetName}.`;
    case "hide_sheet":
      return `Hid sheet ${plan.sheetName}.`;
    case "unhide_sheet":
      return `Unhid sheet ${plan.sheetName}.`;
    default:
      return "Workbook update applied.";
  }
}

export function getRangeTransferStatusSummary(plan) {
  const operation = plan?.transferOperation || plan?.operation;
  const source = plan?.sourceSheet && plan?.sourceRange
    ? `${plan.sourceSheet}!${plan.sourceRange}`
    : "the source range";
  const target = plan?.targetSheet && plan?.targetRange
    ? `${plan.targetSheet}!${plan.targetRange}`
    : "the target range";

  switch (operation) {
    case "copy":
      return `Copied ${source} to ${target}.`;
    case "move":
      return `Moved ${source} to ${target}.`;
    case "append":
      return `Appended ${source} into ${target}.`;
    default:
      return `Transferred ${source} to ${target}.`;
  }
}

export function getDataCleanupStatusSummary(plan) {
  const operation = plan?.cleanupOperation || plan?.operation;
  const target = plan?.targetSheet && plan?.targetRange
    ? `${plan.targetSheet}!${plan.targetRange}`
    : "the target range";

  switch (operation) {
    case "trim_whitespace":
      return `Trimmed whitespace in ${target}.`;
    case "remove_blank_rows":
      return `Removed blank rows from ${target}.`;
    case "remove_duplicate_rows":
      return `Removed duplicate rows from ${target}.`;
    case "normalize_case":
      return `Normalized case in ${target}.`;
    case "split_column":
      return `Split column values in ${target}.`;
    case "join_columns":
      return `Joined column values in ${target}.`;
    case "fill_down":
      return `Filled down values in ${target}.`;
    case "standardize_format":
      return `Standardized format in ${target}.`;
    default:
      return `Applied cleanup in ${target}.`;
  }
}

export function getRangeFormatStatusSummary(plan) {
  const target = plan?.targetSheet && plan?.targetRange
    ? `${plan.targetSheet}!${plan.targetRange}`
    : "the target range";
  return `Applied formatting to ${target}.`;
}

export function expandRangeBorderLines(border) {
  if (!border || typeof border !== "object") {
    return [];
  }

  const lines = [];
  const pushLine = (side, line) => {
    if (line && typeof line === "object" && typeof line.style === "string") {
      lines.push({ side, line });
    }
  };

  if (border.all) {
    ["top", "bottom", "left", "right", "innerHorizontal", "innerVertical"]
      .forEach((side) => pushLine(side, border.all));
  }

  if (border.outer) {
    ["top", "bottom", "left", "right"].forEach((side) => pushLine(side, border.outer));
  }

  if (border.inner) {
    ["innerHorizontal", "innerVertical"].forEach((side) => pushLine(side, border.inner));
  }

  pushLine("top", border.top);
  pushLine("bottom", border.bottom);
  pushLine("left", border.left);
  pushLine("right", border.right);
  pushLine("innerHorizontal", border.innerHorizontal);
  pushLine("innerVertical", border.innerVertical);

  return lines;
}

export function getCompositeStepWritebackStatusLine(plan, result) {
  if (result?.kind === "range_write") {
    const targetSheet = plan?.targetSheet || result.targetSheet;
    const targetRange = plan?.targetRange || result.targetRange;
    const target = targetSheet && targetRange ? `${targetSheet}!${targetRange}` : null;
    const hasValues = Array.isArray(plan?.values);
    const hasFormulas = Array.isArray(plan?.formulas);
    const hasNotes = Array.isArray(plan?.notes);

    if (plan?.sourceAttachmentId && target) {
      return `Inserted imported data into ${target}.`;
    }

    if (hasFormulas && !hasValues && !hasNotes && target) {
      return result.writtenRows === 1 && result.writtenColumns === 1
        ? `Set a formula in ${target}.`
        : `Set formulas in ${target}.`;
    }

    if (hasValues && !hasFormulas && !hasNotes && target) {
      return result.writtenRows === 1 && result.writtenColumns === 1
        ? `Wrote a value to ${target}.`
        : `Wrote values to ${target}.`;
    }

    if (hasNotes && !hasValues && !hasFormulas && target) {
      return `Updated notes in ${target}.`;
    }

    if (target && (hasValues || hasFormulas || hasNotes)) {
      return `Updated cells in ${target}.`;
    }
  }

  return getWritebackStatusLine(result);
}

export function getWritebackStatusLine(result) {
  if (result?.kind === "range_format_update" ||
    result?.kind === "workbook_structure_update" ||
    result?.kind === "sheet_structure_update" ||
    result?.kind === "range_sort" ||
    result?.kind === "range_filter") {
    return result.summary || "Workbook update applied.";
  }

  if (result?.kind === "data_validation_update" ||
    result?.kind === "named_range_update" ||
    result?.kind === "conditional_format_update") {
    return result.summary || "Write applied.";
  }

  if (result?.kind === "range_transfer_update") {
    return result.summary || getRangeTransferStatusSummary(result);
  }

  if (result?.kind === "data_cleanup_update") {
    return result.summary || getDataCleanupStatusSummary(result);
  }

  if (result?.kind === "external_data_update") {
    return result.summary || "Applied external data formula.";
  }

  if (result?.kind === "analysis_report_update") {
    return result.summary || getAnalysisReportStatusSummary(result);
  }

  if (result?.kind === "pivot_table_update") {
    return result.summary || getPivotTableStatusSummary(result);
  }

  if (result?.kind === "chart_update") {
    return result.summary || getChartStatusSummary(result);
  }

  if (result?.kind === "table_update") {
    return result.summary || getTableStatusSummary(result);
  }

  if (result?.kind === "composite_update") {
    return getCompositeStatusSummary(result);
  }

  if (result?.kind === "range_write" &&
    typeof result.targetSheet === "string" &&
    typeof result.targetRange === "string") {
    return `Write applied to ${result.targetSheet}!${result.targetRange}`;
  }

  return "Write applied.";
}

export function getDataValidationStatusSummary(plan) {
  if (plan?.targetSheet && plan?.targetRange) {
    return `Applied validation to ${plan.targetSheet}!${plan.targetRange}.`;
  }

  return "Applied validation.";
}

export function getConditionalFormatStatusSummary(plan) {
  if (plan?.targetSheet && plan?.targetRange) {
    switch (plan.managementMode) {
      case "add":
        return `Added conditional formatting to ${plan.targetSheet}!${plan.targetRange}.`;
      case "replace_all_on_target":
        return `Replaced conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
      case "clear_on_target":
        return `Cleared conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
      default:
        break;
    }
  }

  switch (plan?.managementMode) {
    case "add":
      return "Added conditional formatting.";
    case "replace_all_on_target":
      return "Replaced conditional formatting.";
    case "clear_on_target":
      return "Cleared conditional formatting.";
    default:
      return "Updated conditional formatting.";
  }
}

export function getNamedRangeStatusSummary(plan) {
  if (plan?.operation === "delete") {
    return `Deleted named range ${plan.name}.`;
  }

  if (plan?.operation === "rename" && plan?.newName) {
    return `Renamed named range ${plan.name} to ${plan.newName}.`;
  }

  if (plan?.operation === "create" && plan?.targetSheet && plan?.targetRange) {
    return `Created named range ${plan.name} at ${plan.targetSheet}!${plan.targetRange}.`;
  }

  if (plan?.targetSheet && plan?.targetRange) {
    return `Retargeted ${plan.name} to ${plan.targetSheet}!${plan.targetRange}.`;
  }

  if (plan?.name) {
    return `Updated named range ${plan.name}.`;
  }

  return "Updated named range.";
}

export function mapHorizontalAlignmentToExcel(alignment) {
  switch (alignment) {
    case "left":
      return "Left";
    case "center":
      return "Center";
    case "right":
      return "Right";
    case "justify":
      return "Justify";
    case "general":
      return "General";
    default:
      return undefined;
  }
}

export function mapVerticalAlignmentToExcel(alignment) {
  switch (alignment) {
    case "top":
      return "Top";
    case "middle":
      return "Center";
    case "bottom":
      return "Bottom";
    default:
      return undefined;
  }
}

export function mapWrapStrategyToExcel(strategy) {
  switch (strategy) {
    case "wrap":
      return true;
    case "overflow":
      return false;
    default:
      return undefined;
  }
}
