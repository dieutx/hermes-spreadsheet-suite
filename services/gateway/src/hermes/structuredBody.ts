import { z } from "zod";
import {
  AnalysisReportPlanDataSchema,
  AnalysisReportUpdateDataSchema,
  AttachmentAnalysisDataSchema,
  ChartPlanDataSchema,
  ChartUpdateDataSchema,
  CompositePlanDataSchema,
  ChatDataSchema,
  ConditionalFormatPlanDataSchema,
  DataValidationPlanDataSchema,
  DataCleanupPlanDataSchema,
  DocumentSummaryDataSchema,
  DownstreamProviderSchema,
  ErrorDataSchema,
  ExternalDataPlanDataSchema,
  ExtractedTableDataSchema,
  FormulaDataSchema,
  NamedRangeUpdateDataSchema,
  PivotTablePlanDataSchema,
  PivotTableUpdateDataSchema,
  RangeFormatUpdateDataSchema,
  RangeFilterPlanDataSchema,
  RangeSortPlanDataSchema,
  RangeTransferPlanDataSchema,
  SheetImportPlanDataSchema,
  SheetStructureUpdateDataSchema,
  SheetUpdateDataSchema,
  TablePlanDataSchema,
  TableUpdateDataSchema,
  WarningSchema,
  WorkbookStructureUpdateDataSchema
} from "@hermes/contracts";

type JsonRecord = Record<string, unknown>;

const STRUCTURED_BODY_TYPES = [
  "chat",
  "formula",
  "composite_plan",
  "workbook_structure_update",
  "range_format_update",
  "conditional_format_plan",
  "sheet_structure_update",
  "range_sort_plan",
  "range_filter_plan",
  "data_validation_plan",
  "analysis_report_plan",
  "pivot_table_plan",
  "chart_plan",
  "table_plan",
  "named_range_update",
  "range_transfer_plan",
  "data_cleanup_plan",
  "analysis_report_update",
  "pivot_table_update",
  "chart_update",
  "table_update",
  "sheet_update",
  "sheet_import_plan",
  "external_data_plan",
  "error",
  "attachment_analysis",
  "extracted_table",
  "document_summary"
] as const;

type HermesStructuredBodyType = (typeof STRUCTURED_BODY_TYPES)[number];

export function extractSingleJsonObjectText(content: string): string | null {
  const trimmed = content.trim();
  const fencedMatch = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  const candidate = (fencedMatch ? fencedMatch[1] : trimmed).trim();

  if (candidate.startsWith("{") && candidate.endsWith("}")) {
    return candidate;
  }

  return null;
}

function isObject(value: unknown): value is JsonRecord {
  return typeof value === "object" && value !== null;
}

function hasOwn(source: JsonRecord, key: string): boolean {
  return Object.prototype.hasOwnProperty.call(source, key);
}

function pickFields(source: JsonRecord, keys: readonly string[]): JsonRecord {
  const picked: JsonRecord = {};

  for (const key of keys) {
    if (hasOwn(source, key)) {
      picked[key] = source[key];
    }
  }

  return picked;
}

function normalizeWarningsValue(value: unknown): unknown {
  if (!Array.isArray(value)) {
    return value;
  }

  if (value.every((item) => typeof item === "string")) {
    return value.map((message) => ({
      code: "MODEL_WARNING",
      message,
      severity: "warning" as const
    }));
  }

  return value.map((item) => {
    if (!isObject(item)) {
      return item;
    }

    const normalized = pickFields(item, ["code", "message", "severity", "field"]);
    if (typeof normalized.severity === "string") {
      const severity = normalized.severity.trim().toLowerCase();
      normalized.severity =
        severity === "low" || severity === "info"
          ? "info"
          : severity === "medium" || severity === "warn" || severity === "warning"
          ? "warning"
          : severity === "high" || severity === "critical" || severity === "error"
          ? "error"
          : normalized.severity;
    }

    return normalized;
  });
}

function normalizeShapeValue(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  return pickFields(value, ["rows", "columns"]);
}

function inferShapeFromMatrix(value: unknown): { rows: number; columns: number } | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }

  const firstRow = value[0];
  return {
    rows: value.length,
    columns: Array.isArray(firstRow) ? firstRow.length : 0
  };
}

function buildQualifiedRangeRef(sheet: unknown, range: unknown): string | undefined {
  if (typeof sheet !== "string" || !sheet.trim() || typeof range !== "string" || !range.trim()) {
    return undefined;
  }

  return `${sheet.trim()}!${range.trim()}`;
}

function parseQualifiedRangeRef(value: unknown): { sheet?: string; range?: string } | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  const separatorIndex = trimmed.lastIndexOf("!");
  if (separatorIndex < 1 || separatorIndex === trimmed.length - 1) {
    return { range: trimmed };
  }

  const rawSheet = trimmed.slice(0, separatorIndex).trim();
  const rawRange = trimmed.slice(separatorIndex + 1).trim();
  const sheet = rawSheet.startsWith("'") && rawSheet.endsWith("'")
    ? rawSheet.slice(1, -1).replace(/''/g, "'")
    : rawSheet;

  return {
    sheet: sheet || undefined,
    range: rawRange || undefined
  };
}

function columnLettersToNumber(value: string): number | undefined {
  if (!/^[A-Z]+$/i.test(value)) {
    return undefined;
  }

  let column = 0;
  for (const character of value.toUpperCase()) {
    column = (column * 26) + (character.charCodeAt(0) - 64);
  }
  return column;
}

function columnNumberToLetters(value: number): string {
  let remaining = value;
  let letters = "";

  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return letters;
}

function normalizeStructuredA1Range(value: string): string | undefined {
  const normalized = value.trim().replaceAll("$", "");
  const match = normalized.match(/^([A-Z]+)([1-9][0-9]*)(?::([A-Z]+)([1-9][0-9]*))?$/i);
  if (!match) {
    return undefined;
  }

  const startColumn = columnLettersToNumber(match[1]);
  const startRow = Number.parseInt(match[2], 10);
  const endColumn = match[3] ? columnLettersToNumber(match[3]) : startColumn;
  const endRow = match[4] ? Number.parseInt(match[4], 10) : startRow;
  if (
    !startColumn ||
    !endColumn ||
    !Number.isInteger(startRow) ||
    !Number.isInteger(endRow) ||
    endColumn < startColumn ||
    endRow < startRow
  ) {
    return undefined;
  }

  const start = `${columnNumberToLetters(startColumn)}${startRow}`;
  const end = `${columnNumberToLetters(endColumn)}${endRow}`;
  return start === end ? start : `${start}:${end}`;
}

function normalizeQualifiedAffectedRangeKey(value: string): string {
  const rangeRef = parseQualifiedRangeRef(value);
  if (!rangeRef?.sheet || !rangeRef.range) {
    return value.trim();
  }

  const normalizedRange = normalizeStructuredA1Range(rangeRef.range);
  return normalizedRange ? `${rangeRef.sheet.trim()}!${normalizedRange}` : value.trim();
}

function expandAnalysisReportAnchorRange(targetRange: unknown, sections: unknown): string | undefined {
  if (typeof targetRange !== "string" || !Array.isArray(sections) || sections.length === 0) {
    return undefined;
  }

  const normalized = targetRange.trim().replaceAll("$", "");
  const match = normalized.match(/^([A-Z]+)([1-9][0-9]*)(?::([A-Z]+)([1-9][0-9]*))?$/i);
  if (!match) {
    return undefined;
  }

  const startColumn = columnLettersToNumber(match[1]);
  const startRow = Number.parseInt(match[2], 10);
  const endColumn = match[3] ? columnLettersToNumber(match[3]) : startColumn;
  const endRow = match[4] ? Number.parseInt(match[4], 10) : startRow;
  if (
    !startColumn ||
    !endColumn ||
    !Number.isInteger(startRow) ||
    !Number.isInteger(endRow) ||
    startColumn !== endColumn ||
    startRow !== endRow
  ) {
    return undefined;
  }

  const resolvedEndColumn = startColumn + 3;
  const resolvedEndRow = startRow + sections.length + 3;
  const startCell = `${columnNumberToLetters(startColumn)}${startRow}`;
  const endCell = `${columnNumberToLetters(resolvedEndColumn)}${resolvedEndRow}`;
  return `${startCell}:${endCell}`;
}

function humanizeIdentifier(value: string): string {
  return value
    .trim()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\b\w/g, (match) => match.toUpperCase());
}

function inferAnalysisSectionType(value: string): string {
  const normalized = value.trim().toLowerCase();

  if (
    normalized.includes("trend") ||
    normalized.includes("velocity") ||
    normalized.includes("momentum")
  ) {
    return "trends";
  }

  if (
    normalized.includes("top") ||
    normalized.includes("bottom") ||
    normalized.includes("rank") ||
    normalized.includes("leader")
  ) {
    return "top_bottom";
  }

  if (
    normalized.includes("anomal") ||
    normalized.includes("risk") ||
    normalized.includes("breach") ||
    normalized.includes("outlier") ||
    normalized.includes("alert")
  ) {
    return "anomalies";
  }

  if (
    normalized.includes("group") ||
    normalized.includes("category") ||
    normalized.includes("priority") ||
    normalized.includes("region") ||
    normalized.includes("channel") ||
    normalized.includes("breakdown")
  ) {
    return "group_breakdown";
  }

  if (
    normalized.includes("next") ||
    normalized.includes("action") ||
    normalized.includes("recommend") ||
    normalized.includes("follow up")
  ) {
    return "next_actions";
  }

  return "summary_stats";
}

function normalizeAnalysisSectionLikeValue(
  value: unknown,
  sourceSheet?: unknown,
  sourceRange?: unknown
): unknown {
  if (typeof value === "string" && value.trim()) {
    const title = humanizeIdentifier(value);
    const sourceRef = buildQualifiedRangeRef(sourceSheet, sourceRange);
    return {
      type: inferAnalysisSectionType(value),
      title,
      summary: `${title} section.`,
      sourceRanges: sourceRef ? [sourceRef] : []
    };
  }

  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, ["type", "title", "summary", "sourceRanges"]);
  const inferredTitle = typeof normalized.title === "string" && normalized.title.trim().length > 0
    ? normalized.title.trim()
    : typeof value.type === "string" && value.type.trim().length > 0
    ? humanizeIdentifier(value.type)
    : undefined;

  if (!normalized.type && typeof value.type === "string" && value.type.trim().length > 0) {
    normalized.type = inferAnalysisSectionType(value.type);
  }

  if (!normalized.type && inferredTitle) {
    normalized.type = inferAnalysisSectionType(inferredTitle);
  }

  if (!normalized.title && inferredTitle) {
    normalized.title = inferredTitle;
  }

  if (
    (!Array.isArray(normalized.sourceRanges) || normalized.sourceRanges.length === 0)
  ) {
    const sourceRef = buildQualifiedRangeRef(sourceSheet, sourceRange);
    if (sourceRef) {
      normalized.sourceRanges = [sourceRef];
    }
  }

  if (
    (!normalized.summary || typeof normalized.summary !== "string" || !normalized.summary.trim()) &&
    typeof normalized.title === "string" &&
    normalized.title.trim().length > 0
  ) {
    normalized.summary = `${normalized.title.trim()} section.`;
  }

  return normalized;
}

function normalizeChartSeriesValue(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, ["field", "label"]);
  const fieldCandidate = normalized.field;
  const labelCandidate = normalized.label;
  const nameCandidate = typeof value.name === "string" ? value.name.trim() : undefined;

  if ((!fieldCandidate || typeof fieldCandidate !== "string" || !fieldCandidate.trim()) && nameCandidate) {
    normalized.field = nameCandidate;
  }

  if (
    (!normalized.label || typeof normalized.label !== "string" || !normalized.label.trim()) &&
    nameCandidate
  ) {
    normalized.label = nameCandidate;
  }

  if (
    (!normalized.field || typeof normalized.field !== "string" || !normalized.field.trim()) &&
    typeof labelCandidate === "string" &&
    labelCandidate.trim()
  ) {
    normalized.field = labelCandidate.trim();
  }

  return normalized;
}

function inferCleanupConfirmationLevel(operation: unknown): "standard" | "destructive" | undefined {
  switch (operation) {
    case "remove_blank_rows":
    case "remove_duplicate_rows":
    case "split_column":
    case "join_columns":
      return "destructive";
    case "trim_whitespace":
    case "normalize_case":
    case "fill_down":
    case "standardize_format":
      return "standard";
    default:
      return undefined;
  }
}

function inferCleanupOverwriteRisk(operation: unknown): "low" | "medium" | "high" | undefined {
  switch (operation) {
    case "remove_blank_rows":
    case "remove_duplicate_rows":
    case "split_column":
    case "join_columns":
      return "high";
    case "standardize_format":
      return "medium";
    case "trim_whitespace":
    case "normalize_case":
    case "fill_down":
      return "low";
    default:
      return undefined;
  }
}

function inferStandardizeFormatDetails(explanation: unknown): { formatType?: string; formatPattern?: string } {
  if (typeof explanation !== "string" || !explanation.trim()) {
    return {};
  }

  const normalized = explanation.trim().toLowerCase();
  const mentionsDate = /\b(date|dates|yyyy|mm|dd)\b/.test(normalized);
  const mentionsNumeric = /\b(currency|currencies|number|numeric|price|revenue|amount|amounts|usd|\$)\b/.test(normalized);

  if (mentionsDate && !mentionsNumeric) {
    return {
      formatType: "date_text",
      formatPattern: "yyyy-mm-dd"
    };
  }

  if (mentionsNumeric && !mentionsDate) {
    return {
      formatType: "number_text",
      formatPattern: "#,##0.00"
    };
  }

  return {};
}

function normalizeOverwriteRiskValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }

  const normalized = value.trim().toLowerCase();
  if (
    normalized === "none" ||
    normalized === "low" ||
    normalized === "medium" ||
    normalized === "high"
  ) {
    return normalized;
  }

  if (!normalized) {
    return value;
  }

  if (
    normalized.includes("delete") ||
    normalized.includes("remove") ||
    normalized.includes("clear") ||
    normalized.includes("drop") ||
    normalized.includes("destroy") ||
    normalized.includes("entire sheet") ||
    normalized.includes("entire range") ||
    normalized.includes("all existing")
  ) {
    return "high";
  }

  if (
    normalized.includes("replace") &&
    (normalized.includes("formula") ||
      normalized.includes("cell") ||
      normalized.includes("note") ||
      normalized.includes("format"))
  ) {
    return "low";
  }

  if (
    normalized.includes("anchor") &&
    (normalized.includes("formula") || normalized.includes("cell"))
  ) {
    return "low";
  }

  if (
    normalized.includes("overwrite") ||
    normalized.includes("replace") ||
    normalized.includes("existing")
  ) {
    return "medium";
  }

  if (
    normalized.includes("append") ||
    normalized.includes("add without replacing") ||
    normalized.includes("non-destructive")
  ) {
    return "none";
  }

  return value;
}

function normalizeDownstreamProviderValue(value: unknown): unknown {
  if (value === null) {
    return null;
  }

  if (typeof value === "string" && value.trim().length > 0) {
    return { label: value.trim() };
  }

  if (!isObject(value)) {
    return value;
  }

  return pickFields(value, ["label", "model"]);
}

function normalizeAlternateFormulas(value: unknown): unknown {
  if (!Array.isArray(value)) {
    return value;
  }

  return value.map((item) => {
    if (typeof item === "string" && item.trim().length > 0) {
      return {
        formula: item.trim(),
        explanation: "Alternative formulation."
      };
    }

    return isObject(item)
      ? pickFields(item, ["formula", "explanation"])
      : item;
  });
}

function normalizeChatData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  return pickFields(value, ["message", "followUpSuggestions", "confidence"]);
}

function normalizeFormulaData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "intent",
    "targetCell",
    "formula",
    "formulaLanguage",
    "explanation",
    "confidence",
    "requiresConfirmation"
  ]);

  if (hasOwn(value, "alternateFormulas")) {
    normalized.alternateFormulas = normalizeAlternateFormulas(value.alternateFormulas);
  }

  return normalized;
}

type CompositeStepPlanType =
  | "sheet_update"
  | "sheet_import_plan"
  | "workbook_structure_update"
  | "range_format_update"
  | "conditional_format_plan"
  | "sheet_structure_update"
  | "range_sort_plan"
  | "range_filter_plan"
  | "data_validation_plan"
  | "analysis_report_plan"
  | "pivot_table_plan"
  | "chart_plan"
  | "table_plan"
  | "named_range_update"
  | "range_transfer_plan"
  | "data_cleanup_plan"
  | "external_data_plan";

type CompositeStepPlanNormalizer = {
  type: CompositeStepPlanType;
  matches(value: JsonRecord): boolean;
};

const SHEET_STRUCTURE_UPDATE_OPERATIONS = new Set([
  "insert_rows",
  "delete_rows",
  "hide_rows",
  "unhide_rows",
  "group_rows",
  "ungroup_rows",
  "insert_columns",
  "delete_columns",
  "hide_columns",
  "unhide_columns",
  "group_columns",
  "ungroup_columns",
  "merge_cells",
  "unmerge_cells",
  "freeze_panes",
  "unfreeze_panes",
  "autofit_rows",
  "autofit_columns",
  "set_sheet_tab_color"
]);

const COMPOSITE_STEP_PLAN_NORMALIZERS: CompositeStepPlanNormalizer[] = [
  {
    type: "sheet_update",
    matches: (value) => hasOwn(value, "formulas") || hasOwn(value, "values") || hasOwn(value, "notes")
  },
  {
    type: "sheet_import_plan",
    matches: (value) => hasOwn(value, "sourceAttachmentId") || hasOwn(value, "extractionMode")
  },
  {
    type: "external_data_plan",
    matches: (value) => hasOwn(value, "sourceType") && hasOwn(value, "provider") && hasOwn(value, "formula")
  },
  {
    type: "workbook_structure_update",
    matches: (value) => hasOwn(value, "sheetName") && !hasOwn(value, "targetSheet")
  },
  {
    type: "range_format_update",
    matches: (value) => hasOwn(value, "format")
  },
  {
    type: "conditional_format_plan",
    matches: (value) => hasOwn(value, "managementMode") || hasOwn(value, "replacesExistingRules")
  },
  {
    type: "sheet_structure_update",
    matches: (value) =>
      hasOwn(value, "targetSheet") &&
      typeof value.operation === "string" &&
      SHEET_STRUCTURE_UPDATE_OPERATIONS.has(value.operation)
  },
  {
    type: "range_sort_plan",
    matches: (value) => hasOwn(value, "keys")
  },
  {
    type: "range_filter_plan",
    matches: (value) => hasOwn(value, "conditions") && hasOwn(value, "combiner")
  },
  {
    type: "data_validation_plan",
    matches: (value) => hasOwn(value, "ruleType")
  },
  {
    type: "analysis_report_plan",
    matches: (value) => hasOwn(value, "outputMode")
  },
  {
    type: "pivot_table_plan",
    matches: (value) => hasOwn(value, "rowGroups") || hasOwn(value, "valueAggregations")
  },
  {
    type: "chart_plan",
    matches: (value) => hasOwn(value, "chartType") && hasOwn(value, "series")
  },
  {
    type: "table_plan",
    matches: (value) => hasOwn(value, "hasHeaders") && hasOwn(value, "targetSheet") && hasOwn(value, "targetRange")
  },
  {
    type: "named_range_update",
    matches: (value) => hasOwn(value, "scope") && hasOwn(value, "name")
  },
  {
    type: "range_transfer_plan",
    matches: (value) => hasOwn(value, "pasteMode") || hasOwn(value, "transpose")
  },
  {
    type: "data_cleanup_plan",
    matches: (value) => hasOwn(value, "keyColumns") || hasOwn(value, "mode") || hasOwn(value, "sourceColumn") || hasOwn(value, "formatType")
  }
];

function isWrappedCompositeStepPlanEnvelope(
  value: JsonRecord
): value is JsonRecord & { type: CompositeStepPlanType; data: unknown } {
  return (
    typeof value.type === "string" &&
    hasOwn(value, "data") &&
    COMPOSITE_STEP_PLAN_NORMALIZERS.some((candidate) => candidate.type === value.type)
  );
}

function normalizeStructuredBodyDataByType(type: HermesStructuredBodyType, value: unknown): unknown {
  const normalized = normalizeHermesStructuredBodyInput({
    type,
    data: value
  });

  if (!isObject(normalized) || !hasOwn(normalized, "data")) {
    return value;
  }

  return normalized.data;
}

function normalizeCompositeStepPlanValue(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  if (isWrappedCompositeStepPlanEnvelope(value)) {
    const normalized = normalizeStructuredBodyDataByType(value.type, value.data);

    if (!isObject(normalized)) {
      return normalized;
    }

    switch (value.type) {
      case "sheet_structure_update":
        return pickFields(normalized, [
          "targetSheet",
          "operation",
          "startIndex",
          "count",
          "targetRange",
          "frozenRows",
          "frozenColumns",
          "color",
          "explanation",
          "confidence",
          "requiresConfirmation",
          "confirmationLevel",
          "affectedRanges",
          "overwriteRisk"
        ]);
      case "range_sort_plan":
        return pickFields(normalized, [
          "targetSheet",
          "targetRange",
          "hasHeader",
          "keys",
          "explanation",
          "confidence",
          "requiresConfirmation",
          "affectedRanges"
        ]);
      case "range_filter_plan":
        return pickFields(normalized, [
          "targetSheet",
          "targetRange",
          "hasHeader",
          "conditions",
          "combiner",
          "clearExistingFilters",
          "explanation",
          "confidence",
          "requiresConfirmation",
          "affectedRanges"
        ]);
      default:
        return normalized;
    }
  }

  for (const candidate of COMPOSITE_STEP_PLAN_NORMALIZERS) {
    if (candidate.matches(value)) {
      const normalized = normalizeStructuredBodyDataByType(candidate.type, value);

      if (!isObject(normalized)) {
        return normalized;
      }

      switch (candidate.type) {
        case "sheet_structure_update":
          return pickFields(normalized, [
            "targetSheet",
            "operation",
            "startIndex",
            "count",
            "targetRange",
            "frozenRows",
            "frozenColumns",
            "color",
            "explanation",
            "confidence",
            "requiresConfirmation",
            "confirmationLevel",
            "affectedRanges",
            "overwriteRisk"
          ]);
        case "range_sort_plan":
          return pickFields(normalized, [
            "targetSheet",
            "targetRange",
            "hasHeader",
            "keys",
            "explanation",
            "confidence",
            "requiresConfirmation",
            "affectedRanges"
          ]);
        case "range_filter_plan":
          return pickFields(normalized, [
            "targetSheet",
            "targetRange",
            "hasHeader",
            "conditions",
            "combiner",
            "clearExistingFilters",
            "explanation",
            "confidence",
            "requiresConfirmation",
            "affectedRanges"
          ]);
        default:
          return normalized;
      }
    }
  }

  return value;
}

function normalizeCompositePlanStepValue(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "stepId",
    "dependsOn",
    "continueOnError",
    "plan"
  ]);

  if (!hasOwn(normalized, "stepId") && hasOwn(value, "id")) {
    normalized.stepId = value.id;
  }

  const dependsOnValue = hasOwn(value, "dependsOn")
    ? value.dependsOn
    : hasOwn(value, "depends")
    ? value.depends
    : value.after;

  if (Array.isArray(dependsOnValue)) {
    normalized.dependsOn = [...dependsOnValue];
  } else if (typeof dependsOnValue === "string" && dependsOnValue.trim()) {
    normalized.dependsOn = [dependsOnValue.trim()];
  } else if (!hasOwn(normalized, "dependsOn")) {
    normalized.dependsOn = [];
  }

  if (!hasOwn(normalized, "continueOnError")) {
    if (hasOwn(value, "continueOnFailure")) {
      normalized.continueOnError = value.continueOnFailure;
    } else {
      normalized.continueOnError = false;
    }
  }

  if (hasOwn(value, "plan")) {
    normalized.plan = normalizeCompositeStepPlanValue(value.plan);
  } else if (hasOwn(value, "type") && hasOwn(value, "data")) {
    normalized.plan = normalizeCompositeStepPlanValue({
      type: value.type,
      data: value.data
    });
  }

  return normalized;
}

function normalizeCompositePlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "steps",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel",
    "reversible",
    "dryRunRecommended",
    "dryRunRequired"
  ]);

  const stepsValue = hasOwn(value, "steps")
    ? value.steps
    : hasOwn(value, "actions")
    ? value.actions
    : value.tasks;

  if (Array.isArray(stepsValue)) {
    normalized.steps = stepsValue.map((item) => normalizeCompositePlanStepValue(item));
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  }

  const normalizedSteps = Array.isArray(normalized.steps) ? normalized.steps : [];
  const aggregateAffectedRanges = Array.isArray(normalized.affectedRanges)
    ? [...normalized.affectedRanges]
    : [];
  const seenAffectedRanges = new Set(
    aggregateAffectedRanges
      .filter((range): range is string => typeof range === "string")
      .map((range) => normalizeQualifiedAffectedRangeKey(range))
  );
  normalizedSteps.forEach((step) => {
    if (!isObject(step) || !isObject(step.plan) || !Array.isArray(step.plan.affectedRanges)) {
      return;
    }

    step.plan.affectedRanges.forEach((range) => {
      if (typeof range !== "string") {
        return;
      }
      const rangeKey = normalizeQualifiedAffectedRangeKey(range);
      if (seenAffectedRanges.has(rangeKey)) {
        return;
      }
      seenAffectedRanges.add(rangeKey);
      aggregateAffectedRanges.push(range);
    });
  });
  normalized.affectedRanges = aggregateAffectedRanges;

  const hasDestructiveStep = normalizedSteps.some((step) =>
    isObject(step) &&
    isObject(step.plan) &&
    step.plan.confirmationLevel === "destructive"
  );
  normalized.confirmationLevel = hasDestructiveStep ? "destructive" : "standard";
  if (!hasOwn(normalized, "reversible")) {
    normalized.reversible = false;
  }

  if (!hasOwn(normalized, "dryRunRecommended")) {
    normalized.dryRunRecommended = true;
  }

  if (!hasOwn(normalized, "dryRunRequired")) {
    normalized.dryRunRequired = false;
  }

  return normalized;
}

function normalizeWorkbookStructureUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "operation",
    "sheetName",
    "newSheetName",
    "position",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "confirmationLevel",
    "overwriteRisk"
  ]);

  if (!hasOwn(normalized, "operation")) {
    if (typeof value.action === "string" && value.action.trim()) {
      normalized.operation = value.action.trim();
    } else if (typeof value.op === "string" && value.op.trim()) {
      normalized.operation = value.op.trim();
    }
  }

  if (typeof normalized.operation === "string") {
    const operation = normalized.operation.trim().toLowerCase();
    normalized.operation =
      operation === "add" || operation === "create" || operation === "insert" || operation === "new"
        ? "create_sheet"
        : operation === "delete" || operation === "remove"
        ? "delete_sheet"
        : operation === "rename"
        ? "rename_sheet"
        : operation === "duplicate" || operation === "copy"
        ? "duplicate_sheet"
        : operation === "move"
        ? "move_sheet"
        : operation === "hide"
        ? "hide_sheet"
        : operation === "unhide" || operation === "show"
        ? "unhide_sheet"
        : operation;
  }

  if (!hasOwn(normalized, "sheetName")) {
    if (typeof value.name === "string" && value.name.trim()) {
      normalized.sheetName = value.name.trim();
    } else if (typeof value.sheet === "string" && value.sheet.trim()) {
      normalized.sheetName = value.sheet.trim();
    } else if (typeof value.sheetTitle === "string" && value.sheetTitle.trim()) {
      normalized.sheetName = value.sheetTitle.trim();
    }
  }

  if (!hasOwn(normalized, "newSheetName")) {
    if (typeof value.newName === "string" && value.newName.trim()) {
      normalized.newSheetName = value.newName.trim();
    } else if (typeof value.newTitle === "string" && value.newTitle.trim()) {
      normalized.newSheetName = value.newTitle.trim();
    }
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel = normalized.operation === "delete_sheet" ? "destructive" : "standard";
  }

  return normalized;
}

function normalizeRangeFormatValue(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "numberFormat",
    "backgroundColor",
    "textColor",
    "fontFamily",
    "fontSize",
    "bold",
    "italic",
    "underline",
    "strikethrough",
    "horizontalAlignment",
    "verticalAlignment",
    "wrapStrategy",
    "border",
    "columnWidth",
    "rowHeight"
  ]);

  if (!hasOwn(normalized, "backgroundColor")) {
    if (typeof value.background === "string" && value.background.trim()) {
      normalized.backgroundColor = value.background.trim();
    } else if (typeof value.fillColor === "string" && value.fillColor.trim()) {
      normalized.backgroundColor = value.fillColor.trim();
    }
  }

  if (!hasOwn(normalized, "textColor") && typeof value.fontColor === "string" && value.fontColor.trim()) {
    normalized.textColor = value.fontColor.trim();
  }

  if (!hasOwn(normalized, "wrapStrategy")) {
    if (value.wrapText === true) {
      normalized.wrapStrategy = "wrap";
    } else if (typeof value.wrap === "string" && value.wrap.trim()) {
      normalized.wrapStrategy = value.wrap.trim();
    }
  }

  return normalized;
}

function normalizeRangeFormatUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "targetSheet",
    "targetRange",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "confirmationLevel",
    "overwriteRisk"
  ]);

  if (!hasOwn(normalized, "targetSheet")) {
    if (typeof value.sheet === "string" && value.sheet.trim()) {
      normalized.targetSheet = value.sheet.trim();
    } else if (typeof value.sheetName === "string" && value.sheetName.trim()) {
      normalized.targetSheet = value.sheetName.trim();
    }
  }

  if (!hasOwn(normalized, "targetRange") && hasOwn(value, "range")) {
    const rangeRef = parseQualifiedRangeRef(value.range);
    if (!hasOwn(normalized, "targetSheet") && rangeRef?.sheet) {
      normalized.targetSheet = rangeRef.sheet;
    }
    if (rangeRef?.range) {
      normalized.targetRange = rangeRef.range;
    }
  }

  const formatSource = hasOwn(value, "format") ? value.format : value;
  const normalizedFormat = normalizeRangeFormatValue(formatSource);
  if (isObject(normalizedFormat) && Object.keys(normalizedFormat).length > 0) {
    normalized.format = normalizedFormat;
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (
    (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) &&
    typeof normalized.targetSheet === "string" &&
    typeof normalized.targetRange === "string"
  ) {
    normalized.affectedRanges = [`${normalized.targetSheet}!${normalized.targetRange}`];
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel = "standard";
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  }

  return normalized;
}

function normalizeConditionalFormatStyleValue(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "backgroundColor",
    "textColor",
    "bold",
    "italic",
    "underline",
    "strikethrough",
    "numberFormat"
  ]);

  if (!hasOwn(normalized, "backgroundColor")) {
    if (hasOwn(value, "background")) {
      normalized.backgroundColor = value.background;
    } else if (hasOwn(value, "fillColor")) {
      normalized.backgroundColor = value.fillColor;
    }
  }

  if (!hasOwn(normalized, "textColor")) {
    if (hasOwn(value, "fontColor")) {
      normalized.textColor = value.fontColor;
    } else if (hasOwn(value, "foregroundColor")) {
      normalized.textColor = value.foregroundColor;
    }
  }

  return normalized;
}

function normalizeConditionalFormatColorScalePointValue(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  return pickFields(value, [
    "type",
    "value",
    "color"
  ]);
}

function normalizeConditionalFormatRuleTypeValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }

  const normalized = value
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .replace(/[\s-]+/g, "_")
    .toLowerCase();

  switch (normalized) {
    case "single":
    case "single_color":
    case "cell_value":
      return "single_color";
    case "text":
    case "contains_text":
    case "text_contains":
      return "text_contains";
    case "number":
    case "numeric":
    case "number_compare":
      return "number_compare";
    case "date":
    case "date_compare":
      return "date_compare";
    case "duplicate":
    case "duplicates":
    case "duplicate_values":
      return "duplicate_values";
    case "formula":
    case "custom":
    case "custom_formula":
      return "custom_formula";
    case "top":
    case "top_n":
    case "top_bottom":
      return "top_n";
    case "average":
    case "average_compare":
    case "above_below_average":
      return "average_compare";
    case "color_scale":
      return "color_scale";
    default:
      return value;
  }
}

function normalizeConditionalFormatManagementModeValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }

  const normalized = value
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .replace(/[\s-]+/g, "_")
    .toLowerCase();

  switch (normalized) {
    case "add":
    case "append":
    case "add_rule":
      return "add";
    case "replace":
    case "replace_all":
    case "replace_existing":
    case "replace_existing_rules":
    case "replace_all_rules":
    case "replace_all_on_target":
      return "replace_all_on_target";
    case "clear":
    case "clear_existing":
    case "clear_existing_rules":
    case "clear_rules":
    case "clear_on_target":
      return "clear_on_target";
    default:
      return value;
  }
}

function normalizeSnakeComparatorValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }

  const normalized = value
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .replace(/[\s-]+/g, "_")
    .toLowerCase();

  switch (normalized) {
    case "equal":
    case "equals":
    case "equal_to":
      return "equal_to";
    case "not_equal":
    case "not_equals":
    case "not_equal_to":
      return "not_equal_to";
    case "greater":
    case "greater_than":
      return "greater_than";
    case "greater_or_equal":
    case "greater_than_or_equal":
    case "greater_than_or_equal_to":
      return "greater_than_or_equal_to";
    case "less":
    case "less_than":
      return "less_than";
    case "less_or_equal":
    case "less_than_or_equal":
    case "less_than_or_equal_to":
      return "less_than_or_equal_to";
    case "between":
      return "between";
    case "not_between":
      return "not_between";
    default:
      return value;
  }
}

function normalizeConditionalFormatPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "targetSheet",
    "targetRange",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "replacesExistingRules",
    "managementMode",
    "ruleType",
    "comparator",
    "value",
    "value2",
    "text",
    "formula",
    "rank",
    "direction"
  ]);

  if (!hasOwn(normalized, "targetSheet")) {
    if (hasOwn(value, "sheet")) {
      normalized.targetSheet = value.sheet;
    } else if (hasOwn(value, "sheetName")) {
      normalized.targetSheet = value.sheetName;
    }
  }

  if (!hasOwn(normalized, "targetRange")) {
    const rangeValue = hasOwn(value, "range") ? value.range : value.target;
    const rangeRef = parseQualifiedRangeRef(rangeValue);
    if (!hasOwn(normalized, "targetSheet") && rangeRef?.sheet) {
      normalized.targetSheet = rangeRef.sheet;
    }
    if (rangeRef?.range) {
      normalized.targetRange = rangeRef.range;
    }
  }

  if (!hasOwn(normalized, "managementMode")) {
    if (hasOwn(value, "mode")) {
      normalized.managementMode = normalizeConditionalFormatManagementModeValue(value.mode);
    } else if (hasOwn(value, "action")) {
      normalized.managementMode = normalizeConditionalFormatManagementModeValue(value.action);
    }
  } else {
    normalized.managementMode = normalizeConditionalFormatManagementModeValue(normalized.managementMode);
  }

  if (hasOwn(normalized, "ruleType")) {
    normalized.ruleType = normalizeConditionalFormatRuleTypeValue(normalized.ruleType);
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (hasOwn(value, "style")) {
    normalized.style = normalizeConditionalFormatStyleValue(value.style);
  } else if (hasOwn(value, "format")) {
    normalized.style = normalizeConditionalFormatStyleValue(value.format);
  }

  if (hasOwn(value, "rule") && isObject(value.rule)) {
    const rule = value.rule;
    if (!hasOwn(normalized, "ruleType") && hasOwn(rule, "ruleType")) {
      normalized.ruleType = normalizeConditionalFormatRuleTypeValue(rule.ruleType);
    }
    if (!hasOwn(normalized, "ruleType") && hasOwn(rule, "type")) {
      normalized.ruleType = normalizeConditionalFormatRuleTypeValue(rule.type);
    }
    if (!hasOwn(normalized, "comparator") && hasOwn(rule, "comparator")) {
      normalized.comparator = rule.comparator;
    }
    if (!hasOwn(normalized, "value") && hasOwn(rule, "value")) {
      normalized.value = rule.value;
    }
    if (!hasOwn(normalized, "value2") && hasOwn(rule, "value2")) {
      normalized.value2 = rule.value2;
    }
    if (!hasOwn(normalized, "text") && hasOwn(rule, "text")) {
      normalized.text = rule.text;
    }
    if (!hasOwn(normalized, "formula")) {
      if (hasOwn(rule, "formula")) {
        normalized.formula = rule.formula;
      } else if (hasOwn(rule, "customFormula")) {
        normalized.formula = rule.customFormula;
      }
    }
    if (!hasOwn(normalized, "rank") && hasOwn(rule, "rank")) {
      normalized.rank = rule.rank;
    }
    if (!hasOwn(normalized, "direction") && hasOwn(rule, "direction")) {
      normalized.direction = rule.direction;
    }
    if (!hasOwn(normalized, "style")) {
      if (hasOwn(rule, "style")) {
        normalized.style = normalizeConditionalFormatStyleValue(rule.style);
      } else if (hasOwn(rule, "format")) {
        normalized.style = normalizeConditionalFormatStyleValue(rule.format);
      }
    }
    if (!hasOwn(normalized, "points") && hasOwn(rule, "points") && Array.isArray(rule.points)) {
      normalized.points = rule.points.map((item) => normalizeConditionalFormatColorScalePointValue(item));
    }
  }

  if (hasOwn(value, "points") && Array.isArray(value.points)) {
    normalized.points = value.points.map((item) => normalizeConditionalFormatColorScalePointValue(item));
  }

  if (!hasOwn(normalized, "formula") && hasOwn(value, "customFormula")) {
    normalized.formula = value.customFormula;
  }

  if (hasOwn(normalized, "comparator")) {
    normalized.comparator = normalizeSnakeComparatorValue(normalized.comparator);
  }

  if (!hasOwn(normalized, "replacesExistingRules")) {
    if (normalized.managementMode === "replace_all_on_target" || normalized.managementMode === "clear_on_target") {
      normalized.replacesExistingRules = true;
    } else if (normalized.managementMode === "add") {
      normalized.replacesExistingRules = false;
    }
  }

  if (
    (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0)
  ) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  return normalized;
}

function normalizeSheetStructureUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const knownInputFields = new Set([
    "targetSheet",
    "operation",
    "startIndex",
    "count",
    "targetRange",
    "frozenRows",
    "frozenColumns",
    "color",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "confirmationLevel",
    "affectedRanges",
    "overwriteRisk",
    "sheet",
    "sheetName",
    "action",
    "op",
    "range"
  ]);
  const normalized = pickFields(value, [
    "targetSheet",
    "operation",
    "startIndex",
    "count",
    "targetRange",
    "frozenRows",
    "frozenColumns",
    "color",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "confirmationLevel",
    "affectedRanges",
    "overwriteRisk"
  ]);

  if (!hasOwn(normalized, "targetSheet")) {
    if (typeof value.sheet === "string" && value.sheet.trim()) {
      normalized.targetSheet = value.sheet.trim();
    } else if (typeof value.sheetName === "string" && value.sheetName.trim()) {
      normalized.targetSheet = value.sheetName.trim();
    }
  }

  if (!hasOwn(normalized, "targetRange") && hasOwn(value, "range")) {
    const rangeRef = parseQualifiedRangeRef(value.range);
    if (!hasOwn(normalized, "targetSheet") && rangeRef?.sheet) {
      normalized.targetSheet = rangeRef.sheet;
    }
    if (rangeRef?.range) {
      normalized.targetRange = rangeRef.range;
    }
  }

  if (!hasOwn(normalized, "operation")) {
    if (typeof value.action === "string" && value.action.trim()) {
      normalized.operation = value.action.trim();
    } else if (typeof value.op === "string" && value.op.trim()) {
      normalized.operation = value.op.trim();
    }
  }

  if (typeof normalized.operation === "string") {
    const operation = normalized.operation.trim().toLowerCase();
    normalized.operation =
      operation === "merge"
        ? "merge_cells"
        : operation === "unmerge"
        ? "unmerge_cells"
        : operation === "freeze"
        ? "freeze_panes"
        : operation === "unfreeze"
        ? "unfreeze_panes"
        : operation;
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel =
      normalized.operation === "delete_rows" || normalized.operation === "delete_columns"
        ? "destructive"
        : "standard";
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  }

  for (const key of Object.keys(value)) {
    if (!knownInputFields.has(key) && !hasOwn(normalized, key)) {
      normalized[key] = value[key];
    }
  }

  return normalized;
}

function normalizeRangeSortPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const knownInputFields = new Set([
    "targetSheet",
    "targetRange",
    "hasHeader",
    "keys",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "sheet",
    "sheetName",
    "range",
    "header",
    "includesHeader",
    "sortKeys"
  ]);
  const normalized = pickFields(value, [
    "targetSheet",
    "targetRange",
    "hasHeader",
    "keys",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges"
  ]);

  if (!hasOwn(normalized, "targetSheet")) {
    if (typeof value.sheet === "string" && value.sheet.trim()) {
      normalized.targetSheet = value.sheet.trim();
    } else if (typeof value.sheetName === "string" && value.sheetName.trim()) {
      normalized.targetSheet = value.sheetName.trim();
    }
  }

  if (!hasOwn(normalized, "targetRange")) {
    if (typeof value.range === "string" && value.range.trim()) {
      normalized.targetRange = value.range.trim();
    }
  }

  if (!hasOwn(normalized, "hasHeader")) {
    if (typeof value.header === "boolean") {
      normalized.hasHeader = value.header;
    } else if (typeof value.includesHeader === "boolean") {
      normalized.hasHeader = value.includesHeader;
    }
  }

  if (!hasOwn(normalized, "keys") && hasOwn(value, "sortKeys")) {
    normalized.keys = value.sortKeys;
  }

  if (hasOwn(value, "keys") && Array.isArray(value.keys)) {
    normalized.keys = value.keys;
  }

  if (hasOwn(normalized, "keys") && Array.isArray(normalized.keys)) {
    normalized.keys = normalized.keys.map((item) => {
      if (!isObject(item)) {
        return item;
      }

      const normalizedItem = pickFields(item, ["columnRef", "direction", "sortOn"]);
      if (
        (!normalizedItem.columnRef || typeof normalizedItem.columnRef !== "string") &&
        typeof item.field === "string" && item.field.trim()
      ) {
        normalizedItem.columnRef = item.field.trim();
      }
      if (
        (!normalizedItem.columnRef || typeof normalizedItem.columnRef !== "string") &&
        typeof item.column === "string" && item.column.trim()
      ) {
        normalizedItem.columnRef = item.column.trim();
      } else if (
        !normalizedItem.columnRef &&
        typeof item.column === "number" &&
        Number.isInteger(item.column)
      ) {
        normalizedItem.columnRef = item.column;
      }

      if (!hasOwn(normalizedItem, "direction")) {
        if (typeof item.order === "string") {
          normalizedItem.direction = item.order;
        } else if (typeof item.sortDirection === "string") {
          normalizedItem.direction = item.sortDirection;
        } else if (typeof item.ascending === "boolean") {
          normalizedItem.direction = item.ascending ? "asc" : "desc";
        }
      }

      if (typeof normalizedItem.direction === "string") {
        const direction = normalizedItem.direction.trim().toLowerCase();
        normalizedItem.direction =
          direction === "ascending" ? "asc"
          : direction === "descending" ? "desc"
          : direction;
      }

      return normalizedItem;
    });
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  for (const key of Object.keys(value)) {
    if (!knownInputFields.has(key) && !hasOwn(normalized, key)) {
      normalized[key] = value[key];
    }
  }

  return normalized;
}

function normalizeRangeFilterOperator(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }

  const normalized = value.trim();
  const operatorAliases: Record<string, string> = {
    equal_to: "equals",
    equals: "equals",
    not_equal_to: "notEquals",
    not_equals: "notEquals",
    notEquals: "notEquals",
    contains: "contains",
    starts_with: "startsWith",
    startsWith: "startsWith",
    ends_with: "endsWith",
    endsWith: "endsWith",
    greater_than: "greaterThan",
    greaterThan: "greaterThan",
    greater_than_or_equal_to: "greaterThanOrEqual",
    greaterThanOrEqual: "greaterThanOrEqual",
    less_than: "lessThan",
    lessThan: "lessThan",
    less_than_or_equal_to: "lessThanOrEqual",
    lessThanOrEqual: "lessThanOrEqual",
    is_empty: "isEmpty",
    isEmpty: "isEmpty",
    is_not_empty: "isNotEmpty",
    isNotEmpty: "isNotEmpty",
    top_n: "topN",
    topN: "topN"
  };

  return operatorAliases[normalized] ?? value;
}

function normalizeTopNFilterValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }

  const trimmed = value.trim();
  if (!/^\d+$/.test(trimmed)) {
    return value;
  }

  return Number(trimmed);
}

function normalizeRangeFilterPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized: JsonRecord = { ...value };

  if (hasOwn(value, "conditions") && Array.isArray(value.conditions)) {
    normalized.conditions = value.conditions.map((item) => {
      if (!isObject(item)) {
        return item;
      }

      const condition = pickFields(item, ["columnRef", "operator", "value", "value2"]);
      if (!hasOwn(condition, "columnRef")) {
        if (hasOwn(item, "field")) {
          condition.columnRef = item.field;
        } else if (hasOwn(item, "column")) {
          condition.columnRef = item.column;
        }
      }
      condition.operator = normalizeRangeFilterOperator(condition.operator);
      if (condition.operator === "topN" && hasOwn(condition, "value")) {
        condition.value = normalizeTopNFilterValue(condition.value);
      }
      return condition;
    });
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (!normalized.combiner || typeof normalized.combiner !== "string") {
    normalized.combiner = "and";
  }

  if (!hasOwn(normalized, "clearExistingFilters")) {
    normalized.clearExistingFilters = true;
  }

  if (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  return normalized;
}

function normalizeDataValidationPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "targetSheet",
    "targetRange",
    "ruleType",
    "values",
    "sourceRange",
    "namedRangeName",
    "showDropdown",
    "allowBlank",
    "invalidDataBehavior",
    "checkedValue",
    "uncheckedValue",
    "comparator",
    "value",
    "value2",
    "formula",
    "helpText",
    "inputTitle",
    "inputMessage",
    "errorTitle",
    "errorMessage",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "replacesExistingValidation"
  ]);

  if (hasOwn(value, "values") && Array.isArray(value.values)) {
    normalized.values = [...value.values];
  }

  if (!hasOwn(normalized, "values")) {
    if (hasOwn(value, "options") && Array.isArray(value.options)) {
      normalized.values = [...value.options];
    } else if (hasOwn(value, "allowedValues") && Array.isArray(value.allowedValues)) {
      normalized.values = [...value.allowedValues];
    } else if (hasOwn(value, "listValues") && Array.isArray(value.listValues)) {
      normalized.values = [...value.listValues];
    }
  }

  if (normalized.ruleType === "dropdown" || normalized.ruleType === "pick_list") {
    normalized.ruleType = "list";
  } else if (normalized.ruleType === "integer") {
    normalized.ruleType = "whole_number";
  } else if (normalized.ruleType === "number") {
    normalized.ruleType = "decimal";
  } else if (normalized.ruleType === "formula") {
    normalized.ruleType = "custom_formula";
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  if (!hasOwn(normalized, "allowBlank")) {
    normalized.allowBlank = true;
  }

  if (!hasOwn(normalized, "invalidDataBehavior")) {
    normalized.invalidDataBehavior = "reject";
  }

  if (normalized.ruleType === "list" && !hasOwn(normalized, "showDropdown")) {
    normalized.showDropdown = true;
  }

  if (hasOwn(normalized, "comparator")) {
    normalized.comparator = normalizeSnakeComparatorValue(normalized.comparator);
  }

  if (!hasOwn(normalized, "inputTitle") && hasOwn(value, "promptTitle")) {
    normalized.inputTitle = value.promptTitle;
  }

  if (!hasOwn(normalized, "inputMessage") && hasOwn(value, "promptMessage")) {
    normalized.inputMessage = value.promptMessage;
  }

  if (!hasOwn(normalized, "errorTitle") && hasOwn(value, "errorAlertTitle")) {
    normalized.errorTitle = value.errorAlertTitle;
  }

  if (!hasOwn(normalized, "errorMessage") && hasOwn(value, "errorAlertMessage")) {
    normalized.errorMessage = value.errorAlertMessage;
  }

  return normalized;
}

function normalizeNamedRangeUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "operation",
    "scope",
    "name",
    "sheetName",
    "targetSheet",
    "targetRange",
    "newName",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "confirmationLevel",
    "affectedRanges",
    "overwriteRisk"
  ]);

  if (typeof normalized.operation === "string") {
    const operation = normalized.operation.trim().toLowerCase();
    normalized.operation =
      operation === "define" || operation === "add"
        ? "create"
        : operation === "update" || operation === "change_range" || operation === "set_range"
        ? "retarget"
        : operation === "remove"
        ? "delete"
        : operation;
  }

  if (!hasOwn(normalized, "scope")) {
    normalized.scope = typeof normalized.sheetName === "string" && normalized.sheetName.trim()
      ? "sheet"
      : "workbook";
  }

  if (!hasOwn(normalized, "name")) {
    if (typeof value.rangeName === "string" && value.rangeName.trim()) {
      normalized.name = value.rangeName.trim();
    } else if (typeof value.namedRangeName === "string" && value.namedRangeName.trim()) {
      normalized.name = value.namedRangeName.trim();
    } else if (typeof value.namedRange === "string" && value.namedRange.trim()) {
      normalized.name = value.namedRange.trim();
    }
  }

  if (!hasOwn(normalized, "newName")) {
    if (typeof value.newRangeName === "string" && value.newRangeName.trim()) {
      normalized.newName = value.newRangeName.trim();
    } else if (typeof value.new_name === "string" && value.new_name.trim()) {
      normalized.newName = value.new_name.trim();
    }
  }

  const qualifiedRef =
    parseQualifiedRangeRef(value.refersTo) ??
    parseQualifiedRangeRef(value.range) ??
    parseQualifiedRangeRef(value.target);
  if (qualifiedRef) {
    if (!hasOwn(normalized, "targetSheet")) {
      normalized.targetSheet = qualifiedRef.sheet;
    }
    if (!hasOwn(normalized, "targetRange")) {
      normalized.targetRange = qualifiedRef.range;
    }
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel = normalized.operation === "delete" ? "destructive" : "standard";
  }

  return normalized;
}

function normalizeAnalysisReportSection(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, ["type", "title", "summary", "sourceRanges"]);

  if (hasOwn(value, "sourceRanges") && Array.isArray(value.sourceRanges)) {
    normalized.sourceRanges = [...value.sourceRanges];
  }

  return normalized;
}

function normalizeAnalysisReportPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceSheet",
    "sourceRange",
    "targetSheet",
    "targetRange",
    "outputMode",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel"
  ]);

  if (!hasOwn(normalized, "sourceRange")) {
    const dataRangeRef = parseQualifiedRangeRef(value.dataRange);
    const sourceAliasRef = parseQualifiedRangeRef(value.source);
    const rangeRef = parseQualifiedRangeRef(value.range);
    const sourceRef = dataRangeRef ?? sourceAliasRef ?? rangeRef;
    if (!hasOwn(normalized, "sourceSheet") && sourceRef?.sheet) {
      normalized.sourceSheet = sourceRef.sheet;
    }
    if (sourceRef?.range) {
      normalized.sourceRange = sourceRef.range;
    }
  }

  if (!hasOwn(normalized, "targetRange")) {
    const outputRangeRef = parseQualifiedRangeRef(value.outputRange);
    const destinationRef = parseQualifiedRangeRef(value.destination);
    const targetAliasRef = parseQualifiedRangeRef(value.target);
    const targetRef = outputRangeRef ?? destinationRef ?? targetAliasRef;
    if (!hasOwn(normalized, "targetSheet") && targetRef?.sheet) {
      normalized.targetSheet = targetRef.sheet;
    }
    if (targetRef?.range) {
      normalized.targetRange = targetRef.range;
    }
  }

  if (!hasOwn(normalized, "outputMode")) {
    const outputMode = typeof value.output === "string"
      ? value.output.trim().toLowerCase()
      : typeof value.mode === "string"
      ? value.mode.trim().toLowerCase()
      : undefined;
    normalized.outputMode =
      outputMode === "sheet" ||
      outputMode === "worksheet" ||
      outputMode === "materialized" ||
      outputMode === "materialize" ||
      outputMode === "materialize_report"
        ? "materialize_report"
        : outputMode === "chat" ||
          outputMode === "chat_only" ||
          outputMode === "inline" ||
          outputMode === "summary"
        ? "chat_only"
        : normalized.outputMode;
  }

  const sectionValues = hasOwn(value, "sections") && Array.isArray(value.sections)
    ? value.sections
    : hasOwn(value, "reportSections") && Array.isArray(value.reportSections)
    ? value.reportSections
    : undefined;
  if (sectionValues) {
    normalized.sections = sectionValues.map((item) =>
      normalizeAnalysisSectionLikeValue(item, normalized.sourceSheet, normalized.sourceRange)
    );
  }

  const originalTargetRange = normalized.targetRange;
  const expandedTargetRange = normalized.outputMode === "materialize_report"
    ? expandAnalysisReportAnchorRange(normalized.targetRange, normalized.sections)
    : undefined;
  if (expandedTargetRange) {
    normalized.targetRange = expandedTargetRange;
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    const originalTargetRef = buildQualifiedRangeRef(normalized.targetSheet, originalTargetRange);
    const expandedTargetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    normalized.affectedRanges = value.affectedRanges.map((range) =>
      range === originalTargetRef && expandedTargetRef ? expandedTargetRef : range
    );
  }

  if (!hasOwn(normalized, "requiresConfirmation") && normalized.outputMode === "materialize_report") {
    normalized.requiresConfirmation = true;
  } else if (!hasOwn(normalized, "requiresConfirmation") && normalized.outputMode === "chat_only") {
    normalized.requiresConfirmation = false;
  }

  if (
    (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0)
  ) {
    const refs = [
      buildQualifiedRangeRef(normalized.sourceSheet, normalized.sourceRange),
      normalized.outputMode === "materialize_report"
        ? buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange)
        : undefined
    ].filter((item): item is string => typeof item === "string" && item.length > 0);
    if (refs.length > 0) {
      normalized.affectedRanges = refs;
    }
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  } else if (normalized.outputMode === "chat_only") {
    normalized.overwriteRisk = "none";
  } else if (normalized.outputMode === "materialize_report") {
    normalized.overwriteRisk = "low";
  }

  if (
    (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") &&
    normalized.outputMode === "materialize_report"
  ) {
    normalized.confirmationLevel = "standard";
  }

  if (
    normalized.outputMode === "chat_only" &&
    (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string")
  ) {
    normalized.confirmationLevel = "standard";
  }

  return normalized;
}

function normalizePivotTablePlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceSheet",
    "sourceRange",
    "targetSheet",
    "targetRange",
    "rowGroups",
    "columnGroups",
    "valueAggregations",
    "filters",
    "sort",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel"
  ]);

  if (hasOwn(value, "rowGroups") && Array.isArray(value.rowGroups)) {
    normalized.rowGroups = [...value.rowGroups];
  }

  if (!hasOwn(normalized, "rowGroups") && hasOwn(value, "rows") && Array.isArray(value.rows)) {
    normalized.rowGroups = [...value.rows];
  }

  if (hasOwn(value, "columnGroups") && Array.isArray(value.columnGroups)) {
    normalized.columnGroups = [...value.columnGroups];
  }

  if (!hasOwn(normalized, "columnGroups") && hasOwn(value, "columns") && Array.isArray(value.columns)) {
    normalized.columnGroups = [...value.columns];
  }

  if (hasOwn(value, "valueAggregations") && Array.isArray(value.valueAggregations)) {
    normalized.valueAggregations = value.valueAggregations.map((item) => isObject(item) ? { ...item } : item);
  }

  if (!hasOwn(normalized, "valueAggregations") && hasOwn(value, "values") && Array.isArray(value.values)) {
    const defaultAggregation = typeof value.aggregation === "string" ? value.aggregation : undefined;
    normalized.valueAggregations = value.values.map((item) => {
      if (typeof item === "string") {
        return {
          field: item,
          aggregation: defaultAggregation
        };
      }

      if (isObject(item)) {
        return {
          ...item,
          aggregation: typeof item.aggregation === "string" ? item.aggregation : defaultAggregation
        };
      }

      return item;
    });
  }

  if (hasOwn(value, "filters") && Array.isArray(value.filters)) {
    normalized.filters = value.filters.map((item) => {
      if (!isObject(item)) {
        return item;
      }

      const normalizedFilter = { ...item };
      if (hasOwn(normalizedFilter, "operator")) {
        normalizedFilter.operator = normalizeSnakeComparatorValue(normalizedFilter.operator);
      }
      return normalizedFilter;
    });
  }

  if (hasOwn(value, "sort") && isObject(value.sort)) {
    normalized.sort = { ...value.sort };
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  }

  return normalized;
}

function normalizeChartPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceSheet",
    "sourceRange",
    "targetSheet",
    "targetRange",
    "chartType",
    "categoryField",
    "title",
    "legendPosition",
    "horizontalAxisTitle",
    "verticalAxisTitle",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel"
  ]);

  if (hasOwn(value, "series") && Array.isArray(value.series)) {
    normalized.series = value.series.map((item) => normalizeChartSeriesValue(item));
  }

  if (!hasOwn(normalized, "sourceRange") && hasOwn(value, "dataRange")) {
    const dataRange = parseQualifiedRangeRef(value.dataRange);
    if (!hasOwn(normalized, "sourceSheet") && dataRange?.sheet) {
      normalized.sourceSheet = dataRange.sheet;
    }
    if (dataRange?.range) {
      normalized.sourceRange = dataRange.range;
    }
  }

  if (!hasOwn(normalized, "targetRange") && hasOwn(value, "insertAt")) {
    const insertAt = parseQualifiedRangeRef(value.insertAt);
    if (!hasOwn(normalized, "targetSheet") && insertAt?.sheet) {
      normalized.targetSheet = insertAt.sheet;
    }
    if (insertAt?.range) {
      normalized.targetRange = insertAt.range;
    }
  }

  if (!hasOwn(normalized, "title") && hasOwn(value, "chartTitle")) {
    normalized.title = value.chartTitle;
  }

  if (!hasOwn(normalized, "horizontalAxisTitle") && hasOwn(value, "xAxisTitle")) {
    normalized.horizontalAxisTitle = value.xAxisTitle;
  }

  if (!hasOwn(normalized, "verticalAxisTitle") && hasOwn(value, "yAxisTitle")) {
    normalized.verticalAxisTitle = value.yAxisTitle;
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (hasOwn(normalized, "legendPosition") && normalized.legendPosition === "none") {
    normalized.legendPosition = "hidden";
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    const overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
    normalized.overwriteRisk =
      overwriteRisk === "none" ||
      overwriteRisk === "low" ||
      overwriteRisk === "medium" ||
      overwriteRisk === "high"
        ? overwriteRisk
        : "low";
  } else {
    normalized.overwriteRisk = "low";
  }

  if (
    (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0)
  ) {
    const refs = [
      buildQualifiedRangeRef(normalized.sourceSheet, normalized.sourceRange),
      buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange)
    ].filter((item): item is string => typeof item === "string" && item.length > 0);
    if (refs.length > 0) {
      normalized.affectedRanges = refs;
    }
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel = "standard";
  }

  return normalized;
}

function normalizeTablePlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "targetSheet",
    "targetRange",
    "name",
    "hasHeaders",
    "styleName",
    "showBandedRows",
    "showBandedColumns",
    "showFilterButton",
    "showTotalsRow",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel"
  ]);

  if (!hasOwn(normalized, "targetSheet")) {
    if (typeof value.sheet === "string" && value.sheet.trim()) {
      normalized.targetSheet = value.sheet.trim();
    } else if (typeof value.sheetName === "string" && value.sheetName.trim()) {
      normalized.targetSheet = value.sheetName.trim();
    }
  }

  if (!hasOwn(normalized, "targetRange") && hasOwn(value, "range")) {
    const rangeRef = parseQualifiedRangeRef(value.range);
    if (!hasOwn(normalized, "targetSheet") && rangeRef?.sheet) {
      normalized.targetSheet = rangeRef.sheet;
    }
    if (rangeRef?.range) {
      normalized.targetRange = rangeRef.range;
    }
  }

  if (!hasOwn(normalized, "name") && typeof value.tableName === "string" && value.tableName.trim()) {
    normalized.name = value.tableName.trim();
  }

  if (!hasOwn(normalized, "hasHeaders")) {
    if (typeof value.hasHeader === "boolean") {
      normalized.hasHeaders = value.hasHeader;
    } else if (typeof value.header === "boolean") {
      normalized.hasHeaders = value.header;
    }
  }

  if (!hasOwn(normalized, "styleName")) {
    if (typeof value.tableStyle === "string" && value.tableStyle.trim()) {
      normalized.styleName = value.tableStyle.trim();
    } else if (typeof value.style === "string" && value.style.trim()) {
      normalized.styleName = value.style.trim();
    }
  }

  if (!hasOwn(normalized, "showBandedRows") && typeof value.bandedRows === "boolean") {
    normalized.showBandedRows = value.bandedRows;
  }

  if (!hasOwn(normalized, "showBandedColumns") && typeof value.bandedColumns === "boolean") {
    normalized.showBandedColumns = value.bandedColumns;
  }

  if (!hasOwn(normalized, "showFilterButton")) {
    if (typeof value.filterButton === "boolean") {
      normalized.showFilterButton = value.filterButton;
    } else if (typeof value.filterButtons === "boolean") {
      normalized.showFilterButton = value.filterButtons;
    }
  }

  if (!hasOwn(normalized, "showTotalsRow")) {
    if (typeof value.totalsRow === "boolean") {
      normalized.showTotalsRow = value.totalsRow;
    } else if (typeof value.totalRow === "boolean") {
      normalized.showTotalsRow = value.totalRow;
    }
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    const overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
    normalized.overwriteRisk =
      overwriteRisk === "none" ||
      overwriteRisk === "low" ||
      overwriteRisk === "medium" ||
      overwriteRisk === "high"
        ? overwriteRisk
        : "low";
  } else {
    normalized.overwriteRisk = "low";
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel = "standard";
  }

  return normalized;
}

function normalizeExternalDataProvider(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  return ["googlefinance", "importhtml", "importxml", "importdata"].includes(normalized)
    ? normalized
    : undefined;
}

function normalizeExternalDataSourceType(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const normalized = value
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, "$1_$2")
    .replace(/[\s-]+/g, "_")
    .toLowerCase();

  if (normalized === "market" || normalized === "market_data") {
    return "market_data";
  }

  if (
    normalized === "web" ||
    normalized === "web_import" ||
    normalized === "web_table" ||
    normalized === "web_table_import"
  ) {
    return "web_table_import";
  }

  return undefined;
}

function normalizeExternalDataSelectorType(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  return ["table", "list", "xpath", "direct"].includes(normalized)
    ? normalized
    : undefined;
}

function quoteSpreadsheetFormulaString(value: string): string {
  return `"${value.replace(/"/g, "\"\"")}"`;
}

function buildGoogleFinanceFormula(query: unknown): string | undefined {
  if (!isObject(query) || typeof query.symbol !== "string" || !query.symbol.trim()) {
    return undefined;
  }

  const args = [
    query.symbol,
    query.attribute,
    query.startDate,
    query.endDate,
    query.interval
  ]
    .filter((item): item is string => typeof item === "string" && item.trim().length > 0)
    .map((item) => quoteSpreadsheetFormulaString(item.trim()));

  return `=GOOGLEFINANCE(${args.join(",")})`;
}

function buildWebImportFormula(
  provider: string,
  sourceUrl: unknown,
  selectorType: unknown,
  selector: unknown
): string | undefined {
  if (typeof sourceUrl !== "string" || !sourceUrl.trim()) {
    return undefined;
  }

  const quotedUrl = quoteSpreadsheetFormulaString(sourceUrl.trim());
  if (provider === "importdata") {
    return `=IMPORTDATA(${quotedUrl})`;
  }

  if (provider === "importhtml") {
    if (
      typeof selectorType !== "string" ||
      (selectorType !== "table" && selectorType !== "list") ||
      !(typeof selector === "number" || typeof selector === "string")
    ) {
      return undefined;
    }

    return `=IMPORTHTML(${quotedUrl},${quoteSpreadsheetFormulaString(selectorType)},${selector})`;
  }

  if (provider === "importxml") {
    if (typeof selector !== "string" || !selector.trim()) {
      return undefined;
    }

    return `=IMPORTXML(${quotedUrl},${quoteSpreadsheetFormulaString(selector.trim())})`;
  }

  return undefined;
}

function normalizeExternalDataPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceType",
    "provider",
    "targetSheet",
    "targetRange",
    "formula",
    "sourceUrl",
    "selectorType",
    "selector",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel"
  ]);

  if (hasOwn(value, "query") && isObject(value.query)) {
    normalized.query = pickFields(value.query, [
      "symbol",
      "attribute",
      "startDate",
      "endDate",
      "interval"
    ]);
  }

  const provider = normalizeExternalDataProvider(normalized.provider);
  if (provider) {
    normalized.provider = provider;
  }

  const sourceType = normalizeExternalDataSourceType(normalized.sourceType);
  if (sourceType) {
    normalized.sourceType = sourceType;
  }

  const selectorType = normalizeExternalDataSelectorType(normalized.selectorType);
  if (selectorType) {
    normalized.selectorType = selectorType;
  }

  if (!normalized.sourceType && provider === "googlefinance") {
    normalized.sourceType = "market_data";
  } else if (
    !normalized.sourceType &&
    (provider === "importhtml" || provider === "importxml" || provider === "importdata")
  ) {
    normalized.sourceType = "web_table_import";
  }

  if (!hasOwn(normalized, "selectorType")) {
    if (provider === "importxml") {
      normalized.selectorType = "xpath";
    } else if (provider === "importdata") {
      normalized.selectorType = "direct";
    }
  }

  if (provider === "importhtml" && typeof normalized.selector === "string") {
    const selectorIndex = normalized.selector.trim();
    if (/^[1-9]\d*$/.test(selectorIndex)) {
      normalized.selector = Number(selectorIndex);
    }
  }

  if (!hasOwn(normalized, "formula") && provider === "googlefinance") {
    const formula = buildGoogleFinanceFormula(normalized.query);
    if (formula) {
      normalized.formula = formula;
    }
  } else if (
    !hasOwn(normalized, "formula") &&
    (provider === "importhtml" || provider === "importxml" || provider === "importdata")
  ) {
    const formula = buildWebImportFormula(provider, normalized.sourceUrl, normalized.selectorType, normalized.selector);
    if (formula) {
      normalized.formula = formula;
    }
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  } else {
    normalized.overwriteRisk = "low";
  }

  if (
    (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0) &&
    typeof normalized.targetSheet === "string" &&
    typeof normalized.targetRange === "string"
  ) {
    normalized.affectedRanges = [`${normalized.targetSheet}!${normalized.targetRange}`];
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel = "standard";
  }

  return normalized;
}

function normalizeSheetUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "targetSheet",
    "targetRange",
    "operation",
    "values",
    "formulas",
    "notes",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "overwriteRisk"
  ]);

  if (hasOwn(value, "shape")) {
    normalized.shape = normalizeShapeValue(value.shape);
  }

  if (!hasOwn(normalized, "targetSheet")) {
    if (typeof value.sheet === "string" && value.sheet.trim()) {
      normalized.targetSheet = value.sheet.trim();
    } else if (typeof value.sheetName === "string" && value.sheetName.trim()) {
      normalized.targetSheet = value.sheetName.trim();
    }
  }

  if (!hasOwn(normalized, "targetRange") && hasOwn(value, "range")) {
    const rangeRef = parseQualifiedRangeRef(value.range);
    if (!hasOwn(normalized, "targetSheet") && rangeRef?.sheet) {
      normalized.targetSheet = rangeRef.sheet;
    }
    if (rangeRef?.range) {
      normalized.targetRange = rangeRef.range;
    }
  }

  if (!hasOwn(normalized, "operation")) {
    if (typeof value.action === "string" && value.action.trim()) {
      normalized.operation = value.action.trim();
    } else if (typeof value.op === "string" && value.op.trim()) {
      normalized.operation = value.op.trim();
    }
  }

  if (!hasOwn(normalized, "values") && Array.isArray(value.data)) {
    normalized.values = value.data;
  }

  if (normalized.operation === "set_values") {
    normalized.operation = "replace_range";
  } else if (
    normalized.operation === "set_formulas" &&
    (hasOwn(normalized, "values") || hasOwn(normalized, "notes"))
  ) {
    normalized.operation = "mixed_update";
  } else if (
    normalized.operation === "set_notes" &&
    (hasOwn(normalized, "values") || hasOwn(normalized, "formulas"))
  ) {
    normalized.operation = "mixed_update";
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  }

  if (!hasOwn(normalized, "shape") && hasOwn(value, "data")) {
    const inferredShape =
      inferShapeFromMatrix(normalized.values) ??
      inferShapeFromMatrix(normalized.formulas) ??
      inferShapeFromMatrix(normalized.notes);
    if (inferredShape) {
      normalized.shape = inferredShape;
    }
  }

  return normalized;
}

function normalizeRangeTransferPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceSheet",
    "sourceRange",
    "targetSheet",
    "targetRange",
    "operation",
    "pasteMode",
    "transpose",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel"
  ]);

  if (!hasOwn(normalized, "operation") && hasOwn(value, "transferOperation")) {
    normalized.operation = value.transferOperation;
  }

  if (!hasOwn(normalized, "transpose")) {
    normalized.transpose = false;
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  } else if (normalized.operation === "move") {
    normalized.overwriteRisk = "high";
  } else if (normalized.operation === "append") {
    normalized.overwriteRisk = "medium";
  } else if (normalized.operation === "copy") {
    normalized.overwriteRisk = "low";
  }

  if (
    (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0)
  ) {
    const refs = [
      buildQualifiedRangeRef(normalized.sourceSheet, normalized.sourceRange),
      buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange)
    ].filter((item): item is string => typeof item === "string" && item.length > 0);
    if (refs.length > 0) {
      normalized.affectedRanges = refs;
    }
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    normalized.confirmationLevel = normalized.operation === "move" ? "destructive" : "standard";
  }

  return normalized;
}

function normalizeDataCleanupPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "targetSheet",
    "targetRange",
    "operation",
    "keyColumns",
    "mode",
    "sourceColumn",
    "delimiter",
    "targetStartColumn",
    "sourceColumns",
    "targetColumn",
    "columns",
    "formatType",
    "formatPattern",
    "explanation",
    "confidence",
    "requiresConfirmation",
    "affectedRanges",
    "overwriteRisk",
    "confirmationLevel"
  ]);

  if (!hasOwn(normalized, "targetSheet")) {
    if (typeof value.sheet === "string" && value.sheet.trim()) {
      normalized.targetSheet = value.sheet.trim();
    } else if (typeof value.sheetName === "string" && value.sheetName.trim()) {
      normalized.targetSheet = value.sheetName.trim();
    }
  }

  if (!hasOwn(normalized, "targetRange") && hasOwn(value, "range")) {
    const rangeRef = parseQualifiedRangeRef(value.range);
    if (!hasOwn(normalized, "targetSheet") && rangeRef?.sheet) {
      normalized.targetSheet = rangeRef.sheet;
    }
    if (rangeRef?.range) {
      normalized.targetRange = rangeRef.range;
    }
  }

  if (!hasOwn(normalized, "operation")) {
    if (typeof value.action === "string" && value.action.trim()) {
      normalized.operation = value.action.trim();
    } else if (typeof value.cleanupOperation === "string" && value.cleanupOperation.trim()) {
      normalized.operation = value.cleanupOperation.trim();
    } else if (typeof value.op === "string" && value.op.trim()) {
      normalized.operation = value.op.trim();
    }
  }

  if (!hasOwn(normalized, "sourceColumn")) {
    if (typeof value.column === "string" && value.column.trim()) {
      normalized.sourceColumn = value.column.trim();
    } else if (typeof value.source === "string" && value.source.trim()) {
      normalized.sourceColumn = value.source.trim();
    }
  }

  if (!hasOwn(normalized, "delimiter") && typeof value.separator === "string") {
    normalized.delimiter = value.separator;
  }

  if (!hasOwn(normalized, "targetStartColumn")) {
    if (typeof value.startColumn === "string" && value.startColumn.trim()) {
      normalized.targetStartColumn = value.startColumn.trim();
    } else if (typeof value.outputStartColumn === "string" && value.outputStartColumn.trim()) {
      normalized.targetStartColumn = value.outputStartColumn.trim();
    }
  }

  if (hasOwn(value, "keyColumns") && Array.isArray(value.keyColumns)) {
    normalized.keyColumns = [...value.keyColumns];
  }

  if (hasOwn(value, "sourceColumns") && Array.isArray(value.sourceColumns)) {
    normalized.sourceColumns = [...value.sourceColumns];
  }

  if (hasOwn(value, "columns") && Array.isArray(value.columns)) {
    normalized.columns = [...value.columns];
  }

  if (normalized.operation === "trim") {
    normalized.operation = "trim_whitespace";
  } else if (normalized.operation === "remove_duplicates") {
    normalized.operation = "remove_duplicate_rows";
  } else if (normalized.operation === "standardize_case") {
    normalized.operation = "normalize_case";
  } else if (normalized.operation === "split") {
    normalized.operation = "split_column";
  } else if (normalized.operation === "join") {
    normalized.operation = "join_columns";
  } else if (normalized.operation === "dedupe" || normalized.operation === "deduplicate") {
    normalized.operation = "remove_duplicate_rows";
  }

  if (typeof normalized.formatType === "string") {
    const formatType = normalized.formatType.trim().toLowerCase();
    normalized.formatType =
      formatType.includes("currency") || formatType.includes("number") || formatType.includes("numeric")
        ? "number_text"
        : formatType.includes("date")
        ? "date_text"
        : formatType;
  }

  if (
    normalized.operation === "normalize_case" &&
    (!normalized.mode || typeof normalized.mode !== "string")
  ) {
    normalized.mode = "title";
  }

  if (hasOwn(value, "affectedRanges") && Array.isArray(value.affectedRanges)) {
    normalized.affectedRanges = [...value.affectedRanges];
  }

  if (hasOwn(normalized, "overwriteRisk")) {
    normalized.overwriteRisk = normalizeOverwriteRiskValue(normalized.overwriteRisk);
  } else {
    const inferredOverwriteRisk = inferCleanupOverwriteRisk(normalized.operation);
    if (inferredOverwriteRisk) {
      normalized.overwriteRisk = inferredOverwriteRisk;
    }
  }

  if (
    (!Array.isArray(normalized.affectedRanges) || normalized.affectedRanges.length === 0)
  ) {
    const targetRef = buildQualifiedRangeRef(normalized.targetSheet, normalized.targetRange);
    if (targetRef) {
      normalized.affectedRanges = [targetRef];
    }
  }

  if (!normalized.confirmationLevel || typeof normalized.confirmationLevel !== "string") {
    const inferredConfirmationLevel = inferCleanupConfirmationLevel(normalized.operation);
    if (inferredConfirmationLevel) {
      normalized.confirmationLevel = inferredConfirmationLevel;
    }
  }

  if (normalized.operation === "standardize_format") {
    const inferredFormat = inferStandardizeFormatDetails(normalized.explanation);
    if (
      (!normalized.formatType || typeof normalized.formatType !== "string") &&
      inferredFormat.formatType
    ) {
      normalized.formatType = inferredFormat.formatType;
    }
    if (
      (!normalized.formatPattern || typeof normalized.formatPattern !== "string") &&
      inferredFormat.formatPattern
    ) {
      normalized.formatPattern = inferredFormat.formatPattern;
    }
  }

  return normalized;
}

function normalizeAnalysisReportUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  return pickFields(value, ["operation", "targetSheet", "targetRange", "summary"]);
}

function normalizePivotTableUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  return pickFields(value, ["operation", "targetSheet", "targetRange", "summary"]);
}

function normalizeChartUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "operation",
    "targetSheet",
    "targetRange",
    "chartType",
    "horizontalAxisTitle",
    "verticalAxisTitle",
    "summary"
  ]);

  if (!hasOwn(normalized, "horizontalAxisTitle") && hasOwn(value, "xAxisTitle")) {
    normalized.horizontalAxisTitle = value.xAxisTitle;
  }

  if (!hasOwn(normalized, "verticalAxisTitle") && hasOwn(value, "yAxisTitle")) {
    normalized.verticalAxisTitle = value.yAxisTitle;
  }

  return normalized;
}

function normalizeTableUpdateData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  return pickFields(value, [
    "operation",
    "targetSheet",
    "targetRange",
    "name",
    "hasHeaders",
    "styleName",
    "showBandedRows",
    "showBandedColumns",
    "showFilterButton",
    "showTotalsRow",
    "summary"
  ]);
}

function getSheetImportRowsMatrix(value: JsonRecord): unknown[][] | undefined {
  const rowsCandidate = hasOwn(value, "rows")
    ? value.rows
    : hasOwn(value, "data")
    ? value.data
    : value.table;

  if (
    !Array.isArray(rowsCandidate) ||
    rowsCandidate.length === 0 ||
    !rowsCandidate.every((row) => Array.isArray(row))
  ) {
    return undefined;
  }

  return rowsCandidate as unknown[][];
}

function normalizeExtractionModeValue(value: unknown): unknown {
  if (typeof value !== "string") {
    return value;
  }

  const normalized = value.trim().toLowerCase();
  return normalized === "real" || normalized === "demo" || normalized === "unavailable"
    ? normalized
    : value;
}

function normalizeSheetImportPlanData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceAttachmentId",
    "targetSheet",
    "targetRange",
    "headers",
    "values",
    "confidence",
    "requiresConfirmation",
    "extractionMode"
  ]);

  const rowsMatrix = getSheetImportRowsMatrix(value);

  if (!hasOwn(normalized, "sourceAttachmentId")) {
    if (hasOwn(value, "attachmentId")) {
      normalized.sourceAttachmentId = value.attachmentId;
    } else if (hasOwn(value, "sourceId")) {
      normalized.sourceAttachmentId = value.sourceId;
    }
  }

  if (!hasOwn(normalized, "targetSheet")) {
    if (hasOwn(value, "sheet")) {
      normalized.targetSheet = value.sheet;
    } else if (hasOwn(value, "sheetName")) {
      normalized.targetSheet = value.sheetName;
    }
  }

  if (!hasOwn(normalized, "targetRange")) {
    const rangeValue = hasOwn(value, "range") ? value.range : value.target;
    const rangeRef = parseQualifiedRangeRef(rangeValue);
    if (rangeRef) {
      if (!hasOwn(normalized, "targetSheet") && rangeRef.sheet) {
        normalized.targetSheet = rangeRef.sheet;
      }
      if (rangeRef.range) {
        normalized.targetRange = rangeRef.range;
      }
    }
  }

  if (!hasOwn(normalized, "headers") && rowsMatrix) {
    normalized.headers = rowsMatrix[0].map((cell) => String(cell ?? ""));
  }

  if (!hasOwn(normalized, "values") && rowsMatrix) {
    normalized.values = rowsMatrix.slice(1);
  }

  if (!hasOwn(normalized, "extractionMode")) {
    if (hasOwn(value, "mode")) {
      normalized.extractionMode = normalizeExtractionModeValue(value.mode);
    } else if (hasOwn(value, "extraction")) {
      normalized.extractionMode = normalizeExtractionModeValue(value.extraction);
    }
  } else {
    normalized.extractionMode = normalizeExtractionModeValue(normalized.extractionMode);
  }

  if (hasOwn(value, "warnings")) {
    normalized.warnings = normalizeWarningsValue(value.warnings);
  }

  if (hasOwn(value, "shape")) {
    normalized.shape = normalizeShapeValue(value.shape);
  } else if (Array.isArray(normalized.headers) && Array.isArray(normalized.values)) {
    normalized.shape = {
      rows: 1 + normalized.values.length,
      columns: normalized.headers.length
    };
  }

  return normalized;
}

function normalizeErrorData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, ["code", "message", "retryable", "userAction"]);

  if (normalized.code === "MISSING_REQUIRED_CONTEXT") {
    normalized.code = "SPREADSHEET_CONTEXT_MISSING";
  }

  return normalized;
}

function normalizeAttachmentAnalysisData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceAttachmentId",
    "contentKind",
    "summary",
    "confidence",
    "extractionMode"
  ]);

  if (hasOwn(value, "warnings")) {
    normalized.warnings = normalizeWarningsValue(value.warnings);
  }

  return normalized;
}

function normalizeExtractedTableData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceAttachmentId",
    "headers",
    "rows",
    "confidence",
    "extractionMode"
  ]);

  if (hasOwn(value, "warnings")) {
    normalized.warnings = normalizeWarningsValue(value.warnings);
  }

  if (hasOwn(value, "shape")) {
    normalized.shape = normalizeShapeValue(value.shape);
  }

  return normalized;
}

function normalizeDocumentSummaryData(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = pickFields(value, [
    "sourceAttachmentId",
    "summary",
    "contentKind",
    "keyPoints",
    "confidence",
    "extractionMode"
  ]);

  if (hasOwn(value, "warnings")) {
    normalized.warnings = normalizeWarningsValue(value.warnings);
  }

  return normalized;
}

function normalizeDataByType(type: HermesStructuredBodyType, value: unknown): unknown {
  switch (type) {
    case "chat":
      return normalizeChatData(value);
    case "formula":
      return normalizeFormulaData(value);
    case "composite_plan":
      return normalizeCompositePlanData(value);
    case "workbook_structure_update":
      return normalizeWorkbookStructureUpdateData(value);
    case "range_format_update":
      return normalizeRangeFormatUpdateData(value);
    case "conditional_format_plan":
      return normalizeConditionalFormatPlanData(value);
    case "sheet_structure_update":
      return normalizeSheetStructureUpdateData(value);
    case "range_sort_plan":
      return normalizeRangeSortPlanData(value);
    case "range_filter_plan":
      return normalizeRangeFilterPlanData(value);
    case "data_validation_plan":
      return normalizeDataValidationPlanData(value);
    case "analysis_report_plan":
      return normalizeAnalysisReportPlanData(value);
    case "pivot_table_plan":
      return normalizePivotTablePlanData(value);
    case "chart_plan":
      return normalizeChartPlanData(value);
    case "table_plan":
      return normalizeTablePlanData(value);
    case "external_data_plan":
      return normalizeExternalDataPlanData(value);
    case "named_range_update":
      return normalizeNamedRangeUpdateData(value);
    case "range_transfer_plan":
      return normalizeRangeTransferPlanData(value);
    case "data_cleanup_plan":
      return normalizeDataCleanupPlanData(value);
    case "analysis_report_update":
      return normalizeAnalysisReportUpdateData(value);
    case "pivot_table_update":
      return normalizePivotTableUpdateData(value);
    case "chart_update":
      return normalizeChartUpdateData(value);
    case "table_update":
      return normalizeTableUpdateData(value);
    case "sheet_update":
      return normalizeSheetUpdateData(value);
    case "sheet_import_plan":
      return normalizeSheetImportPlanData(value);
    case "error":
      return normalizeErrorData(value);
    case "attachment_analysis":
      return normalizeAttachmentAnalysisData(value);
    case "extracted_table":
      return normalizeExtractedTableData(value);
    case "document_summary":
      return normalizeDocumentSummaryData(value);
  }
}

export function normalizeHermesStructuredBodyInput(value: unknown): unknown {
  if (!isObject(value) || typeof value.type !== "string") {
    return value;
  }

  if (!STRUCTURED_BODY_TYPES.includes(value.type as HermesStructuredBodyType)) {
    return value;
  }

  const type = value.type as HermesStructuredBodyType;
  const normalized: JsonRecord = {
    type
  };

  if (hasOwn(value, "data")) {
    normalized.data = normalizeDataByType(type, value.data);
  }

  if (hasOwn(value, "warnings")) {
    normalized.warnings = normalizeWarningsValue(value.warnings);
  }

  if (hasOwn(value, "skillsUsed")) {
    normalized.skillsUsed = value.skillsUsed;
  }

  if (hasOwn(value, "downstreamProvider")) {
    normalized.downstreamProvider = normalizeDownstreamProviderValue(value.downstreamProvider);
  }

  return normalized;
}

const optionalWarningsSchema = z.preprocess(
  normalizeWarningsValue,
  z.array(WarningSchema).optional()
);
const optionalSkillsUsedSchema = z.array(z.string().min(1).max(128)).optional();
const optionalDownstreamProviderSchema = DownstreamProviderSchema.optional();

function createStructuredBodySchema<
  TypeName extends string,
  DataSchema extends z.ZodTypeAny
>(typeName: TypeName, dataSchema: DataSchema) {
  return z.object({
    type: z.literal(typeName),
    data: dataSchema,
    warnings: optionalWarningsSchema,
    skillsUsed: optionalSkillsUsedSchema,
    downstreamProvider: optionalDownstreamProviderSchema
  }).strict();
}

export const HermesStructuredBodySchema = z.discriminatedUnion("type", [
  createStructuredBodySchema("chat", ChatDataSchema),
  createStructuredBodySchema("formula", FormulaDataSchema),
  createStructuredBodySchema("composite_plan", CompositePlanDataSchema),
  createStructuredBodySchema("workbook_structure_update", WorkbookStructureUpdateDataSchema),
  createStructuredBodySchema("range_format_update", RangeFormatUpdateDataSchema),
  createStructuredBodySchema("conditional_format_plan", ConditionalFormatPlanDataSchema),
  createStructuredBodySchema("sheet_structure_update", SheetStructureUpdateDataSchema),
  createStructuredBodySchema("range_sort_plan", RangeSortPlanDataSchema),
  createStructuredBodySchema("range_filter_plan", RangeFilterPlanDataSchema),
  createStructuredBodySchema("data_validation_plan", DataValidationPlanDataSchema),
  createStructuredBodySchema("analysis_report_plan", AnalysisReportPlanDataSchema),
  createStructuredBodySchema("pivot_table_plan", PivotTablePlanDataSchema),
  createStructuredBodySchema("chart_plan", ChartPlanDataSchema),
  createStructuredBodySchema("table_plan", TablePlanDataSchema),
  createStructuredBodySchema("external_data_plan", ExternalDataPlanDataSchema),
  createStructuredBodySchema("named_range_update", NamedRangeUpdateDataSchema),
  createStructuredBodySchema("range_transfer_plan", RangeTransferPlanDataSchema),
  createStructuredBodySchema("data_cleanup_plan", DataCleanupPlanDataSchema),
  createStructuredBodySchema("analysis_report_update", AnalysisReportUpdateDataSchema),
  createStructuredBodySchema("pivot_table_update", PivotTableUpdateDataSchema),
  createStructuredBodySchema("chart_update", ChartUpdateDataSchema),
  createStructuredBodySchema("table_update", TableUpdateDataSchema),
  createStructuredBodySchema("sheet_update", SheetUpdateDataSchema),
  createStructuredBodySchema("sheet_import_plan", SheetImportPlanDataSchema),
  createStructuredBodySchema("error", ErrorDataSchema),
  createStructuredBodySchema("attachment_analysis", AttachmentAnalysisDataSchema),
  createStructuredBodySchema("extracted_table", ExtractedTableDataSchema),
  createStructuredBodySchema("document_summary", DocumentSummaryDataSchema)
]);

export type HermesStructuredBody = z.infer<typeof HermesStructuredBodySchema>;
