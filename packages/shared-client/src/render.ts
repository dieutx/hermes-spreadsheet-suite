import { matrixShape } from "@hermes/contracts";
import type {
  AttachmentAnalysisResponse,
  AnalysisReportPlanData,
  AnalysisReportPlanResponse,
  AnalysisReportUpdateData,
  AnalysisReportUpdateResponse,
  ChartPlanData,
  ChartPlanResponse,
  ChartUpdateData,
  ChartUpdateResponse,
  CompositePlanData,
  CompositePlanResponse,
  CompositeUpdateData,
  ConditionalFormatPlanResponse,
  ConditionalFormatPlanData,
  ConditionalFormatUpdateData,
  DataValidationPlanData,
  DataValidationPlanResponse,
  DataCleanupPlanData,
  DataCleanupPlanResponse,
  DataCleanupUpdateData,
  DocumentSummaryResponse,
  ExtractedTableResponse,
  FormulaResponse,
  DryRunResult,
  PlanHistoryEntry,
  PlanHistoryPage,
  HermesResponse,
  RangeFormatUpdateData,
  RangeFormatUpdateResponse,
  RangeFilterPlanData,
  RangeFilterPlanResponse,
  RangeSortPlanData,
  RangeSortPlanResponse,
  NamedRangeUpdateData,
  NamedRangeUpdateResponse,
  RangeTransferPlanData,
  RangeTransferPlanResponse,
  RangeTransferUpdateData,
  PivotTablePlanData,
  PivotTablePlanResponse,
  PivotTableUpdateData,
  PivotTableUpdateResponse,
  SheetImportPlanData,
  SheetImportPlanResponse,
  SheetStructureUpdateData,
  SheetStructureUpdateResponse,
  SheetUpdateData,
  SheetUpdateResponse,
  WorkbookStructurePosition,
  WorkbookStructureUpdateData,
  WorkbookStructureUpdateResponse,
  Warning
} from "@hermes/contracts";
type RenderCell = string | number | boolean | null;
type MatrixWritePlan = SheetImportPlanData | SheetUpdateData;
type AnalysisReportWritePlanData = Extract<
  AnalysisReportPlanData,
  { outputMode: "materialize_report" }
>;

export type PreviewTable = {
  headers: string[];
  rows: RenderCell[][];
  shape: {
    rows: number;
    columns: number;
  };
};

export type StructuredPreview =
  | {
      kind: "composite_plan";
      stepCount: number;
      destructiveConfirmationRequired: boolean;
      steps: Array<{
        stepId: string;
        dependsOn: string[];
        continueOnError: boolean;
        destructive: boolean;
        reversible: boolean;
        skippedIfDependenciesFail: boolean;
        summary: string;
      }>;
      explanation: string;
      confidence: number;
      reversible: boolean;
      dryRunRecommended: boolean;
      dryRunRequired: boolean;
      affectedRanges: string[];
      overwriteRisk: CompositePlanData["overwriteRisk"];
      confirmationLevel: CompositePlanData["confirmationLevel"];
      summary: string;
      details: string[];
    }
  | {
      kind: "composite_update";
      executionId: string;
      stepResults: CompositeUpdateData["stepResults"];
      summary: string;
    }
  | {
      kind: "dry_run_result";
      simulated: boolean;
      reversible: boolean;
      overwriteRisk: DryRunResult["overwriteRisk"];
      expiresAt: string;
      predictedAffectedRanges: string[];
      predictedSummaries: string[];
      unsupportedReason?: string;
      steps: NonNullable<DryRunResult["steps"]>;
      summary: string;
      details: string[];
    }
  | {
      kind: "plan_history_page";
      entries: PlanHistoryEntry[];
      summary: string;
      details: string[];
      nextCursor?: string;
    }
  | {
      kind: "formula";
      intent: FormulaResponse["data"]["intent"];
      formula: string;
      formulaLanguage: FormulaResponse["data"]["formulaLanguage"];
      targetCell?: string;
      explanation: string;
      alternateFormulas: FormulaResponse["data"]["alternateFormulas"];
    }
  | {
      kind: "workbook_structure_update";
      operation: WorkbookStructureUpdateResponse["data"]["operation"];
      sheetName: string;
      position?: WorkbookStructurePosition;
      newSheetName?: string;
      explanation: string;
      overwriteRisk?: WorkbookStructureUpdateResponse["data"]["overwriteRisk"];
    }
  | {
      kind: "sheet_structure_update";
      targetSheet: string;
      operation: SheetStructureUpdateResponse["data"]["operation"];
      summary: string;
      confirmationLevel: SheetStructureUpdateData["confirmationLevel"];
    }
  | {
      kind: "analysis_report_plan";
    } & AnalysisReportPlanData & {
      summary: string;
      details: string[];
    }
  | {
      kind: "pivot_table_plan";
    } & PivotTablePlanData & {
      summary: string;
      details: string[];
    }
  | {
      kind: "chart_plan";
    } & ChartPlanData & {
      summary: string;
      details: string[];
    }
  | {
      kind: "range_format_update";
      targetSheet: string;
      targetRange: string;
      format: RangeFormatUpdateResponse["data"]["format"];
      explanation: string;
      overwriteRisk?: RangeFormatUpdateResponse["data"]["overwriteRisk"];
    }
  | ({
      kind: "conditional_format_plan";
    } & ConditionalFormatPlanData & {
      summary: string;
      details: string[];
    })
  | {
      kind: "range_sort_plan";
      targetSheet: string;
      targetRange: string;
      hasHeader: boolean;
      keys: RangeSortPlanResponse["data"]["keys"];
      summary: string;
    }
  | {
      kind: "range_filter_plan";
      targetSheet: string;
      targetRange: string;
      hasHeader: boolean;
      conditions: RangeFilterPlanResponse["data"]["conditions"];
      combiner: RangeFilterPlanResponse["data"]["combiner"];
      clearExistingFilters: boolean;
      summary: string;
    }
  | {
      kind: "range_transfer_plan";
    } & RangeTransferPlanData & {
      summary: string;
      details: string[];
    }
  | {
      kind: "data_cleanup_plan";
    } & DataCleanupPlanData & {
      summary: string;
      details: string[];
    }
  | {
      kind: "data_validation_plan";
    } & DataValidationPlanData
  | {
      kind: "named_range_update";
    } & NamedRangeUpdateData
  | {
      kind: "range_transfer_update";
    } & RangeTransferUpdateData
  | {
      kind: "data_cleanup_update";
    } & DataCleanupUpdateData
  | {
      kind: "analysis_report_update";
    } & AnalysisReportUpdateData
  | {
      kind: "pivot_table_update";
    } & PivotTableUpdateData
  | {
      kind: "chart_update";
    } & ChartUpdateData
  | {
    kind: "conditional_format_update";
      operation: "conditional_format_update";
      targetSheet: string;
      targetRange: string;
      managementMode: ConditionalFormatUpdateData["managementMode"];
      summary: string;
    }
  | {
      kind: "sheet_update";
      targetSheet: string;
      targetRange: string;
      shape: SheetUpdateResponse["data"]["shape"];
      operation: SheetUpdateResponse["data"]["operation"];
      overwriteRisk?: SheetUpdateResponse["data"]["overwriteRisk"];
      matrixKind: "values" | "formulas" | "notes" | "mixed_update";
      table: PreviewTable;
    }
  | {
      kind: "sheet_import_plan";
      targetSheet: string;
      targetRange: string;
      shape: SheetImportPlanResponse["data"]["shape"];
      extractionMode: SheetImportPlanResponse["data"]["extractionMode"];
      sourceAttachmentId: string;
      table: PreviewTable;
    }
  | {
      kind: "attachment_analysis";
      sourceAttachmentId: string;
      contentKind: AttachmentAnalysisResponse["data"]["contentKind"];
      extractionMode: AttachmentAnalysisResponse["data"]["extractionMode"];
      summary: string;
    }
  | {
      kind: "extracted_table";
      sourceAttachmentId: string;
      extractionMode: ExtractedTableResponse["data"]["extractionMode"];
      table: PreviewTable;
    }
  | {
      kind: "document_summary";
      sourceAttachmentId: string;
      contentKind: DocumentSummaryResponse["data"]["contentKind"];
      extractionMode: DocumentSummaryResponse["data"]["extractionMode"];
      keyPoints: string[];
    };

type WarningCarrier = {
  warnings?: Warning[];
};

function hasOwnWarnings(value: unknown): value is WarningCarrier {
  return typeof value === "object" && value !== null && "warnings" in value;
}

function formatConditionalFormatValue(value: unknown): string {
  if (typeof value === "string") {
    return JSON.stringify(value);
  }

  return String(value);
}

function formatConditionalFormatStyle(plan: ConditionalFormatPlanData): string {
  if (!("ruleType" in plan)) {
    return "Style: not applicable.";
  }

  if (plan.ruleType === "color_scale") {
    const points = plan.points
      .map((point) => {
        const pointValue = point.value === undefined
          ? point.type
          : `${point.type} ${formatConditionalFormatValue(point.value)}`;
        return `${pointValue} ${point.color}`;
      })
      .join(", ");

    return `Style: color scale points ${points}.`;
  }

  const styleParts: string[] = [];
  if (plan.style.backgroundColor) {
    styleParts.push(`background ${plan.style.backgroundColor}`);
  }
  if (plan.style.textColor) {
    styleParts.push(`text ${plan.style.textColor}`);
  }
  if (plan.style.bold) {
    styleParts.push("bold");
  }
  if (plan.style.italic) {
    styleParts.push("italic");
  }
  if (plan.style.underline) {
    styleParts.push("underline");
  }
  if (plan.style.strikethrough) {
    styleParts.push("strikethrough");
  }
  if (plan.style.numberFormat) {
    styleParts.push(`number format ${plan.style.numberFormat}`);
  }

  return styleParts.length > 0
    ? `Style: ${styleParts.join(", ")}.`
    : "Style: no formatting fields specified.";
}

function formatConditionalFormatRule(plan: ConditionalFormatPlanData): string {
  if (!("ruleType" in plan)) {
    return "Rule: clear existing conditional formatting.";
  }

  switch (plan.ruleType) {
    case "single_color":
    case "number_compare":
    case "date_compare": {
      const comparator = plan.comparator.replace(/_/g, " ");
      const value = plan.value !== undefined ? ` ${formatConditionalFormatValue(plan.value)}` : "";
      const value2 = plan.value2 !== undefined ? ` and ${formatConditionalFormatValue(plan.value2)}` : "";
      return `Rule: ${plan.ruleType.replace(/_/g, " ")} when value is ${comparator}${value}${value2}.`;
    }
    case "text_contains":
      return `Rule: text contains ${formatConditionalFormatValue(plan.text)}.`;
    case "duplicate_values":
      return "Rule: duplicate values.";
    case "custom_formula":
      return `Rule: custom formula ${plan.formula}.`;
    case "top_n":
      return `Rule: ${plan.direction} ${plan.rank} values.`;
    case "average_compare":
      return `Rule: values ${plan.direction} average.`;
    case "color_scale":
      return `Rule: ${plan.points.length}-point color scale.`;
  }
}

function buildConditionalFormatDetails(plan: ConditionalFormatPlanData): string[] {
  const modeDescription = plan.managementMode === "add"
    ? "This will keep existing conditional formatting on the target range."
    : plan.managementMode === "replace_all_on_target"
      ? "This will replace existing conditional formatting on the target range."
      : "This will clear existing conditional formatting from the target range.";

  const details = [
    `Target sheet: ${plan.targetSheet}.`,
    `Target range: ${plan.targetRange}.`,
    modeDescription,
    formatConditionalFormatRule(plan),
    formatConditionalFormatStyle(plan)
  ];

  if (plan.affectedRanges.length > 0) {
    details.push(`Affected ranges: ${plan.affectedRanges.join(", ")}.`);
  }

  return details;
}

function buildConditionalFormatPreview(
  plan: ConditionalFormatPlanData
): Extract<StructuredPreview, { kind: "conditional_format_plan" }> {
  const verb = plan.managementMode === "add"
    ? "Add conditional formatting"
    : plan.managementMode === "replace_all_on_target"
      ? "Replace all conditional formatting"
      : "Clear conditional formatting";

  return {
    kind: "conditional_format_plan",
    ...plan,
    summary: `Will ${verb.toLowerCase()} on ${plan.targetSheet}!${plan.targetRange}.`,
    details: buildConditionalFormatDetails(plan)
  };
}

function formatAnalysisReportSection(section: AnalysisReportPlanData["sections"][number]): string {
  return `${section.type}: ${section.title} - ${section.summary} (sources: ${section.sourceRanges.join(", ")}).`;
}

function buildAnalysisReportDetails(plan: AnalysisReportPlanData): string[] {
  const details = [
    `Source sheet: ${plan.sourceSheet}.`,
    `Source range: ${plan.sourceRange}.`,
    `Output mode: ${plan.outputMode}.`
  ];

  if ("targetSheet" in plan && plan.targetSheet) {
    details.push(`Target sheet: ${plan.targetSheet}.`);
  }

  if ("targetRange" in plan && plan.targetRange) {
    details.push(`Target range: ${plan.targetRange}.`);
  }

  details.push(...plan.sections.map(formatAnalysisReportSection));
  details.push(`Affected ranges: ${plan.affectedRanges.join(", ")}.`);
  details.push(`Overwrite risk: ${plan.overwriteRisk}.`);
  details.push(`Confirmation level: ${plan.confirmationLevel}.`);

  return details;
}

function buildAnalysisReportPreview(
  plan: AnalysisReportPlanData
): Extract<StructuredPreview, { kind: "analysis_report_plan" }> {
  const target = "targetSheet" in plan && plan.targetSheet
    ? `${plan.targetSheet}${"targetRange" in plan && plan.targetRange ? `!${plan.targetRange}` : ""}`
    : null;

  return {
    kind: "analysis_report_plan",
    ...plan,
    summary: plan.outputMode === "chat_only"
      ? `Will answer with a chat-only analysis of ${plan.sourceSheet}!${plan.sourceRange}.`
      : `Will materialize an analysis report on ${target ?? `${plan.sourceSheet}!${plan.sourceRange}`}.`,
    details: buildAnalysisReportDetails(plan)
  };
}

function buildCompositeStepSummary(step: CompositePlanData["steps"][number]): string {
  if ("explanation" in step.plan) {
    return step.plan.explanation;
  }

  if ("sourceAttachmentId" in step.plan) {
    return `Import extracted data into ${step.plan.targetSheet}!${step.plan.targetRange}.`;
  }

  return "Execute composite step.";
}

function isDestructiveCompositeStep(step: CompositePlanData["steps"][number]): boolean {
  return "confirmationLevel" in step.plan && step.plan.confirmationLevel === "destructive";
}

function isLikelyReversibleCompositeStep(step: CompositePlanData["steps"][number]): boolean {
  if ("rowGroups" in step.plan) {
    return false;
  }

  if ("chartType" in step.plan && "series" in step.plan) {
    return false;
  }

  return true;
}

function buildCompositePlanDetails(plan: CompositePlanData): string[] {
  const details = [
    `Steps: ${plan.steps.length}.`,
    `Reversible: ${plan.reversible ? "yes" : "no"}.`,
    `Dry run recommended: ${plan.dryRunRecommended ? "yes" : "no"}.`,
    `Dry run required: ${plan.dryRunRequired ? "yes" : "no"}.`
  ];

  for (const step of plan.steps) {
    const dependencyText = step.dependsOn.length > 0 ? step.dependsOn.join(", ") : "none";
    details.push(
      `${step.stepId}: depends on ${dependencyText}; continue on error ${step.continueOnError ? "yes" : "no"}.`
    );
    details.push(
      `Step flags: destructive ${isDestructiveCompositeStep(step) ? "yes" : "no"}, reversible ${isLikelyReversibleCompositeStep(step) ? "yes" : "no"}, skipped if dependency fails ${step.dependsOn.length > 0 ? "yes" : "no"}.`
    );
    details.push(`Step summary: ${buildCompositeStepSummary(step)}.`);
  }

  details.push(`Affected ranges: ${plan.affectedRanges.join(", ")}.`);
  details.push(`Overwrite risk: ${plan.overwriteRisk}.`);
  details.push(`Confirmation level: ${plan.confirmationLevel}.`);

  return details;
}

export function buildCompositePlanPreview(
  plan: CompositePlanData
): Extract<StructuredPreview, { kind: "composite_plan" }> {
  return {
    kind: "composite_plan",
    stepCount: plan.steps.length,
    destructiveConfirmationRequired: plan.confirmationLevel === "destructive",
    steps: plan.steps.map((step) => ({
      stepId: step.stepId,
      dependsOn: step.dependsOn,
      continueOnError: step.continueOnError,
      destructive: isDestructiveCompositeStep(step),
      reversible: isLikelyReversibleCompositeStep(step),
      skippedIfDependenciesFail: step.dependsOn.length > 0,
      summary: buildCompositeStepSummary(step)
    })),
    explanation: plan.explanation,
    confidence: plan.confidence,
    reversible: plan.reversible,
    dryRunRecommended: plan.dryRunRecommended,
    dryRunRequired: plan.dryRunRequired,
    affectedRanges: plan.affectedRanges,
    overwriteRisk: plan.overwriteRisk,
    confirmationLevel: plan.confirmationLevel,
    summary: `Will run ${plan.steps.length} workflow step${plan.steps.length === 1 ? "" : "s"}.`,
    details: buildCompositePlanDetails(plan)
  };
}

export function buildCompositeUpdatePreview(
  result: CompositeUpdateData
): Extract<StructuredPreview, { kind: "composite_update" }> {
  return {
    kind: "composite_update",
    executionId: result.executionId,
    stepResults: result.stepResults,
    summary: result.summary
  };
}

const DRY_RUN_INTERNAL_LANGUAGE_PATTERN = /\b(contract|schema|structured body|validation|json|payload|parse|parser|normaliz(?:e|ation)|exact-safe|live demo subset)\b/i;

function sanitizeDryRunUnsupportedReason(reason?: string): string | undefined {
  const resolvedReason = typeof reason === "string" ? reason.trim() : "";
  if (!resolvedReason) {
    return undefined;
  }

  if (/cannot hide the only visible worksheet/i.test(resolvedReason)) {
    return "At least one worksheet must stay visible.";
  }

  if (/invalid date literal/i.test(resolvedReason)) {
    return "The request includes a date that is not valid.";
  }

  if (
    /duplicate header/i.test(resolvedReason) ||
    /requires a header row/i.test(resolvedReason) ||
    /cannot find .* field in header row/i.test(resolvedReason) ||
    /could not resolve any valid sort keys/i.test(resolvedReason) ||
    /could not resolve a filter column inside the target range/i.test(resolvedReason) ||
    /column .* is outside /i.test(resolvedReason)
  ) {
    return "This preview needs a table with clear, matching column headers.";
  }

  if (
    /unsupported filter operator/i.test(resolvedReason) ||
    /unsupported filter combiner/i.test(resolvedReason) ||
    /filter combiners other than and/i.test(resolvedReason) ||
    /multiple conditions for the same column/i.test(resolvedReason) ||
    /cannot represent combiner "or" exactly/i.test(resolvedReason)
  ) {
    return "This preview can't represent that filter logic yet.";
  }

  if (
    /named ranges? on this scope/i.test(resolvedReason) ||
    /sheet-scoped named ranges/i.test(resolvedReason) ||
    /does not support creating named ranges/i.test(resolvedReason) ||
    /does not support renaming named ranges/i.test(resolvedReason) ||
    /does not support deleting named ranges/i.test(resolvedReason) ||
    /does not support retargeting named ranges/i.test(resolvedReason) ||
    /unsupported named range update/i.test(resolvedReason)
  ) {
    return "This preview can't represent that named range action here.";
  }

  if (
    /checkbox/i.test(resolvedReason) ||
    /data validation/i.test(resolvedReason) ||
    /allowBlank/i.test(resolvedReason) ||
    /invalidDataBehavior/i.test(resolvedReason) ||
    /validation comparator/i.test(resolvedReason) ||
    /Custom formula validation requires/i.test(resolvedReason)
  ) {
    return "This preview can't represent that validation setup safely.";
  }

  if (/pivot/i.test(resolvedReason) && /(does not support|requires|cannot find|unsupported)/i.test(resolvedReason)) {
    return "This preview can't represent that pivot configuration yet.";
  }

  if (/chart/i.test(resolvedReason) && /(does not support|requires|cannot find|unsupported)/i.test(resolvedReason)) {
    return "This preview can't represent that chart configuration yet.";
  }

  if (
    /cleanup/i.test(resolvedReason) ||
    /transfer/i.test(resolvedReason) ||
    /append/i.test(resolvedReason) ||
    /targetRange/i.test(resolvedReason) ||
    /range contains formulas/i.test(resolvedReason)
  ) {
    return "This preview can't represent that write safely on the current range.";
  }

  if (DRY_RUN_INTERNAL_LANGUAGE_PATTERN.test(resolvedReason)) {
    return "Dry-run preview isn't available for this plan in this spreadsheet app.";
  }

  return resolvedReason
    .replace(/^(Google Sheets|Excel)\s+host\s+/i, "This spreadsheet app ")
    .replace(/\bexact-safe\b/gi, "")
    .replace(/\bin the live demo subset\b/gi, "")
    .replace(/\blive demo subset\b/gi, "")
    .replace(/\s{2,}/g, " ")
    .replace(/\s+\./g, ".")
    .trim();
}

export function formatDryRunSummary(result: DryRunResult): string {
  const unsupportedReason = sanitizeDryRunUnsupportedReason(result.unsupportedReason);
  if (!result.simulated) {
    return unsupportedReason || "Dry-run preview isn't available for this plan.";
  }

  return result.predictedSummaries.join(" ");
}

export function formatHistoryEntrySummary(entry: PlanHistoryEntry): string {
  return entry.summary;
}

export function buildDryRunPreview(
  result: DryRunResult
): Extract<StructuredPreview, { kind: "dry_run_result" }> {
  const unsupportedReason = sanitizeDryRunUnsupportedReason(result.unsupportedReason);
  const details = [
    `Simulated: ${result.simulated ? "yes" : "no"}.`,
    `Overwrite risk: ${result.overwriteRisk}.`,
    `Reversible: ${result.reversible ? "yes" : "no"}.`,
    `Expires at: ${result.expiresAt}.`
  ];

  if (result.predictedAffectedRanges.length > 0) {
    details.push(`Predicted affected ranges: ${result.predictedAffectedRanges.join(", ")}.`);
  }

  if (result.predictedSummaries.length > 0) {
    details.push(`Predicted summaries: ${result.predictedSummaries.join(" ")}`);
  }

  if (result.steps && result.steps.length > 0) {
    details.push(`Predicted step outcomes: ${result.steps.map((step) => `${step.stepId}=${step.status}`).join(", ")}.`);
  }

  if (unsupportedReason) {
    details.push(`Reason: ${unsupportedReason}`);
  }

  return {
    kind: "dry_run_result",
    simulated: result.simulated,
    reversible: result.reversible,
    overwriteRisk: result.overwriteRisk,
    expiresAt: result.expiresAt,
    predictedAffectedRanges: result.predictedAffectedRanges,
    predictedSummaries: result.predictedSummaries,
    unsupportedReason,
    steps: result.steps ?? [],
    summary: formatDryRunSummary(result),
    details
  };
}

export function buildPlanHistoryPreview(
  page: PlanHistoryPage
): Extract<StructuredPreview, { kind: "plan_history_page" }> {
  return {
    kind: "plan_history_page",
    entries: page.entries,
    summary: page.entries.length === 0
      ? "No plan history entries."
      : `Loaded ${page.entries.length} plan history entr${page.entries.length === 1 ? "y" : "ies"}.`,
    details: page.entries.map((entry) => {
      const lineage = entry.linkedExecutionId ? ` linked=${entry.linkedExecutionId}` : "";
      return `${entry.timestamp} ${entry.summary} (undo=${entry.undoEligible ? "eligible" : "ineligible"}, redo=${entry.redoEligible ? "eligible" : "ineligible"})${lineage}`;
    }),
    nextCursor: page.nextCursor
  };
}

function formatPivotAggregation(aggregation: PivotTablePlanData["valueAggregations"][number]): string {
  return `${aggregation.field} ${aggregation.aggregation}`;
}

function formatPivotFilter(filter: NonNullable<PivotTablePlanData["filters"]>[number]): string {
  const valuePart = filter.value === undefined ? "" : ` ${JSON.stringify(filter.value)}`;
  const value2Part = filter.value2 === undefined ? "" : ` and ${JSON.stringify(filter.value2)}`;
  return `${filter.field} ${filter.operator}${valuePart}${value2Part}`;
}

function formatPivotSort(plan: PivotTablePlanData["sort"]): string {
  if (!plan) {
    return "";
  }

  return `${plan.field} ${plan.direction} on ${plan.sortOn.replace(/_/g, " ")}`;
}

function buildPivotTableDetails(plan: PivotTablePlanData): string[] {
  const details = [
    `Source sheet: ${plan.sourceSheet}.`,
    `Source range: ${plan.sourceRange}.`,
    `Target sheet: ${plan.targetSheet}.`,
    `Target range: ${plan.targetRange}.`,
    `Row groups: ${joinNaturalList(plan.rowGroups)}.`,
    `Column groups: ${plan.columnGroups && plan.columnGroups.length > 0 ? joinNaturalList(plan.columnGroups) : "none"}.`,
    `Value aggregations: ${joinNaturalList(plan.valueAggregations.map(formatPivotAggregation))}.`
  ];

  if (plan.filters && plan.filters.length > 0) {
    details.push(`Filters: ${joinNaturalList(plan.filters.map(formatPivotFilter))}.`);
  }

  if (plan.sort) {
    details.push(`Sort: ${formatPivotSort(plan.sort)}.`);
  }

  details.push(`Affected ranges: ${plan.affectedRanges.join(", ")}.`);
  details.push(`Overwrite risk: ${plan.overwriteRisk}.`);
  details.push(`Confirmation level: ${plan.confirmationLevel}.`);

  return details;
}

function buildPivotTablePreview(
  plan: PivotTablePlanData
): Extract<StructuredPreview, { kind: "pivot_table_plan" }> {
  return {
    kind: "pivot_table_plan",
    ...plan,
    summary: `Will create a pivot table on ${plan.targetSheet}!${plan.targetRange} from ${plan.sourceSheet}!${plan.sourceRange}.`,
    details: buildPivotTableDetails(plan)
  };
}

function formatChartSeries(series: ChartPlanData["series"][number]): string {
  return series.label ? `${series.field} as ${series.label}` : series.field;
}

function buildChartDetails(plan: ChartPlanData): string[] {
  const details = [
    `Source sheet: ${plan.sourceSheet}.`,
    `Source range: ${plan.sourceRange}.`,
    `Target sheet: ${plan.targetSheet}.`,
    `Target range: ${plan.targetRange}.`,
    `Chart type: ${plan.chartType}.`,
    `Category field: ${plan.categoryField ?? "none"}.`,
    `Series: ${joinNaturalList(plan.series.map(formatChartSeries))}.`
  ];

  if (plan.title) {
    details.push(`Title: ${plan.title}.`);
  }

  if (plan.legendPosition) {
    details.push(`Legend position: ${plan.legendPosition}.`);
  }

  details.push(`Affected ranges: ${plan.affectedRanges.join(", ")}.`);
  details.push(`Overwrite risk: ${plan.overwriteRisk}.`);
  details.push(`Confirmation level: ${plan.confirmationLevel}.`);

  return details;
}

function buildChartPreview(
  plan: ChartPlanData
): Extract<StructuredPreview, { kind: "chart_plan" }> {
  return {
    kind: "chart_plan",
    ...plan,
    summary: `Will create a ${plan.chartType} chart on ${plan.targetSheet}!${plan.targetRange} from ${plan.sourceSheet}!${plan.sourceRange}.`,
    details: buildChartDetails(plan)
  };
}

function formatRangeTransferSummary(plan: RangeTransferPlanData): string {
  const action = plan.operation === "copy"
    ? "copy"
    : plan.operation === "move"
      ? "move"
      : "append";
  const source = `${plan.sourceSheet}!${plan.sourceRange}`;
  const target = `${plan.targetSheet}!${plan.targetRange}`;
  const mode = `${plan.transpose ? "transposed " : ""}${plan.pasteMode}`;
  const suffix = plan.operation === "move"
    ? " and clear the source after success"
    : "";

  return `Will ${action} ${mode} from ${source} to ${target}${suffix}.`;
}

function joinNaturalList(values: string[]): string {
  if (values.length <= 1) {
    return values[0] ?? "";
  }

  if (values.length === 2) {
    return `${values[0]} and ${values[1]}`;
  }

  return `${values.slice(0, -1).join(", ")}, and ${values[values.length - 1]}`;
}

function formatCleanupOperationLabel(operation: DataCleanupPlanData["operation"]): string {
  switch (operation) {
    case "trim_whitespace":
      return "trim whitespace";
    case "remove_blank_rows":
      return "remove blank rows";
    case "remove_duplicate_rows":
      return "remove duplicate rows";
    case "normalize_case":
      return "normalize case";
    case "split_column":
      return "split column";
    case "join_columns":
      return "join columns";
    case "fill_down":
      return "fill down";
    case "standardize_format":
      return "standardize format";
  }
}

function buildRangeTransferDetails(plan: RangeTransferPlanData): string[] {
  const details = [
    `Source sheet: ${plan.sourceSheet}.`,
    `Source range: ${plan.sourceRange}.`,
    `Target sheet: ${plan.targetSheet}.`,
    `Target range: ${plan.targetRange}.`,
    `Operation: ${plan.operation}.`,
    `Paste mode: ${plan.pasteMode}.`,
    `Transpose: ${plan.transpose ? "on" : "off"}.`
  ];

  details.push(
    plan.operation === "move"
      ? "This will clear the source after success."
      : "This will leave the source unchanged."
  );
  details.push(`Affected ranges: ${plan.affectedRanges.join(", ")}.`);
  details.push(`Overwrite risk: ${plan.overwriteRisk}.`);
  details.push(`Confirmation level: ${plan.confirmationLevel}.`);

  return details;
}

function formatDataCleanupSummary(plan: DataCleanupPlanData): string {
  switch (plan.operation) {
    case "trim_whitespace":
      return `Will trim whitespace in ${plan.targetSheet}!${plan.targetRange}.`;
    case "remove_blank_rows":
      return plan.keyColumns && plan.keyColumns.length > 0
        ? `Will remove blank rows in ${plan.targetSheet}!${plan.targetRange} using key columns ${joinNaturalList(plan.keyColumns)}.`
        : `Will remove blank rows in ${plan.targetSheet}!${plan.targetRange}.`;
    case "remove_duplicate_rows":
      return plan.keyColumns && plan.keyColumns.length > 0
        ? `Will remove duplicate rows in ${plan.targetSheet}!${plan.targetRange} using key columns ${joinNaturalList(plan.keyColumns)}.`
        : `Will remove duplicate rows in ${plan.targetSheet}!${plan.targetRange}.`;
    case "normalize_case":
      return `Will normalize case to ${plan.mode} in ${plan.targetSheet}!${plan.targetRange}.`;
    case "split_column":
      return `Will split column ${plan.sourceColumn} on ${JSON.stringify(plan.delimiter)} into columns starting at ${plan.targetStartColumn} in ${plan.targetSheet}!${plan.targetRange}.`;
    case "join_columns":
      return `Will join columns ${plan.sourceColumns.join(", ")} with ${JSON.stringify(plan.delimiter)} into column ${plan.targetColumn} in ${plan.targetSheet}!${plan.targetRange}.`;
    case "fill_down":
      return plan.columns && plan.columns.length > 0
        ? `Will fill down columns ${plan.columns.join(", ")} in ${plan.targetSheet}!${plan.targetRange}.`
        : `Will fill down values in ${plan.targetSheet}!${plan.targetRange}.`;
    case "standardize_format":
      return `Will standardize format to ${plan.formatType} using ${plan.formatPattern} in ${plan.targetSheet}!${plan.targetRange}.`;
  }
}

function buildDataCleanupDetails(plan: DataCleanupPlanData): string[] {
  const details = [
    `Target sheet: ${plan.targetSheet}.`,
    `Target range: ${plan.targetRange}.`,
    `Operation: ${formatCleanupOperationLabel(plan.operation)}.`
  ];

  switch (plan.operation) {
    case "remove_blank_rows":
    case "remove_duplicate_rows":
      if (plan.keyColumns && plan.keyColumns.length > 0) {
        details.push(`Key columns: ${joinNaturalList(plan.keyColumns)}.`);
      }
      details.push("This will remove rows from the target range.");
      break;
    case "normalize_case":
      details.push(`Mode: ${plan.mode}.`);
      details.push("This will rewrite cell values in the target range.");
      break;
    case "split_column":
      details.push(`Source column: ${plan.sourceColumn}.`);
      details.push(`Delimiter: ${JSON.stringify(plan.delimiter)}.`);
      details.push(`Target start column: ${plan.targetStartColumn}.`);
      details.push("This will rewrite target columns in the target range.");
      break;
    case "join_columns":
      details.push(`Source columns: ${joinNaturalList(plan.sourceColumns)}.`);
      details.push(`Delimiter: ${JSON.stringify(plan.delimiter)}.`);
      details.push(`Target column: ${plan.targetColumn}.`);
      details.push("This will rewrite target columns in the target range.");
      break;
    case "fill_down":
      if (plan.columns && plan.columns.length > 0) {
        details.push(`Columns: ${joinNaturalList(plan.columns)}.`);
      }
      details.push("This will rewrite cell values in the target range.");
      break;
    case "standardize_format":
      details.push(`Format type: ${plan.formatType}.`);
      details.push(`Format pattern: ${plan.formatPattern}.`);
      details.push("This will rewrite cell values in the target range.");
      break;
    case "trim_whitespace":
      details.push("This will rewrite cell values in the target range.");
      break;
  }

  details.push(`Affected ranges: ${plan.affectedRanges.join(", ")}.`);
  details.push(`Overwrite risk: ${plan.overwriteRisk}.`);
  details.push(`Confirmation level: ${plan.confirmationLevel}.`);

  return details;
}

export function buildRangeTransferPreview(
  plan: RangeTransferPlanData
): Extract<StructuredPreview, { kind: "range_transfer_plan" }> {
  return {
    kind: "range_transfer_plan",
    ...plan,
    summary: formatRangeTransferSummary(plan),
    details: buildRangeTransferDetails(plan)
  };
}

export function buildDataCleanupPreview(
  plan: DataCleanupPlanData
): Extract<StructuredPreview, { kind: "data_cleanup_plan" }> {
  return {
    kind: "data_cleanup_plan",
    ...plan,
    summary: formatDataCleanupSummary(plan),
    details: buildDataCleanupDetails(plan)
  };
}

export function buildRangeTransferUpdatePreview(
  result: RangeTransferUpdateData
): Extract<StructuredPreview, { kind: "range_transfer_update" }> {
  return {
    kind: "range_transfer_update",
    ...result
  };
}

export function buildDataCleanupUpdatePreview(
  result: DataCleanupUpdateData
): Extract<StructuredPreview, { kind: "data_cleanup_update" }> {
  return {
    kind: "data_cleanup_update",
    ...result
  };
}

export function buildWriteMatrix(plan: MatrixWritePlan): RenderCell[][] {
  if ("headers" in plan) {
    return [plan.headers, ...plan.values];
  }

  const sections: Array<{ label: string; rows: RenderCell[][] } > = [];

  if ("values" in plan && plan.values !== undefined) {
    sections.push({ label: "values", rows: plan.values });
  }

  if ("formulas" in plan && plan.formulas !== undefined) {
    sections.push({ label: "formulas", rows: plan.formulas });
  }

  if ("notes" in plan && plan.notes !== undefined) {
    sections.push({ label: "notes", rows: plan.notes });
  }

  if (sections.length === 0) {
    return [];
  }

  if (sections.length === 1) {
    return sections[0].rows;
  }

  const combined: RenderCell[][] = [];
  for (const section of sections) {
    combined.push([section.label]);
    combined.push(...section.rows);
  }

  return combined;
}

export function buildWorkbookStructurePreview(
  plan: WorkbookStructureUpdateData
): Extract<StructuredPreview, { kind: "workbook_structure_update" }> {
  return {
    kind: "workbook_structure_update",
    operation: plan.operation,
    sheetName: plan.sheetName,
    position: "position" in plan ? plan.position : undefined,
    newSheetName: "newSheetName" in plan ? plan.newSheetName : undefined,
    explanation: plan.explanation,
    overwriteRisk: plan.overwriteRisk
  };
}

export function buildSheetStructurePreview(
  plan: SheetStructureUpdateData
): Extract<StructuredPreview, { kind: "sheet_structure_update" }> {
  return {
    kind: "sheet_structure_update",
    targetSheet: plan.targetSheet,
    operation: plan.operation,
    summary: plan.explanation,
    confirmationLevel: plan.confirmationLevel
  };
}

export function buildRangeFormatPreview(
  plan: RangeFormatUpdateData
): Extract<StructuredPreview, { kind: "range_format_update" }> {
  return {
    kind: "range_format_update",
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    format: plan.format,
    explanation: plan.explanation,
    overwriteRisk: plan.overwriteRisk
  };
}

export function buildRangeSortPreview(
  plan: RangeSortPlanData
): Extract<StructuredPreview, { kind: "range_sort_plan" }> {
  return {
    kind: "range_sort_plan",
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    hasHeader: plan.hasHeader,
    keys: plan.keys,
    summary: plan.explanation
  };
}

export function buildRangeFilterPreview(
  plan: RangeFilterPlanData
): Extract<StructuredPreview, { kind: "range_filter_plan" }> {
  return {
    kind: "range_filter_plan",
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    hasHeader: plan.hasHeader,
    conditions: plan.conditions,
    combiner: plan.combiner,
    clearExistingFilters: plan.clearExistingFilters,
    summary: plan.explanation
  };
}

export function buildDataValidationPreview(
  plan: DataValidationPlanData
): Extract<StructuredPreview, { kind: "data_validation_plan" }> {
  return {
    kind: "data_validation_plan",
    ...plan
  };
}

export function buildNamedRangeUpdatePreview(
  plan: NamedRangeUpdateData
): Extract<StructuredPreview, { kind: "named_range_update" }> {
  return {
    kind: "named_range_update",
    ...plan
  };
}

export function buildSheetImportPreview(plan: SheetImportPlanData): PreviewTable {
  return {
    headers: plan.headers,
    rows: plan.values,
    shape: plan.shape
  };
}

export function buildSheetUpdatePreview(plan: SheetUpdateData): {
  matrixKind: "values" | "formulas" | "notes" | "mixed_update";
  table: PreviewTable;
} {
  const hasValues = "values" in plan && plan.values !== undefined;
  const hasFormulas = "formulas" in plan && plan.formulas !== undefined;
  const hasNotes = "notes" in plan && plan.notes !== undefined;
  const matrices = [
    hasValues ? "values" : null,
    hasFormulas ? "formulas" : null,
    hasNotes ? "notes" : null
  ].filter(Boolean) as Array<"values" | "formulas" | "notes">;

  const matrixKind = matrices.length > 1
    ? "mixed_update"
    : (matrices[0] ?? "values");
  const rows = buildWriteMatrix(plan);

  return {
    matrixKind,
    table: {
      headers: [],
      rows,
      shape: plan.shape
    }
  };
}

export function buildExtractedTablePreview(response: ExtractedTableResponse): PreviewTable {
  const inferred = response.data.shape ?? {
    rows: response.data.rows.length,
    columns: response.data.headers.length > 0
      ? response.data.headers.length
      : matrixShape(response.data.rows).columns
  };

  return {
    headers: response.data.headers,
    rows: response.data.rows,
    shape: inferred
  };
}

export function getStructuredPreview(response: HermesResponse): StructuredPreview | null {
  switch (response.type) {
    case "composite_plan":
      return buildCompositePlanPreview(response.data);
    case "composite_update":
      return buildCompositeUpdatePreview(response.data);
    case "formula":
      return {
        kind: "formula",
        intent: response.data.intent,
        formula: response.data.formula,
        formulaLanguage: response.data.formulaLanguage,
        targetCell: response.data.targetCell,
        explanation: response.data.explanation,
        alternateFormulas: response.data.alternateFormulas
      };
    case "workbook_structure_update":
      return buildWorkbookStructurePreview(response.data);
    case "sheet_structure_update":
      return buildSheetStructurePreview(response.data);
    case "range_format_update":
      return buildRangeFormatPreview(response.data);
    case "analysis_report_plan":
      return buildAnalysisReportPreview(response.data);
    case "pivot_table_plan":
      return buildPivotTablePreview(response.data);
    case "chart_plan":
      return buildChartPreview(response.data);
    case "conditional_format_plan":
      return buildConditionalFormatPreview(response.data);
    case "range_sort_plan":
      return buildRangeSortPreview(response.data);
    case "range_filter_plan":
      return buildRangeFilterPreview(response.data);
    case "range_transfer_plan":
      return buildRangeTransferPreview(response.data);
    case "data_cleanup_plan":
      return buildDataCleanupPreview(response.data);
    case "data_validation_plan":
      return buildDataValidationPreview(response.data);
    case "named_range_update":
      return buildNamedRangeUpdatePreview(response.data);
    case "range_transfer_update":
      return buildRangeTransferUpdatePreview(response.data);
    case "data_cleanup_update":
      return buildDataCleanupUpdatePreview(response.data);
    case "analysis_report_update":
      return {
        kind: "analysis_report_update",
        ...response.data
      };
    case "pivot_table_update":
      return {
        kind: "pivot_table_update",
        ...response.data
      };
    case "chart_update":
      return {
        kind: "chart_update",
        ...response.data
      };
    case "conditional_format_update":
      return {
        kind: "conditional_format_update",
        ...response.data
      };
    case "sheet_update": {
      const preview = buildSheetUpdatePreview(response.data);
      return {
        kind: "sheet_update",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        shape: response.data.shape,
        operation: response.data.operation,
        overwriteRisk: response.data.overwriteRisk,
        matrixKind: preview.matrixKind,
        table: preview.table
      };
    }
    case "sheet_import_plan":
      return {
        kind: "sheet_import_plan",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        shape: response.data.shape,
        extractionMode: response.data.extractionMode,
        sourceAttachmentId: response.data.sourceAttachmentId,
        table: buildSheetImportPreview(response.data)
      };
    case "attachment_analysis":
      return {
        kind: "attachment_analysis",
        sourceAttachmentId: response.data.sourceAttachmentId,
        contentKind: response.data.contentKind,
        extractionMode: response.data.extractionMode,
        summary: response.data.summary
      };
    case "extracted_table":
      return {
        kind: "extracted_table",
        sourceAttachmentId: response.data.sourceAttachmentId,
        extractionMode: response.data.extractionMode,
        table: buildExtractedTablePreview(response)
      };
    case "document_summary":
      return {
        kind: "document_summary",
        sourceAttachmentId: response.data.sourceAttachmentId,
        contentKind: response.data.contentKind,
        extractionMode: response.data.extractionMode,
        keyPoints: response.data.keyPoints ?? []
      };
    default:
      return null;
  }
}

export function buildStructuredPreview(
  response: HermesResponse
): StructuredPreview | null {
  return getStructuredPreview(response);
}

function formatUserFacingErrorText(message: string, userAction?: string): string {
  const resolvedMessage = String(message ?? "").trim();
  const resolvedUserAction = typeof userAction === "string" ? userAction.trim() : "";

  if (!resolvedUserAction || resolvedUserAction === resolvedMessage) {
    return resolvedMessage;
  }

  return `${resolvedMessage}\n\n${resolvedUserAction}`;
}

export function getResponseBodyText(response: HermesResponse): string {
  switch (response.type) {
    case "composite_plan":
      return buildCompositePlanPreview(response.data).summary;
    case "composite_update":
      return response.data.summary;
    case "chat":
      return response.data.message;
    case "formula":
      return response.data.explanation;
    case "analysis_report_plan":
      return buildAnalysisReportPreview(response.data).summary;
    case "pivot_table_plan":
      return buildPivotTablePreview(response.data).summary;
    case "chart_plan":
      return buildChartPreview(response.data).summary;
    case "workbook_structure_update":
      switch (response.data.operation) {
        case "create_sheet":
          return `Prepared a workbook update to create sheet ${response.data.sheetName}.`;
        case "delete_sheet":
          return `Prepared a workbook update to delete sheet ${response.data.sheetName}.`;
        case "rename_sheet":
          return `Prepared a workbook update to rename ${response.data.sheetName} to ${response.data.newSheetName}.`;
        case "duplicate_sheet":
          return `Prepared a workbook update to duplicate sheet ${response.data.sheetName}.`;
        case "move_sheet":
          return `Prepared a workbook update to move sheet ${response.data.sheetName}.`;
        case "hide_sheet":
          return `Prepared a workbook update to hide sheet ${response.data.sheetName}.`;
        case "unhide_sheet":
          return `Prepared a workbook update to unhide sheet ${response.data.sheetName}.`;
        default:
          return "Prepared a workbook structure update.";
      }
    case "sheet_structure_update":
      return `Prepared a sheet structure update for ${response.data.targetSheet}.`;
    case "range_format_update":
      return `Prepared a formatting update for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "conditional_format_plan":
      return buildConditionalFormatPreview(response.data).summary;
    case "range_sort_plan":
      return `Prepared a sort plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "range_filter_plan":
      return `Prepared a filter plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "range_transfer_plan":
      return buildRangeTransferPreview(response.data).summary;
    case "data_cleanup_plan":
      return buildDataCleanupPreview(response.data).summary;
    case "data_validation_plan":
      return `Prepared a validation plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "named_range_update":
      return `Prepared a named range update for ${response.data.name}.`;
    case "range_transfer_update":
      return response.data.summary;
    case "data_cleanup_update":
      return response.data.summary;
    case "analysis_report_update":
      return response.data.summary;
    case "pivot_table_update":
      return response.data.summary;
    case "chart_update":
      return response.data.summary;
    case "conditional_format_update":
      return response.data.summary;
    case "sheet_update":
      return response.data.explanation;
    case "sheet_import_plan":
      return `Prepared an import preview for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "error":
      return formatUserFacingErrorText(response.data.message, response.data.userAction);
    case "attachment_analysis":
      return response.data.summary;
    case "extracted_table":
      return "Prepared an extracted table preview from the uploaded image.";
    case "document_summary":
      return response.data.summary;
  }

  return "Hermes response received.";
}

export function getResponseConfidence(response: HermesResponse): number | undefined {
  switch (response.type) {
    case "composite_plan":
    case "chat":
    case "formula":
    case "workbook_structure_update":
    case "sheet_structure_update":
    case "range_format_update":
    case "conditional_format_plan":
    case "range_sort_plan":
    case "range_filter_plan":
    case "range_transfer_plan":
    case "data_cleanup_plan":
    case "sheet_update":
    case "sheet_import_plan":
    case "attachment_analysis":
    case "extracted_table":
    case "document_summary":
    case "data_validation_plan":
    case "named_range_update":
    case "analysis_report_plan":
    case "pivot_table_plan":
    case "chart_plan":
      return response.data.confidence;
    default:
      return undefined;
  }
}

export function getResponseWarnings(response: HermesResponse): Warning[] {
  const topLevel = response.warnings ?? [];
  const dataWarnings = hasOwnWarnings(response.data) ? response.data.warnings ?? [] : [];
  return [...topLevel, ...dataWarnings];
}

export function getRequiresConfirmation(response: HermesResponse): boolean {
  switch (response.type) {
    case "composite_plan":
      return response.data.requiresConfirmation;
    case "formula":
      return response.data.requiresConfirmation ?? false;
    case "workbook_structure_update":
    case "sheet_structure_update":
    case "range_format_update":
    case "conditional_format_plan":
    case "range_sort_plan":
    case "range_filter_plan":
    case "range_transfer_plan":
    case "data_cleanup_plan":
    case "sheet_update":
    case "sheet_import_plan":
    case "data_validation_plan":
    case "named_range_update":
    case "pivot_table_plan":
    case "chart_plan":
      return response.data.requiresConfirmation;
    case "analysis_report_plan":
      return response.data.outputMode === "materialize_report";
    default:
      return false;
  }
}

export function getResponseMetaLine(response: HermesResponse): string {
  const parts: string[] = [];

  if (response.skillsUsed && response.skillsUsed.length > 0) {
    parts.push(`skills ${response.skillsUsed.join(", ")}`);
  }

  if (response.downstreamProvider?.label) {
    const provider = response.downstreamProvider.model
      ? `${response.downstreamProvider.label}/${response.downstreamProvider.model}`
      : response.downstreamProvider.label;
    parts.push(`provider ${provider}`);
  }

  if (response.type === "composite_plan") {
    parts.push(`workflow steps ${response.data.steps.length}`);
    if (response.data.dryRunRequired) {
      parts.push("dry-run required");
    } else if (response.data.dryRunRecommended) {
      parts.push("dry-run recommended");
    }
  } else if (response.type === "composite_update") {
    parts.push(`composite update ${response.data.executionId}`);
  }

  const confidence = getResponseConfidence(response);
  if (typeof confidence === "number" && response.ui.showConfidence) {
    parts.push(`confidence ${(confidence * 100).toFixed(0)}%`);
  }

  if (response.ui.showRequiresConfirmation && getRequiresConfirmation(response)) {
    parts.push("confirmation required");
  }

  switch (response.type) {
    case "sheet_import_plan":
    case "attachment_analysis":
    case "extracted_table":
    case "document_summary":
      parts.push(`extraction ${response.data.extractionMode}`);
      break;
    default:
      break;
  }

  return parts.join(" • ");
}

export function getFollowUpSuggestions(response: HermesResponse): string[] {
  return response.type === "chat"
    ? response.data.followUpSuggestions ?? []
    : [];
}

export function isWritePlanResponse(
  response: HermesResponse
): response is
  | CompositePlanResponse
  | SheetImportPlanResponse
  | SheetUpdateResponse
  | WorkbookStructureUpdateResponse
  | SheetStructureUpdateResponse
  | RangeFormatUpdateResponse
  | ConditionalFormatPlanResponse
  | RangeSortPlanResponse
  | RangeFilterPlanResponse
  | RangeTransferPlanResponse
  | DataCleanupPlanResponse
  | DataValidationPlanResponse
  | NamedRangeUpdateResponse
  | Extract<AnalysisReportPlanResponse, { data: AnalysisReportWritePlanData }>
  | PivotTablePlanResponse
  | ChartPlanResponse {
  return response.type === "sheet_import_plan" ||
    response.type === "composite_plan" ||
    response.type === "sheet_update" ||
    response.type === "workbook_structure_update" ||
    response.type === "sheet_structure_update" ||
    response.type === "range_format_update" ||
    response.type === "conditional_format_plan" ||
    response.type === "range_sort_plan" ||
    response.type === "range_filter_plan" ||
    response.type === "range_transfer_plan" ||
    response.type === "data_cleanup_plan" ||
    response.type === "data_validation_plan" ||
    response.type === "named_range_update" ||
    response.type === "pivot_table_plan" ||
    response.type === "chart_plan" ||
    (response.type === "analysis_report_plan" && response.data.outputMode === "materialize_report");
}
