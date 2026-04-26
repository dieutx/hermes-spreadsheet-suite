import type {
  HermesResponse,
  HermesTraceEvent
} from "@hermes/contracts";
import type { WritebackResult } from "./types.js";

const EVENT_LABELS: Record<string, string> = {
  request_received: "Request received by Hermes",
  spreadsheet_context_received: "Spreadsheet context received",
  attachment_received: "Attachment received",
  image_received: "Image received",
  skill_selected: "Skill selected",
  tool_selected: "Tool selected",
  downstream_provider_called: "Downstream provider called",
  ocr_started: "OCR started",
  table_extraction_started: "Table extraction started",
  result_generated: "Result generated",
  sheet_update_plan_ready: "Sheet update plan ready",
  sheet_import_plan_ready: "Sheet import plan ready",
  workbook_structure_update_ready: "Workbook structure update ready",
  sheet_structure_update_ready: "Sheet structure update ready",
  range_format_update_ready: "Range format update ready",
  conditional_format_plan_ready: "Conditional format plan ready",
  range_sort_plan_ready: "Range sort plan ready",
  range_filter_plan_ready: "Range filter plan ready",
  data_validation_plan_ready: "Data validation plan ready",
  named_range_update_ready: "Named range update ready",
  range_transfer_plan_ready: "Range transfer plan ready",
  data_cleanup_plan_ready: "Data cleanup plan ready",
  range_transfer_update_ready: "Range transfer update ready",
  data_cleanup_update_ready: "Data cleanup update ready",
  analysis_report_plan_ready: "Analysis report plan ready",
  pivot_table_plan_ready: "Pivot table plan ready",
  chart_plan_ready: "Chart plan ready",
  table_plan_ready: "Table plan ready",
  analysis_report_update_ready: "Analysis report update ready",
  pivot_table_update_ready: "Pivot table update ready",
  chart_update_ready: "Chart update ready",
  table_update_ready: "Table update ready",
  composite_plan_ready: "Composite plan ready",
  composite_update_ready: "Composite update ready",
  dry_run_requested: "Dry run requested",
  dry_run_completed: "Dry run completed",
  history_requested: "History requested",
  undo_requested: "Undo requested",
  redo_requested: "Redo requested",
  completed: "Completed",
  failed: "Failed"
};

export function formatTraceEvent(event: HermesTraceEvent): string {
  if (event.label && event.label.trim().length > 0) {
    return event.label;
  }

  return EVENT_LABELS[event.event] ?? "Waiting for Hermes";
}

export function summarizeLatestTrace(trace: HermesTraceEvent[]): string {
  if (trace.length === 0) {
    return "Waiting for Hermes";
  }

  return formatTraceEvent(trace[trace.length - 1]);
}

export function formatTraceTimeline(trace: HermesTraceEvent[]): string {
  if (trace.length === 0) {
    return "";
  }

  return trace
    .map(formatTraceEvent)
    .filter((label, index, labels) => index === 0 || label !== labels[index - 1])
    .join(" -> ");
}

export function formatProofLine(response: HermesResponse): string {
  return [
    "Processed by Hermes",
    `requestId ${response.requestId}`,
    `hermesRunId ${response.hermesRunId}`,
    `service ${response.serviceLabel}`,
    `environment ${response.environmentLabel}`,
    `${response.durationMs}ms`
  ].join(" • ");
}

export function formatWritebackStatusLine(result: WritebackResult): string {
  switch (result.kind) {
    case "range_write":
      return `Wrote ${result.writtenRows} rows x ${result.writtenColumns} columns to ${result.targetSheet}!${result.targetRange}.`;
    case "range_format_update":
    case "workbook_structure_update":
    case "sheet_structure_update":
    case "range_sort":
    case "range_filter":
    case "data_validation_update":
    case "conditional_format_update":
    case "named_range_update":
    case "range_transfer_update":
    case "data_cleanup_update":
    case "analysis_report_update":
    case "pivot_table_update":
    case "chart_update":
    case "table_update":
    case "composite_update":
      return result.summary;
  }
}
