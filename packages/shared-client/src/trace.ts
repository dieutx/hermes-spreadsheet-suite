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
  external_data_plan_ready: "External data plan ready",
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

const UNSAFE_PROOF_VALUE_PATTERN = /(?:APPROVAL_SECRET|HERMES_API_SERVER_KEY|HERMES_AGENT_API_KEY|HERMES_AGENT_BASE_URL|OPENAI_API_KEY|ANTHROPIC_API_KEY|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)|\/(?:root|srv|home|tmp|var|opt|workspace|app|mnt|Users)\/[^\s]+|[A-Za-z]:\\[^\s]+|(?:^|[\s=:])\\\\[^\s]+|https?:\/\/(?:internal(?:[.\w-]*)?|localhost|127\.0\.0\.1|10\.\d{1,3}\.\d{1,3}\.\d{1,3}|169\.254\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[0-1])\.\d{1,3}\.\d{1,3}|\[(?:::ffff:|(?:0:){5}ffff:)(?:(?:127|10)(?:\.\d{1,3}){3}|169\.254\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])\.\d{1,3}\.\d{1,3}|(?:0{1,4}|7f[0-9a-f]{2}|0?a[0-9a-f]{2}|a9fe|c0a8|ac1[0-9a-f]):[0-9a-f]{1,4})\]|\[(?:::|::1|f[cd][0-9a-f:]*|fe[89ab][0-9a-f:]*)\])[^\s]*/i;
const UNSAFE_NUMERIC_IPV4_URL_PATTERN = /https?:\/\/(?:0x[0-9a-f]+|0[0-7]+|\d+)(?:\.(?:0x[0-9a-f]+|0[0-7]+|\d+)){0,3}(?::\d+)?(?:[/?#]|\s|$)/i;
const PUBLIC_PROOF_IDENTIFIER_PATTERN = /^[A-Za-z0-9._:-]+$/;

function safeProofValue(value: unknown, fallback = "unavailable"): string {
  const normalized = typeof value === "string" ? value.trim() : "";
  if (
    !normalized ||
    normalized.length > 256 ||
    UNSAFE_PROOF_VALUE_PATTERN.test(normalized) ||
    UNSAFE_NUMERIC_IPV4_URL_PATTERN.test(normalized)
  ) {
    return fallback;
  }

  return normalized;
}

function safeProofIdentifier(value: unknown, fallback = "unavailable"): string {
  const normalized = safeProofValue(value, fallback);
  return PUBLIC_PROOF_IDENTIFIER_PATTERN.test(normalized) ? normalized : fallback;
}

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
    `requestId ${safeProofIdentifier(response.requestId)}`,
    `hermesRunId ${safeProofIdentifier(response.hermesRunId)}`,
    `service ${safeProofValue(response.serviceLabel)}`,
    `environment ${safeProofValue(response.environmentLabel)}`,
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
    case "external_data_update":
    case "composite_update":
      return result.summary;
  }
}
