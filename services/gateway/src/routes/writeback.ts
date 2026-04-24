import { randomUUID } from "node:crypto";
import { Router } from "express";
import { z } from "zod";
import {
  AnalysisReportPlanDataSchema,
  ChartPlanDataSchema,
  CompositePlanDataSchema,
  CompositeUpdateDataSchema,
  ConditionalFormatManagementModeSchema,
  DataCleanupPlanDataSchema,
  DataValidationPlanDataSchema,
  ConditionalFormatPlanDataSchema,
  ExternalDataPlanDataSchema,
  PivotTablePlanDataSchema,
  NamedRangeUpdateDataSchema,
  RangeTransferPlanDataSchema,
  RangeTransferUpdateDataSchema,
  RangeFormatUpdateDataSchema,
  RangeFilterPlanDataSchema,
  RangeSortPlanDataSchema,
  SheetImportPlanDataSchema,
  SheetStructureUpdateDataSchema,
  SheetUpdateDataSchema,
  WorkbookStructureUpdateDataSchema,
  parseA1Range
} from "@hermes/contracts";
import type {
  AnalysisReportPlanData,
  ChartPlanData,
  CompositePlanData,
  ExternalDataPlanData,
  HermesResponse,
  NamedRangeUpdateData,
  PivotTablePlanData,
  RangeTransferPlanData,
  SheetStructureUpdateData
} from "@hermes/contracts";
import {
  canonicalizePlan,
  createApprovalToken,
  digestCanonicalPlan,
  verifyApprovalToken
} from "../lib/approval.js";
import type { GatewayConfig } from "../lib/config.js";
import type { TraceBus } from "../lib/traceBus.js";
import {
  ExecutionLedger,
  FreshDryRunRequiredError
} from "../lib/executionLedger.js";
import { normalizeCompositePlanForDigest } from "../lib/planNormalization.js";

const APPROVAL_TOKEN_MAX_AGE_MS = 15 * 60 * 1000;

const ApprovalRequestSchema = z.object({
  requestId: z.string().min(1),
  runId: z.string().min(1),
  workbookSessionKey: z.string().min(1).max(256).optional(),
  destructiveConfirmation: z.object({
    confirmed: z.literal(true)
  }).optional(),
  plan: z.union([
    SheetUpdateDataSchema,
    SheetImportPlanDataSchema,
    WorkbookStructureUpdateDataSchema,
    RangeFormatUpdateDataSchema,
    ConditionalFormatPlanDataSchema,
    SheetStructureUpdateDataSchema,
    RangeSortPlanDataSchema,
    RangeFilterPlanDataSchema,
    DataValidationPlanDataSchema,
    NamedRangeUpdateDataSchema,
    RangeTransferPlanDataSchema,
    DataCleanupPlanDataSchema,
    ExternalDataPlanDataSchema,
    AnalysisReportPlanDataSchema.refine(
      (
        plan
      ): plan is Extract<AnalysisReportPlanData, { outputMode: "materialize_report" }> =>
        plan.outputMode === "materialize_report",
      {
        message: "Only materialized analysis reports are writeback eligible."
      }
    ),
    PivotTablePlanDataSchema,
    ChartPlanDataSchema,
    CompositePlanDataSchema
  ])
});

const HostPlatformSchema = z.enum(["google_sheets", "excel_windows", "excel_macos"]);
const CompletionSummarySchema = z.string().min(1).max(12000);
function stripCompletionEnvelopeInput(value: unknown): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return value;
  }

  const {
    kind: _kind,
    hostPlatform: _hostPlatform,
    summary: _summary,
    operation: _operation,
    undoReady: _undoReady,
    ...rest
  } = value as Record<string, unknown>;
  return rest;
}

function stripCompletionEnvelopeKeepingOperationInput(value: unknown): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return value;
  }

  const {
    kind: _kind,
    hostPlatform: _hostPlatform,
    summary: _summary,
    undoReady: _undoReady,
    ...rest
  } = value as Record<string, unknown>;
  return rest;
}

function stripRangeWriteCompletionEnvelopeInput(value: unknown): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return value;
  }

  const {
    kind: _kind,
    hostPlatform: _hostPlatform,
    writtenRows: _writtenRows,
    writtenColumns: _writtenColumns,
    undoReady: _undoReady,
    ...rest
  } = value as Record<string, unknown>;
  return rest;
}
const sheetStructureOperations = [
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
] as const satisfies readonly SheetStructureUpdateData["operation"][];
const SheetStructureOperationSchema = z.enum(sheetStructureOperations);
const A1ResultRangeSchema = z.string().min(1).refine(
  (value) => parseA1Range(value) !== null,
  { message: "targetRange must be a valid A1 range." }
);

const RangeWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("range_write"),
    hostPlatform: HostPlatformSchema,
    writtenRows: z.number().int().nonnegative(),
    writtenColumns: z.number().int().nonnegative(),
    undoReady: z.boolean().optional()
  }),
  z.preprocess(
    stripRangeWriteCompletionEnvelopeInput,
    z.union([SheetUpdateDataSchema, SheetImportPlanDataSchema])
  )
);

const RangeFormatWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("range_format_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema
  }),
  z.preprocess(stripCompletionEnvelopeInput, RangeFormatUpdateDataSchema)
);

const WorkbookPositionResolutionFields = {
  positionResolved: z.number().int().min(0),
  sheetCount: z.number().int().min(1)
} satisfies z.ZodRawShape;

const WorkbookStructureWritebackResultSchema = z.discriminatedUnion("operation", [
  z.object({
    kind: z.literal("workbook_structure_update"),
    hostPlatform: HostPlatformSchema,
    operation: z.literal("create_sheet"),
    sheetName: z.string().min(1),
    ...WorkbookPositionResolutionFields,
    summary: CompletionSummarySchema
  }),
  z.object({
    kind: z.literal("workbook_structure_update"),
    hostPlatform: HostPlatformSchema,
    operation: z.literal("delete_sheet"),
    sheetName: z.string().min(1),
    summary: CompletionSummarySchema
  }),
  z.object({
    kind: z.literal("workbook_structure_update"),
    hostPlatform: HostPlatformSchema,
    operation: z.literal("rename_sheet"),
    sheetName: z.string().min(1),
    newSheetName: z.string().min(1),
    summary: CompletionSummarySchema
  }),
  z.object({
    kind: z.literal("workbook_structure_update"),
    hostPlatform: HostPlatformSchema,
    operation: z.literal("duplicate_sheet"),
    sheetName: z.string().min(1),
    newSheetName: z.string().min(1),
    ...WorkbookPositionResolutionFields,
    summary: CompletionSummarySchema
  }),
  z.object({
    kind: z.literal("workbook_structure_update"),
    hostPlatform: HostPlatformSchema,
    operation: z.literal("move_sheet"),
    sheetName: z.string().min(1),
    ...WorkbookPositionResolutionFields,
    summary: CompletionSummarySchema
  }),
  z.object({
    kind: z.literal("workbook_structure_update"),
    hostPlatform: HostPlatformSchema,
    operation: z.literal("hide_sheet"),
    sheetName: z.string().min(1),
    summary: CompletionSummarySchema
  }),
  z.object({
    kind: z.literal("workbook_structure_update"),
    hostPlatform: HostPlatformSchema,
    operation: z.literal("unhide_sheet"),
    sheetName: z.string().min(1),
    summary: CompletionSummarySchema
  })
]);

const SheetStructureWritebackResultSchema = z.object({
  kind: z.literal("sheet_structure_update"),
  hostPlatform: HostPlatformSchema,
  targetSheet: z.string().min(1),
  operation: SheetStructureOperationSchema,
  startIndex: z.number().int().min(0).optional(),
  count: z.number().int().min(1).optional(),
  targetRange: A1ResultRangeSchema.optional(),
  frozenRows: z.number().int().min(0).optional(),
  frozenColumns: z.number().int().min(0).optional(),
  color: z.string().regex(/^#[0-9a-fA-F]{6}$/).optional(),
  summary: CompletionSummarySchema
}).superRefine((data, ctx) => {
  switch (data.operation) {
    case "insert_rows":
    case "delete_rows":
    case "hide_rows":
    case "unhide_rows":
    case "group_rows":
    case "ungroup_rows":
    case "insert_columns":
    case "delete_columns":
    case "hide_columns":
    case "unhide_columns":
    case "group_columns":
    case "ungroup_columns":
      if (data.startIndex === undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "startIndex is required for row and column structure updates.",
          path: ["startIndex"]
        });
      }
      if (data.count === undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "count is required for row and column structure updates.",
          path: ["count"]
        });
      }
      return;
    case "merge_cells":
    case "unmerge_cells":
    case "autofit_rows":
    case "autofit_columns":
      if (data.targetRange === undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "targetRange is required for range-based structure updates.",
          path: ["targetRange"]
        });
      }
      return;
    case "freeze_panes":
    case "unfreeze_panes":
      if (data.frozenRows === undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "frozenRows is required for pane freeze updates.",
          path: ["frozenRows"]
        });
      }
      if (data.frozenColumns === undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "frozenColumns is required for pane freeze updates.",
          path: ["frozenColumns"]
        });
      }
      return;
    case "set_sheet_tab_color":
      if (data.color === undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "color is required for sheet tab color updates.",
          path: ["color"]
        });
      }
      return;
  }
});

const RangeSortWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("range_sort"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema,
    undoReady: z.boolean().optional()
  }),
  z.preprocess(stripCompletionEnvelopeInput, RangeSortPlanDataSchema)
);

const RangeFilterWritebackResultSchema = z.object({
  kind: z.literal("range_filter"),
  hostPlatform: HostPlatformSchema,
  summary: CompletionSummarySchema
}).and(
  z.preprocess(stripCompletionEnvelopeInput, RangeFilterPlanDataSchema)
);

const DataValidationWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("data_validation_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema
  }),
  z.preprocess(stripCompletionEnvelopeInput, DataValidationPlanDataSchema)
);

const ConditionalFormatWritebackResultSchema = z.object({
  kind: z.literal("conditional_format_update"),
  hostPlatform: HostPlatformSchema,
  summary: CompletionSummarySchema
}).and(
  z.preprocess(stripCompletionEnvelopeInput, ConditionalFormatPlanDataSchema)
);

const NamedRangeWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("named_range_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema
  }),
  z.preprocess(stripCompletionEnvelopeKeepingOperationInput, NamedRangeUpdateDataSchema)
);

const RangeTransferWritebackResultSchema = RangeTransferUpdateDataSchema.extend({
  kind: z.literal("range_transfer_update"),
  hostPlatform: HostPlatformSchema
});

const DataCleanupWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("data_cleanup_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema,
    undoReady: z.boolean().optional()
  }),
  z.preprocess(stripCompletionEnvelopeKeepingOperationInput, DataCleanupPlanDataSchema)
);

const AnalysisReportWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("analysis_report_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema,
    undoReady: z.boolean().optional()
  }),
  z.preprocess(
    stripCompletionEnvelopeInput,
    AnalysisReportPlanDataSchema.refine(
      (
        plan
      ): plan is Extract<AnalysisReportPlanData, { outputMode: "materialize_report" }> =>
        plan.outputMode === "materialize_report",
      {
        message: "Only materialized analysis reports are writeback eligible."
      }
    )
  )
);

const ExternalDataWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("external_data_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema
  }),
  z.preprocess(stripCompletionEnvelopeInput, ExternalDataPlanDataSchema)
);

const PivotTableWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("pivot_table_update"),
    operation: z.literal("pivot_table_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema
  }),
  z.preprocess(stripCompletionEnvelopeInput, PivotTablePlanDataSchema)
);

const ChartWritebackResultSchema = z.intersection(
  z.object({
    kind: z.literal("chart_update"),
    operation: z.literal("chart_update"),
    hostPlatform: HostPlatformSchema,
    summary: CompletionSummarySchema
  }),
  z.preprocess(stripCompletionEnvelopeInput, ChartPlanDataSchema)
);

const CompositeWritebackResultSchema = CompositeUpdateDataSchema.extend({
  kind: z.literal("composite_update"),
  hostPlatform: HostPlatformSchema
});

const CompletionRequestSchema = z.object({
  requestId: z.string().min(1),
  runId: z.string().min(1),
  workbookSessionKey: z.string().min(1).max(256).optional(),
  approvalToken: z.string().min(1),
  planDigest: z.string().min(1),
  result: z.union([
    RangeWritebackResultSchema,
    RangeFormatWritebackResultSchema,
    WorkbookStructureWritebackResultSchema,
    SheetStructureWritebackResultSchema,
    RangeSortWritebackResultSchema,
    RangeFilterWritebackResultSchema,
    DataValidationWritebackResultSchema,
    ConditionalFormatWritebackResultSchema,
    NamedRangeWritebackResultSchema,
    RangeTransferWritebackResultSchema,
    DataCleanupWritebackResultSchema,
    AnalysisReportWritebackResultSchema,
    ExternalDataWritebackResultSchema,
    PivotTableWritebackResultSchema,
    ChartWritebackResultSchema,
    CompositeWritebackResultSchema
  ])
});

type ApprovalPlan = z.infer<typeof ApprovalRequestSchema>["plan"];
type CompletionResult = z.infer<typeof CompletionRequestSchema>["result"];
type WorkbookStructureCompletionResult = z.infer<typeof WorkbookStructureWritebackResultSchema>;
type SheetStructureCompletionResult = z.infer<typeof SheetStructureWritebackResultSchema>;
type RangeTransferCompletionResult = z.infer<typeof RangeTransferWritebackResultSchema>;
type CompositeCompletionResult = z.infer<typeof CompositeWritebackResultSchema>;
type MaterializedAnalysisReportPlan = Extract<AnalysisReportPlanData, { outputMode: "materialize_report" }>;
type ResolvedMaterializedAnalysisReportPlan =
  MaterializedAnalysisReportPlan & { targetSheet: string; targetRange: string };

type RouteErrorPayload = {
  error: {
    code: string;
    message: string;
    userAction?: string;
    issues?: Array<{ path: string; message: string }>;
  };
};

function formatIssues(issues: z.ZodIssue[]): Array<{ path: string; message: string }> {
  return issues.map((issue) => ({
    path: issue.path.join("."),
    message: issue.message
  }));
}

function invalidWritebackRequest(
  message: string,
  userAction: string,
  issues?: z.ZodIssue[]
): RouteErrorPayload {
  return {
    error: {
      code: "INVALID_REQUEST",
      message,
      userAction,
      ...(issues ? { issues: formatIssues(issues) } : {})
    }
  };
}

function formatWritebackRouteError(error: unknown): {
  status: number;
  body: RouteErrorPayload;
} {
  const message = error instanceof Error ? error.message : "Write-back request failed.";

  if (error instanceof FreshDryRunRequiredError) {
    return {
      status: 409,
      body: {
        error: {
          code: "STALE_PREVIEW",
          message: "That workflow preview is stale and must be regenerated before approval.",
          userAction: "Preview the workflow again, then confirm the latest version."
        }
      }
    };
  }

  if (message === "Run not found.") {
    return {
      status: 404,
      body: {
        error: {
          code: "RUN_NOT_FOUND",
          message: "That Hermes request is no longer available.",
          userAction: "Send the request again from the spreadsheet, then retry the approval or apply step."
        }
      }
    };
  }

  if (message === "Request ID does not match the stored run.") {
    return {
      status: 409,
      body: {
        error: {
          code: "STALE_REQUEST",
          message: "This approval no longer matches the current Hermes request.",
          userAction: "Refresh the chat state and ask Hermes to prepare the update again."
        }
      }
    };
  }

  if (message === "Destructive confirmation required.") {
    return {
      status: 400,
      body: {
        error: {
          code: "DESTRUCTIVE_CONFIRMATION_REQUIRED",
          message: "This update needs an explicit destructive confirmation before it can run.",
          userAction: "Confirm the destructive step, then retry the approval."
        }
      }
    };
  }

  if (
    message === "Approved plan does not match the stored Hermes response." ||
    message === "Writeback result does not match the approved plan family." ||
    message === "Writeback result does not match the approved plan details."
  ) {
    return {
      status: 409,
      body: {
        error: {
          code: "STALE_APPROVAL",
          message: "The approved update no longer matches the current Hermes plan.",
          userAction: "Refresh the spreadsheet state and ask Hermes to prepare a fresh update."
        }
      }
    };
  }

  if (message === "Writeback approval not found.") {
    return {
      status: 409,
      body: {
        error: {
          code: "APPROVAL_NOT_FOUND",
          message: "This update is no longer awaiting approval.",
          userAction: "Ask Hermes for a fresh writeback proposal, then confirm it again."
        }
      }
    };
  }

  if (message === "Writeback already completed for this run.") {
    return {
      status: 409,
      body: {
        error: {
          code: "ALREADY_COMPLETED",
          message: "This update was already applied.",
          userAction: "Refresh the sheet. If you need another change, ask Hermes for a new update."
        }
      }
    };
  }

  if (message === "Writeback approval already pending for this run.") {
    return {
      status: 409,
      body: {
        error: {
          code: "APPROVAL_ALREADY_PENDING",
          message: "This update is already awaiting completion in another spreadsheet session.",
          userAction: "Finish or retry the existing confirm flow, or ask Hermes for a fresh update from this sheet."
        }
      }
    };
  }

  if (
    message === "Approval token already consumed." ||
    message === "Approval token does not match the current writeback approval." ||
    message === "Writeback completion does not match the approved workbook session." ||
    message === "Invalid approval token."
  ) {
    return {
      status: message === "Invalid approval token." ? 403 : 409,
      body: {
        error: {
          code: "INVALID_APPROVAL",
          message: "This approval is no longer valid for the current update.",
          userAction: "Refresh the chat state and confirm the latest Hermes update again."
        }
      }
    };
  }

  if (message === "Approval token expired.") {
    return {
      status: 403,
      body: {
        error: {
          code: "APPROVAL_EXPIRED",
          message: "This approval expired before it was applied.",
          userAction: "Ask Hermes to generate a fresh update, then confirm it again."
        }
      }
    };
  }

  return {
    status: 500,
    body: {
      error: {
        code: "INTERNAL_ERROR",
        message: "The gateway couldn't complete that write-back request.",
        userAction: "Retry the action. If it keeps failing, refresh the spreadsheet session and try again."
      }
    }
  };
}

function assertMaterializedAnalysisReportPlan(
  plan: AnalysisReportPlanData
): asserts plan is ResolvedMaterializedAnalysisReportPlan {
  if (plan.outputMode !== "materialize_report" ||
    typeof plan.targetSheet !== "string" ||
    typeof plan.targetRange !== "string") {
    throw new Error("Materialized analysis report plan is missing target location.");
  }
}

function columnLettersToNumber(columnLetters: string): number {
  let column = 0;
  for (const character of columnLetters.trim().toUpperCase()) {
    column = (column * 26) + (character.charCodeAt(0) - 64);
  }
  return column;
}

function columnNumberToLetters(columnNumber: number): string {
  let value = Number(columnNumber);
  let letters = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }

  return letters;
}

function parseA1Anchor(range: string): { startColumn: number; startRow: number } {
  const normalized = String(range).trim().toUpperCase().replaceAll("$", "");
  const withoutSheet = normalized.includes("!") ? normalized.split("!").pop() || normalized : normalized;
  const anchor = withoutSheet.split(":")[0];
  const match = anchor.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Unsupported analysis report targetRange anchor: ${range}`);
  }

  return {
    startColumn: columnLettersToNumber(match[1]),
    startRow: Number(match[2])
  };
}

function buildAnalysisReportMatrixShape(plan: ResolvedMaterializedAnalysisReportPlan): { rows: number; columns: number } {
  return {
    rows: 4 + plan.sections.length,
    columns: 4
  };
}

function resolveAnalysisReportTargetRange(plan: ResolvedMaterializedAnalysisReportPlan): string {
  const anchor = parseA1Anchor(plan.targetRange);
  const shape = buildAnalysisReportMatrixShape(plan);
  const endRow = anchor.startRow + shape.rows - 1;
  const endColumn = anchor.startColumn + shape.columns - 1;
  const startCell = `${columnNumberToLetters(anchor.startColumn)}${anchor.startRow}`;
  const endCell = `${columnNumberToLetters(endColumn)}${endRow}`;

  return startCell === endCell ? startCell : `${startCell}:${endCell}`;
}

function normalizeMaterializedAnalysisReportPlan(
  plan: AnalysisReportPlanData
): ResolvedMaterializedAnalysisReportPlan {
  assertMaterializedAnalysisReportPlan(plan);
  return {
    ...plan,
    targetRange: resolveAnalysisReportTargetRange(plan)
  };
}

function normalizeApprovalPlan(plan: ApprovalPlan): ApprovalPlan {
  if ("steps" in plan) {
    return normalizeCompositePlanForDigest(plan);
  }

  if ("outputMode" in plan && plan.outputMode === "materialize_report") {
    return normalizeMaterializedAnalysisReportPlan(plan);
  }

  return plan;
}

function buildA1RangeFromBounds(bounds: {
  column: number;
  row: number;
  endColumn: number;
  endRow: number;
}): string {
  const startCell = `${columnNumberToLetters(bounds.column)}${bounds.row}`;
  const endCell = `${columnNumberToLetters(bounds.endColumn)}${bounds.endRow}`;
  return startCell === endCell ? startCell : `${startCell}:${endCell}`;
}

function a1RangesEqual(left: string, right: string): boolean {
  const leftRange = parseA1Range(left);
  const rightRange = parseA1Range(right);
  if (!leftRange || !rightRange) {
    return false;
  }

  return leftRange.column === rightRange.column &&
    leftRange.row === rightRange.row &&
    leftRange.endColumn === rightRange.endColumn &&
    leftRange.endRow === rightRange.endRow;
}

function getRangeTransferShape(plan: RangeTransferPlanData): { rows: number; columns: number } {
  const parsedSourceRange = parseA1Range(plan.sourceRange);
  if (!parsedSourceRange) {
    throw new Error(`Unsupported range transfer sourceRange: ${plan.sourceRange}`);
  }

  return {
    rows: plan.transpose ? parsedSourceRange.columns : parsedSourceRange.rows,
    columns: plan.transpose ? parsedSourceRange.rows : parsedSourceRange.columns
  };
}

function resolveExpectedRangeTransferTargetRange(plan: RangeTransferPlanData): string {
  const parsedTargetRange = parseA1Range(plan.targetRange);
  if (!parsedTargetRange) {
    throw new Error(`Unsupported range transfer targetRange: ${plan.targetRange}`);
  }

  const shape = getRangeTransferShape(plan);
  if (parsedTargetRange.rows === shape.rows && parsedTargetRange.columns === shape.columns) {
    return buildA1RangeFromBounds(parsedTargetRange);
  }

  if (parsedTargetRange.rows === 1 && parsedTargetRange.columns === 1) {
    return buildA1RangeFromBounds({
      column: parsedTargetRange.column,
      row: parsedTargetRange.row,
      endColumn: parsedTargetRange.column + shape.columns - 1,
      endRow: parsedTargetRange.row + shape.rows - 1
    });
  }

  throw new Error("Approved range transfer targetRange does not match an exact or anchor destination.");
}

function matchesExpectedAppendTargetRange(
  plan: Extract<RangeTransferPlanData, { operation: "append" }>,
  actualTargetRange: string
): boolean {
  const approvedTargetRange = parseA1Range(plan.targetRange);
  const actualRange = parseA1Range(actualTargetRange);
  if (!approvedTargetRange || !actualRange) {
    return false;
  }

  const shape = getRangeTransferShape(plan);
  if (approvedTargetRange.columns !== shape.columns ||
    actualRange.rows !== shape.rows ||
    actualRange.columns !== shape.columns) {
    return false;
  }

  if (actualRange.column !== approvedTargetRange.column ||
    actualRange.endColumn !== approvedTargetRange.column + shape.columns - 1) {
    return false;
  }

  if (approvedTargetRange.rows === 1 && shape.rows > 1) {
    return actualRange.row === approvedTargetRange.row &&
      actualRange.endRow === approvedTargetRange.row + shape.rows - 1;
  }

  return actualRange.row >= approvedTargetRange.row &&
    actualRange.endRow <= approvedTargetRange.endRow;
}

function assertSheetStructureCompletionMatchesApprovedPlan(
  plan: SheetStructureUpdateData,
  result: SheetStructureCompletionResult
): void {
  if (
    result.kind !== "sheet_structure_update" ||
    result.operation !== plan.operation ||
    result.targetSheet !== plan.targetSheet
  ) {
    throw new Error("Writeback result does not match the approved plan details.");
  }

  switch (plan.operation) {
    case "insert_rows":
    case "delete_rows":
    case "hide_rows":
    case "unhide_rows":
    case "group_rows":
    case "ungroup_rows":
    case "insert_columns":
    case "delete_columns":
    case "hide_columns":
    case "unhide_columns":
    case "group_columns":
    case "ungroup_columns":
      if (result.startIndex !== plan.startIndex || result.count !== plan.count) {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      return;
    case "merge_cells":
    case "unmerge_cells":
    case "autofit_rows":
    case "autofit_columns":
      if (!result.targetRange || !a1RangesEqual(result.targetRange, plan.targetRange)) {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      return;
    case "freeze_panes":
    case "unfreeze_panes":
      if (result.frozenRows !== plan.frozenRows || result.frozenColumns !== plan.frozenColumns) {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      return;
    case "set_sheet_tab_color":
      if (
        typeof result.color !== "string" ||
        result.color.toUpperCase() !== plan.color.toUpperCase()
      ) {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      return;
  }
}

function assertRangeTransferCompletionMatchesApprovedPlan(
  plan: RangeTransferPlanData,
  result: RangeTransferCompletionResult
): void {
  if (
    result.kind !== "range_transfer_update" ||
    result.sourceSheet !== plan.sourceSheet ||
    result.sourceRange !== plan.sourceRange ||
    result.targetSheet !== plan.targetSheet ||
    result.transferOperation !== plan.operation ||
    result.pasteMode !== plan.pasteMode ||
    result.transpose !== plan.transpose
  ) {
    throw new Error("Writeback result does not match the approved plan details.");
  }

  if (plan.operation === "append") {
    if (!matchesExpectedAppendTargetRange(plan, result.targetRange)) {
      throw new Error("Writeback result does not match the approved plan details.");
    }
    return;
  }

  if (!a1RangesEqual(result.targetRange, resolveExpectedRangeTransferTargetRange(plan))) {
    throw new Error("Writeback result does not match the approved plan details.");
  }
}

function assertCompositeCompletionMatchesApprovedPlan(
  plan: CompositePlanData,
  result: CompositeCompletionResult
): void {
  if (result.kind !== "composite_update") {
    throw new Error("Writeback result does not match the approved plan details.");
  }

  if (result.stepResults.length !== plan.steps.length) {
    throw new Error("Writeback result does not match the approved plan details.");
  }

  const completedSteps = new Set<string>();
  const failedSteps = new Set<string>();
  const skippedSteps = new Set<string>();
  let halted = false;

  for (let index = 0; index < plan.steps.length; index += 1) {
    const expectedStep = plan.steps[index];
    const stepResult = result.stepResults[index];
    if (!stepResult || stepResult.stepId !== expectedStep.stepId) {
      throw new Error("Writeback result does not match the approved plan details.");
    }

    if (halted) {
      if (stepResult.status !== "skipped") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      skippedSteps.add(stepResult.stepId);
      continue;
    }

    const dependencyStatuses = expectedStep.dependsOn.map((dependency) => {
      if (completedSteps.has(dependency)) {
        return "completed";
      }
      if (failedSteps.has(dependency)) {
        return "failed";
      }
      if (skippedSteps.has(dependency)) {
        return "skipped";
      }
      return "pending";
    });

    if (dependencyStatuses.includes("pending")) {
      throw new Error("Writeback result does not match the approved plan details.");
    }

    if (dependencyStatuses.includes("failed") || dependencyStatuses.includes("skipped")) {
      if (stepResult.status !== "skipped") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      skippedSteps.add(stepResult.stepId);
      continue;
    }

    switch (stepResult.status) {
      case "completed":
        completedSteps.add(stepResult.stepId);
        break;
      case "failed":
        failedSteps.add(stepResult.stepId);
        if (!expectedStep.continueOnError) {
          halted = true;
        }
        break;
      case "skipped":
        throw new Error("Writeback result does not match the approved plan details.");
    }
  }
}

function getCompositeExecutionOutcome(
  result: z.infer<typeof CompositeWritebackResultSchema>
): { status: "completed" | "failed"; undoEligible: boolean } {
  const allCompleted = result.stepResults.every((stepResult) => stepResult.status === "completed");
  return allCompleted
    ? {
      status: "completed",
      // Composite host results do not yet carry an exact rollback snapshot contract.
      // Fail closed so history does not promise undo for workflows the host cannot restore exactly.
      undoEligible: false
    }
    : { status: "failed", undoEligible: false };
}

function canonicalizeForDigest(value: unknown): unknown {
  return canonicalizePlan(value);
}

function clampHistorySummary(summary: string): string {
  return summary.length > 12000 ? summary.slice(0, 12000) : summary;
}

function assertCanonicalPlanSpecMatchesApprovedPlan(plan: unknown, resultSpec: unknown): void {
  if (digestCanonicalPlan(plan) !== digestCanonicalPlan(resultSpec)) {
    throw new Error("Writeback result does not match the approved plan details.");
  }
}

function extractAnalysisReportExecutionSpec(
  plan: ResolvedMaterializedAnalysisReportPlan
): Pick<
  ResolvedMaterializedAnalysisReportPlan,
  "sourceSheet" | "sourceRange" | "outputMode" | "targetSheet" | "targetRange" | "sections"
> {
  return {
    sourceSheet: plan.sourceSheet,
    sourceRange: plan.sourceRange,
    outputMode: plan.outputMode,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    sections: plan.sections
  };
}

function resolveWorkbookStructurePosition(
  requestedPosition: number | "start" | "end" | undefined,
  sheetCount: number
): number {
  if (requestedPosition === "start") {
    return 0;
  }

  if (requestedPosition === "end" || requestedPosition === undefined) {
    return Math.max(0, sheetCount - 1);
  }

  return Math.max(0, Math.min(requestedPosition, Math.max(0, sheetCount - 1)));
}

function stripCompletionEnvelope<
  T extends {
    kind: string;
    hostPlatform: z.infer<typeof HostPlatformSchema>;
    summary: string;
    operation?: string;
    undoReady?: boolean;
  }
>(result: T): Omit<T, "kind" | "hostPlatform" | "summary" | "operation" | "undoReady"> {
  const {
    kind: _kind,
    hostPlatform: _hostPlatform,
    summary: _summary,
    operation: _operation,
    undoReady: _undoReady,
    ...rest
  } = result;
  return rest;
}

function stripCompletionEnvelopeKeepingOperation<
  T extends {
    kind: string;
    hostPlatform: z.infer<typeof HostPlatformSchema>;
    summary: string;
    undoReady?: boolean;
  }
>(result: T): Omit<T, "kind" | "hostPlatform" | "summary" | "undoReady"> {
  const {
    kind: _kind,
    hostPlatform: _hostPlatform,
    summary: _summary,
    undoReady: _undoReady,
    ...rest
  } = result;
  return rest;
}

function stripRangeWriteCompletionEnvelope<
  T extends {
    kind: "range_write";
    hostPlatform: z.infer<typeof HostPlatformSchema>;
    writtenRows: number;
    writtenColumns: number;
    undoReady?: boolean;
  }
>(result: T): Omit<T, "kind" | "hostPlatform" | "writtenRows" | "writtenColumns" | "undoReady"> {
  const {
    kind: _kind,
    hostPlatform: _hostPlatform,
    writtenRows: _writtenRows,
    writtenColumns: _writtenColumns,
    undoReady: _undoReady,
    ...rest
  } = result;
  return rest;
}

function completionResultsMatch(
  left: CompletionResult | undefined,
  right: CompletionResult
): boolean {
  if (!left) {
    return false;
  }

  return digestCanonicalPlan(left) === digestCanonicalPlan(right);
}

function digestApprovalPlan(plan: ApprovalPlan): string {
  return digestCanonicalPlan(normalizeApprovalPlan(plan));
}

function getExpectedResultKind(
  response: HermesResponse | undefined
): CompletionResult["kind"] | undefined {
  switch (response?.type) {
    case "sheet_update":
    case "sheet_import_plan":
      return "range_write";
    case "external_data_plan":
      return "external_data_update";
    case "range_format_update":
      return "range_format_update";
    case "conditional_format_plan":
      return "conditional_format_update";
    case "workbook_structure_update":
      return "workbook_structure_update";
    case "sheet_structure_update":
      return "sheet_structure_update";
    case "range_sort_plan":
      return "range_sort";
    case "range_filter_plan":
      return "range_filter";
    case "data_validation_plan":
      return "data_validation_update";
    case "named_range_update":
      return "named_range_update";
    case "range_transfer_plan":
      return "range_transfer_update";
    case "data_cleanup_plan":
      return "data_cleanup_update";
    case "analysis_report_plan":
      return response.data.outputMode === "materialize_report"
        ? "analysis_report_update"
        : undefined;
    case "pivot_table_plan":
      return "pivot_table_update";
    case "chart_plan":
      return "chart_update";
    case "composite_plan":
      return "composite_update";
    default:
      return undefined;
  }
}

function getStoredPlan(run: ReturnType<TraceBus["getRun"]>): ApprovalPlan | undefined {
  if (!run?.response) {
    return undefined;
  }

  switch (run.response.type) {
    case "sheet_update":
    case "sheet_import_plan":
    case "external_data_plan":
    case "workbook_structure_update":
    case "range_format_update":
    case "sheet_structure_update":
    case "range_sort_plan":
    case "range_filter_plan":
    case "data_validation_plan":
    case "named_range_update":
    case "conditional_format_plan":
    case "range_transfer_plan":
    case "data_cleanup_plan":
    case "pivot_table_plan":
    case "chart_plan":
    case "composite_plan":
      return run.response.data;
    case "analysis_report_plan":
      return run.response.data.outputMode === "materialize_report"
        ? run.response.data
        : undefined;
    default:
      return undefined;
  }
}

function requiresDestructiveConfirmation(plan: ApprovalPlan): boolean {
  if ("confirmationLevel" in plan && plan.confirmationLevel === "destructive") {
    return true;
  }

  if ("steps" in plan) {
    return plan.steps.some((step) =>
      requiresDestructiveConfirmation(step.plan as ApprovalPlan)
    );
  }

  return false;
}

function getPlanTypeFromResponse(response: HermesResponse | undefined): string {
  return response?.type ?? "unknown_plan";
}

function buildHistorySummary(
  planType: string,
  result?: CompletionResult
): string {
  if (result && "summary" in result && typeof result.summary === "string") {
    return clampHistorySummary(result.summary);
  }

  if (result?.kind === "range_write") {
    return `Applied range write to ${result.targetSheet}!${result.targetRange}.`;
  }

  return `Approved ${planType}.`;
}

function buildExecutionId(runId: string): string {
  return `exec_${randomUUID()}_${runId}`.slice(0, 128);
}

function isCompositePlan(plan: ApprovalPlan | undefined): plan is CompositePlanData {
  return Boolean(plan && "steps" in plan);
}

function hasNonEmptyNotesMatrix(notes: unknown): boolean {
  return Array.isArray(notes) && notes.some((row) =>
    Array.isArray(row) && row.some((cell) => cell !== null && cell !== undefined && String(cell).length > 0)
  );
}

function isPlanReversible(plan: ApprovalPlan | undefined): boolean {
  if (!plan) {
    return false;
  }

  return (
    "targetSheet" in plan &&
    typeof plan.targetSheet === "string" &&
    "targetRange" in plan &&
    typeof plan.targetRange === "string" &&
    (
      ("shape" in plan && !("notes" in plan && hasNonEmptyNotesMatrix(plan.notes))) ||
      ("outputMode" in plan && plan.outputMode === "materialize_report") ||
      ("operation" in plan && (
        plan.operation === "normalize_case" ||
        plan.operation === "trim_whitespace" ||
        plan.operation === "remove_blank_rows" ||
        plan.operation === "remove_duplicate_rows" ||
        plan.operation === "split_column" ||
        plan.operation === "join_columns" ||
        plan.operation === "fill_down" ||
        plan.operation === "standardize_format"
      )) ||
      ("keys" in plan && Array.isArray(plan.keys))
    )
  );
}

function isCompletionUndoReady(result: CompletionResult): boolean {
  switch (result.kind) {
    case "range_write":
    case "range_sort":
    case "data_cleanup_update":
    case "analysis_report_update":
      return result.undoReady === true;
    default:
      return false;
  }
}

function assertMatchingRequestId(
  run: ReturnType<TraceBus["getRun"]>,
  requestId: string
): void {
  if (run?.requestId && run.requestId !== requestId) {
    throw new Error("Request ID does not match the stored run.");
  }
}

function assertCurrentWritebackApproval(
  run: ReturnType<TraceBus["getRun"]>,
  approvalToken: string,
  planDigest: string
): void {
  if (!run?.writeback) {
    throw new Error("Writeback approval not found.");
  }

  if (
    run.writeback.approvalToken !== approvalToken ||
    run.writeback.approvedPlanDigest !== planDigest
  ) {
    throw new Error("Approval token does not match the current writeback approval.");
  }
}

function throwWritebackPlanDetailsMismatch(): never {
  throw new Error("Writeback result does not match the approved plan details.");
}

function assertWorkbookStructureCompletionResult(
  result: CompletionResult
): asserts result is WorkbookStructureCompletionResult {
  if (result.kind !== "workbook_structure_update") {
    throwWritebackPlanDetailsMismatch();
  }
}

function assertSheetStructureCompletionResult(
  result: CompletionResult
): asserts result is SheetStructureCompletionResult {
  if (result.kind !== "sheet_structure_update") {
    throwWritebackPlanDetailsMismatch();
  }
}

function assertRangeTransferCompletionResult(
  result: CompletionResult
): asserts result is RangeTransferCompletionResult {
  if (result.kind !== "range_transfer_update") {
    throwWritebackPlanDetailsMismatch();
  }
}

function assertCompositeCompletionResult(
  result: CompletionResult
): asserts result is CompositeCompletionResult {
  if (result.kind !== "composite_update") {
    throwWritebackPlanDetailsMismatch();
  }
}

function assertCompletionMatchesApprovedPlan(
  response: HermesResponse | undefined,
  result: CompletionResult
): void {
  switch (response?.type) {
    case "sheet_update":
    case "sheet_import_plan":
      {
      const expectedTargetRange = parseA1Range(response.data.targetRange);
      if (!expectedTargetRange) {
        throw new Error("Approved plan does not have a valid target rectangle.");
      }
      if (
        result.kind !== "range_write" ||
        result.targetSheet !== response.data.targetSheet ||
        result.targetRange !== response.data.targetRange ||
        result.writtenRows !== expectedTargetRange.rows ||
        result.writtenColumns !== expectedTargetRange.columns
      ) {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripRangeWriteCompletionEnvelope(result)
      );
      return;
      }
    case "external_data_plan":
      if (result.kind !== "external_data_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result) as ExternalDataPlanData
      );
      return;
    case "range_format_update":
      if (result.kind !== "range_format_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result)
      );
      return;
    case "workbook_structure_update":
      assertWorkbookStructureCompletionResult(result);
      if (result.operation !== response.data.operation) {
        throwWritebackPlanDetailsMismatch();
      }
      switch (response.data.operation) {
        case "create_sheet":
          if (result.operation !== "create_sheet") {
            throwWritebackPlanDetailsMismatch();
          }
          if (
            result.sheetName !== response.data.sheetName ||
            result.positionResolved !== resolveWorkbookStructurePosition(
              response.data.position,
              result.sheetCount
            )
          ) {
            throwWritebackPlanDetailsMismatch();
          }
          return;
        case "delete_sheet":
        case "hide_sheet":
        case "unhide_sheet":
          if (result.sheetName !== response.data.sheetName) {
            throwWritebackPlanDetailsMismatch();
          }
          return;
        case "rename_sheet":
          if (result.operation !== "rename_sheet") {
            throwWritebackPlanDetailsMismatch();
          }
          if (
            result.sheetName !== response.data.sheetName ||
            result.newSheetName !== response.data.newSheetName
          ) {
            throwWritebackPlanDetailsMismatch();
          }
          return;
        case "duplicate_sheet":
          if (result.operation !== "duplicate_sheet") {
            throwWritebackPlanDetailsMismatch();
          }
          if (
            result.sheetName !== response.data.sheetName ||
            result.positionResolved !== resolveWorkbookStructurePosition(
              response.data.position,
              result.sheetCount
            )
          ) {
            throwWritebackPlanDetailsMismatch();
          }
          if (
            typeof response.data.newSheetName === "string" &&
            result.newSheetName !== response.data.newSheetName
          ) {
            throwWritebackPlanDetailsMismatch();
          }
          if (typeof result.newSheetName !== "string" || result.newSheetName.trim().length === 0) {
            throwWritebackPlanDetailsMismatch();
          }
          return;
        case "move_sheet":
          if (result.operation !== "move_sheet") {
            throwWritebackPlanDetailsMismatch();
          }
          if (
            result.sheetName !== response.data.sheetName ||
            result.positionResolved !== resolveWorkbookStructurePosition(
              response.data.position,
              result.sheetCount
            )
          ) {
            throwWritebackPlanDetailsMismatch();
          }
          return;
        default:
          return;
      }
    case "sheet_structure_update":
      assertSheetStructureCompletionResult(result);
      assertSheetStructureCompletionMatchesApprovedPlan(response.data, result);
      return;
    case "range_sort_plan":
      if (result.kind !== "range_sort") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result)
      );
      return;
    case "range_filter_plan":
      if (result.kind !== "range_filter") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result)
      );
      return;
    case "data_validation_plan":
      if (result.kind !== "data_validation_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result)
      );
      return;
    case "conditional_format_plan":
      if (result.kind !== "conditional_format_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result)
      );
      return;
    case "named_range_update":
      if (result.kind !== "named_range_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelopeKeepingOperation(result)
      );
      return;
    case "range_transfer_plan":
      assertRangeTransferCompletionResult(result);
      assertRangeTransferCompletionMatchesApprovedPlan(response.data, result);
      return;
    case "data_cleanup_plan":
      if (result.kind !== "data_cleanup_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelopeKeepingOperation(result)
      );
      return;
    case "analysis_report_plan":
      {
        const normalizedPlan = normalizeMaterializedAnalysisReportPlan(response.data);
        if (
          response.data.outputMode !== "materialize_report" ||
          result.kind !== "analysis_report_update"
        ) {
          throw new Error("Writeback result does not match the approved plan details.");
        }
        assertCanonicalPlanSpecMatchesApprovedPlan(
          extractAnalysisReportExecutionSpec(normalizedPlan),
          extractAnalysisReportExecutionSpec(stripCompletionEnvelope(result) as ResolvedMaterializedAnalysisReportPlan)
        );
        return;
      }
    case "pivot_table_plan":
      if (result.kind !== "pivot_table_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result)
      );
      return;
    case "chart_plan":
      if (result.kind !== "chart_update") {
        throw new Error("Writeback result does not match the approved plan details.");
      }
      assertCanonicalPlanSpecMatchesApprovedPlan(
        response.data,
        stripCompletionEnvelope(result)
      );
      return;
    case "composite_plan":
      assertCompositeCompletionResult(result);
      assertCompositeCompletionMatchesApprovedPlan(response.data, result);
      return;
    default:
      return;
  }
}

export function approveWriteback(input: {
  requestId: string;
  runId: string;
  workbookSessionKey?: string;
  plan: ApprovalPlan;
  destructiveConfirmation?: z.infer<typeof ApprovalRequestSchema>["destructiveConfirmation"];
  traceBus: TraceBus;
  executionLedger?: ExecutionLedger;
  config: GatewayConfig;
}) {
  const executionLedger = input.executionLedger ?? new ExecutionLedger();
  const run = input.traceBus.getRun(input.runId);
  if (!run) {
    throw new Error("Run not found.");
  }
  assertMatchingRequestId(run, input.requestId);

  const storedPlan = getStoredPlan(run);
  if (!storedPlan || digestApprovalPlan(storedPlan) !== digestApprovalPlan(input.plan)) {
    throw new Error("Approved plan does not match the stored Hermes response.");
  }

  if (requiresDestructiveConfirmation(input.plan) && !input.destructiveConfirmation?.confirmed) {
    throw new Error("Destructive confirmation required.");
  }

  const planDigest = digestApprovalPlan(input.plan);
  const workbookSessionKey = input.workbookSessionKey ?? `run::${input.runId}`;
  const existingWriteback = run.writeback;
  if (existingWriteback) {
    if (existingWriteback.completedAt) {
      throw new Error("Writeback already completed for this run.");
    }

    if (
      existingWriteback.approvedPlanDigest === planDigest &&
      existingWriteback.workbookSessionKey === workbookSessionKey
    ) {
      return {
        requestId: input.requestId,
        runId: input.runId,
        executionId: existingWriteback.executionId,
        approvalToken: existingWriteback.approvalToken,
        planDigest,
        approvedAt: existingWriteback.approvedAt
      };
    }

    throw new Error("Writeback approval already pending for this run.");
  }

  if ("steps" in input.plan) {
    executionLedger.assertFreshDryRun({
      workbookSessionKey,
      planDigest,
      required: input.plan.dryRunRequired
    });
  }

  const issuedAt = executionLedger.isoTimestamp();
  const executionId = buildExecutionId(input.runId);
  const approvalToken = createApprovalToken({
    requestId: input.requestId,
    runId: input.runId,
    planDigest,
    issuedAt,
    secret: input.config.approvalSecret
  });

  input.traceBus.recordWritebackApproval({
    runId: input.runId,
    executionId,
    workbookSessionKey,
    approvedAt: issuedAt,
    approvedPlanDigest: planDigest,
    approvalToken,
    destructiveConfirmation: input.destructiveConfirmation
  });
  executionLedger.recordApproved({
    executionId,
    workbookSessionKey,
    requestId: input.requestId,
    runId: input.runId,
    planType: getPlanTypeFromResponse(run.response),
    planDigest,
    status: "approved",
    timestamp: issuedAt,
    reversible: isPlanReversible(input.plan),
    undoEligible: false,
    redoEligible: false,
    summary: buildHistorySummary(getPlanTypeFromResponse(run.response))
  });

  return {
    requestId: input.requestId,
    runId: input.runId,
    executionId,
    approvalToken,
    planDigest,
    approvedAt: issuedAt
  };
}

export function completeWriteback(input: {
  requestId: string;
  runId: string;
  workbookSessionKey?: string;
  approvalToken: string;
  planDigest: string;
  result: CompletionResult;
  traceBus: TraceBus;
  executionLedger?: ExecutionLedger;
  config: GatewayConfig;
}) {
  const executionLedger = input.executionLedger ?? new ExecutionLedger();
  const run = input.traceBus.getRun(input.runId);
  if (!run) {
    throw new Error("Run not found.");
  }
  assertMatchingRequestId(run, input.requestId);
  assertCurrentWritebackApproval(run, input.approvalToken, input.planDigest);

  const verified = verifyApprovalToken({
    token: input.approvalToken,
    requestId: input.requestId,
    runId: input.runId,
    planDigest: input.planDigest,
    secret: input.config.approvalSecret,
    maxAgeMs: APPROVAL_TOKEN_MAX_AGE_MS
  });

  if (verified.expired) {
    throw new Error("Approval token expired.");
  }

  if (!verified.valid) {
    throw new Error("Invalid approval token.");
  }

  const expectedResultKind = getExpectedResultKind(run.response);
  if (!expectedResultKind || input.result.kind !== expectedResultKind) {
    throw new Error("Writeback result does not match the approved plan family.");
  }
  assertCompletionMatchesApprovedPlan(run.response, input.result);
  const writeback = run.writeback;
  if (!writeback) {
    throw new Error("Writeback approval not found.");
  }
  const completionWorkbookSessionKey = input.workbookSessionKey ?? `run::${input.runId}`;

  if (writeback.workbookSessionKey !== completionWorkbookSessionKey) {
    throw new Error("Writeback completion does not match the approved workbook session.");
  }

  if (writeback.completedAt) {
    if (
      writeback.completedPlanDigest === input.planDigest &&
      completionResultsMatch(writeback.result as CompletionResult | undefined, input.result)
    ) {
      return { ok: true as const };
    }

    throw new Error("Approval token already consumed.");
  }

  if (input.result.kind === "composite_update" && writeback.executionId !== input.result.executionId) {
    throw new Error("Writeback result does not match the approved plan details.");
  }
  const storedPlan = getStoredPlan(run);
  const reversible = isPlanReversible(storedPlan);
  const compositeOutcome = input.result.kind === "composite_update"
    ? getCompositeExecutionOutcome(input.result)
    : undefined;
  const completedStatus = compositeOutcome?.status ?? "completed";
  const undoEligible = compositeOutcome
    ? reversible && compositeOutcome.undoEligible
    : reversible && isCompletionUndoReady(input.result);

  const completedAt = executionLedger.isoTimestamp();
  executionLedger.recordCompleted({
    executionId: writeback.executionId,
    workbookSessionKey: writeback.workbookSessionKey,
    requestId: input.requestId,
    runId: input.runId,
    planType: getPlanTypeFromResponse(run.response),
    planDigest: input.planDigest,
    status: completedStatus,
    timestamp: completedAt,
    reversible,
    undoEligible,
    redoEligible: false,
    summary: buildHistorySummary(getPlanTypeFromResponse(run.response), input.result),
    stepEntries: input.result.kind === "composite_update"
      ? input.result.stepResults.map((stepResult) => ({
        stepId: stepResult.stepId,
        planType: "composite_step",
        status: stepResult.status,
        summary: stepResult.summary
      }))
      : undefined
  });
  input.traceBus.recordWritebackCompletion({
    runId: input.runId,
    completedAt,
    completedPlanDigest: input.planDigest,
    result: input.result
  });

  return { ok: true as const };
}

export function createWritebackRouter(input: {
  traceBus: TraceBus;
  executionLedger: ExecutionLedger;
  config: GatewayConfig;
}): Router {
  const router = Router();

  router.post("/approve", (req, res) => {
    const parsed = ApprovalRequestSchema.safeParse(req.body);
    if (!parsed.success) {
      res.status(400).json(invalidWritebackRequest(
        "That writeback approval request is invalid.",
        "Refresh the spreadsheet state and retry the approval from the latest Hermes response.",
        parsed.error.issues
      ));
      return;
    }

    try {
      const approved = approveWriteback({
        requestId: parsed.data.requestId,
        runId: parsed.data.runId,
        workbookSessionKey: parsed.data.workbookSessionKey,
        plan: parsed.data.plan,
        destructiveConfirmation: parsed.data.destructiveConfirmation,
        traceBus: input.traceBus,
        executionLedger: input.executionLedger,
        config: input.config
      });
      res.json(approved);
    } catch (error) {
      const formatted = formatWritebackRouteError(error);
      res.status(formatted.status).json(formatted.body);
    }
  });

  router.post("/complete", (req, res) => {
    const parsed = CompletionRequestSchema.safeParse(req.body);
    if (!parsed.success) {
      res.status(400).json(invalidWritebackRequest(
        "That writeback completion request is invalid.",
        "Retry the apply step from the latest approved Hermes update.",
        parsed.error.issues
      ));
      return;
    }

    try {
      const completed = completeWriteback({
        requestId: parsed.data.requestId,
        runId: parsed.data.runId,
        workbookSessionKey: parsed.data.workbookSessionKey,
        approvalToken: parsed.data.approvalToken,
        planDigest: parsed.data.planDigest,
        result: parsed.data.result,
        traceBus: input.traceBus,
        executionLedger: input.executionLedger,
        config: input.config
      });
      res.json(completed);
    } catch (error) {
      const formatted = formatWritebackRouteError(error);
      res.status(formatted.status).json(formatted.body);
    }
  });

  return router;
}
