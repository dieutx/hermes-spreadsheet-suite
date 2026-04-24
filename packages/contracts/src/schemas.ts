import { z } from "zod";
import {
  matrixShape,
  parseA1Range,
  validateRectangularMatrix,
  validateTargetRangeMatchesShape
} from "./validators.js";

const strictObject = <Shape extends z.ZodRawShape>(shape: Shape) =>
  z.object(shape).strict();

export const MvpImageMimeTypes = [
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/webp"
] as const;

export const SpreadsheetPlatformSchema = z.enum([
  "google_sheets",
  "excel_windows",
  "excel_macos"
]);

export const ExtractionModeSchema = z.enum(["real", "demo", "unavailable"]);

const StrictA1RangeStringSchema = z.string().min(1).max(128).refine(
  (value) => parseA1Range(value) !== null,
  { message: "must be a valid A1 range." }
);

const MAX_CONTEXT_CELL_TEXT_LENGTH = 4000;
const MAX_CONTEXT_FORMULA_TEXT_LENGTH = 16000;

export const CellValueSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null()
]);

export const TableRowSchema = z.array(CellValueSchema);
export const SheetValues2DSchema = z.array(TableRowSchema);
const NullableTextCellSchema = z.union([z.string(), z.null()]);
const NullableText2DSchema = z.array(z.array(NullableTextCellSchema));

function validateMatrixHeaderWidths(
  matrix: unknown[][] | undefined,
  headersLength: number | undefined,
  ctx: z.RefinementCtx,
  field: string
): void {
  if (matrix === undefined || headersLength === undefined) {
    return;
  }

  matrix.forEach((row, rowIndex) => {
    if (row.length !== headersLength) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `${field} rows must match headers.length.`,
        path: [field, rowIndex]
      });
    }
  });
}

function validateMatrixStringLengths(
  matrix: unknown[][] | undefined,
  maxLength: number,
  ctx: z.RefinementCtx,
  field: string
): void {
  if (matrix === undefined) {
    return;
  }

  matrix.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      if (typeof cell === "string" && cell.length > maxLength) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${field} string cells must be ${maxLength} characters or shorter.`,
          path: [field, rowIndex, columnIndex]
        });
      }
    });
  });
}

export const ShapeSchema = strictObject({
  rows: z.number().int().min(1),
  columns: z.number().int().min(1)
});

export const PreviewShapeSchema = strictObject({
  rows: z.number().int().min(0),
  columns: z.number().int().min(0)
});

export const WarningSchema = strictObject({
  code: z.string().min(1).max(64),
  message: z.string().min(1).max(4000),
  severity: z.enum(["info", "warning", "error"]).optional(),
  field: z.string().min(1).max(128).optional()
});

export const OverwriteRiskSchema = z.enum(["none", "low", "medium", "high"]);
export const ConfirmationLevelSchema = z.enum(["standard", "destructive"]);

export const DownstreamProviderSchema = z.union([
  z.null(),
  strictObject({
    label: z.string().min(1).max(128),
    model: z.string().min(1).max(256).optional()
  })
]);

export const TraceDetailsSchema = strictObject({
  range: z.string().max(64).optional(),
  sheet: z.string().max(128).optional(),
  attachmentId: z.string().max(128).optional(),
  mode: ExtractionModeSchema.optional()
});

export const HermesTraceEventSchema = strictObject({
  event: z.enum([
    "request_received",
    "spreadsheet_context_received",
    "attachment_received",
    "image_received",
    "skill_selected",
    "tool_selected",
    "downstream_provider_called",
    "ocr_started",
    "table_extraction_started",
    "result_generated",
    "composite_plan_ready",
    "sheet_update_plan_ready",
    "sheet_import_plan_ready",
    "workbook_structure_update_ready",
    "sheet_structure_update_ready",
    "range_format_update_ready",
    "conditional_format_plan_ready",
    "range_sort_plan_ready",
    "range_filter_plan_ready",
    "data_validation_plan_ready",
    "named_range_update_ready",
    "range_transfer_plan_ready",
    "data_cleanup_plan_ready",
    "external_data_plan_ready",
    "range_transfer_update_ready",
    "data_cleanup_update_ready",
    "analysis_report_plan_ready",
    "pivot_table_plan_ready",
    "chart_plan_ready",
    "analysis_report_update_ready",
    "pivot_table_update_ready",
    "chart_update_ready",
    "completed",
    "failed"
  ]),
  timestamp: z.string().datetime({ offset: true }),
  label: z.string().max(256).optional(),
  skillName: z.string().max(128).optional(),
  toolName: z.string().max(128).optional(),
  providerLabel: z.string().max(128).optional(),
  details: TraceDetailsSchema.optional()
});

export const UiContractSchema = strictObject({
  displayMode: z.enum(["chat-first", "structured-preview", "error"]),
  showTrace: z.boolean(),
  showWarnings: z.boolean(),
  showConfidence: z.boolean(),
  showRequiresConfirmation: z.boolean()
});

export const ConversationMessageSchema = strictObject({
  role: z.enum(["user", "assistant", "system"]),
  content: z.string().min(1).max(16000)
});

export const AttachmentSchema = strictObject({
  id: z.string().min(1).max(128),
  type: z.literal("image"),
  mimeType: z.enum(MvpImageMimeTypes),
  fileName: z.string().max(512).optional(),
  size: z.number().int().min(0).optional(),
  source: z.enum(["upload", "clipboard", "drag_drop"]),
  previewUrl: z.string().max(4000).optional(),
  uploadToken: z.string().max(1024).optional(),
  storageRef: z.string().max(1024).optional(),
  extractedText: z.string().max(50000).optional(),
  extractedTables: z.array(z.object({}).passthrough()).optional(),
  metadata: z.record(z.unknown()).optional()
});

export const ImageAttachmentSchema = AttachmentSchema;

export const SourceSchema = strictObject({
  channel: SpreadsheetPlatformSchema,
  clientVersion: z.string().min(1).max(64),
  sessionId: z.string().max(128).optional()
});

export const HostSchema = strictObject({
  platform: SpreadsheetPlatformSchema,
  workbookTitle: z.string().min(1).max(512),
  workbookId: z.string().max(256).optional(),
  activeSheet: z.string().min(1).max(128),
  selectedRange: StrictA1RangeStringSchema.optional(),
  locale: z.string().max(32).optional(),
  timeZone: z.string().max(64).optional()
});

export const SelectionContextSchema = strictObject({
  range: StrictA1RangeStringSchema.optional(),
  headers: z.array(z.string().max(256)).optional(),
  values: SheetValues2DSchema.optional(),
  formulas: NullableText2DSchema.optional()
}).superRefine((data, ctx) => {
  validateRectangularMatrix(data.values, ctx, "values");
  validateRectangularMatrix(data.formulas, ctx, "formulas");

  const headersLength = data.headers?.length;
  validateMatrixHeaderWidths(data.values, headersLength, ctx, "values");
  validateMatrixHeaderWidths(data.formulas, headersLength, ctx, "formulas");
  validateMatrixStringLengths(data.values, MAX_CONTEXT_CELL_TEXT_LENGTH, ctx, "values");
  validateMatrixStringLengths(data.formulas, MAX_CONTEXT_FORMULA_TEXT_LENGTH, ctx, "formulas");

  if (data.range && data.headers) {
    const parsedRange = parseA1Range(data.range);
    if (parsedRange && parsedRange.columns !== data.headers.length) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "headers.length must match selection.range width.",
        path: ["headers"]
      });
    }
  }

  if (data.range && data.values) {
    validateTargetRangeMatchesShape(data.range, matrixShape(data.values), ctx, "range");
  }

  if (data.range && data.formulas) {
    validateTargetRangeMatchesShape(data.range, matrixShape(data.formulas), ctx, "range");
  }

  if (data.values && data.formulas) {
    const valuesShape = matrixShape(data.values);
    const formulasShape = matrixShape(data.formulas);
    if (
      valuesShape.rows !== formulasShape.rows ||
      valuesShape.columns !== formulasShape.columns
    ) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "values and formulas must have the same shape.",
        path: ["formulas"]
      });
    }
  }
});

export const ContextCellSchema = strictObject({
  a1Notation: z.string().max(32).optional(),
  displayValue: z.union([z.string().max(MAX_CONTEXT_CELL_TEXT_LENGTH), z.number(), z.boolean(), z.null()]).optional(),
  value: z.unknown().optional(),
  formula: z.string().max(16000).optional(),
  note: z.string().max(4000).optional()
});

export const SheetsPreviewSchema = strictObject({
  sheetName: z.string().max(128),
  headers: z.array(z.string().max(256)).optional(),
  values: SheetValues2DSchema.optional()
});

export const SpreadsheetContextSchema = strictObject({
  selection: SelectionContextSchema.optional(),
  currentRegion: SelectionContextSchema.optional(),
  currentRegionArtifactTarget: StrictA1RangeStringSchema.optional(),
  currentRegionAppendTarget: StrictA1RangeStringSchema.optional(),
  activeCell: ContextCellSchema.optional(),
  referencedCells: z.array(ContextCellSchema).max(20).optional(),
  sheetsPreview: z.array(SheetsPreviewSchema).max(10).optional(),
  attachments: z.array(AttachmentSchema).max(10).optional()
});

export const CapabilitiesSchema = strictObject({
  canRenderTrace: z.boolean(),
  canRenderStructuredPreview: z.boolean(),
  canConfirmWriteBack: z.boolean(),
  supportsStructureEdits: z.boolean().optional(),
  supportsAutofit: z.boolean().optional(),
  supportsSortFilter: z.boolean().optional(),
  supportsImageInputs: z.boolean().optional(),
  supportsWriteBackExecution: z.boolean().optional(),
  supportsNoteWrites: z.boolean().optional()
});

export const ReviewerSchema = strictObject({
  reviewerSafeMode: z.boolean(),
  forceExtractionMode: z.union([ExtractionModeSchema, z.null()]).optional()
});

export const ConfirmationSchema = strictObject({
  state: z.enum(["none", "requested", "confirmed", "rejected"]),
  confirmedPlanId: z.string().max(128).optional()
});

export const HermesRequestSchema = strictObject({
  schemaVersion: z.literal("1.0.0"),
  requestId: z.string().min(1).max(128),
  source: SourceSchema,
  host: HostSchema,
  userMessage: z.string().min(1).max(16000),
  conversation: z.array(ConversationMessageSchema).max(50),
  context: SpreadsheetContextSchema,
  capabilities: CapabilitiesSchema,
  reviewer: ReviewerSchema,
  confirmation: ConfirmationSchema
}).superRefine((data, ctx) => {
  const selectedRange = data.host.selectedRange;
  const selectionRange = data.context.selection?.range;
  if (selectedRange && selectionRange && selectedRange !== selectionRange) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "host.selectedRange must match context.selection.range.",
      path: ["host", "selectedRange"]
    });
  }
});

const BaseResponseEnvelopeSchema = {
  schemaVersion: z.literal("1.0.0"),
  requestId: z.string().min(1).max(128),
  hermesRunId: z.string().min(1).max(128),
  processedBy: z.literal("hermes"),
  serviceLabel: z.string().min(1).max(128),
  environmentLabel: z.string().min(1).max(128),
  startedAt: z.string().datetime({ offset: true }),
  completedAt: z.string().datetime({ offset: true }),
  durationMs: z.number().int().min(0),
  skillsUsed: z.array(z.string().min(1).max(128)).optional(),
  downstreamProvider: DownstreamProviderSchema.optional(),
  warnings: z.array(WarningSchema).optional(),
  trace: z.array(HermesTraceEventSchema).min(1),
  ui: UiContractSchema
} satisfies z.ZodRawShape;

export const ChatDataSchema = strictObject({
  message: z.string().min(1).max(12000),
  followUpSuggestions: z.array(z.string().min(1).max(256)).max(5).optional(),
  confidence: z.number().min(0).max(1).optional()
});

export const FormulaDataSchema = strictObject({
  intent: z.enum(["suggest", "fix", "explain", "translate"]),
  targetCell: z.string().min(1).max(128).optional(),
  formula: z.string().min(1).max(16000),
  formulaLanguage: z.enum(["excel", "google_sheets"]),
  explanation: z.string().min(1).max(12000),
  alternateFormulas: z.array(
    strictObject({
      formula: z.string().min(1).max(16000),
      explanation: z.string().min(1).max(4000)
    })
  ).max(5).optional(),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.boolean().optional()
});

export const WorkbookStructurePositionSchema = z.union([
  z.enum(["start", "end"]),
  z.number().int().min(0)
]);

const workbookStructureSharedFields = {
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  overwriteRisk: OverwriteRiskSchema.optional()
} satisfies z.ZodRawShape;

export const WorkbookStructureUpdateDataSchema = z.discriminatedUnion("operation", [
  strictObject({
    operation: z.literal("create_sheet"),
    sheetName: z.string().min(1).max(128),
    position: WorkbookStructurePositionSchema.optional(),
    ...workbookStructureSharedFields
  }),
  strictObject({
    operation: z.literal("delete_sheet"),
    sheetName: z.string().min(1).max(128),
    ...workbookStructureSharedFields
  }),
  strictObject({
    operation: z.literal("rename_sheet"),
    sheetName: z.string().min(1).max(128),
    newSheetName: z.string().min(1).max(128),
    ...workbookStructureSharedFields
  }),
  strictObject({
    operation: z.literal("duplicate_sheet"),
    sheetName: z.string().min(1).max(128),
    newSheetName: z.string().min(1).max(128).optional(),
    position: WorkbookStructurePositionSchema.optional(),
    ...workbookStructureSharedFields
  }),
  strictObject({
    operation: z.literal("move_sheet"),
    sheetName: z.string().min(1).max(128),
    position: WorkbookStructurePositionSchema,
    ...workbookStructureSharedFields
  }),
  strictObject({
    operation: z.literal("hide_sheet"),
    sheetName: z.string().min(1).max(128),
    ...workbookStructureSharedFields
  }),
  strictObject({
    operation: z.literal("unhide_sheet"),
    sheetName: z.string().min(1).max(128),
    ...workbookStructureSharedFields
  })
]);

const sheetStructureSharedFields = {
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  confirmationLevel: ConfirmationLevelSchema,
  affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional(),
  overwriteRisk: OverwriteRiskSchema.optional()
} satisfies z.ZodRawShape;

const sheetStructureRowColumnOperationFields = {
  startIndex: z.number().int().min(0),
  count: z.number().int().min(1)
} satisfies z.ZodRawShape;

const sheetStructureRangeOperationFields = {
  targetRange: z.string().min(1).max(128)
} satisfies z.ZodRawShape;

const sheetStructureFreezeOperationFields = {
  frozenRows: z.number().int().min(0).optional(),
  frozenColumns: z.number().int().min(0).optional()
} satisfies z.ZodRawShape;

const sheetStructureTabColorOperationFields = {
  color: z.string().regex(/^#[0-9a-fA-F]{6}$/)
} satisfies z.ZodRawShape;

export const SheetStructureUpdateDataSchema = z.discriminatedUnion("operation", [
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("insert_rows"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("delete_rows"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("hide_rows"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("unhide_rows"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("group_rows"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("ungroup_rows"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("insert_columns"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("delete_columns"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("hide_columns"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("unhide_columns"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("group_columns"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("ungroup_columns"),
    ...sheetStructureRowColumnOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("merge_cells"),
    ...sheetStructureRangeOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("unmerge_cells"),
    ...sheetStructureRangeOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("freeze_panes"),
    ...sheetStructureFreezeOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("unfreeze_panes"),
    ...sheetStructureFreezeOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("autofit_rows"),
    ...sheetStructureRangeOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("autofit_columns"),
    ...sheetStructureRangeOperationFields,
    ...sheetStructureSharedFields
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("set_sheet_tab_color"),
    ...sheetStructureTabColorOperationFields,
    ...sheetStructureSharedFields
  })
]).superRefine((data, ctx) => {
  const isDestructive = data.operation === "delete_rows" || data.operation === "delete_columns";
  if (data.confirmationLevel !== (isDestructive ? "destructive" : "standard")) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: isDestructive
        ? "delete_rows and delete_columns require destructive confirmation."
        : "Wave 1 sheet structure operations require standard confirmation.",
      path: ["confirmationLevel"]
    });
  }

  if (data.operation === "freeze_panes" || data.operation === "unfreeze_panes") {
    if (data.frozenRows === undefined || data.frozenColumns === undefined) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "freeze_panes and unfreeze_panes require frozenRows and frozenColumns.",
        path: ["frozenRows"]
      });
    }
  }

  if (data.operation === "unfreeze_panes") {
    if (data.frozenRows !== 0 || data.frozenColumns !== 0) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "unfreeze_panes requires frozenRows and frozenColumns to resolve to 0.",
        path: ["frozenRows"]
      });
    }
  }
});

export const RangeSortKeySchema = strictObject({
  columnRef: z.union([z.string().min(1).max(128), z.number().int().min(1)]),
  direction: z.enum(["asc", "desc"]),
  sortOn: z.string().min(1).max(128).optional()
});

export const RangeSortPlanDataSchema = strictObject({
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  hasHeader: z.boolean(),
  keys: z.array(RangeSortKeySchema).min(1).max(5),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional()
});

export const RangeFilterConditionSchema = strictObject({
  columnRef: z.union([z.string().min(1).max(128), z.number().int().min(1)]),
  operator: z.enum([
    "equals",
    "notEquals",
    "contains",
    "startsWith",
    "endsWith",
    "greaterThan",
    "greaterThanOrEqual",
    "lessThan",
    "lessThanOrEqual",
    "isEmpty",
    "isNotEmpty",
    "topN"
  ]),
  value: CellValueSchema.optional(),
  value2: CellValueSchema.optional()
}).superRefine((data, ctx) => {
  const hasValue = data.value !== undefined;
  const hasValue2 = data.value2 !== undefined;
  const stringOperators = new Set(["equals", "notEquals", "contains", "startsWith", "endsWith"]);
  const numericOperators = new Set([
    "greaterThan",
    "greaterThanOrEqual",
    "lessThan",
    "lessThanOrEqual"
  ]);

  if (hasValue2) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "value2 is reserved for future range/between-style operators.",
      path: ["value2"]
    });
  }

  if (data.operator === "isEmpty" || data.operator === "isNotEmpty") {
    if (hasValue) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `${data.operator} does not accept value.`,
        path: ["value"]
      });
    }

    return;
  }

  if (data.operator === "topN") {
    if (!hasValue || typeof data.value !== "number" || !Number.isFinite(data.value) || data.value <= 0) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "topN requires a positive numeric value.",
        path: ["value"]
      });
    }
    return;
  }

  if (stringOperators.has(data.operator)) {
    if (!hasValue || typeof data.value !== "string") {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `${data.operator} requires a string value.`,
        path: ["value"]
      });
    }
    return;
  }

  if (numericOperators.has(data.operator)) {
    if (!hasValue || typeof data.value !== "number" || !Number.isFinite(data.value)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `${data.operator} requires a numeric value.`,
        path: ["value"]
      });
    }
    return;
  }

  if (!hasValue) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `${data.operator} requires value.`,
      path: ["value"]
    });
  }
});

export const RangeFilterPlanDataSchema = strictObject({
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  hasHeader: z.boolean(),
  conditions: z.array(RangeFilterConditionSchema).min(1).max(10),
  combiner: z.enum(["and", "or"]),
  clearExistingFilters: z.boolean(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional()
});

export const ValidationComparatorSchema = z.enum([
  "between",
  "not_between",
  "equal_to",
  "not_equal_to",
  "greater_than",
  "greater_than_or_equal_to",
  "less_than",
  "less_than_or_equal_to"
]);

export const InvalidDataBehaviorSchema = z.enum(["warn", "reject"]);

const A1TargetRangeSchema = z.string().min(1).max(128).refine(
  (value) => parseA1Range(value) !== null,
  { message: "targetRange must be a valid A1 range." }
);

const SingleCellA1TargetSchema = z.string().min(1).max(128).refine((value) => {
  const parsed = parseA1Range(value);
  return parsed !== null && parsed.rows === 1 && parsed.columns === 1;
}, {
  message: "targetRange must be a single-cell A1 anchor."
});

const SpreadsheetFormulaStringSchema = z.string().min(2).max(4000).refine(
  (value) => value.trim().startsWith("="),
  { message: "formula must start with =." }
);

function isValidDateLiteral(value: string): boolean {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return false;
  }

  const [, yearString, monthString, dayString] = match;
  const year = Number.parseInt(yearString, 10);
  const month = Number.parseInt(monthString, 10);
  const day = Number.parseInt(dayString, 10);

  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return false;
  }

  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

const DateLiteralSchema = z.string().min(1).max(128).refine(
  isValidDateLiteral,
  { message: "Date validation values must be valid YYYY-MM-DD literals." }
);

const dataValidationSharedFields = {
  targetSheet: z.string().min(1).max(128),
  targetRange: A1TargetRangeSchema,
  allowBlank: z.boolean(),
  invalidDataBehavior: InvalidDataBehaviorSchema,
  helpText: z.string().min(1).max(500).optional(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional(),
  replacesExistingValidation: z.boolean().optional()
} satisfies z.ZodRawShape;

const CheckboxValueSchema = z.union([
  z.string().min(1).max(256),
  z.number(),
  z.boolean()
]);

export const DataValidationPlanDataSchema = z.union([
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("list"),
    values: z.array(z.string().min(1).max(256)).min(1).max(500),
    showDropdown: z.boolean().optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("list"),
    sourceRange: z.string().min(1).max(128),
    showDropdown: z.boolean().optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("list"),
    namedRangeName: z.string().min(1).max(255),
    showDropdown: z.boolean().optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("checkbox"),
    checkedValue: CheckboxValueSchema.optional(),
    uncheckedValue: CheckboxValueSchema.optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("whole_number"),
    comparator: ValidationComparatorSchema,
    value: z.number(),
    value2: z.number().optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("decimal"),
    comparator: ValidationComparatorSchema,
    value: z.number(),
    value2: z.number().optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("date"),
    comparator: ValidationComparatorSchema,
    value: DateLiteralSchema,
    value2: DateLiteralSchema.optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("text_length"),
    comparator: ValidationComparatorSchema,
    value: z.number(),
    value2: z.number().optional()
  }),
  strictObject({
    ...dataValidationSharedFields,
    ruleType: z.literal("custom_formula"),
    formula: z.string().min(1).max(16000)
  })
]).superRefine((data, ctx) => {
  if (!("comparator" in data)) {
    return;
  }

  const requiresSecondValue =
    data.comparator === "between" || data.comparator === "not_between";
  const hasSecondValue = "value2" in data && data.value2 !== undefined;

  if (requiresSecondValue && !hasSecondValue) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `${data.comparator} requires value2.`,
      path: ["value2"]
    });
  }

  if (!requiresSecondValue && hasSecondValue) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `${data.comparator} must not include value2.`,
      path: ["value2"]
    });
  }

  if (data.ruleType === "whole_number" || data.ruleType === "text_length") {
    if (!Number.isInteger(data.value)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `${data.ruleType} requires an integer value.`,
        path: ["value"]
      });
    }

    if (data.value2 !== undefined && !Number.isInteger(data.value2)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `${data.ruleType} requires an integer value2.`,
        path: ["value2"]
      });
    }
  }
});

const AnalysisSectionTypeSchema = z.enum([
  "summary_stats",
  "trends",
  "top_bottom",
  "anomalies",
  "group_breakdown",
  "next_actions"
]);

const AnalysisReportSectionSchema = strictObject({
  type: AnalysisSectionTypeSchema,
  title: z.string().min(1).max(256),
  summary: z.string().min(1).max(4000),
  sourceRanges: z.array(z.string().min(1).max(128)).min(1).max(10)
});

const analysisReportPlanSharedFields = {
  sourceSheet: z.string().min(1).max(128),
  sourceRange: A1TargetRangeSchema,
  targetRange: A1TargetRangeSchema.optional(),
  sections: z.array(AnalysisReportSectionSchema).min(1).max(12),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  affectedRanges: z.array(z.string().min(1).max(128)).max(12),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
} satisfies z.ZodRawShape;

export const AnalysisReportPlanDataSchema = z.discriminatedUnion("outputMode", [
  strictObject({
    ...analysisReportPlanSharedFields,
    outputMode: z.literal("chat_only"),
    targetSheet: z.string().min(1).max(128).optional(),
    requiresConfirmation: z.literal(false)
  }),
  strictObject({
    ...analysisReportPlanSharedFields,
    outputMode: z.literal("materialize_report"),
    targetSheet: z.string().min(1).max(128),
    targetRange: A1TargetRangeSchema,
    requiresConfirmation: z.literal(true)
  })
]);

const PivotAggregationSchema = strictObject({
  field: z.string().min(1).max(128),
  aggregation: z.enum(["sum", "count", "average", "min", "max"])
});

const ExternalDataPlanSharedFields = {
  targetSheet: z.string().min(1).max(128),
  targetRange: SingleCellA1TargetSchema,
  formula: SpreadsheetFormulaStringSchema,
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(12),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
} satisfies z.ZodRawShape;

const MarketDataQuerySchema = strictObject({
  symbol: z.string().min(1).max(128),
  attribute: z.string().min(1).max(128).optional(),
  startDate: z.string().min(1).max(64).optional(),
  endDate: z.string().min(1).max(64).optional(),
  interval: z.enum(["DAILY", "WEEKLY"]).optional()
});

const WebImportProviderSchema = z.enum(["importhtml", "importxml", "importdata"]);
const MarketDataPlanSchema = strictObject({
  ...ExternalDataPlanSharedFields,
  sourceType: z.literal("market_data"),
  provider: z.literal("googlefinance"),
  query: MarketDataQuerySchema
});

const WebTableImportPlanSchema = strictObject({
  ...ExternalDataPlanSharedFields,
  sourceType: z.literal("web_table_import"),
  provider: WebImportProviderSchema,
  sourceUrl: z.string().url().max(4000),
  selectorType: z.enum(["table", "list", "xpath", "direct"]),
  selector: z.union([z.string().min(1).max(2000), z.number().int().min(1)]).optional()
});

export const ExternalDataPlanDataSchema = z.union([
  MarketDataPlanSchema,
  WebTableImportPlanSchema
]).superRefine((data, ctx) => {
  if (data.sourceType === "market_data") {
    if (!/GOOGLEFINANCE\s*\(/i.test(data.formula)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "market_data formulas must contain GOOGLEFINANCE(...).",
        path: ["formula"]
      });
    }
    return;
  }

  if (data.provider === "importhtml") {
    if (!/IMPORTHTML\s*\(/i.test(data.formula)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "web_table_import formulas using importhtml must contain IMPORTHTML(...).",
        path: ["formula"]
      });
    }
    if (data.selectorType !== "table" && data.selectorType !== "list") {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "importhtml requires selectorType table or list.",
        path: ["selectorType"]
      });
    }
    if (!Number.isInteger(data.selector) || Number(data.selector) < 1) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "importhtml requires a positive numeric selector index.",
        path: ["selector"]
      });
    }
    return;
  }

  if (data.provider === "importxml") {
    if (!/IMPORTXML\s*\(/i.test(data.formula)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "web_table_import formulas using importxml must contain IMPORTXML(...).",
        path: ["formula"]
      });
    }
    if (data.selectorType !== "xpath") {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "importxml requires selectorType xpath.",
        path: ["selectorType"]
      });
    }
    if (typeof data.selector !== "string" || data.selector.trim().length === 0) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "importxml requires a non-empty xpath selector.",
        path: ["selector"]
      });
    }
    return;
  }

  if (!/IMPORTDATA\s*\(/i.test(data.formula)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "web_table_import formulas using importdata must contain IMPORTDATA(...).",
      path: ["formula"]
    });
  }
  if (data.selectorType !== "direct") {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "importdata requires selectorType direct.",
      path: ["selectorType"]
    });
  }
  if (data.selector !== undefined) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "importdata does not use selector.",
      path: ["selector"]
    });
  }
});

const PivotFilterSchema = strictObject({
  field: z.string().min(1).max(128),
  operator: z.enum([
    "equal_to",
    "not_equal_to",
    "greater_than",
    "greater_than_or_equal_to",
    "less_than",
    "less_than_or_equal_to"
  ]),
  value: CellValueSchema.optional(),
  value2: CellValueSchema.optional()
});

export const PivotTablePlanDataSchema = strictObject({
  sourceSheet: z.string().min(1).max(128),
  sourceRange: A1TargetRangeSchema,
  targetSheet: z.string().min(1).max(128),
  targetRange: A1TargetRangeSchema,
  rowGroups: z.array(z.string().min(1).max(128)).min(1).max(10),
  columnGroups: z.array(z.string().min(1).max(128)).max(10).optional(),
  valueAggregations: z.array(PivotAggregationSchema).min(1).max(10),
  filters: z.array(PivotFilterSchema).max(10).optional(),
  sort: strictObject({
    field: z.string().min(1).max(128),
    direction: z.enum(["asc", "desc"]),
    sortOn: z.enum(["group_field", "aggregated_value"])
  }).optional(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(12),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
});

export const ChartPlanDataSchema = strictObject({
  sourceSheet: z.string().min(1).max(128),
  sourceRange: A1TargetRangeSchema,
  targetSheet: z.string().min(1).max(128),
  targetRange: A1TargetRangeSchema,
  chartType: z.enum([
    "bar",
    "column",
    "stacked_bar",
    "stacked_column",
    "line",
    "area",
    "pie",
    "scatter"
  ]),
  categoryField: z.string().min(1).max(128).optional(),
  series: z.array(
    strictObject({
      field: z.string().min(1).max(128),
      label: z.string().min(1).max(128).optional()
    })
  ).min(1).max(10),
  title: z.string().min(1).max(256).optional(),
  legendPosition: z.enum(["top", "bottom", "left", "right", "hidden"]).optional(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(12),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
});

export const AnalysisReportUpdateDataSchema = strictObject({
  operation: z.literal("analysis_report_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: A1TargetRangeSchema,
  summary: z.string().min(1).max(500)
});

export const PivotTableUpdateDataSchema = strictObject({
  operation: z.literal("pivot_table_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: A1TargetRangeSchema,
  summary: z.string().min(1).max(500)
});

export const ChartUpdateDataSchema = strictObject({
  operation: z.literal("chart_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: A1TargetRangeSchema,
  chartType: ChartPlanDataSchema.shape.chartType,
  summary: z.string().min(1).max(500)
});

const namedRangeUpdateSharedFields = {
  name: z.string().min(1).max(255),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional(),
  overwriteRisk: OverwriteRiskSchema.optional()
} satisfies z.ZodRawShape;

export const NamedRangeUpdateDataSchema = z.union([
  strictObject({
    operation: z.literal("create"),
    scope: z.literal("workbook"),
    targetSheet: z.string().min(1).max(128),
    targetRange: A1TargetRangeSchema,
    ...namedRangeUpdateSharedFields
  }),
  strictObject({
    operation: z.literal("create"),
    scope: z.literal("sheet"),
    sheetName: z.string().min(1).max(128),
    targetSheet: z.string().min(1).max(128),
    targetRange: A1TargetRangeSchema,
    ...namedRangeUpdateSharedFields
  }),
  strictObject({
    operation: z.literal("rename"),
    scope: z.literal("workbook"),
    newName: z.string().min(1).max(255),
    ...namedRangeUpdateSharedFields
  }),
  strictObject({
    operation: z.literal("rename"),
    scope: z.literal("sheet"),
    sheetName: z.string().min(1).max(128),
    newName: z.string().min(1).max(255),
    ...namedRangeUpdateSharedFields
  }),
  strictObject({
    operation: z.literal("delete"),
    scope: z.literal("workbook"),
    ...namedRangeUpdateSharedFields
  }),
  strictObject({
    operation: z.literal("delete"),
    scope: z.literal("sheet"),
    sheetName: z.string().min(1).max(128),
    ...namedRangeUpdateSharedFields
  }),
  strictObject({
    operation: z.literal("retarget"),
    scope: z.literal("workbook"),
    targetSheet: z.string().min(1).max(128),
    targetRange: A1TargetRangeSchema,
    ...namedRangeUpdateSharedFields
  }),
  strictObject({
    operation: z.literal("retarget"),
    scope: z.literal("sheet"),
    sheetName: z.string().min(1).max(128),
    targetSheet: z.string().min(1).max(128),
    targetRange: A1TargetRangeSchema,
    ...namedRangeUpdateSharedFields
  })
]);

export const TransferPasteModeSchema = z.enum(["values", "formulas", "formats"]);

const rangeTransferPlanSharedFields = {
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  pasteMode: TransferPasteModeSchema,
  transpose: z.boolean(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).min(1).max(10),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
} satisfies z.ZodRawShape;

export const RangeTransferPlanDataSchema = z.discriminatedUnion("operation", [
  strictObject({
    sourceSheet: z.string().min(1).max(128),
    sourceRange: z.string().min(1).max(128),
    operation: z.literal("copy"),
    ...rangeTransferPlanSharedFields
  }),
  strictObject({
    sourceSheet: z.string().min(1).max(128),
    sourceRange: z.string().min(1).max(128),
    operation: z.literal("move"),
    ...rangeTransferPlanSharedFields
  }),
  strictObject({
    sourceSheet: z.string().min(1).max(128),
    sourceRange: z.string().min(1).max(128),
    operation: z.literal("append"),
    ...rangeTransferPlanSharedFields
  })
]).superRefine((data, ctx) => {
  const shouldBeDestructive = data.operation === "move";
  if (data.confirmationLevel !== (shouldBeDestructive ? "destructive" : "standard")) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: shouldBeDestructive
        ? "move transfer plans require destructive confirmation."
        : "copy and append transfer plans require standard confirmation.",
      path: ["confirmationLevel"]
    });
  }
});

const dataCleanupPlanSharedFields = {
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).min(1).max(10),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
} satisfies z.ZodRawShape;

const dataCleanupKeyColumnsSchema = z.array(z.string().min(1).max(16)).max(50).optional();
export const Wave4CleanupOperationSchema = z.enum([
  "trim_whitespace",
  "remove_blank_rows",
  "remove_duplicate_rows",
  "normalize_case",
  "split_column",
  "join_columns",
  "fill_down",
  "standardize_format"
]);

export const DataCleanupPlanDataSchema = z.discriminatedUnion("operation", [
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("trim_whitespace")
  }),
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("remove_blank_rows"),
    keyColumns: dataCleanupKeyColumnsSchema
  }),
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("remove_duplicate_rows"),
    keyColumns: dataCleanupKeyColumnsSchema
  }),
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("normalize_case"),
    mode: z.enum(["upper", "lower", "title"])
  }),
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("split_column"),
    sourceColumn: z.string().min(1).max(16),
    delimiter: z.string().min(1).max(128),
    targetStartColumn: z.string().min(1).max(16)
  }),
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("join_columns"),
    sourceColumns: z.array(z.string().min(1).max(16)).min(2).max(50),
    delimiter: z.string().min(1).max(128),
    targetColumn: z.string().min(1).max(16)
  }),
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("fill_down"),
    columns: z.array(z.string().min(1).max(16)).min(1).max(50).optional()
  }),
  strictObject({
    ...dataCleanupPlanSharedFields,
    operation: z.literal("standardize_format"),
    formatType: z.enum(["date_text", "number_text"]),
    formatPattern: z.string().min(1).max(128)
  })
]).superRefine((data, ctx) => {
  const isDestructive =
    data.operation === "remove_blank_rows" ||
    data.operation === "remove_duplicate_rows" ||
    data.operation === "split_column" ||
    data.operation === "join_columns";

  if (data.confirmationLevel !== (isDestructive ? "destructive" : "standard")) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: isDestructive
        ? "This cleanup operation requires destructive confirmation."
        : "This cleanup operation requires standard confirmation.",
      path: ["confirmationLevel"]
    });
  }
});

export const HexColorSchema = z.string().regex(/^#[0-9a-fA-F]{6}$/);
export const HorizontalAlignmentSchema = z.enum([
  "left",
  "center",
  "right",
  "justify",
  "general"
]);
export const VerticalAlignmentSchema = z.enum([
  "top",
  "middle",
  "bottom"
]);
export const WrapStrategySchema = z.enum([
  "wrap",
  "clip",
  "overflow"
]);

export const RangeFormatSchema = strictObject({
  numberFormat: z.string().min(1).max(128).optional(),
  backgroundColor: HexColorSchema.optional(),
  textColor: HexColorSchema.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  horizontalAlignment: HorizontalAlignmentSchema.optional(),
  verticalAlignment: VerticalAlignmentSchema.optional(),
  wrapStrategy: WrapStrategySchema.optional(),
  columnWidth: z.number().positive().max(1000).optional(),
  rowHeight: z.number().positive().max(1000).optional()
}).superRefine((data, ctx) => {
  if (!Object.values(data).some((value) => value !== undefined)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "format must include at least one formatting field."
    });
  }
});

export const RangeFormatUpdateDataSchema = strictObject({
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  format: RangeFormatSchema,
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  overwriteRisk: OverwriteRiskSchema.optional()
});

export const ConditionalFormatManagementModeSchema = z.enum([
  "add",
  "replace_all_on_target",
  "clear_on_target"
]);

export const ConditionalFormatComparatorSchema = z.enum([
  "between",
  "not_between",
  "equal_to",
  "not_equal_to",
  "greater_than",
  "greater_than_or_equal_to",
  "less_than",
  "less_than_or_equal_to"
]);

export const ConditionalFormatRuleTypeSchema = z.enum([
  "single_color",
  "text_contains",
  "number_compare",
  "date_compare",
  "duplicate_values",
  "custom_formula",
  "top_n",
  "average_compare",
  "color_scale"
]);

export const ConditionalFormatColorScalePointTypeSchema = z.enum([
  "min",
  "max",
  "number",
  "percent",
  "percentile"
]);

export const ConditionalFormatStyleSchema = strictObject({
  backgroundColor: HexColorSchema.optional(),
  textColor: HexColorSchema.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  numberFormat: z.string().min(1).max(128).optional()
}).superRefine((data, ctx) => {
  if (!Object.values(data).some((value) => value !== undefined)) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "style must include at least one formatting field."
    });
  }
});

export const ConditionalFormatColorScalePointSchema = strictObject({
  type: ConditionalFormatColorScalePointTypeSchema,
  value: z.number().finite().optional(),
  color: HexColorSchema
}).superRefine((data, ctx) => {
  const requiresValue = data.type === "number" || data.type === "percent" || data.type === "percentile";
  if (requiresValue && data.value === undefined) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `${data.type} color scale points require value.`,
      path: ["value"]
    });
  }

  if (!requiresValue && data.value !== undefined) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `${data.type} color scale points must not include value.`,
      path: ["value"]
    });
  }
});

const ConditionalFormatPlanSharedFields = {
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).min(1).max(10),
  replacesExistingRules: z.boolean()
} satisfies z.ZodRawShape;

const conditionalFormatCompareValueSchema = z.union([
  z.string().min(1).max(128),
  z.number().finite()
]);

const conditionalFormatAddOrReplaceModeSchema = z.enum([
  "add",
  "replace_all_on_target"
]);

function conditionalFormatComparatorRefinement(
  data: { comparator: z.infer<typeof ConditionalFormatComparatorSchema>; value2?: unknown },
  ctx: z.RefinementCtx
): void {
  const requiresSecondValue = data.comparator === "between" || data.comparator === "not_between";
  const hasSecondValue = data.value2 !== undefined;

  if (requiresSecondValue && !hasSecondValue) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `${data.comparator} requires value2.`,
      path: ["value2"]
    });
  }

  if (!requiresSecondValue && hasSecondValue) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: `${data.comparator} must not include value2.`,
      path: ["value2"]
    });
  }
}

export const ConditionalFormatPlanDataSchema = z.union([
  strictObject({
    ...ConditionalFormatPlanSharedFields,
    managementMode: z.literal("clear_on_target")
  }).superRefine((data, ctx) => {
    if (!data.replacesExistingRules) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "clear_on_target replaces existing conditional-format rules.",
        path: ["replacesExistingRules"]
      });
    }
  }),
  z.union([
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("single_color"),
      comparator: ConditionalFormatComparatorSchema,
      value: conditionalFormatCompareValueSchema,
      value2: conditionalFormatCompareValueSchema.optional(),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      conditionalFormatComparatorRefinement(data, ctx);
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("text_contains"),
      text: z.string().min(1).max(512),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("number_compare"),
      comparator: ConditionalFormatComparatorSchema,
      value: conditionalFormatCompareValueSchema,
      value2: conditionalFormatCompareValueSchema.optional(),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      conditionalFormatComparatorRefinement(data, ctx);
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("date_compare"),
      comparator: ConditionalFormatComparatorSchema,
      value: conditionalFormatCompareValueSchema,
      value2: conditionalFormatCompareValueSchema.optional(),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      conditionalFormatComparatorRefinement(data, ctx);
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("duplicate_values"),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("custom_formula"),
      formula: z.string().min(1).max(16000),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("top_n"),
      rank: z.number().int().positive(),
      direction: z.enum(["top", "bottom"]),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("average_compare"),
      direction: z.enum(["above", "below"]),
      style: ConditionalFormatStyleSchema
    }).superRefine((data, ctx) => {
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    }),
    strictObject({
      ...ConditionalFormatPlanSharedFields,
      managementMode: conditionalFormatAddOrReplaceModeSchema,
      ruleType: z.literal("color_scale"),
      points: z.array(ConditionalFormatColorScalePointSchema).min(2).max(3)
    }).superRefine((data, ctx) => {
      const shouldReplace = data.managementMode === "replace_all_on_target";
      if (data.replacesExistingRules !== shouldReplace) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: `${data.managementMode} must set replacesExistingRules to ${shouldReplace}.`,
          path: ["replacesExistingRules"]
        });
      }
    })
  ])
]);

export const ConditionalFormatUpdateDataSchema = strictObject({
  operation: z.literal("conditional_format_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  managementMode: ConditionalFormatManagementModeSchema,
  summary: z.string().min(1).max(500)
});

export const SheetUpdateDataSchema = strictObject({
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  operation: z.enum([
    "replace_range",
    "append_rows",
    "set_formulas",
    "set_notes",
    "mixed_update"
  ]),
  values: SheetValues2DSchema.optional(),
  formulas: NullableText2DSchema.optional(),
  notes: NullableText2DSchema.optional(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  overwriteRisk: OverwriteRiskSchema.optional(),
  shape: ShapeSchema
}).superRefine((data, ctx) => {
  validateRectangularMatrix(data.values, ctx, "values");
  validateRectangularMatrix(data.formulas, ctx, "formulas");
  validateRectangularMatrix(data.notes, ctx, "notes");

  const hasValues = data.values !== undefined;
  const hasFormulas = data.formulas !== undefined;
  const hasNotes = data.notes !== undefined;
  const hasAnyMatrix = hasValues || hasFormulas || hasNotes;

  if (data.operation === "replace_range" && !hasValues && !hasFormulas && !hasNotes) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "replace_range requires values, formulas, or notes.",
      path: ["operation"]
    });
  }

  if (data.operation === "append_rows" && !hasValues) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "append_rows requires values.",
      path: ["values"]
    });
  }

  if (data.operation === "set_formulas") {
    if (!hasFormulas) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "set_formulas requires formulas.",
        path: ["formulas"]
      });
    }

    if (hasValues) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "set_formulas must not include values.",
        path: ["values"]
      });
    }

    if (hasNotes) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "set_formulas must not include notes.",
        path: ["notes"]
      });
    }
  }

  if (data.operation === "set_notes") {
    if (!hasNotes) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "set_notes requires notes.",
        path: ["notes"]
      });
    }

    if (hasValues) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "set_notes must not include values.",
        path: ["values"]
      });
    }

    if (hasFormulas) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "set_notes must not include formulas.",
        path: ["formulas"]
      });
    }
  }

  if (data.operation === "mixed_update" && !hasAnyMatrix) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "mixed_update requires values, formulas, or notes.",
      path: ["operation"]
    });
  }

  for (const [field, matrix] of [
    ["values", data.values],
    ["formulas", data.formulas],
    ["notes", data.notes]
  ] as const) {
    if (!matrix) {
      continue;
    }

    const shape = matrixShape(matrix);
    if (shape.rows !== data.shape.rows || shape.columns !== data.shape.columns) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: `${field} must match shape ${data.shape.rows}x${data.shape.columns}.`,
        path: [field]
      });
    }
  }

  validateTargetRangeMatchesShape(data.targetRange, data.shape, ctx);
});

export const SheetImportPlanDataSchema = strictObject({
  sourceAttachmentId: z.string().min(1).max(128),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  headers: z.array(z.string().min(1).max(256)).min(1),
  values: SheetValues2DSchema,
  confidence: z.number().min(0).max(1),
  warnings: z.array(WarningSchema).optional(),
  requiresConfirmation: z.literal(true),
  extractionMode: ExtractionModeSchema,
  shape: ShapeSchema
}).superRefine((data, ctx) => {
  validateRectangularMatrix(data.values, ctx, "values");

  if (data.shape.columns !== data.headers.length) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "shape.columns must equal headers.length.",
      path: ["shape", "columns"]
    });
  }

  const expectedRows = 1 + data.values.length;
  if (data.shape.rows !== expectedRows) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "shape.rows must equal 1 + values.length when headers are present.",
      path: ["shape", "rows"]
    });
  }

  data.values.forEach((row, rowIndex) => {
    if (row.length !== data.headers.length) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Each values row must match headers.length.",
        path: ["values", rowIndex]
      });
    }
  });

  validateTargetRangeMatchesShape(data.targetRange, data.shape, ctx);
});

const AnalysisReportCompositePlanDataSchema = AnalysisReportPlanDataSchema.refine(
  (plan) => plan.outputMode === "materialize_report",
  {
    message: "analysis_report_plan(chat_only) is not allowed in composite plans."
  }
);

const CompositeExecutablePlanSchema = z.union([
  SheetUpdateDataSchema,
  SheetImportPlanDataSchema,
  ExternalDataPlanDataSchema,
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
  AnalysisReportCompositePlanDataSchema,
  PivotTablePlanDataSchema,
  ChartPlanDataSchema
]);

const CompositePlanStepSchema = strictObject({
  stepId: z.string().min(1).max(128),
  dependsOn: z.array(z.string().min(1).max(128)).max(32),
  continueOnError: z.boolean(),
  plan: CompositeExecutablePlanSchema
});

export const CompositePlanDataSchema = strictObject({
  steps: z.array(CompositePlanStepSchema).min(1).max(32),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(128),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema,
  reversible: z.boolean(),
  dryRunRecommended: z.boolean(),
  dryRunRequired: z.boolean()
}).superRefine((data, ctx) => {
  const stepIndexById = new Map<string, number>();
  data.steps.forEach((step, index) => {
    if (stepIndexById.has(step.stepId)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Composite stepIds must be unique.",
        path: ["steps", index, "stepId"]
      });
      return;
    }

    stepIndexById.set(step.stepId, index);
  });

  data.steps.forEach((step, index) => {
    step.dependsOn.forEach((dependency) => {
      const dependencyIndex = stepIndexById.get(dependency);
      if (dependencyIndex === undefined) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Composite dependsOn targets must reference existing stepIds.",
          path: ["steps", index, "dependsOn"]
        });
        return;
      }

      if (dependencyIndex >= index) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Composite dependencies must appear earlier than the step that depends on them.",
          path: ["steps", index, "dependsOn"]
        });
      }
    });
  });

  const hasDestructiveStep = data.steps.some(
    (step) => "confirmationLevel" in step.plan && step.plan.confirmationLevel === "destructive"
  );
  if (data.confirmationLevel !== (hasDestructiveStep ? "destructive" : "standard")) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: hasDestructiveStep
        ? "Composite confirmationLevel must escalate to destructive when any step is destructive."
        : "Composite confirmationLevel must remain standard when no step is destructive.",
      path: ["confirmationLevel"]
    });
  }

  const hasKnownNonReversibleStep = data.steps.some(
    (step) => ("rowGroups" in step.plan) || ("chartType" in step.plan && "series" in step.plan)
  );
  if (hasKnownNonReversibleStep && data.reversible) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Composite reversible must be false when the plan contains known non-reversible steps.",
      path: ["reversible"]
    });
  }

  const visiting = new Set<string>();
  const visited = new Set<string>();

  const visit = (stepId: string): boolean => {
    if (visited.has(stepId)) {
      return false;
    }
    if (visiting.has(stepId)) {
      return true;
    }

    visiting.add(stepId);
    const step = data.steps[stepIndexById.get(stepId) ?? -1];
    if (step) {
      for (const dependency of step.dependsOn) {
        if (stepIndexById.has(dependency) && visit(dependency)) {
          return true;
        }
      }
    }
    visiting.delete(stepId);
    visited.add(stepId);
    return false;
  };

  for (const step of data.steps) {
    if (visit(step.stepId)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "Composite step dependencies must be acyclic.",
        path: ["steps"]
      });
      break;
    }
  }
});

const compositeStepResultSchema = strictObject({
  stepId: z.string().min(1).max(128),
  status: z.enum(["completed", "failed", "skipped"]),
  summary: z.string().min(1).max(12000)
});

export const CompositeUpdateDataSchema = strictObject({
  operation: z.literal("composite_update"),
  executionId: z.string().min(1).max(128),
  stepResults: z.array(compositeStepResultSchema).min(1),
  summary: z.string().min(1).max(12000)
});

export const UndoRequestSchema = strictObject({
  executionId: z.string().min(1).max(128),
  requestId: z.string().min(1).max(128),
  workbookSessionKey: z.string().min(1).max(256),
  reason: z.string().min(1).max(4000).optional()
});

export const RedoRequestSchema = strictObject({
  executionId: z.string().min(1).max(128),
  requestId: z.string().min(1).max(128),
  workbookSessionKey: z.string().min(1).max(256),
  reason: z.string().min(1).max(4000).optional()
});

export const WritebackApprovalResponseSchema = strictObject({
  requestId: z.string().min(1).max(128),
  runId: z.string().min(1).max(128),
  executionId: z.string().min(1).max(128),
  approvalToken: z.string().min(1).max(4096),
  planDigest: z.string().min(1).max(256),
  approvedAt: z.string().datetime({ offset: true })
});

export const WritebackCompletionResponseSchema = strictObject({
  ok: z.literal(true)
});

const DryRunStepSchema = strictObject({
  stepId: z.string().min(1).max(128),
  status: z.enum(["simulated", "unsupported", "skipped"]),
  summary: z.string().min(1).max(12000),
  predictedAffectedRanges: z.array(z.string().min(1).max(128)).optional(),
  predictedSummaries: z.array(z.string().min(1).max(4000)).optional()
});

export const DryRunResultSchema = strictObject({
  planDigest: z.string().min(1).max(256),
  workbookSessionKey: z.string().min(1).max(256),
  simulated: z.boolean(),
  steps: z.array(DryRunStepSchema).optional(),
  predictedAffectedRanges: z.array(z.string().min(1).max(128)).max(128),
  predictedSummaries: z.array(z.string().min(1).max(4000)).max(128),
  overwriteRisk: OverwriteRiskSchema,
  reversible: z.boolean(),
  expiresAt: z.string().datetime({ offset: true }),
  unsupportedReason: z.string().min(1).max(4000).optional()
}).superRefine((data, ctx) => {
  if (!data.simulated && !data.unsupportedReason) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "non-simulated dry runs require an unsupportedReason.",
      path: ["unsupportedReason"]
    });
  }
});

const PlanHistoryStepEntrySchema = strictObject({
  stepId: z.string().min(1).max(128),
  planType: z.string().min(1).max(128),
  status: z.enum(["completed", "failed", "skipped"]),
  summary: z.string().min(1).max(12000),
  linkedExecutionId: z.string().min(1).max(128).optional()
});

export const PlanHistoryEntrySchema = strictObject({
  executionId: z.string().min(1).max(128),
  requestId: z.string().min(1).max(128),
  runId: z.string().min(1).max(128),
  planType: z.string().min(1).max(128),
  planDigest: z.string().min(1).max(256),
  status: z.enum(["approved", "completed", "failed", "undone", "redone"]),
  timestamp: z.string().datetime({ offset: true }),
  reversible: z.boolean(),
  undoEligible: z.boolean(),
  redoEligible: z.boolean(),
  summary: z.string().min(1).max(12000),
  stepEntries: z.array(PlanHistoryStepEntrySchema).optional(),
  linkedExecutionId: z.string().min(1).max(128).optional()
});

export const PlanHistoryPageSchema = strictObject({
  entries: z.array(PlanHistoryEntrySchema),
  nextCursor: z.string().min(1).max(256).regex(/^(0|[1-9]\d*)$/).optional()
});

export const ErrorDataSchema = strictObject({
  code: z.enum([
    "INVALID_REQUEST",
    "UNSUPPORTED_ATTACHMENT_TYPE",
    "ATTACHMENT_UNAVAILABLE",
    "UNSUPPORTED_OPERATION",
    "SPREADSHEET_CONTEXT_MISSING",
    "EXTRACTION_UNAVAILABLE",
    "CONFIRMATION_REQUIRED",
    "PROVIDER_ERROR",
    "TIMEOUT",
    "INTERNAL_ERROR"
  ]),
  message: z.string().min(1).max(8000),
  retryable: z.boolean(),
  userAction: z.string().max(2000).optional()
});

export const AttachmentContentKindSchema = z.enum([
  "plain_text",
  "table",
  "list",
  "key_value",
  "semi_structured_document",
  "unknown"
]);

export const AttachmentAnalysisDataSchema = strictObject({
  sourceAttachmentId: z.string().min(1).max(128),
  contentKind: AttachmentContentKindSchema,
  summary: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  warnings: z.array(WarningSchema).optional(),
  extractionMode: ExtractionModeSchema
});

export const ExtractedTableDataSchema = strictObject({
  sourceAttachmentId: z.string().min(1).max(128),
  headers: z.array(z.string().min(1).max(256)),
  rows: SheetValues2DSchema,
  confidence: z.number().min(0).max(1),
  warnings: z.array(WarningSchema).optional(),
  extractionMode: ExtractionModeSchema,
  shape: PreviewShapeSchema.optional()
}).superRefine((data, ctx) => {
  validateRectangularMatrix(data.rows, ctx, "rows");

  if (data.headers.length > 0) {
    data.rows.forEach((row, rowIndex) => {
      if (row.length !== data.headers.length) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: "Each extracted row must match headers.length.",
          path: ["rows", rowIndex]
        });
      }
    });
  }

  if (data.shape) {
    const inferredColumns = data.headers.length > 0
      ? data.headers.length
      : matrixShape(data.rows).columns;

    if (data.shape.rows !== data.rows.length) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "shape.rows must equal rows.length.",
        path: ["shape", "rows"]
      });
    }

    if (data.shape.columns !== inferredColumns) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "shape.columns must match the extracted table width.",
        path: ["shape", "columns"]
      });
    }
  }
});

export const DocumentSummaryContentKindSchema = z.enum([
  "plain_text",
  "list",
  "key_value",
  "semi_structured_document",
  "unknown"
]);

export const DocumentSummaryDataSchema = strictObject({
  sourceAttachmentId: z.string().min(1).max(128),
  summary: z.string().min(1).max(12000),
  contentKind: DocumentSummaryContentKindSchema,
  keyPoints: z.array(z.string().min(1).max(1000)).max(20).optional(),
  confidence: z.number().min(0).max(1),
  warnings: z.array(WarningSchema).optional(),
  extractionMode: ExtractionModeSchema
});

function createResponseSchema<
  TypeName extends string,
  DataSchema extends z.ZodTypeAny
>(
  typeName: TypeName,
  dataSchema: DataSchema,
  processedBySchema: z.ZodTypeAny = z.literal("hermes")
) {
  return strictObject({
    ...BaseResponseEnvelopeSchema,
    processedBy: processedBySchema,
    type: z.literal(typeName),
    data: dataSchema
  });
}

const hostOrHermesProcessedBySchema = z.enum(["hermes", "host"]);

export const ChatResponseSchema = createResponseSchema("chat", ChatDataSchema);
export const FormulaResponseSchema = createResponseSchema("formula", FormulaDataSchema);
export const CompositePlanResponseSchema = createResponseSchema(
  "composite_plan",
  CompositePlanDataSchema
);
export const CompositeUpdateResponseSchema = createResponseSchema(
  "composite_update",
  CompositeUpdateDataSchema,
  hostOrHermesProcessedBySchema
);
export const WorkbookStructureUpdateResponseSchema = createResponseSchema(
  "workbook_structure_update",
  WorkbookStructureUpdateDataSchema
);
export const RangeFormatUpdateResponseSchema = createResponseSchema(
  "range_format_update",
  RangeFormatUpdateDataSchema
);
export const ConditionalFormatPlanResponseSchema = createResponseSchema(
  "conditional_format_plan",
  ConditionalFormatPlanDataSchema
);
export const ConditionalFormatUpdateResponseSchema = createResponseSchema(
  "conditional_format_update",
  ConditionalFormatUpdateDataSchema
);
export const SheetStructureUpdateResponseSchema = createResponseSchema(
  "sheet_structure_update",
  SheetStructureUpdateDataSchema
);
export const RangeSortPlanResponseSchema = createResponseSchema(
  "range_sort_plan",
  RangeSortPlanDataSchema
);
export const RangeFilterPlanResponseSchema = createResponseSchema(
  "range_filter_plan",
  RangeFilterPlanDataSchema
);
export const DataValidationPlanResponseSchema = createResponseSchema(
  "data_validation_plan",
  DataValidationPlanDataSchema
);
export const AnalysisReportPlanResponseSchema = createResponseSchema(
  "analysis_report_plan",
  AnalysisReportPlanDataSchema
);
export const PivotTablePlanResponseSchema = createResponseSchema(
  "pivot_table_plan",
  PivotTablePlanDataSchema
);
export const ChartPlanResponseSchema = createResponseSchema("chart_plan", ChartPlanDataSchema);
export const NamedRangeUpdateResponseSchema = createResponseSchema(
  "named_range_update",
  NamedRangeUpdateDataSchema
);
export const RangeTransferPlanResponseSchema = createResponseSchema(
  "range_transfer_plan",
  RangeTransferPlanDataSchema
);
export const DataCleanupPlanResponseSchema = createResponseSchema(
  "data_cleanup_plan",
  DataCleanupPlanDataSchema
);
export const AnalysisReportUpdateResponseSchema = createResponseSchema(
  "analysis_report_update",
  AnalysisReportUpdateDataSchema,
  hostOrHermesProcessedBySchema
);
export const PivotTableUpdateResponseSchema = createResponseSchema(
  "pivot_table_update",
  PivotTableUpdateDataSchema,
  hostOrHermesProcessedBySchema
);
export const ChartUpdateResponseSchema = createResponseSchema(
  "chart_update",
  ChartUpdateDataSchema,
  hostOrHermesProcessedBySchema
);
export const RangeTransferUpdateDataSchema = strictObject({
  operation: z.literal("range_transfer_update"),
  sourceSheet: z.string().min(1).max(128),
  sourceRange: z.string().min(1).max(128),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  transferOperation: z.enum(["copy", "move", "append"]),
  pasteMode: TransferPasteModeSchema,
  transpose: z.boolean(),
  summary: z.string().min(1).max(500)
});

export const DataCleanupUpdateDataSchema = strictObject({
  operation: z.literal("data_cleanup_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  cleanupOperation: Wave4CleanupOperationSchema,
  summary: z.string().min(1).max(500)
});

export const RangeTransferUpdateResponseSchema = createResponseSchema(
  "range_transfer_update",
  RangeTransferUpdateDataSchema
);
export const DataCleanupUpdateResponseSchema = createResponseSchema(
  "data_cleanup_update",
  DataCleanupUpdateDataSchema
);
export const SheetUpdateResponseSchema = createResponseSchema("sheet_update", SheetUpdateDataSchema);
export const SheetImportPlanResponseSchema = createResponseSchema(
  "sheet_import_plan",
  SheetImportPlanDataSchema
);
export const ExternalDataPlanResponseSchema = createResponseSchema(
  "external_data_plan",
  ExternalDataPlanDataSchema
);
export const ErrorResponseSchema = createResponseSchema("error", ErrorDataSchema);
export const AttachmentAnalysisResponseSchema = createResponseSchema(
  "attachment_analysis",
  AttachmentAnalysisDataSchema
);
export const ExtractedTableResponseSchema = createResponseSchema(
  "extracted_table",
  ExtractedTableDataSchema
);
export const DocumentSummaryResponseSchema = createResponseSchema(
  "document_summary",
  DocumentSummaryDataSchema
);

export const HermesResponseSchema = z.discriminatedUnion("type", [
  ChatResponseSchema,
  FormulaResponseSchema,
  CompositePlanResponseSchema,
  CompositeUpdateResponseSchema,
  WorkbookStructureUpdateResponseSchema,
  RangeFormatUpdateResponseSchema,
  ConditionalFormatPlanResponseSchema,
  ConditionalFormatUpdateResponseSchema,
  SheetStructureUpdateResponseSchema,
  RangeSortPlanResponseSchema,
  RangeFilterPlanResponseSchema,
  DataValidationPlanResponseSchema,
  AnalysisReportPlanResponseSchema,
  PivotTablePlanResponseSchema,
  ChartPlanResponseSchema,
  NamedRangeUpdateResponseSchema,
  RangeTransferPlanResponseSchema,
  DataCleanupPlanResponseSchema,
  AnalysisReportUpdateResponseSchema,
  PivotTableUpdateResponseSchema,
  ChartUpdateResponseSchema,
  RangeTransferUpdateResponseSchema,
  DataCleanupUpdateResponseSchema,
  SheetUpdateResponseSchema,
  SheetImportPlanResponseSchema,
  ExternalDataPlanResponseSchema,
  ErrorResponseSchema,
  AttachmentAnalysisResponseSchema,
  ExtractedTableResponseSchema,
  DocumentSummaryResponseSchema
]);

export type ConversationMessage = z.infer<typeof ConversationMessageSchema>;
export type Attachment = z.infer<typeof AttachmentSchema>;
export type ImageAttachment = z.infer<typeof ImageAttachmentSchema>;
export type Source = z.infer<typeof SourceSchema>;
export type Host = z.infer<typeof HostSchema>;
export type SelectionContext = z.infer<typeof SelectionContextSchema>;
export type SpreadsheetContext = z.infer<typeof SpreadsheetContextSchema>;
export type Capabilities = z.infer<typeof CapabilitiesSchema>;
export type Reviewer = z.infer<typeof ReviewerSchema>;
export type Confirmation = z.infer<typeof ConfirmationSchema>;
export type HermesRequest = z.infer<typeof HermesRequestSchema>;
export type Warning = z.infer<typeof WarningSchema>;
export type OverwriteRisk = z.infer<typeof OverwriteRiskSchema>;
export type ConfirmationLevel = z.infer<typeof ConfirmationLevelSchema>;
export type ExtractionMode = z.infer<typeof ExtractionModeSchema>;
export type HermesTraceEvent = z.infer<typeof HermesTraceEventSchema>;
export type DownstreamProvider = z.infer<typeof DownstreamProviderSchema>;
export type Shape = z.infer<typeof ShapeSchema>;
export type WorkbookStructurePosition = z.infer<typeof WorkbookStructurePositionSchema>;
export type SheetStructureUpdateData = z.infer<typeof SheetStructureUpdateDataSchema>;
export type RangeSortKey = z.infer<typeof RangeSortKeySchema>;
export type RangeSortPlanData = z.infer<typeof RangeSortPlanDataSchema>;
export type RangeFilterCondition = z.infer<typeof RangeFilterConditionSchema>;
export type RangeFilterPlanData = z.infer<typeof RangeFilterPlanDataSchema>;
export type ValidationComparator = z.infer<typeof ValidationComparatorSchema>;
export type InvalidDataBehavior = z.infer<typeof InvalidDataBehaviorSchema>;
export type DataValidationPlanData = z.infer<typeof DataValidationPlanDataSchema>;
export type AnalysisReportSection = z.infer<typeof AnalysisReportSectionSchema>;
export type AnalysisReportPlanData = z.infer<typeof AnalysisReportPlanDataSchema>;
export type PivotAggregation = z.infer<typeof PivotAggregationSchema>;
export type PivotFilter = z.infer<typeof PivotFilterSchema>;
export type PivotTablePlanData = z.infer<typeof PivotTablePlanDataSchema>;
export type ChartPlanData = z.infer<typeof ChartPlanDataSchema>;
export type ExternalDataPlanData = z.infer<typeof ExternalDataPlanDataSchema>;
export type AnalysisReportUpdateData = z.infer<typeof AnalysisReportUpdateDataSchema>;
export type PivotTableUpdateData = z.infer<typeof PivotTableUpdateDataSchema>;
export type ChartUpdateData = z.infer<typeof ChartUpdateDataSchema>;
export type NamedRangeUpdateData = z.infer<typeof NamedRangeUpdateDataSchema>;
export type TransferPasteMode = z.infer<typeof TransferPasteModeSchema>;
export type Wave4CleanupOperation = z.infer<typeof Wave4CleanupOperationSchema>;
export type RangeTransferPlanData = z.infer<typeof RangeTransferPlanDataSchema>;
export type DataCleanupPlanData = z.infer<typeof DataCleanupPlanDataSchema>;
export type RangeTransferUpdateData = z.infer<typeof RangeTransferUpdateDataSchema>;
export type DataCleanupUpdateData = z.infer<typeof DataCleanupUpdateDataSchema>;
export type HexColor = z.infer<typeof HexColorSchema>;
export type HorizontalAlignment = z.infer<typeof HorizontalAlignmentSchema>;
export type VerticalAlignment = z.infer<typeof VerticalAlignmentSchema>;
export type WrapStrategy = z.infer<typeof WrapStrategySchema>;
export type ChatData = z.infer<typeof ChatDataSchema>;
export type FormulaData = z.infer<typeof FormulaDataSchema>;
export type CompositePlanData = z.infer<typeof CompositePlanDataSchema>;
export type CompositeUpdateData = z.infer<typeof CompositeUpdateDataSchema>;
export type UndoRequest = z.infer<typeof UndoRequestSchema>;
export type RedoRequest = z.infer<typeof RedoRequestSchema>;
export type WritebackApprovalResponse = z.infer<typeof WritebackApprovalResponseSchema>;
export type WritebackCompletionResponse = z.infer<typeof WritebackCompletionResponseSchema>;
export type DryRunResult = z.infer<typeof DryRunResultSchema>;
export type PlanHistoryEntry = z.infer<typeof PlanHistoryEntrySchema>;
export type PlanHistoryPage = z.infer<typeof PlanHistoryPageSchema>;
export type WorkbookStructureUpdateData = z.infer<typeof WorkbookStructureUpdateDataSchema>;
export type RangeFormat = z.infer<typeof RangeFormatSchema>;
export type RangeFormatUpdateData = z.infer<typeof RangeFormatUpdateDataSchema>;
export type ConditionalFormatManagementMode = z.infer<
  typeof ConditionalFormatManagementModeSchema
>;
export type ConditionalFormatComparator = z.infer<typeof ConditionalFormatComparatorSchema>;
export type ConditionalFormatRuleType = z.infer<typeof ConditionalFormatRuleTypeSchema>;
export type ConditionalFormatColorScalePointType = z.infer<
  typeof ConditionalFormatColorScalePointTypeSchema
>;
export type ConditionalFormatStyle = z.infer<typeof ConditionalFormatStyleSchema>;
export type ConditionalFormatColorScalePoint = z.infer<
  typeof ConditionalFormatColorScalePointSchema
>;
export type ConditionalFormatPlanData = z.infer<typeof ConditionalFormatPlanDataSchema>;
export type ConditionalFormatUpdateData = z.infer<typeof ConditionalFormatUpdateDataSchema>;
export type SheetUpdateData = z.infer<typeof SheetUpdateDataSchema>;
export type SheetImportPlanData = z.infer<typeof SheetImportPlanDataSchema>;
export type ErrorData = z.infer<typeof ErrorDataSchema>;
export type AttachmentAnalysisData = z.infer<typeof AttachmentAnalysisDataSchema>;
export type ExtractedTableData = z.infer<typeof ExtractedTableDataSchema>;
export type DocumentSummaryData = z.infer<typeof DocumentSummaryDataSchema>;
export type ChatResponse = z.infer<typeof ChatResponseSchema>;
export type FormulaResponse = z.infer<typeof FormulaResponseSchema>;
export type CompositePlanResponse = z.infer<typeof CompositePlanResponseSchema>;
export type CompositeUpdateResponse = z.infer<typeof CompositeUpdateResponseSchema>;
export type WorkbookStructureUpdateResponse = z.infer<typeof WorkbookStructureUpdateResponseSchema>;
export type RangeFormatUpdateResponse = z.infer<typeof RangeFormatUpdateResponseSchema>;
export type ConditionalFormatPlanResponse = z.infer<typeof ConditionalFormatPlanResponseSchema>;
export type ConditionalFormatUpdateResponse = z.infer<typeof ConditionalFormatUpdateResponseSchema>;
export type SheetStructureUpdateResponse = z.infer<typeof SheetStructureUpdateResponseSchema>;
export type RangeSortPlanResponse = z.infer<typeof RangeSortPlanResponseSchema>;
export type RangeFilterPlanResponse = z.infer<typeof RangeFilterPlanResponseSchema>;
export type DataValidationPlanResponse = z.infer<typeof DataValidationPlanResponseSchema>;
export type AnalysisReportPlanResponse = z.infer<typeof AnalysisReportPlanResponseSchema>;
export type PivotTablePlanResponse = z.infer<typeof PivotTablePlanResponseSchema>;
export type ChartPlanResponse = z.infer<typeof ChartPlanResponseSchema>;
export type NamedRangeUpdateResponse = z.infer<typeof NamedRangeUpdateResponseSchema>;
export type RangeTransferPlanResponse = z.infer<typeof RangeTransferPlanResponseSchema>;
export type DataCleanupPlanResponse = z.infer<typeof DataCleanupPlanResponseSchema>;
export type AnalysisReportUpdateResponse = z.infer<typeof AnalysisReportUpdateResponseSchema>;
export type PivotTableUpdateResponse = z.infer<typeof PivotTableUpdateResponseSchema>;
export type ChartUpdateResponse = z.infer<typeof ChartUpdateResponseSchema>;
export type RangeTransferUpdateResponse = z.infer<typeof RangeTransferUpdateResponseSchema>;
export type DataCleanupUpdateResponse = z.infer<typeof DataCleanupUpdateResponseSchema>;
export type SheetUpdateResponse = z.infer<typeof SheetUpdateResponseSchema>;
export type SheetImportPlanResponse = z.infer<typeof SheetImportPlanResponseSchema>;
export type ExternalDataPlanResponse = z.infer<typeof ExternalDataPlanResponseSchema>;
export type ErrorResponse = z.infer<typeof ErrorResponseSchema>;
export type AttachmentAnalysisResponse = z.infer<typeof AttachmentAnalysisResponseSchema>;
export type ExtractedTableResponse = z.infer<typeof ExtractedTableResponseSchema>;
export type DocumentSummaryResponse = z.infer<typeof DocumentSummaryResponseSchema>;
export type HermesResponse = z.infer<typeof HermesResponseSchema>;
