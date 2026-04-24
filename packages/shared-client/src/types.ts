import type {
  Confirmation,
  ConversationMessage,
  AnalysisReportPlanData,
  CompositePlanData,
  CompositeUpdateData,
  ExtractionMode,
  HermesRequest,
  HermesResponse,
  HermesTraceEvent,
  ImageAttachment,
  ChartPlanData,
  ChartUpdateData,
  ConditionalFormatPlanData,
  DataValidationPlanData,
  DataCleanupPlanData,
  RangeFormatUpdateData,
  RangeFilterPlanData,
  RangeSortPlanData,
  PivotTablePlanData,
  PivotTableUpdateData,
  NamedRangeUpdateData,
  RangeTransferPlanData,
  RangeTransferUpdateData,
  Reviewer,
  DryRunResult,
  PlanHistoryEntry,
  PlanHistoryPage,
  RedoRequest,
  SheetImportPlanData,
  SheetStructureUpdateData,
  SheetUpdateData,
  SpreadsheetContext,
  UndoRequest,
  WritebackApprovalResponse,
  WritebackCompletionResponse,
  WorkbookStructureUpdateData
} from "@hermes/contracts";

type AnalysisReportWritePlanData = Extract<
  AnalysisReportPlanData,
  { outputMode: "materialize_report" }
>;

export type WritePlan =
  | SheetImportPlanData
  | SheetUpdateData
  | WorkbookStructureUpdateData
  | RangeFormatUpdateData
  | ConditionalFormatPlanData
  | SheetStructureUpdateData
  | RangeSortPlanData
  | RangeFilterPlanData
  | DataValidationPlanData
  | NamedRangeUpdateData
  | RangeTransferPlanData
  | DataCleanupPlanData
  | AnalysisReportWritePlanData
  | PivotTablePlanData
  | ChartPlanData;

export type CompositeWritePlan = CompositePlanData;

export type RangeWritebackResult = (SheetUpdateData | SheetImportPlanData) & {
  kind: "range_write";
  hostPlatform: HermesRequest["host"]["platform"];
  writtenRows: number;
  writtenColumns: number;
  undoReady?: boolean;
};

export type RangeFormatUpdateWritebackResult = RangeFormatUpdateData & {
  kind: "range_format_update";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

type WorkbookStructureWritebackBase = {
  kind: "workbook_structure_update";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

export type WorkbookStructureWritebackResult =
  | (WorkbookStructureWritebackBase & {
      operation: "create_sheet";
      sheetName: string;
      positionResolved: number;
      sheetCount: number;
    })
  | (WorkbookStructureWritebackBase & {
      operation: "delete_sheet" | "hide_sheet" | "unhide_sheet";
      sheetName: string;
    })
  | (WorkbookStructureWritebackBase & {
      operation: "rename_sheet";
      sheetName: string;
      newSheetName: string;
    })
  | (WorkbookStructureWritebackBase & {
      operation: "duplicate_sheet";
      sheetName: string;
      newSheetName: string;
      positionResolved: number;
      sheetCount: number;
    })
  | (WorkbookStructureWritebackBase & {
      operation: "move_sheet";
      sheetName: string;
      positionResolved: number;
      sheetCount: number;
    });

export type SheetStructureWritebackResult = {
  kind: "sheet_structure_update";
  hostPlatform: HermesRequest["host"]["platform"];
  targetSheet: string;
  operation: SheetStructureUpdateData["operation"];
  startIndex?: number;
  count?: number;
  targetRange?: string;
  frozenRows?: number;
  frozenColumns?: number;
  color?: string;
  summary: string;
};

export type RangeSortWritebackResult = RangeSortPlanData & {
  kind: "range_sort";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

export type RangeFilterWritebackResult = {
  kind: "range_filter";
  hostPlatform: HermesRequest["host"]["platform"];
} & RangeFilterPlanData & {
  summary: string;
};

export type DataValidationWritebackResult = DataValidationPlanData & {
  kind: "data_validation_update";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

export type ConditionalFormatUpdateWritebackResult = ConditionalFormatPlanData & {
  kind: "conditional_format_update";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

export type NamedRangeUpdateWritebackResult = NamedRangeUpdateData & {
  kind: "named_range_update";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

export type RangeTransferUpdateWritebackResult = {
  kind: "range_transfer_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: RangeTransferUpdateData["operation"];
  sourceSheet: RangeTransferUpdateData["sourceSheet"];
  sourceRange: RangeTransferUpdateData["sourceRange"];
  targetSheet: RangeTransferUpdateData["targetSheet"];
  targetRange: RangeTransferUpdateData["targetRange"];
  transferOperation: RangeTransferUpdateData["transferOperation"];
  summary: string;
};

export type DataCleanupUpdateWritebackResult = DataCleanupPlanData & {
  kind: "data_cleanup_update";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

export type AnalysisReportUpdateWritebackResult = AnalysisReportWritePlanData & {
  kind: "analysis_report_update";
  hostPlatform: HermesRequest["host"]["platform"];
  summary: string;
};

export type PivotTableUpdateWritebackResult = PivotTablePlanData & {
  kind: "pivot_table_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: PivotTableUpdateData["operation"];
  summary: string;
};

export type ChartUpdateWritebackResult = ChartPlanData & {
  kind: "chart_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: ChartUpdateData["operation"];
  summary: string;
};

export type CompositeUpdateWritebackResult = {
  kind: "composite_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: CompositeUpdateData["operation"];
  executionId: CompositeUpdateData["executionId"];
  stepResults: CompositeUpdateData["stepResults"];
  summary: string;
};

export type WritebackResult =
  | RangeWritebackResult
  | RangeFormatUpdateWritebackResult
  | WorkbookStructureWritebackResult
  | SheetStructureWritebackResult
  | RangeSortWritebackResult
  | RangeFilterWritebackResult
  | DataValidationWritebackResult
  | ConditionalFormatUpdateWritebackResult
  | NamedRangeUpdateWritebackResult
  | RangeTransferUpdateWritebackResult
  | DataCleanupUpdateWritebackResult
  | AnalysisReportUpdateWritebackResult
  | PivotTableUpdateWritebackResult
  | ChartUpdateWritebackResult
  | CompositeUpdateWritebackResult;

export type WritebackDestructiveConfirmation = {
  confirmed: true;
};

export type WritebackApprovalRequest = {
  requestId: string;
  runId: string;
  workbookSessionKey?: string;
  destructiveConfirmation?: WritebackDestructiveConfirmation;
  plan: WritePlan;
};

export type WritebackCompletionRequest = {
  requestId: string;
  runId: string;
  workbookSessionKey?: string;
  approvalToken: string;
  planDigest: string;
  result: WritebackResult;
};

export type HostSnapshot = {
  source: HermesRequest["source"];
  host: HermesRequest["host"];
  context: SpreadsheetContext;
};

export type HostBridge = {
  platform: HermesRequest["host"]["platform"];
  getSnapshot(): Promise<HostSnapshot>;
  applyWritePlan(input: {
    plan: WritePlan;
    requestId: string;
    runId: string;
    approvalToken: string;
  }): Promise<WritebackResult>;
};

export type StartRunAccepted = {
  requestId: string;
  runId: string;
  status: "accepted";
};

export type TracePollResult = {
  runId: string;
  requestId?: string;
  hermesRunId?: string;
  status: "accepted" | "processing" | "completed" | "failed";
  nextIndex: number;
  events: HermesTraceEvent[];
};

export type RunPollResult = {
  runId: string;
  requestId?: string;
  hermesRunId?: string;
  status: "accepted" | "processing" | "completed" | "failed";
  startedAt: string;
  completedAt?: string;
  response?: HermesResponse;
  error?: string;
};

export type GatewayClient = {
  uploadImage(input: {
    file: Blob;
    fileName: string;
    source: ImageAttachment["source"];
    sessionId: string;
    workbookId: string;
  }): Promise<ImageAttachment>;
  startRun(request: HermesRequest): Promise<StartRunAccepted>;
  pollRun(runId: string, requestId?: string): Promise<RunPollResult>;
  pollTrace(runId: string, after?: number, requestId?: string): Promise<TracePollResult>;
  dryRunPlan(input: {
    requestId: string;
    runId: string;
    plan: CompositeWritePlan;
  }): Promise<DryRunResult>;
  listPlanHistory(input: {
    workbookSessionKey: string;
    cursor?: string;
    limit?: number;
  }): Promise<PlanHistoryPage>;
  prepareUndoExecution(input: UndoRequest): Promise<CompositeUpdateData>;
  undoExecution(input: UndoRequest): Promise<CompositeUpdateData>;
  prepareRedoExecution(input: RedoRequest): Promise<CompositeUpdateData>;
  redoExecution(input: RedoRequest): Promise<CompositeUpdateData>;
  approveWrite(input: WritebackApprovalRequest): Promise<WritebackApprovalResponse>;
  completeWrite(input: WritebackCompletionRequest): Promise<WritebackCompletionResponse>;
};

export type HostRuntimeConfig = {
  gatewayBaseUrl: string;
  reviewerSafeMode: boolean;
  forceExtractionMode: ExtractionMode | null;
  clientVersion: string;
};

export type RequestEnvelopeInput = {
  source: HermesRequest["source"];
  host: HermesRequest["host"];
  userMessage: string;
  conversation: ConversationMessage[];
  context: SpreadsheetContext;
  capabilities: HermesRequest["capabilities"];
  reviewer: Reviewer;
  confirmation: Confirmation;
};

export type {
  DryRunResult,
  PlanHistoryEntry,
  PlanHistoryPage,
  RedoRequest,
  UndoRequest
};
