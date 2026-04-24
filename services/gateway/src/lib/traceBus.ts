import type {
  AnalysisReportPlanData,
  ChartUpdateData,
  CompositeUpdateData,
  DataCleanupPlanData,
  ConditionalFormatPlanData,
  ExternalDataPlanData,
  HermesResponse,
  HermesTraceEvent,
  NamedRangeUpdateData,
  PivotTableUpdateData,
  RangeFormatUpdateData,
  RangeTransferUpdateData,
  SheetImportPlanData,
  SheetStructureUpdateData,
  SheetUpdateData
} from "@hermes/contracts";

export type RunStatus = "accepted" | "processing" | "completed" | "failed";

export type WritebackResult =
  | ((SheetUpdateData | SheetImportPlanData) & {
      kind: "range_write";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      writtenRows: number;
      writtenColumns: number;
      undoReady?: boolean;
    })
  | {
      kind: "workbook_structure_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      sheetName: string;
      operation: string;
      summary: string;
    }
  | {
      kind: "sheet_structure_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      targetSheet: string;
      operation: SheetStructureUpdateData["operation"];
      summary: string;
    }
  | (RangeFormatUpdateData & {
      kind: "range_format_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      summary: string;
    })
  | (ExternalDataPlanData & {
      kind: "external_data_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      summary: string;
    })
  | {
      kind: "range_sort";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      targetSheet: string;
      targetRange: string;
      summary: string;
    }
  | {
      kind: "range_filter";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      targetSheet: string;
      targetRange: string;
      summary: string;
    }
  | {
      kind: "data_validation_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      targetSheet: string;
      targetRange: string;
      summary: string;
    }
  | (ConditionalFormatPlanData & {
      kind: "conditional_format_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      summary: string;
    })
  | (NamedRangeUpdateData & {
      kind: "named_range_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      summary: string;
    })
  | {
      kind: "range_transfer_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      operation: "range_transfer_update";
      sourceSheet: RangeTransferUpdateData["sourceSheet"];
      sourceRange: RangeTransferUpdateData["sourceRange"];
      targetSheet: RangeTransferUpdateData["targetSheet"];
      targetRange: RangeTransferUpdateData["targetRange"];
      transferOperation: RangeTransferUpdateData["transferOperation"];
      summary: string;
    }
  | (DataCleanupPlanData & {
      kind: "data_cleanup_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      summary: string;
      undoReady?: boolean;
    })
  | (Extract<AnalysisReportPlanData, { outputMode: "materialize_report" }> & {
      kind: "analysis_report_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      summary: string;
      undoReady?: boolean;
    })
  | {
      kind: "pivot_table_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      operation: PivotTableUpdateData["operation"];
      targetSheet: PivotTableUpdateData["targetSheet"];
      targetRange: PivotTableUpdateData["targetRange"];
      summary: string;
    }
  | {
      kind: "composite_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      operation: CompositeUpdateData["operation"];
      executionId: CompositeUpdateData["executionId"];
      stepResults: CompositeUpdateData["stepResults"];
      summary: string;
    }
  | {
      kind: "chart_update";
      hostPlatform: "google_sheets" | "excel_windows" | "excel_macos";
      operation: ChartUpdateData["operation"];
      targetSheet: ChartUpdateData["targetSheet"];
      targetRange: ChartUpdateData["targetRange"];
      chartType: ChartUpdateData["chartType"];
      summary: string;
    };

export type WritebackState = {
  executionId: string;
  workbookSessionKey: string;
  approvedAt: string;
  approvedPlanDigest: string;
  approvalToken: string;
  destructiveConfirmation?: {
    confirmed: true;
  };
  completedAt?: string;
  completedPlanDigest?: string;
  result?: WritebackResult;
};

export type StoredRun = {
  runId: string;
  requestId?: string;
  hermesRunId?: string;
  status: RunStatus;
  events: HermesTraceEvent[];
  firstEventIndex: number;
  response?: HermesResponse;
  error?: string;
  startedAt: string;
  completedAt?: string;
  writeback?: WritebackState;
  touchedAtMs: number;
};

function eventKey(event: HermesTraceEvent): string {
  return JSON.stringify(event);
}

function mergeTrace(
  existing: HermesTraceEvent[],
  incoming: HermesTraceEvent[]
): HermesTraceEvent[] {
  const seen = new Set(existing.map(eventKey));
  const merged = [...existing];

  for (const event of incoming) {
    const key = eventKey(event);
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    merged.push(event);
  }

  return merged;
}

export class TraceBus {
  private readonly maxRuns: number;
  private readonly maxEventsPerRun: number;
  private readonly ttlMs: number;
  private readonly now: () => number;
  private readonly runs = new Map<string, StoredRun>();

  constructor(options?: {
    maxRuns?: number;
    maxEventsPerRun?: number;
    ttlMs?: number;
    now?: () => number;
  }) {
    this.maxRuns = Math.max(1, options?.maxRuns ?? 1000);
    this.maxEventsPerRun = Math.max(1, options?.maxEventsPerRun ?? 500);
    this.ttlMs = Math.max(1_000, options?.ttlMs ?? 60 * 60 * 1000);
    this.now = options?.now ?? (() => Date.now());
  }

  private touchRun(run: StoredRun): void {
    run.touchedAtMs = this.now();
  }

  nowIso(): string {
    return new Date(this.now()).toISOString();
  }

  private isExpired(run: StoredRun): boolean {
    return (this.now() - run.touchedAtMs) > this.ttlMs;
  }

  private trimRunEvents(run: StoredRun): void {
    if (run.events.length <= this.maxEventsPerRun) {
      return;
    }

    const overflow = run.events.length - this.maxEventsPerRun;
    run.events = run.events.slice(overflow);
    run.firstEventIndex += overflow;

    if (run.response?.trace) {
      run.response = {
        ...run.response,
        trace: run.events
      };
    }
  }

  private pruneExpiredRuns(): void {
    for (const [runId, run] of this.runs.entries()) {
      if (this.isExpired(run)) {
        this.runs.delete(runId);
      }
    }
  }

  private evictOldestRuns(): void {
    this.pruneExpiredRuns();

    while (this.runs.size >= this.maxRuns) {
      const candidates = [...this.runs.entries()];
      const completedCandidates = candidates.filter(([, run]) =>
        run.status === "completed" || run.status === "failed"
      );
      const evictionPool = completedCandidates.length > 0 ? completedCandidates : candidates;
      const oldestEntry = evictionPool
        .sort((left, right) => left[1].touchedAtMs - right[1].touchedAtMs)[0];

      if (!oldestEntry) {
        return;
      }

      this.runs.delete(oldestEntry[0]);
    }
  }

  ensureRun(runId: string, requestId?: string): StoredRun {
    this.pruneExpiredRuns();
    const existing = this.runs.get(runId);
    if (existing) {
      if (requestId && !existing.requestId) {
        existing.requestId = requestId;
      }
      this.touchRun(existing);
      return existing;
    }

    this.evictOldestRuns();
    const nowMs = this.now();

    const created: StoredRun = {
      runId,
      requestId,
      status: "accepted",
      events: [],
      firstEventIndex: 0,
      startedAt: new Date(nowMs).toISOString(),
      touchedAtMs: nowMs
    };

    this.runs.set(runId, created);
    return created;
  }

  append(runId: string, event: HermesTraceEvent): void {
    const run = this.ensureRun(runId);
    run.events.push(event);
    this.trimRunEvents(run);
    this.touchRun(run);
  }

  list(runId: string, afterIndex = 0): HermesTraceEvent[] {
    const run = this.getRun(runId);
    if (!run) {
      return [];
    }

    return run.events.slice(Math.max(0, afterIndex - run.firstEventIndex));
  }

  markStatus(runId: string, status: RunStatus): void {
    const run = this.ensureRun(runId);
    run.status = status;
    if (status === "completed" || status === "failed") {
      run.completedAt = this.nowIso();
    }
    this.touchRun(run);
  }

  setResponse(runId: string, response: HermesResponse): void {
    const run = this.ensureRun(runId, response.requestId);
    const mergedTrace = mergeTrace(run.events, response.trace);
    run.response = {
      ...response,
      trace: mergedTrace
    };
    run.hermesRunId = response.hermesRunId;
    run.events = mergedTrace;
    run.firstEventIndex = Math.max(run.firstEventIndex, 0);
    this.trimRunEvents(run);
    run.status = "completed";
    run.completedAt = this.nowIso();
    this.touchRun(run);
  }

  setError(runId: string, error: string): void {
    const run = this.ensureRun(runId);
    run.error = error;
    run.status = "failed";
    run.completedAt = this.nowIso();
    this.touchRun(run);
  }

  getRun(runId: string): StoredRun | undefined {
    this.pruneExpiredRuns();
    const run = this.runs.get(runId);
    if (!run) {
      return undefined;
    }

    this.touchRun(run);
    return run;
  }

  peekRun(runId: string): StoredRun | undefined {
    this.pruneExpiredRuns();
    return this.runs.get(runId);
  }

  recordWritebackApproval(input: {
    runId: string;
    executionId: string;
    workbookSessionKey: string;
    approvedAt: string;
    approvedPlanDigest: string;
    approvalToken: string;
    destructiveConfirmation?: {
      confirmed: true;
    };
  }): void {
    const run = this.getRun(input.runId);
    if (!run) {
      throw new Error("Run not found.");
    }

    if (run.writeback?.completedAt) {
      throw new Error("Writeback already completed for this run.");
    }

    run.writeback = {
      executionId: input.executionId,
      workbookSessionKey: input.workbookSessionKey,
      approvedAt: input.approvedAt,
      approvedPlanDigest: input.approvedPlanDigest,
      approvalToken: input.approvalToken,
      destructiveConfirmation: input.destructiveConfirmation
    };
    this.touchRun(run);
  }

  recordWritebackCompletion(input: {
    runId: string;
    completedAt: string;
    completedPlanDigest: string;
    result: WritebackResult;
  }): void {
    const run = this.getRun(input.runId);
    if (!run) {
      throw new Error("Run not found.");
    }

    if (!run.writeback) {
      throw new Error("Writeback approval not found.");
    }

    if (run.writeback.completedAt) {
      throw new Error("Approval token already consumed.");
    }

    run.writeback.completedAt = input.completedAt;
    run.writeback.completedPlanDigest = input.completedPlanDigest;
    run.writeback.result = input.result;
    this.touchRun(run);
  }
}
