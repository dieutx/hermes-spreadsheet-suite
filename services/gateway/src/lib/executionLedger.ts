import type {
  CompositeUpdateData,
  DryRunResult,
  PlanHistoryEntry,
  PlanHistoryPage,
  RedoRequest,
  UndoRequest
} from "@hermes/contracts";

export class StaleExecutionError extends Error {}

export class UnsupportedExecutionControlError extends Error {}

export class FreshDryRunRequiredError extends Error {}

const SANITIZED_EXECUTION_SUMMARY =
  "Execution summary hidden because it contained internal details.";
const UNSAFE_EXECUTION_SUMMARY_PATTERN = /(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|approval_secret|APPROVAL_SECRET|HERMES_[A-Z0-9_]+|OPENAI_API_KEY|ANTHROPIC_API_KEY|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)|\/(?:root|srv|home|tmp|var|opt|workspace|app|mnt)\/[^\s]+|https?:\/\/[^\s]+/i;

type StoredDryRunResult = DryRunResult & { sessionId?: string };

type HistoryRecord = {
  workbookSessionKey: string;
  sessionId?: string;
  entry: PlanHistoryEntry;
};

type ExecutionLedgerOptions = {
  maxHistoryEntriesPerWorkbook?: number;
  maxDryRuns?: number;
  now?: () => number;
};

export function sanitizeExecutionSummary(value: string): string {
  const summary = String(value || "")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!summary || UNSAFE_EXECUTION_SUMMARY_PATTERN.test(summary)) {
    return SANITIZED_EXECUTION_SUMMARY;
  }

  return summary;
}

function sanitizeExecutionSummaryList(values: string[] | undefined): string[] | undefined {
  if (!values) {
    return undefined;
  }

  return values.map((value) => sanitizeExecutionSummary(value).slice(0, 4000));
}

function sanitizeHistoryEntry(entry: PlanHistoryEntry): PlanHistoryEntry {
  return {
    ...entry,
    summary: sanitizeExecutionSummary(entry.summary),
    ...(entry.stepEntries
      ? {
          stepEntries: entry.stepEntries.map((step) => ({
            ...step,
            summary: sanitizeExecutionSummary(step.summary)
          }))
        }
      : {})
  };
}

function sanitizeDryRunResult(result: StoredDryRunResult): StoredDryRunResult {
  return {
    ...result,
    predictedSummaries: sanitizeExecutionSummaryList(result.predictedSummaries) ?? [],
    ...(result.unsupportedReason
      ? { unsupportedReason: sanitizeExecutionSummary(result.unsupportedReason).slice(0, 4000) }
      : {}),
    ...(result.steps
      ? {
          steps: result.steps.map((step) => ({
            ...step,
            summary: sanitizeExecutionSummary(step.summary),
            ...(step.predictedSummaries
              ? { predictedSummaries: sanitizeExecutionSummaryList(step.predictedSummaries) }
              : {})
          }))
        }
      : {})
  };
}

export class ExecutionLedger {
  private readonly history = new Map<string, PlanHistoryEntry[]>();
  private readonly historyByExecutionId = new Map<string, HistoryRecord>();
  private readonly dryRuns = new Map<string, DryRunResult>();
  private readonly maxHistoryEntriesPerWorkbook: number;
  private readonly maxDryRuns: number;
  private readonly now: () => number;

  constructor(options: ExecutionLedgerOptions = {}) {
    this.maxHistoryEntriesPerWorkbook = Math.max(1, options.maxHistoryEntriesPerWorkbook ?? 500);
    this.maxDryRuns = Math.max(1, options.maxDryRuns ?? 500);
    this.now = options.now ?? (() => Date.now());
  }

  nowMs(): number {
    return this.now();
  }

  isoTimestamp(offsetMs = 0): string {
    return new Date(this.now() + offsetMs).toISOString();
  }

  private pruneExpiredDryRuns(): void {
    for (const [key, result] of this.dryRuns.entries()) {
      if (Date.parse(result.expiresAt) <= this.now()) {
        this.dryRuns.delete(key);
      }
    }
  }

  private evictOverflowDryRuns(): void {
    while (this.dryRuns.size > this.maxDryRuns) {
      const oldestKey = this.dryRuns.keys().next().value;
      if (!oldestKey) {
        return;
      }

      this.dryRuns.delete(oldestKey);
    }
  }

  private pruneWorkbookHistory(
    workbookSessionKey: string,
    sessionId: string | undefined,
    bucket: PlanHistoryEntry[]
  ): PlanHistoryEntry[] {
    if (bucket.length <= this.maxHistoryEntriesPerWorkbook) {
      return bucket;
    }

    const retained = [...bucket]
      .sort((left, right) => Date.parse(left.timestamp) - Date.parse(right.timestamp))
      .slice(-this.maxHistoryEntriesPerWorkbook);
    const retainedIds = new Set(retained.map((entry) => entry.executionId));

    for (const entry of bucket) {
      if (!retainedIds.has(entry.executionId)) {
        this.historyByExecutionId.delete(this.getExecutionKey(
          workbookSessionKey,
          entry.executionId,
          sessionId
        ));
      }
    }

    return retained;
  }

  listHistory(
    workbookSessionKey: string,
    limit?: number,
    cursor?: string,
    sessionId?: string
  ): PlanHistoryPage {
    const entries = [...(this.history.get(this.getHistoryKey(workbookSessionKey, sessionId)) ?? [])]
      .sort((left, right) => Date.parse(right.timestamp) - Date.parse(left.timestamp));

    const startIndex = this.parseCursor(cursor);
    const pageSize = typeof limit === "number" && Number.isFinite(limit)
      ? Math.max(1, Math.min(limit, 100))
      : entries.length;

    const pageEntries = entries.slice(startIndex, startIndex + pageSize);
    const nextCursor = startIndex + pageSize < entries.length
      ? String(startIndex + pageSize)
      : undefined;

    return nextCursor
      ? { entries: pageEntries, nextCursor }
      : { entries: pageEntries };
  }

  recordApproved(input: { workbookSessionKey: string; sessionId?: string } & PlanHistoryEntry): void {
    this.upsertHistoryEntry(input);
  }

  recordCompleted(input: { workbookSessionKey: string; sessionId?: string } & PlanHistoryEntry): void {
    this.upsertHistoryEntry(input);
  }

  assertFreshDryRun(input: {
    workbookSessionKey: string;
    sessionId?: string;
    planDigest: string;
    required: boolean;
  }): void {
    if (!input.required) {
      return;
    }

    const result = this.getDryRun(input.workbookSessionKey, input.planDigest, input.sessionId);
    if (!result || !result.simulated || Date.parse(result.expiresAt) <= this.now()) {
      throw new FreshDryRunRequiredError("Required dry-run is missing, stale, or unusable.");
    }
  }

  private upsertHistoryEntry(input: {
    workbookSessionKey: string;
    sessionId?: string;
  } & PlanHistoryEntry): void {
    const { workbookSessionKey, sessionId, ...entry } = input;
    const sanitizedEntry = sanitizeHistoryEntry(entry);
    const historyKey = this.getHistoryKey(workbookSessionKey, sessionId);
    const bucket = this.history.get(historyKey) ?? [];
    const existingIndex = bucket.findIndex(
      (candidate) => candidate.executionId === sanitizedEntry.executionId
    );
    if (existingIndex >= 0) {
      bucket[existingIndex] = sanitizedEntry;
    } else {
      bucket.push(sanitizedEntry);
    }
    this.history.set(historyKey, this.pruneWorkbookHistory(workbookSessionKey, sessionId, bucket));
    this.historyByExecutionId.set(this.getExecutionKey(workbookSessionKey, input.executionId, sessionId), {
      workbookSessionKey,
      sessionId,
      entry: sanitizedEntry
    });
  }

  storeDryRun(result: StoredDryRunResult): void {
    const sanitizedResult = sanitizeDryRunResult(result);
    this.pruneExpiredDryRuns();
    this.dryRuns.set(
      this.getDryRunKey(
        sanitizedResult.workbookSessionKey,
        sanitizedResult.planDigest,
        sanitizedResult.sessionId
      ),
      sanitizedResult
    );
    this.evictOverflowDryRuns();
  }

  getDryRun(
    workbookSessionKey: string,
    planDigest: string,
    sessionId?: string
  ): DryRunResult | undefined {
    this.pruneExpiredDryRuns();
    return this.dryRuns.get(this.getDryRunKey(workbookSessionKey, planDigest, sessionId));
  }

  prepareUndoExecution(request: UndoRequest): CompositeUpdateData {
    const record = this.getUndoRecord(request);
    return this.buildUndoResult(request, record.entry);
  }

  undoExecution(request: UndoRequest): CompositeUpdateData {
    const record = this.getUndoRecord(request);
    const result = this.buildUndoResult(request, record.entry);

    this.upsertHistoryEntry({
      workbookSessionKey: record.workbookSessionKey,
      sessionId: record.sessionId,
      ...record.entry,
      undoEligible: false,
      redoEligible: false
    });

    this.upsertHistoryEntry({
      workbookSessionKey: record.workbookSessionKey,
      sessionId: record.sessionId,
      executionId: result.executionId,
      requestId: request.requestId,
      runId: this.buildControlRunId("undo", request.requestId),
      planType: "undo_request",
      planDigest: this.buildControlPlanDigest("undo", record.entry.planDigest),
      status: "undone",
      timestamp: this.isoTimestamp(),
      reversible: true,
      undoEligible: false,
      redoEligible: true,
      summary: result.summary,
      linkedExecutionId: record.entry.executionId
    });

    return result;
  }

  prepareRedoExecution(request: RedoRequest): CompositeUpdateData {
    const record = this.getRedoRecord(request);
    return this.buildRedoResult(request, record.entry);
  }

  redoExecution(request: RedoRequest): CompositeUpdateData {
    const record = this.getRedoRecord(request);
    const result = this.buildRedoResult(request, record.entry);

    this.upsertHistoryEntry({
      workbookSessionKey: record.workbookSessionKey,
      sessionId: record.sessionId,
      ...record.entry,
      undoEligible: false,
      redoEligible: false
    });

    this.upsertHistoryEntry({
      workbookSessionKey: record.workbookSessionKey,
      sessionId: record.sessionId,
      executionId: result.executionId,
      requestId: request.requestId,
      runId: this.buildControlRunId("redo", request.requestId),
      planType: "redo_request",
      planDigest: this.buildControlPlanDigest("redo", record.entry.planDigest),
      status: "redone",
      timestamp: this.isoTimestamp(),
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: result.summary,
      linkedExecutionId: record.entry.executionId
    });

    return result;
  }

  private getUndoRecord(request: UndoRequest): HistoryRecord {
    const record = this.historyByExecutionId.get(
      this.getExecutionKey(request.workbookSessionKey, request.executionId, request.sessionId)
    );
    if (
      !record
      || !record.entry.undoEligible
      || !record.entry.reversible
    ) {
      throw new StaleExecutionError("Undo target is missing, stale, or not reversible.");
    }

    return record;
  }

  private getRedoRecord(request: RedoRequest): HistoryRecord {
    const record = this.historyByExecutionId.get(
      this.getExecutionKey(request.workbookSessionKey, request.executionId, request.sessionId)
    );
    if (
      !record
      || !record.entry.redoEligible
      || !record.entry.reversible
      || record.entry.status !== "undone"
    ) {
      throw new StaleExecutionError("Redo target is missing, stale, or not reversible.");
    }

    return record;
  }

  private buildUndoResult(request: UndoRequest, entry: PlanHistoryEntry): CompositeUpdateData {
    const summary = `Undid execution ${entry.executionId}.`;
    const executionId = this.buildControlExecutionId("undo", request.requestId, entry.executionId);

    return {
      operation: "composite_update",
      executionId,
      stepResults: [
        {
          stepId: `undo_${entry.executionId}`.slice(0, 128),
          status: "completed",
          summary
        }
      ],
      summary
    };
  }

  private buildRedoResult(request: RedoRequest, entry: PlanHistoryEntry): CompositeUpdateData {
    const summary = `Redid execution ${entry.executionId}.`;
    const executionId = this.buildControlExecutionId("redo", request.requestId, entry.executionId);

    return {
      operation: "composite_update",
      executionId,
      stepResults: [
        {
          stepId: `redo_${entry.executionId}`.slice(0, 128),
          status: "completed",
          summary
        }
      ],
      summary
    };
  }

  private getHistoryKey(workbookSessionKey: string, sessionId?: string): string {
    return sessionId ? `${workbookSessionKey}::session::${sessionId}` : workbookSessionKey;
  }

  private getDryRunKey(workbookSessionKey: string, planDigest: string, sessionId?: string): string {
    return `${this.getHistoryKey(workbookSessionKey, sessionId)}::${planDigest}`;
  }

  private getExecutionKey(workbookSessionKey: string, executionId: string, sessionId?: string): string {
    return `${this.getHistoryKey(workbookSessionKey, sessionId)}::${executionId}`;
  }

  private buildControlExecutionId(
    verb: "undo" | "redo",
    requestId: string,
    targetExecutionId: string
  ): string {
    return `exec_${verb}_${this.now()}_${requestId}_${targetExecutionId}`.slice(0, 128);
  }

  private buildControlRunId(verb: "undo" | "redo", requestId: string): string {
    return `run_${verb}_${requestId}`.slice(0, 128);
  }

  private buildControlPlanDigest(verb: "undo" | "redo", planDigest: string): string {
    return `${verb}::${planDigest}`.slice(0, 256);
  }

  private parseCursor(cursor?: string): number {
    if (cursor === undefined) {
      return 0;
    }

    if (!/^(0|[1-9]\d*)$/.test(cursor)) {
      throw new Error("History cursor must be a non-negative integer.");
    }

    const parsed = Number(cursor);
    if (!Number.isSafeInteger(parsed)) {
      throw new Error("History cursor must be a safe integer.");
    }

    return parsed;
  }
}
