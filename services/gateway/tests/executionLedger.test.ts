import { afterEach, describe, expect, it, vi } from "vitest";
import { ExecutionLedger } from "../src/lib/executionLedger.ts";

afterEach(() => {
  vi.useRealTimers();
});

describe("ExecutionLedger", () => {
  it("stores workbook/session-scoped history with undo/redo lineage", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_001",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_001",
      runId: "run_001",
      planType: "sheet_update",
      planDigest: "digest_001",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated Sales!A1:B2."
    });

    const entry = ledger.listHistory("excel_windows::workbook-123").entries[0];

    expect(entry?.executionId).toBe("exec_001");
    expect(entry).not.toHaveProperty("workbookSessionKey");
  });

  it("does not expose session-scoped history through delimiter-colliding unscoped workbook keys", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_session_scoped",
      workbookSessionKey: "excel_windows::workbook-123",
      sessionId: "sess_abc",
      requestId: "req_001",
      runId: "run_001",
      planType: "sheet_update",
      planDigest: "digest_001",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated Sales!A1:B2."
    });

    expect(
      ledger.listHistory("excel_windows::workbook-123", undefined, undefined, "sess_abc")
        .entries.map((entry) => entry.executionId)
    ).toEqual(["exec_session_scoped"]);
    expect(
      ledger.listHistory("excel_windows::workbook-123::session::sess_abc")
        .entries
    ).toEqual([]);
  });

  it("does not expose session-scoped dry-runs through delimiter-colliding unscoped workbook keys", () => {
    const nowMs = Date.UTC(2026, 3, 20, 13, 0, 0);
    const ledger = new ExecutionLedger({ now: () => nowMs });
    const expiresAt = "2026-04-20T13:05:00.000Z";

    ledger.storeDryRun({
      simulated: true,
      planDigest: "digest_session_scoped",
      workbookSessionKey: "excel_windows::workbook-123",
      sessionId: "sess_abc",
      steps: [],
      predictedAffectedRanges: [],
      predictedSummaries: [],
      overwriteRisk: "low",
      reversible: true,
      expiresAt
    });

    expect(
      ledger.getDryRun(
        "excel_windows::workbook-123",
        "digest_session_scoped",
        "sess_abc"
      )?.planDigest
    ).toBe("digest_session_scoped");
    expect(
      ledger.getDryRun(
        "excel_windows::workbook-123::session::sess_abc",
        "digest_session_scoped"
      )
    ).toBeUndefined();
  });

  it("sanitizes unsafe execution history summaries before storage", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_unsafe",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_unsafe",
      runId: "run_unsafe",
      planType: "sheet_update",
      planDigest: "digest_unsafe",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "ReferenceError stack trace /srv/hermes/internal.ts HERMES_API_SERVER_KEY=secret",
      stepEntries: [
        {
          stepId: "step_unsafe",
          planType: "sheet_update",
          status: "completed",
          summary: "http://internal.example/provider"
        }
      ]
    });

    const entry = ledger.listHistory("excel_windows::workbook-123").entries[0];

    expect(entry.summary).toBe("Execution summary hidden because it contained internal details.");
    expect(entry.stepEntries?.[0]?.summary).toBe(
      "Execution summary hidden because it contained internal details."
    );
  });

  it("sanitizes embedded secret markers in execution history summaries", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_embedded_unsafe",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_embedded_unsafe",
      runId: "run_embedded_unsafe",
      planType: "sheet_update",
      planDigest: "digest_embedded_unsafe",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated_sales_HERMES_API_SERVER_KEY",
      stepEntries: [
        {
          stepId: "step_embedded_unsafe",
          planType: "sheet_update",
          status: "completed",
          summary: "Checked_qa_APPROVAL_SECRET"
        }
      ]
    });

    const entry = ledger.listHistory("excel_windows::workbook-123").entries[0];

    expect(entry.summary).toBe("Execution summary hidden because it contained internal details.");
    expect(entry.stepEntries?.[0]?.summary).toBe(
      "Execution summary hidden because it contained internal details."
    );
  });

  it("sanitizes Windows local paths in execution history summaries", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_windows_unsafe",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_windows_unsafe",
      runId: "run_windows_unsafe",
      planType: "sheet_update",
      planDigest: "digest_windows_unsafe",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated by C:\\Users\\runner\\work\\internal.ts",
      stepEntries: [
        {
          stepId: "step_windows_unsafe",
          planType: "sheet_update",
          status: "completed",
          summary: "Checked \\\\server\\share\\step.ts"
        }
      ]
    });

    const entry = ledger.listHistory("excel_windows::workbook-123").entries[0];

    expect(entry.summary).toBe("Execution summary hidden because it contained internal details.");
    expect(entry.stepEntries?.[0]?.summary).toBe(
      "Execution summary hidden because it contained internal details."
    );
  });

  it("rejects invalid pagination cursors instead of silently paging from the wrong index", () => {
    const ledger = new ExecutionLedger();

    expect(() => ledger.listHistory("excel_windows::workbook-123", 10, "-1")).toThrow(
      "History cursor must be a non-negative integer."
    );
    expect(() => ledger.listHistory("excel_windows::workbook-123", 10, "1.5")).toThrow(
      "History cursor must be a non-negative integer."
    );
  });

  it("defaults history pages to the contract page size", () => {
    const ledger = new ExecutionLedger();

    for (let index = 0; index < 101; index += 1) {
      ledger.recordCompleted({
        executionId: `exec_${index}`,
        workbookSessionKey: "excel_windows::workbook-123",
        requestId: `req_${index}`,
        runId: `run_${index}`,
        planType: "sheet_update",
        planDigest: `digest_${index}`,
        status: "completed",
        timestamp: new Date(Date.UTC(2026, 3, 20, 13, 0, index)).toISOString(),
        reversible: true,
        undoEligible: true,
        redoEligible: false,
        summary: "Updated cells."
      });
    }

    const page = ledger.listHistory("excel_windows::workbook-123");

    expect(page.entries).toHaveLength(100);
    expect(page.nextCursor).toBe("100");
  });

  it("sorts history chronologically across mixed timestamp offsets", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_late",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_late",
      runId: "run_late",
      planType: "sheet_update",
      planDigest: "digest_late",
      status: "completed",
      timestamp: "2026-04-20T10:00:00+07:00",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Later execution."
    });
    ledger.recordCompleted({
      executionId: "exec_early",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_early",
      runId: "run_early",
      planType: "sheet_update",
      planDigest: "digest_early",
      status: "completed",
      timestamp: "2026-04-20T04:30:00Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Earlier execution."
    });

    expect(
      ledger.listHistory("excel_windows::workbook-123").entries.map((entry) => entry.executionId)
    ).toEqual(["exec_early", "exec_late"]);
  });

  it("keeps same executionId values isolated by workbook/session during undo lookup", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_shared",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_001",
      runId: "run_001",
      planType: "sheet_update",
      planDigest: "digest_001",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Workbook A update."
    });
    ledger.recordCompleted({
      executionId: "exec_shared",
      workbookSessionKey: "excel_windows::workbook-456",
      requestId: "req_002",
      runId: "run_002",
      planType: "sheet_update",
      planDigest: "digest_002",
      status: "completed",
      timestamp: "2026-04-20T14:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Workbook B update."
    });

    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-04-20T15:00:00.000Z"));

    const undo = ledger.undoExecution({
      executionId: "exec_shared",
      requestId: "req_undo_001",
      workbookSessionKey: "excel_windows::workbook-123"
    });

    expect(undo).toMatchObject({
      operation: "composite_update",
      summary: "Undid execution exec_shared.",
      stepResults: [
        {
          stepId: "undo_exec_shared",
          status: "completed",
          summary: "Undid execution exec_shared."
        }
      ]
    });
    expect(undo.executionId).not.toBe("exec_shared");

    const workbookAHistory = ledger.listHistory("excel_windows::workbook-123").entries;
    expect(workbookAHistory[0]).toMatchObject({
      executionId: undo.executionId,
      requestId: "req_undo_001",
      planType: "undo_request",
      status: "undone",
      reversible: true,
      undoEligible: false,
      redoEligible: true,
      linkedExecutionId: "exec_shared",
      summary: "Undid execution exec_shared."
    });
    expect(workbookAHistory.find((entry) => entry.executionId === "exec_shared")).toMatchObject({
      status: "completed",
      undoEligible: false,
      redoEligible: false
    });

    expect(ledger.listHistory("excel_windows::workbook-456").entries[0]).toMatchObject({
      executionId: "exec_shared",
      status: "completed",
      undoEligible: true,
      redoEligible: false
    });
  });

  it("does not undo session-scoped executions through delimiter-colliding unscoped workbook keys", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_session_scoped",
      workbookSessionKey: "excel_windows::workbook-123",
      sessionId: "sess_abc",
      requestId: "req_001",
      runId: "run_001",
      planType: "sheet_update",
      planDigest: "digest_001",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated Sales!A1:B2."
    });

    expect(() => ledger.undoExecution({
      executionId: "exec_session_scoped",
      requestId: "req_undo_collide",
      workbookSessionKey: "excel_windows::workbook-123::session::sess_abc"
    })).toThrow("Undo target is missing, stale, or not reversible.");
  });

  it("creates redo lineage from an undo execution and rejects superseded lineage points", () => {
    const ledger = new ExecutionLedger();

    ledger.recordCompleted({
      executionId: "exec_001",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_001",
      runId: "run_001",
      planType: "sheet_update",
      planDigest: "digest_001",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Updated Sales!A1:B2."
    });

    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-04-20T15:00:00.000Z"));
    const undo = ledger.undoExecution({
      executionId: "exec_001",
      requestId: "req_undo_001",
      workbookSessionKey: "excel_windows::workbook-123"
    });

    expect(() => ledger.undoExecution({
      executionId: "exec_001",
      requestId: "req_undo_002",
      workbookSessionKey: "excel_windows::workbook-123"
    })).toThrow("Undo target is missing, stale, or not reversible.");

    vi.setSystemTime(new Date("2026-04-20T16:00:00.000Z"));
    const redo = ledger.redoExecution({
      executionId: undo.executionId,
      requestId: "req_redo_001",
      workbookSessionKey: "excel_windows::workbook-123"
    });

    expect(redo).toMatchObject({
      operation: "composite_update",
      summary: `Redid execution ${undo.executionId}.`,
      stepResults: [
        {
          stepId: `redo_${undo.executionId}`,
          status: "completed",
          summary: `Redid execution ${undo.executionId}.`
        }
      ]
    });

    expect(() => ledger.redoExecution({
      executionId: undo.executionId,
      requestId: "req_redo_002",
      workbookSessionKey: "excel_windows::workbook-123"
    })).toThrow("Redo target is missing, stale, or not reversible.");

    const history = ledger.listHistory("excel_windows::workbook-123").entries;
    expect(history[0]).toMatchObject({
      executionId: redo.executionId,
      requestId: "req_redo_001",
      planType: "redo_request",
      status: "redone",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      linkedExecutionId: undo.executionId
    });
    expect(history.find((entry) => entry.executionId === undo.executionId)).toMatchObject({
      status: "undone",
      undoEligible: false,
      redoEligible: false
    });
  });

  it("evicts stale dry-runs and caps dry-run storage size", () => {
    let nowMs = Date.UTC(2026, 3, 22, 6, 0, 0);
    const ledger = new ExecutionLedger({
      maxDryRuns: 2,
      now: () => nowMs
    });

    ledger.storeDryRun({
      simulated: true,
      planDigest: "digest_1",
      workbookSessionKey: "excel_windows::book_1",
      steps: [],
      predictedAffectedRanges: [],
      predictedSummaries: [],
      overwriteRisk: "low",
      reversible: true,
      expiresAt: new Date(nowMs + 1_000).toISOString()
    });

    nowMs += 2_000;
    expect(ledger.getDryRun("excel_windows::book_1", "digest_1")).toBeUndefined();

    ledger.storeDryRun({
      simulated: true,
      planDigest: "digest_2",
      workbookSessionKey: "excel_windows::book_1",
      steps: [],
      predictedAffectedRanges: [],
      predictedSummaries: [],
      overwriteRisk: "low",
      reversible: true,
      expiresAt: new Date(nowMs + 60_000).toISOString()
    });
    ledger.storeDryRun({
      simulated: true,
      planDigest: "digest_3",
      workbookSessionKey: "excel_windows::book_1",
      steps: [],
      predictedAffectedRanges: [],
      predictedSummaries: [],
      overwriteRisk: "low",
      reversible: true,
      expiresAt: new Date(nowMs + 60_000).toISOString()
    });
    ledger.storeDryRun({
      simulated: true,
      planDigest: "digest_4",
      workbookSessionKey: "excel_windows::book_1",
      steps: [],
      predictedAffectedRanges: [],
      predictedSummaries: [],
      overwriteRisk: "low",
      reversible: true,
      expiresAt: new Date(nowMs + 60_000).toISOString()
    });

    expect(ledger.getDryRun("excel_windows::book_1", "digest_2")).toBeUndefined();
    expect(ledger.getDryRun("excel_windows::book_1", "digest_3")?.planDigest).toBe("digest_3");
    expect(ledger.getDryRun("excel_windows::book_1", "digest_4")?.planDigest).toBe("digest_4");
  });

  it("uses the injected ledger clock when checking whether a required dry-run is stale", () => {
    let nowMs = Date.UTC(2026, 3, 22, 6, 0, 0);
    const ledger = new ExecutionLedger({
      now: () => nowMs
    });

    ledger.storeDryRun({
      simulated: true,
      planDigest: "digest_required",
      workbookSessionKey: "excel_windows::book_1",
      steps: [],
      predictedAffectedRanges: [],
      predictedSummaries: [],
      overwriteRisk: "low",
      reversible: true,
      expiresAt: new Date(nowMs + 60_000).toISOString()
    });

    nowMs += 30_000;
    expect(() => ledger.assertFreshDryRun({
      workbookSessionKey: "excel_windows::book_1",
      planDigest: "digest_required",
      required: true
    })).not.toThrow();

    nowMs += 40_000;
    expect(() => ledger.assertFreshDryRun({
      workbookSessionKey: "excel_windows::book_1",
      planDigest: "digest_required",
      required: true
    })).toThrow("Required dry-run is missing, stale, or unusable.");
  });

  it("caps per-workbook history so stale entries do not accumulate forever", () => {
    const ledger = new ExecutionLedger({ maxHistoryEntriesPerWorkbook: 2 });

    ledger.recordCompleted({
      executionId: "exec_1",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_1",
      runId: "run_1",
      planType: "sheet_update",
      planDigest: "digest_1",
      status: "completed",
      timestamp: "2026-04-20T13:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "First execution."
    });
    ledger.recordCompleted({
      executionId: "exec_2",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_2",
      runId: "run_2",
      planType: "sheet_update",
      planDigest: "digest_2",
      status: "completed",
      timestamp: "2026-04-20T14:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Second execution."
    });
    ledger.recordCompleted({
      executionId: "exec_3",
      workbookSessionKey: "excel_windows::workbook-123",
      requestId: "req_3",
      runId: "run_3",
      planType: "sheet_update",
      planDigest: "digest_3",
      status: "completed",
      timestamp: "2026-04-20T15:00:00.000Z",
      reversible: true,
      undoEligible: true,
      redoEligible: false,
      summary: "Third execution."
    });

    expect(
      ledger.listHistory("excel_windows::workbook-123").entries.map((entry) => entry.executionId)
    ).toEqual(["exec_3", "exec_2"]);

    expect(() => ledger.undoExecution({
      executionId: "exec_1",
      requestId: "req_undo_1",
      workbookSessionKey: "excel_windows::workbook-123"
    })).toThrow("Undo target is missing, stale, or not reversible.");
  });
});
