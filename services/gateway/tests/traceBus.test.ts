import { describe, expect, it } from "vitest";
import type { HermesResponse, HermesTraceEvent } from "@hermes/contracts";
import { TraceBus } from "../src/lib/traceBus.ts";

describe("TraceBus", () => {
  it("stores public trace events in order and returns only new events after an index", () => {
    const bus = new TraceBus();

    const baseEvent: HermesTraceEvent = {
      event: "request_received",
      timestamp: "2026-04-19T09:00:00.000Z"
    };

    bus.append("run_1", baseEvent);
    bus.append("run_1", {
      event: "result_generated",
      timestamp: "2026-04-19T09:00:01.000Z",
      label: "Result generated"
    });

    expect(bus.list("run_1", 0)).toHaveLength(2);
    expect(bus.list("run_1", 1)).toEqual([
      expect.objectContaining({ event: "result_generated" })
    ]);
  });

  it("sanitizes unsafe optional trace metadata before storage", () => {
    const bus = new TraceBus();

    bus.append("run_sanitized_trace", {
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      label: "ReferenceError stack trace /srv/hermes/internal.ts",
      skillName: "HERMES_API_SERVER_KEY=secret",
      toolName: "SelectionExplainerSkill",
      providerLabel: "http://internal.example/provider",
      details: {
        range: "A1:B2",
        sheet: "/root/secrets",
        attachmentId: "att_public_001",
        mode: "demo"
      }
    });

    const [event] = bus.list("run_sanitized_trace", 0);

    expect(event).toMatchObject({
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      toolName: "SelectionExplainerSkill",
      details: {
        range: "A1:B2",
        attachmentId: "att_public_001",
        mode: "demo"
      }
    });
    expect(event).not.toHaveProperty("label");
    expect(event).not.toHaveProperty("skillName");
    expect(event).not.toHaveProperty("providerLabel");
    expect(event.details).not.toHaveProperty("sheet");
  });

  it("sanitizes embedded secret markers in optional trace metadata", () => {
    const bus = new TraceBus();

    bus.append("run_embedded_trace", {
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      label: "Result_HERMES_API_SERVER_KEY",
      skillName: "SelectionExplainerSkill",
      toolName: "Tool_APPROVAL_SECRET",
      providerLabel: "Gateway provider",
      details: {
        range: "A1:B2_HERMES_AGENT_BASE_URL",
        sheet: "Budget",
        attachmentId: "att_public_001",
        mode: "demo"
      }
    });

    const [event] = bus.list("run_embedded_trace", 0);

    expect(event).toMatchObject({
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      skillName: "SelectionExplainerSkill",
      providerLabel: "Gateway provider",
      details: {
        sheet: "Budget",
        attachmentId: "att_public_001",
        mode: "demo"
      }
    });
    expect(event).not.toHaveProperty("label");
    expect(event).not.toHaveProperty("toolName");
    expect(event.details).not.toHaveProperty("range");
  });

  it("sanitizes unsafe trace detail modes before storage", () => {
    const bus = new TraceBus();

    bus.append("run_mode_trace", {
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      label: "Result generated",
      details: {
        range: "A1:B2",
        mode: "APPROVAL_SECRET=secret"
      }
    });

    const [event] = bus.list("run_mode_trace", 0);

    expect(event).toMatchObject({
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      label: "Result generated",
      details: {
        range: "A1:B2"
      }
    });
    expect(event.details).not.toHaveProperty("mode");
    expect(JSON.stringify(event)).not.toContain("secret");
  });

  it("sanitizes Windows local paths in optional trace metadata", () => {
    const bus = new TraceBus();

    bus.append("run_windows_trace", {
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      label: "C:\\Users\\runner\\work\\internal.ts",
      skillName: "SelectionExplainerSkill",
      toolName: "C:\\Users\\runner\\tools\\extractor.ts",
      providerLabel: "\\\\server\\share\\provider.ts",
      details: {
        range: "A1:B2",
        sheet: "C:\\Users\\runner\\sheets\\budget.xlsx",
        attachmentId: "att_public_001",
        mode: "demo"
      }
    });

    const [event] = bus.list("run_windows_trace", 0);

    expect(event).toMatchObject({
      event: "skill_selected",
      timestamp: "2026-04-19T09:00:00.000Z",
      skillName: "SelectionExplainerSkill",
      details: {
        range: "A1:B2",
        attachmentId: "att_public_001",
        mode: "demo"
      }
    });
    expect(event).not.toHaveProperty("label");
    expect(event).not.toHaveProperty("toolName");
    expect(event).not.toHaveProperty("providerLabel");
    expect(event.details).not.toHaveProperty("sheet");
  });

  it("keeps requestId, hermesRunId, and the validated final response together", () => {
    const bus = new TraceBus();
    bus.ensureRun("run_2", "req_2");
    bus.markStatus("run_2", "processing");
    bus.setResponse("run_2", {
      schemaVersion: "1.0.0",
      type: "chat",
      requestId: "req_2",
      hermesRunId: "hermes_run_2",
      processedBy: "hermes",
      serviceLabel: "spreadsheet-gateway",
      environmentLabel: "review",
      startedAt: "2026-04-19T09:00:00.000Z",
      completedAt: "2026-04-19T09:00:01.000Z",
      durationMs: 1000,
      trace: [
        { event: "request_received", timestamp: "2026-04-19T09:00:00.000Z" },
        { event: "completed", timestamp: "2026-04-19T09:00:01.000Z" }
      ],
      ui: {
        displayMode: "chat-first",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        message: "Processed by Hermes"
      }
    } satisfies HermesResponse);

    const run = bus.getRun("run_2");
    expect(run?.status).toBe("completed");
    expect(run?.requestId).toBe("req_2");
    expect(run?.hermesRunId).toBe("hermes_run_2");
    expect(run?.response?.hermesRunId).toBe("hermes_run_2");
    expect(run?.response?.trace).toBe(run?.events);
  });

  it("evicts the oldest runs when the configured run cap is exceeded", () => {
    const bus = new TraceBus({ maxRuns: 2 });

    bus.ensureRun("run_1", "req_1");
    bus.ensureRun("run_2", "req_2");
    bus.ensureRun("run_3", "req_3");

    expect(bus.getRun("run_1")).toBeUndefined();
    expect(bus.getRun("run_2")?.requestId).toBe("req_2");
    expect(bus.getRun("run_3")?.requestId).toBe("req_3");
  });

  it("prefers evicting completed runs before active runs when the run cap is exceeded", () => {
    let nowMs = Date.UTC(2026, 3, 22, 12, 0, 0);
    const bus = new TraceBus({
      maxRuns: 2,
      now: () => nowMs
    });

    bus.ensureRun("run_completed", "req_completed");
    bus.markStatus("run_completed", "completed");

    nowMs += 1_000;
    bus.ensureRun("run_processing", "req_processing");
    bus.markStatus("run_processing", "processing");

    nowMs += 1_000;
    bus.ensureRun("run_new", "req_new");

    expect(bus.getRun("run_completed")).toBeUndefined();
    expect(bus.getRun("run_processing")?.status).toBe("processing");
    expect(bus.getRun("run_new")?.requestId).toBe("req_new");
  });

  it("expires stale runs after the configured TTL", () => {
    let nowMs = Date.UTC(2026, 3, 22, 12, 0, 0);
    const bus = new TraceBus({
      ttlMs: 1_000,
      now: () => nowMs
    });

    bus.ensureRun("run_stale", "req_stale");
    expect(bus.getRun("run_stale")?.requestId).toBe("req_stale");

    nowMs += 1_500;
    expect(bus.getRun("run_stale")).toBeUndefined();
  });

  it("uses the injected clock for lifecycle timestamps beyond run creation", () => {
    let nowMs = Date.UTC(2026, 3, 23, 0, 0, 0);
    const bus = new TraceBus({
      now: () => nowMs
    });

    bus.ensureRun("run_clocked", "req_clocked");
    expect(bus.getRun("run_clocked")?.startedAt).toBe("2026-04-23T00:00:00.000Z");

    nowMs += 1_000;
    bus.markStatus("run_clocked", "completed");
    expect(bus.getRun("run_clocked")?.completedAt).toBe("2026-04-23T00:00:01.000Z");
  });

  it("does not refresh run TTL when a caller only peeks for auth checks", () => {
    let nowMs = Date.UTC(2026, 3, 22, 12, 0, 0);
    const bus = new TraceBus({
      ttlMs: 1_000,
      now: () => nowMs
    });

    bus.ensureRun("run_peek", "req_peek");

    nowMs += 800;
    expect(bus.peekRun("run_peek")?.requestId).toBe("req_peek");

    nowMs += 300;
    expect(bus.peekRun("run_peek")).toBeUndefined();
  });

  it("caps per-run trace retention and advances the first event index", () => {
    const bus = new TraceBus({ maxEventsPerRun: 2 });

    bus.append("run_trimmed", {
      event: "request_received",
      timestamp: "2026-04-22T00:00:00.000Z"
    });
    bus.append("run_trimmed", {
      event: "result_generated",
      timestamp: "2026-04-22T00:00:01.000Z",
      label: "one"
    });
    bus.append("run_trimmed", {
      event: "completed",
      timestamp: "2026-04-22T00:00:02.000Z"
    });

    const run = bus.peekRun("run_trimmed");
    expect(run?.firstEventIndex).toBe(1);
    expect(run?.events).toHaveLength(2);
    expect(run?.events[0]).toMatchObject({ event: "result_generated" });
    expect(bus.list("run_trimmed", 0)).toHaveLength(2);
  });
});
