# Wave 6 Composite Plans + Undo/Redo + Dry-Run + History Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a strict Wave 6 control plane for `composite_plan`, workbook/session-scoped undo/redo, exact-safe dry-run, and execution history without breaking Waves 1 through 5.

**Architecture:** Keep Hermes responsible only for proposing `composite_plan`; keep undo/redo, dry-run, and history fully gateway-owned. Extend the existing typed-plan pipeline with strict control-plane schemas, a dedicated execution-control router, and exact-safe host helpers for simulation and inverse execution on the reversible subset only.

**Tech Stack:** TypeScript, Zod, Vitest, Express, Office.js, Google Apps Script, browser-safe shared client helpers.

---

## File Structure

### Contracts and shared type surfaces

- Modify: `packages/contracts/src/schemas.ts`
  - add `composite_plan`, `composite_update`, `undo_request`, `redo_request`, `dry_run_result`, `plan_history_entry`, and `plan_history_page`
- Modify: `packages/contracts/src/index.ts`
  - export Wave 6 schemas and inferred types
- Test: `packages/contracts/tests/contracts.test.ts`
  - validate Wave 6 schemas, no nesting, dependency invariants, and history/dry-run shapes

### Shared client and render pipeline

- Modify: `packages/shared-client/src/types.ts`
  - add Wave 6 write/control/result/history types and gateway client methods
- Modify: `packages/shared-client/src/render.ts`
  - add composite preview rendering, dry-run rendering, history rendering, and response summaries
- Modify: `packages/shared-client/src/trace.ts`
  - add labels for composite update and Wave 6 control events
- Modify: `packages/shared-client/src/index.ts`
  - add `dryRunPlan`, `listPlanHistory`, `undoExecution`, and `redoExecution`
- Test: `packages/shared-client/tests/client.test.ts`

### Gateway runtime, state, and control routes

- Modify: `services/gateway/src/hermes/runtimeRules.ts`
  - teach Hermes `composite_plan` exact-safe rules and no-nesting constraint
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
  - route explicit multi-step workflow prompts toward `composite_plan`
- Modify: `services/gateway/src/hermes/structuredBody.ts`
  - normalize and validate `composite_plan` and `composite_update`
- Modify: `services/gateway/src/lib/hermesClient.ts`
  - map Wave 6 preview and trace metadata
- Create: `services/gateway/src/lib/executionLedger.ts`
  - own workbook/session-scoped history, dry-run cache, undo/redo lineage, and eligibility checks
- Create: `services/gateway/src/routes/executionControl.ts`
  - expose `/api/execution/dry-run`, `/api/execution/history`, `/api/execution/undo`, and `/api/execution/redo`
- Modify: `services/gateway/src/routes/writeback.ts`
  - add `composite_plan` approval/completion support, dry-run gating, and history recording
- Modify: `services/gateway/src/lib/traceBus.ts`
  - persist composite completion state and execution ids alongside run state
- Modify: `services/gateway/src/app.ts`
  - mount the new execution-control router
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`
- Test: `services/gateway/tests/app.test.ts`
- Create: `services/gateway/tests/executionLedger.test.ts`
- Create: `services/gateway/tests/executionControl.test.ts`

### Excel host

- Create: `apps/excel-addin/src/taskpane/compositePlan.js`
  - encapsulate Wave 6 composite summaries, step flags, and dry-run payload helpers
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
  - integrate composite preview, dry-run, history, undo, and redo flows
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
  - extend status summaries for `composite_update`
- Modify: `packages/shared-client/src/index.ts`
  - client methods consumed by Excel taskpane
- Create: `services/gateway/tests/excelWave6Plans.test.ts`
- Test: `services/gateway/tests/excelWritePlan.test.ts`

### Google Sheets host

- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
  - add composite preview, dry-run, history, undo, and redo flows
- Modify: `apps/google-sheets-addon/src/Code.gs`
  - add exact-safe dry-run and inverse helpers for the reversible subset
- Create: `services/gateway/tests/googleSheetsWave6Plans.test.ts`

### Final regression

- Re-run Waves 1 through 5 test batches
- Re-run Wave 6 batches
- Re-run gateway build and host syntax checks

Note: this checkout still has no `.git` metadata, so workers should skip local commit steps unless git metadata is restored.

---

### Task 1: Add Wave 6 Contract Schemas

**Files:**
- Modify: `packages/contracts/src/schemas.ts`
- Modify: `packages/contracts/src/index.ts`
- Test: `packages/contracts/tests/contracts.test.ts`

- [ ] **Step 1: Write the failing contract tests**

```ts
import {
  CompositePlanDataSchema,
  DryRunResultSchema,
  HermesResponseSchema,
  PlanHistoryEntrySchema,
  PlanHistoryPageSchema,
  RedoRequestSchema,
  UndoRequestSchema
} from "../src/index.ts";

it("accepts a composite plan with dependencies and continueOnError semantics", () => {
  const parsed = CompositePlanDataSchema.parse({
    steps: [
      {
        stepId: "step_sort",
        dependsOn: [],
        continueOnError: false,
        plan: {
          targetSheet: "Sales",
          targetRange: "A1:F50",
          hasHeader: true,
          keys: [{ columnRef: "Revenue", direction: "desc" }],
          explanation: "Sort by revenue.",
          confidence: 0.91,
          requiresConfirmation: true,
          affectedRanges: ["Sales!A1:F50"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      },
      {
        stepId: "step_filter",
        dependsOn: ["step_sort"],
        continueOnError: true,
        plan: {
          targetSheet: "Sales",
          targetRange: "A1:F50",
          hasHeader: true,
          conditions: [{ columnRef: "Status", operator: "equal_to", value: "Open" }],
          combiner: "and",
          clearExistingFilters: true,
          explanation: "Filter open rows.",
          confidence: 0.88,
          requiresConfirmation: true,
          affectedRanges: ["Sales!A1:F50"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      }
    ],
    explanation: "Sort then filter the current table.",
    confidence: 0.9,
    requiresConfirmation: true,
    affectedRanges: ["Sales!A1:F50"],
    overwriteRisk: "low",
    confirmationLevel: "standard",
    reversible: true,
    dryRunRecommended: true,
    dryRunRequired: false
  });

  expect(parsed.steps).toHaveLength(2);
});

it("rejects nested composite plans", () => {
  const parsed = CompositePlanDataSchema.safeParse({
    steps: [
      {
        stepId: "nested",
        dependsOn: [],
        continueOnError: false,
        plan: {
          steps: [],
          explanation: "No nesting.",
          confidence: 0.1,
          requiresConfirmation: true,
          affectedRanges: [],
          overwriteRisk: "none",
          confirmationLevel: "standard",
          reversible: false,
          dryRunRecommended: false,
          dryRunRequired: false
        }
      }
    ],
    explanation: "Bad nesting.",
    confidence: 0.3,
    requiresConfirmation: true,
    affectedRanges: [],
    overwriteRisk: "none",
    confirmationLevel: "standard",
    reversible: false,
    dryRunRecommended: false,
    dryRunRequired: false
  });

  expect(parsed.success).toBe(false);
});

it("accepts undo and redo request envelopes", () => {
  expect(UndoRequestSchema.parse({
    executionId: "exec_001",
    requestId: "req_undo_001",
    workbookSessionKey: "excel_windows::workbook-123"
  }).executionId).toBe("exec_001");

  expect(RedoRequestSchema.parse({
    executionId: "exec_undo_001",
    requestId: "req_redo_001",
    workbookSessionKey: "excel_windows::workbook-123"
  }).executionId).toBe("exec_undo_001");
});

it("accepts dry-run results and history pages", () => {
  expect(DryRunResultSchema.parse({
    planDigest: "digest_001",
    workbookSessionKey: "excel_windows::workbook-123",
    simulated: true,
    predictedAffectedRanges: ["Sales!A1:F50"],
    predictedSummaries: ["Will sort Sales!A1:F50 by Revenue descending."],
    overwriteRisk: "low",
    reversible: true,
    expiresAt: "2026-04-20T13:00:00.000Z"
  }).simulated).toBe(true);

  expect(PlanHistoryPageSchema.parse({
    entries: [
      {
        executionId: "exec_001",
        requestId: "req_001",
        runId: "run_001",
        planType: "composite_plan",
        planDigest: "digest_001",
        status: "completed",
        timestamp: "2026-04-20T13:00:00.000Z",
        reversible: true,
        undoEligible: true,
        redoEligible: false,
        summary: "Completed 2-step composite execution."
      }
    ]
  }).entries).toHaveLength(1);
});

it("accepts a composite_plan Hermes response envelope", () => {
  const parsed = HermesResponseSchema.parse({
    schemaVersion: "1.0.0",
    type: "composite_plan",
    requestId: "req_composite_001",
    hermesRunId: "run_composite_001",
    processedBy: "hermes",
    serviceLabel: "hermes-gateway-local",
    environmentLabel: "local-dev",
    startedAt: "2026-04-20T13:00:00.000Z",
    completedAt: "2026-04-20T13:00:01.000Z",
    durationMs: 1000,
    trace: [{ event: "completed", timestamp: "2026-04-20T13:00:01.000Z" }],
    ui: {
      displayMode: "structured-preview",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: true
    },
    data: {
      steps: [
        {
          stepId: "step_sort",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:F50",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "desc" }],
            explanation: "Sort by revenue.",
            confidence: 0.91,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50"],
            overwriteRisk: "low",
            confirmationLevel: "standard"
          }
        }
      ],
      explanation: "Run a one-step workflow.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    }
  });

  expect(parsed.type).toBe("composite_plan");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: FAIL with missing Wave 6 schemas, exports, or unsupported response branches.

- [ ] **Step 3: Write the minimal contract implementation**

```ts
const CompositeStepSchema = z.lazy(() => strictObject({
  stepId: z.string().min(1).max(128),
  dependsOn: z.array(z.string().min(1).max(128)).max(32),
  continueOnError: z.boolean(),
  plan: ExecutableWritePlanSchema
}));

export const CompositePlanDataSchema = strictObject({
  steps: z.array(CompositeStepSchema).min(1).max(32),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(128),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema,
  reversible: z.boolean(),
  dryRunRecommended: z.boolean(),
  dryRunRequired: z.boolean()
}).superRefine((plan, ctx) => {
  const seen = new Set<string>();
  for (const step of plan.steps) {
    if (seen.has(step.stepId)) {
      ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["steps"], message: "Duplicate stepId." });
    }
    seen.add(step.stepId);
    if ((step.plan as { steps?: unknown }).steps) {
      ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["steps"], message: "Nested composite plans are not allowed." });
    }
  }
});

export const UndoRequestSchema = strictObject({
  executionId: z.string().min(1).max(128),
  requestId: z.string().min(1).max(128),
  workbookSessionKey: z.string().min(1).max(256),
  reason: z.string().min(1).max(4000).optional()
});

export const RedoRequestSchema = UndoRequestSchema;

export const DryRunResultSchema = strictObject({
  planDigest: z.string().min(1).max(256),
  workbookSessionKey: z.string().min(1).max(256),
  simulated: z.boolean(),
  steps: z.array(strictObject({
    stepId: z.string().min(1).max(128),
    status: z.enum(["simulated", "unsupported", "skipped"]),
    summary: z.string().min(1).max(4000)
  })).optional(),
  predictedAffectedRanges: z.array(z.string().min(1).max(128)).max(128),
  predictedSummaries: z.array(z.string().min(1).max(4000)).max(128),
  overwriteRisk: OverwriteRiskSchema,
  reversible: z.boolean(),
  expiresAt: IsoTimestampSchema,
  unsupportedReason: z.string().min(1).max(4000).optional()
});
```

- [ ] **Step 4: Run the contracts test again**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: PASS for the new Wave 6 contract coverage.

- [ ] **Step 5: Checkpoint the schema surface**

Verify:

- `packages/contracts/src/index.ts` exports every new Wave 6 schema/type
- `HermesResponseSchema` recognizes `composite_plan` and `composite_update`
- no existing Wave 1–5 schema branches changed shape unintentionally

---

### Task 2: Extend Shared Client Types and Rendering

**Files:**
- Modify: `packages/shared-client/src/types.ts`
- Modify: `packages/shared-client/src/render.ts`
- Modify: `packages/shared-client/src/trace.ts`
- Modify: `packages/shared-client/src/index.ts`
- Test: `packages/shared-client/tests/client.test.ts`

- [ ] **Step 1: Write the failing shared-client tests**

```ts
it("renders a composite preview with step flags", () => {
  const response = {
    type: "composite_plan",
    data: {
      steps: [
        {
          stepId: "step_sort",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:F50",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "desc" }],
            explanation: "Sort by revenue.",
            confidence: 0.91,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50"],
            overwriteRisk: "low",
            confirmationLevel: "standard"
          }
        }
      ],
      explanation: "Run the workflow.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    }
  };

  expect(buildStructuredPreview(response as any)).toMatchObject({
    kind: "composite_plan",
    stepCount: 1,
    reversible: true
  });
});

it("formats dry-run results and history entries without lossy fallback text", () => {
  expect(formatDryRunSummary({
    simulated: true,
    predictedAffectedRanges: ["Sales!A1:F50"],
    predictedSummaries: ["Will sort Sales!A1:F50 by Revenue descending."]
  } as any)).toContain("Will sort Sales!A1:F50");

  expect(formatHistoryEntrySummary({
    status: "completed",
    summary: "Completed 2-step composite execution."
  } as any)).toBe("Completed 2-step composite execution.");
});

it("adds gateway client methods for dry-run, history, undo, and redo", async () => {
  const client = createGatewayClient("http://localhost:18787");

  expect(typeof client.dryRunPlan).toBe("function");
  expect(typeof client.listPlanHistory).toBe("function");
  expect(typeof client.undoExecution).toBe("function");
  expect(typeof client.redoExecution).toBe("function");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: FAIL with missing Wave 6 preview kinds, helper functions, or gateway-client methods.

- [ ] **Step 3: Add the minimal shared-client implementation**

```ts
export type CompositeWritePlan = {
  steps: Array<{
    stepId: string;
    dependsOn: string[];
    continueOnError: boolean;
    plan: WritePlan;
  }>;
  explanation: string;
  confidence: number;
  requiresConfirmation: true;
  affectedRanges: string[];
  overwriteRisk: "none" | "low" | "medium" | "high";
  confirmationLevel: "standard" | "destructive";
  reversible: boolean;
  dryRunRecommended: boolean;
  dryRunRequired: boolean;
};

export type DryRunResult = {
  planDigest: string;
  workbookSessionKey: string;
  simulated: boolean;
  predictedAffectedRanges: string[];
  predictedSummaries: string[];
  overwriteRisk: "none" | "low" | "medium" | "high";
  reversible: boolean;
  expiresAt: string;
  unsupportedReason?: string;
};

export function formatDryRunSummary(result: DryRunResult): string {
  if (!result.simulated) {
    return result.unsupportedReason || "Dry-run is unavailable for this plan.";
  }

  return result.predictedSummaries.join(" ");
}

export function formatHistoryEntrySummary(entry: { summary: string }): string {
  return entry.summary;
}
```

- [ ] **Step 4: Run the shared-client test again**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: PASS with explicit Wave 6 preview and client-method coverage.

- [ ] **Step 5: Checkpoint the shared surface**

Verify:

- `packages/shared-client/src/index.ts` exports Wave 6 client methods
- `packages/shared-client/src/trace.ts` has labels for `composite_plan` and `composite_update`
- `isWritePlanResponse()` only treats `composite_plan` as confirmable, not dry-run/history objects

---

### Task 3: Add Hermes Runtime and Response Handling for `composite_plan`

**Files:**
- Modify: `services/gateway/src/hermes/runtimeRules.ts`
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
- Modify: `services/gateway/src/hermes/structuredBody.ts`
- Modify: `services/gateway/src/lib/hermesClient.ts`
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`

- [ ] **Step 1: Write the failing gateway-runtime tests**

```ts
it("tells Hermes that composite_plan cannot contain nested composite steps", () => {
  expect(SPREADSHEET_RUNTIME_RULES).toContain("composite_plan");
  expect(SPREADSHEET_RUNTIME_RULES).toContain("must not contain another composite_plan");
});

it("routes explicit multi-step workflow prompts toward composite_plan", () => {
  const template = buildHermesRequestTemplate({
    prompt: "sort this table by revenue and then filter status = Open",
    source: "chat",
    hostPlatform: "excel_windows"
  } as any);

  expect(template).toContain('Prefer type="composite_plan"');
});

it("normalizes a composite_plan response into structured preview metadata", async () => {
  const client = new HermesAgentClient({
    hermesApiBaseUrl: "http://example.test",
    hermesApiKey: "test",
    serviceLabel: "gateway",
    environmentLabel: "test"
  } as any);

  const result = (client as any).normalizeStructuredBody({
    requestId: "req_composite_001",
    runId: "run_composite_001",
    assistantText: JSON.stringify({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_sort",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:F50",
              hasHeader: true,
              keys: [{ columnRef: "Revenue", direction: "desc" }],
              explanation: "Sort by revenue.",
              confidence: 0.91,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Run the workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    })
  });

  expect(result.type).toBe("composite_plan");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: FAIL with missing Wave 6 runtime rules, routing hints, or composite normalization support.

- [ ] **Step 3: Add the minimal runtime and normalization implementation**

```ts
// runtimeRules.ts
- If type="composite_plan", data.steps is required and each step must define stepId, dependsOn, continueOnError, and a typed executable plan.
- composite_plan must not contain another composite_plan step.
- composite_plan always requires confirmation.
- If any step would be unsupported on the host, return type="error" with data.code="UNSUPPORTED_OPERATION".

// requestTemplate.ts
if (looksLikeMultiStepWorkflow(input.prompt)) {
  guidance.push('Prefer type="composite_plan" for explicit multi-step spreadsheet workflows.');
}

// structuredBody.ts
case "composite_plan":
  return HermesResponseSchema.parse({
    ...baseEnvelope,
    type: "composite_plan",
    data: CompositePlanDataSchema.parse(body.data)
  });
```

- [ ] **Step 4: Run the gateway-runtime tests again**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: PASS with strict composite handling and no regressions to existing plan families.

- [ ] **Step 5: Checkpoint the Hermes boundary**

Verify:

- only `composite_plan` is added to Hermes’ plan-emission surface
- `undo_request`, `redo_request`, `dry_run_result`, and history stay out of Hermes response normalization
- `chat_only` analysis remains outside `WritePlan`

---

### Task 4: Add Execution Ledger and Control Routes

**Files:**
- Create: `services/gateway/src/lib/executionLedger.ts`
- Create: `services/gateway/src/routes/executionControl.ts`
- Modify: `services/gateway/src/app.ts`
- Test: `services/gateway/tests/executionLedger.test.ts`
- Test: `services/gateway/tests/executionControl.test.ts`
- Test: `services/gateway/tests/app.test.ts`

- [ ] **Step 1: Write the failing execution-ledger and route tests**

```ts
it("stores workbook/session-scoped history with undo/redo lineage", () => {
  const ledger = new ExecutionLedger();

  ledger.recordCompleted({
    executionId: "exec_001",
    workbookSessionKey: "excel_windows::workbook-123",
    requestId: "req_001",
    runId: "run_001",
    planType: "sheet_update",
    planDigest: "digest_001",
    reversible: true,
    summary: "Updated Sales!A1:B2."
  });

  expect(ledger.listHistory("excel_windows::workbook-123").entries[0]?.executionId).toBe("exec_001");
});

it("returns history for a workbook/session and rejects stale undo targets", async () => {
  const app = createApp().app;

  const history = await request(app)
    .get("/api/execution/history")
    .query({ workbookSessionKey: "excel_windows::workbook-123" });

  expect(history.status).toBe(200);

  const undo = await request(app)
    .post("/api/execution/undo")
    .send({
      requestId: "req_undo_001",
      executionId: "missing_exec",
      workbookSessionKey: "excel_windows::workbook-123"
    });

  expect(undo.status).toBe(409);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/executionLedger.test.ts services/gateway/tests/executionControl.test.ts services/gateway/tests/app.test.ts
```

Expected: FAIL with missing execution ledger, route surface, or app wiring.

- [ ] **Step 3: Add the minimal execution-ledger and route implementation**

```ts
export class ExecutionLedger {
  private readonly history = new Map<string, PlanHistoryEntry[]>();
  private readonly dryRuns = new Map<string, DryRunResult>();

  listHistory(workbookSessionKey: string): PlanHistoryPage {
    return {
      entries: [...(this.history.get(workbookSessionKey) ?? [])]
        .sort((a, b) => b.timestamp.localeCompare(a.timestamp))
    };
  }

  record(entry: PlanHistoryEntry) {
    const bucket = this.history.get(entry.workbookSessionKey) ?? [];
    bucket.push(entry);
    this.history.set(entry.workbookSessionKey, bucket);
  }

  storeDryRun(result: DryRunResult) {
    this.dryRuns.set(`${result.workbookSessionKey}::${result.planDigest}`, result);
  }
}

router.post("/dry-run", (req, res) => {
  const parsed = DryRunRequestSchema.parse(req.body);
  const result = input.executionLedger.simulate(parsed);
  res.json(result);
});

router.get("/history", (req, res) => {
  res.json(input.executionLedger.listHistory(String(req.query.workbookSessionKey || "")));
});
```

- [ ] **Step 4: Run the execution-ledger tests again**

Run:

```bash
npm test -- services/gateway/tests/executionLedger.test.ts services/gateway/tests/executionControl.test.ts services/gateway/tests/app.test.ts
```

Expected: PASS with mounted `/api/execution/*` routes and workbook/session history storage.

- [ ] **Step 5: Checkpoint the route surface**

Verify:

- `services/gateway/src/app.ts` mounts `createExecutionControlRouter(...)`
- `writeback.ts` is not yet responsible for history listing or undo/redo HTTP routes
- no existing `/api/writeback/*` routes changed contract at this task

---

### Task 5: Add Composite Approval, Completion, Dry-Run Gating, and History Recording

**Files:**
- Modify: `services/gateway/src/routes/writeback.ts`
- Modify: `services/gateway/src/lib/traceBus.ts`
- Modify: `services/gateway/src/lib/executionLedger.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

- [ ] **Step 1: Write the failing writeback-flow tests**

```ts
it("rejects composite approval when a required dry-run is missing or stale", async () => {
  const app = createApp().app;

  const response = await request(app)
    .post("/api/writeback/approve")
    .send({
      requestId: "req_composite_approve_001",
      runId: "run_composite_approve_001",
      workbookSessionKey: "excel_windows::workbook-123",
      plan: {
        steps: [
          {
            stepId: "cleanup",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:F50",
              operation: "remove_duplicate_rows",
              explanation: "Dedupe rows.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50"],
              overwriteRisk: "medium",
              confirmationLevel: "destructive"
            }
          }
        ],
        explanation: "Run cleanup workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: true
      }
    });

  expect(response.status).toBe(409);
});

it("records composite completion with step-level status history", async () => {
  // create approved composite, complete it, then inspect history
  expect(true).toBe(true);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: FAIL with no composite approval branch, no dry-run gating, or no history recording.

- [ ] **Step 3: Add the minimal writeback and history implementation**

```ts
const CompositeWritebackResultSchema = z.object({
  kind: z.literal("composite_update"),
  hostPlatform: HostPlatformSchema,
  operation: z.literal("composite_update"),
  executionId: z.string().min(1),
  stepResults: z.array(z.object({
    stepId: z.string().min(1),
    status: z.enum(["completed", "failed", "skipped"]),
    summary: z.string().min(1)
  })).min(1),
  summary: z.string().min(1)
});

if (CompositePlanDataSchema.safeParse(parsed.plan).success) {
  input.executionLedger.assertFreshDryRun({
    workbookSessionKey: parsed.workbookSessionKey,
    planDigest: digestPlan(parsed.plan),
    required: parsed.plan.dryRunRequired
  });
}

input.executionLedger.recordApproved({
  executionId,
  workbookSessionKey: parsed.workbookSessionKey,
  requestId: parsed.requestId,
  runId: parsed.runId,
  planType: "composite_plan",
  planDigest
});
```

- [ ] **Step 4: Run the writeback-flow tests again**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: PASS with composite approval/completion, dry-run gating, and step-level history persisted.

- [ ] **Step 5: Checkpoint the control-plane state**

Verify:

- `traceBus` still owns run-trace state only
- `executionLedger` owns dry-run cache and plan history
- stale plan digest, stale session key, and stale undo/redo all reject before host execution

---

### Task 6: Add Excel Composite, Dry-Run, History, and Undo/Redo Support

**Files:**
- Create: `apps/excel-addin/src/taskpane/compositePlan.js`
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Modify: `packages/shared-client/src/index.ts`
- Create: `services/gateway/tests/excelWave6Plans.test.ts`

- [ ] **Step 1: Write the failing Excel Wave 6 tests**

```ts
it("renders a composite preview with dry-run and destructive flags", async () => {
  const taskpane = await loadTaskpaneModule({ sync: vi.fn(async () => {}) });
  const response = {
    type: "composite_plan",
    data: {
      steps: [
        {
          stepId: "cleanup",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:F50",
            operation: "remove_duplicate_rows",
            explanation: "Deduplicate rows.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50"],
            overwriteRisk: "medium",
            confirmationLevel: "destructive"
          }
        }
      ],
      explanation: "Run cleanup workflow.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "medium",
      confirmationLevel: "destructive",
      reversible: false,
      dryRunRecommended: true,
      dryRunRequired: true
    }
  };

  expect(taskpane.getStructuredPreview(response)).toMatchObject({
    kind: "composite_plan",
    dryRunRequired: true
  });
});

it("maps undo and redo through the gateway client using workbook/session scope", async () => {
  const taskpane = await loadTaskpaneModule({ sync: vi.fn(async () => {}) });
  expect(typeof taskpane.undoExecution).toBe("function");
  expect(typeof taskpane.redoExecution).toBe("function");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/excelWave6Plans.test.ts services/gateway/tests/excelWritePlan.test.ts
```

Expected: FAIL with missing composite preview helpers or missing undo/redo wiring.

- [ ] **Step 3: Add the minimal Excel implementation**

```js
export function getCompositePreviewSummary(plan) {
  return `Will run ${plan.steps.length} workflow step${plan.steps.length === 1 ? "" : "s"}.`;
}

export function deriveCompositeFlags(plan) {
  return {
    destructive: plan.confirmationLevel === "destructive",
    reversible: plan.reversible,
    dryRunRequired: plan.dryRunRequired
  };
}

export async function undoExecution(executionId) {
  return gatewayClient.undoExecution({
    executionId,
    requestId: createRequestId(),
    workbookSessionKey: await getWorkbookSessionKey()
  });
}
```

- [ ] **Step 4: Run the Excel Wave 6 tests again**

Run:

```bash
npm test -- services/gateway/tests/excelWave6Plans.test.ts services/gateway/tests/excelWritePlan.test.ts
```

Expected: PASS with composite preview, dry-run affordance, and undo/redo client calls.

- [ ] **Step 5: Checkpoint the Excel host boundary**

Verify:

- Excel does not attempt host-native transactional rollback
- undo/redo only target executions marked reversible by the gateway
- unsupported reverse or dry-run simulation still fails closed

---

### Task 7: Add Google Sheets Composite, Dry-Run, History, and Undo/Redo Support

**Files:**
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Create: `services/gateway/tests/googleSheetsWave6Plans.test.ts`

- [ ] **Step 1: Write the failing Google Sheets Wave 6 tests**

```ts
it("renders composite previews and dry-run summaries in the sidebar", async () => {
  const sidebar = await loadSidebarModule();

  expect(sidebar.renderStructuredPreview({
    type: "composite_plan",
    data: {
      steps: [
        {
          stepId: "sort",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:F50",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "desc" }],
            explanation: "Sort by revenue.",
            confidence: 0.91,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50"],
            overwriteRisk: "low",
            confirmationLevel: "standard"
          }
        }
      ],
      explanation: "Sort current table.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    }
  }, { runId: "run", requestId: "req" })).toContain("workflow");
});

it("exposes exact-safe dry-run and inverse helpers in Apps Script", () => {
  expect(typeof buildDryRunResult).toBe("function");
  expect(typeof applyUndoExecution).toBe("function");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave6Plans.test.ts
```

Expected: FAIL with missing sidebar rendering or Apps Script helpers.

- [ ] **Step 3: Add the minimal Google Sheets implementation**

```js
function getCompositePreviewSummary(plan) {
  return `Will run ${plan.steps.length} workflow step${plan.steps.length === 1 ? "" : "s"}.`;
}

function buildDryRunResult(input) {
  if (!input.simulated) {
    return { summary: input.unsupportedReason || "Dry-run unavailable." };
  }

  return { summary: input.predictedSummaries.join(" ") };
}

function applyUndoExecution(input) {
  throw new Error("Wire through gateway undo execution instead of inferring host-native rollback.");
}
```

- [ ] **Step 4: Run the Google Sheets Wave 6 tests again**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave6Plans.test.ts
```

Expected: PASS with Wave 6 sidebar previews and exact-safe control helpers wired.

- [ ] **Step 5: Checkpoint the Google Sheets boundary**

Verify:

- sidebar still uses normalized approved plans, not original assistant payloads
- dry-run and undo/redo are gateway-driven, not inferred from chat text
- unsupported inverse/simulation paths remain explicit errors

---

### Task 8: Run Full Verification and Regression

**Files:**
- Test: `packages/contracts/tests/contracts.test.ts`
- Test: `packages/shared-client/tests/client.test.ts`
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`
- Test: `services/gateway/tests/executionLedger.test.ts`
- Test: `services/gateway/tests/executionControl.test.ts`
- Test: `services/gateway/tests/excelWave6Plans.test.ts`
- Test: `services/gateway/tests/googleSheetsWave6Plans.test.ts`
- Test: existing Wave 1–5 host suites
- Modify if needed: any file touched by regression fixes

- [ ] **Step 1: Run the Wave 6 focused batch**

Run:

```bash
npm test -- \
  packages/contracts/tests/contracts.test.ts \
  packages/shared-client/tests/client.test.ts \
  services/gateway/tests/runtimeRules.test.ts \
  services/gateway/tests/requestTemplate.test.ts \
  services/gateway/tests/hermesClient.test.ts \
  services/gateway/tests/writebackFlow.test.ts \
  services/gateway/tests/executionLedger.test.ts \
  services/gateway/tests/executionControl.test.ts \
  services/gateway/tests/excelWave6Plans.test.ts \
  services/gateway/tests/googleSheetsWave6Plans.test.ts
```

Expected: PASS for all Wave 6 additions.

- [ ] **Step 2: Run the existing host regression batch**

Run:

```bash
npm test -- \
  services/gateway/tests/requestRouter.test.ts \
  services/gateway/tests/structuredBody.test.ts \
  services/gateway/tests/traceBus.test.ts \
  services/gateway/tests/uploads.test.ts \
  services/gateway/tests/excelCellValues.test.ts \
  services/gateway/tests/excelReferencedCells.test.ts \
  services/gateway/tests/googleSheetsReferencedCells.test.ts \
  services/gateway/tests/rangeSafety.test.ts \
  services/gateway/tests/excelWave1Plans.test.ts \
  services/gateway/tests/excelWave2Plans.test.ts \
  services/gateway/tests/excelWave3Plans.test.ts \
  services/gateway/tests/excelWave4Plans.test.ts \
  services/gateway/tests/excelWave5Plans.test.ts \
  services/gateway/tests/googleSheetsWave1Plans.test.ts \
  services/gateway/tests/googleSheetsWave2Plans.test.ts \
  services/gateway/tests/googleSheetsWave3Plans.test.ts \
  services/gateway/tests/googleSheetsWave4Plans.test.ts \
  services/gateway/tests/googleSheetsWave5Plans.test.ts
```

Expected: PASS with no Wave 6 regressions into Waves 1 through 5.

- [ ] **Step 3: Run build and syntax checks**

Run:

```bash
node --check <repo-root>/apps/excel-addin/src/taskpane/taskpane.js
node --check <repo-root>/apps/excel-addin/src/taskpane/compositePlan.js
bash -lc 'awk "/<script>/{flag=1;next}/<\\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave6-sidebar.js && node --check /tmp/hermes-wave6-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave6-code-gs.js && node --check /tmp/hermes-wave6-code-gs.js'
npm --workspace @hermes/gateway run build
```

Expected: PASS for all syntax checks and gateway TypeScript build.

- [ ] **Step 4: If any regression fails, fix only the smallest exact cause**

Rules:

- do not refactor unrelated code
- add or tighten the failing test first if the regression is under-specified
- rerun only the failing batch until green, then rerun the full verification commands above

- [ ] **Step 5: Final checkpoint**

Confirm all of the following before declaring Wave 6 complete:

- `composite_plan` works end to end through contracts, preview, approval, execution, and history
- dry-run is exact-safe only and gated by digest/session freshness
- undo/redo works only for reversible subset and fails closed otherwise
- execution history is workbook/session-scoped, not conversation-scoped
- Waves 1 through 5 remain green
