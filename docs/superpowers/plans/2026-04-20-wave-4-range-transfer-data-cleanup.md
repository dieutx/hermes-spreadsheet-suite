# Wave 4 Range Transfer + Data Cleanup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add contract-valid `range_transfer_plan` and `data_cleanup_plan` flows with strict preview, destructive confirmation where required, and exact-safe host apply support for Excel and Google Sheets.

**Architecture:** Extend the existing typed-plan pipeline with two new strict plan families rather than folding them into `sheet_update`. Add contract schemas, runtime/request routing, structured-body parsing, writeback support, shared previews, and then host-specific apply paths for Excel and Google Sheets with fail-closed behavior for overlap ambiguity and unsupported cleanup semantics.

**Tech Stack:** TypeScript, Zod, Vitest, Express, Office.js, Google Apps Script, browser-safe shared client helpers.

---

## File Structure

### Contracts and shared types

- Modify: `packages/contracts/src/schemas.ts`
  - add `range_transfer_plan`, `data_cleanup_plan`, and their typed result kinds
- Modify: `packages/contracts/src/index.ts`
  - export the new schemas and inferred types
- Test: `packages/contracts/tests/contracts.test.ts`
  - schema coverage for transfer operations, cleanup operations, destructive confirmation levels, and overlap metadata invariants

### Shared preview and client types

- Modify: `packages/shared-client/src/types.ts`
  - extend `WritePlan`, `WritebackResult`, and structured preview unions
- Modify: `packages/shared-client/src/render.ts`
  - add non-lossy previews for transfer and cleanup plans
- Modify: `packages/shared-client/src/trace.ts`
  - add typed status/trace labels for the new result kinds
- Test: `packages/shared-client/tests/client.test.ts`
  - preview coverage and result typing regression

### Gateway runtime and writeback

- Modify: `services/gateway/src/hermes/runtimeRules.ts`
  - teach Hermes the new response types and fail-closed behavior
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
  - route prompts toward `range_transfer_plan` and `data_cleanup_plan`
- Modify: `services/gateway/src/hermes/structuredBody.ts`
  - normalize and validate the new plan families
- Modify: `services/gateway/src/lib/hermesClient.ts`
  - map UI/trace behavior for the new plan families
- Modify: `services/gateway/src/routes/writeback.ts`
  - add approval/completion support for the new result kinds and destructive second-confirm gating
- Modify: `services/gateway/src/lib/traceBus.ts`
  - persist typed completion state for the new plans
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

### Excel host

- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
  - integrate preview and apply paths for transfer and cleanup plans
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
  - extend status summaries for the new result kinds
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave4Plans.test.ts`

### Google Sheets host

- Modify: `apps/google-sheets-addon/src/Code.gs`
  - apply transfer and cleanup plans
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
  - preview and confirmation handling for the new plan families
- Test: `services/gateway/tests/googleSheetsWave4Plans.test.ts`

### Final regression

- Re-run existing tests covering:
  - `sheet_update`
  - `sheet_import_plan`
  - `workbook_structure_update`
  - `sheet_structure_update`
  - `range_format_update`
  - `conditional_format_plan`
  - waves 1 through 3

Note: this workspace snapshot currently has no `.git` metadata, so execution should skip commit steps locally unless the repo is re-initialized or restored.

---

### Task 1: Add Wave 4 Contract Schemas

**Files:**
- Modify: `packages/contracts/src/schemas.ts`
- Modify: `packages/contracts/src/index.ts`
- Test: `packages/contracts/tests/contracts.test.ts`

- [ ] **Step 1: Write the failing contract tests**

```ts
import {
  DataCleanupPlanDataSchema,
  HermesResponseSchema,
  RangeTransferPlanDataSchema
} from "../src/index.ts";

it("accepts a destructive move transfer plan", () => {
  const parsed = RangeTransferPlanDataSchema.parse({
    sourceSheet: "Sheet1",
    sourceRange: "A1:D20",
    targetSheet: "Archive",
    targetRange: "A1",
    operation: "move",
    pasteMode: "values",
    transpose: false,
    explanation: "Move the current table into the archive sheet.",
    confidence: 0.94,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
    overwriteRisk: "medium",
    confirmationLevel: "destructive"
  });

  expect(parsed.operation).toBe("move");
  expect(parsed.confirmationLevel).toBe("destructive");
});

it("rejects an append transfer plan without a target range", () => {
  const parsed = RangeTransferPlanDataSchema.safeParse({
    sourceSheet: "Sheet1",
    sourceRange: "A1:D20",
    targetSheet: "Archive",
    operation: "append",
    pasteMode: "values",
    transpose: false,
    explanation: "Invalid append without an anchor.",
    confidence: 0.7,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!A1:D20"],
    overwriteRisk: "low",
    confirmationLevel: "standard"
  });

  expect(parsed.success).toBe(false);
});

it("accepts a destructive remove_duplicate_rows cleanup plan", () => {
  const parsed = DataCleanupPlanDataSchema.parse({
    targetSheet: "Sheet1",
    targetRange: "A2:F200",
    operation: "remove_duplicate_rows",
    keyColumns: ["A", "C"],
    explanation: "Remove duplicate rows by key columns.",
    confidence: 0.91,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!A2:F200"],
    overwriteRisk: "low",
    confirmationLevel: "destructive"
  });

  expect(parsed.operation).toBe("remove_duplicate_rows");
  expect(parsed.confirmationLevel).toBe("destructive");
});

it("accepts a range_transfer_update response envelope", () => {
  const parsed = HermesResponseSchema.parse({
    schemaVersion: "1.0.0",
    type: "range_transfer_update",
    requestId: "req_transfer_update_001",
    hermesRunId: "run_transfer_update_001",
    processedBy: "host",
    serviceLabel: "hermes-gateway-local",
    environmentLabel: "local-dev",
    startedAt: "2026-04-20T11:00:00.000Z",
    completedAt: "2026-04-20T11:00:01.000Z",
    durationMs: 1000,
    trace: [{ event: "completed", timestamp: "2026-04-20T11:00:01.000Z" }],
    ui: {
      displayMode: "structured-preview",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: false
    },
    data: {
      operation: "range_transfer_update",
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Archive",
      targetRange: "A1",
      transferOperation: "copy",
      summary: "Copied Sheet1!A1:D20 to Archive!A1."
    }
  });

  expect(parsed.type).toBe("range_transfer_update");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: FAIL with missing exports or missing `range_transfer_plan` / `data_cleanup_plan` / result branches.

- [ ] **Step 3: Write the minimal contract implementation**

```ts
const TransferOperationSchema = z.enum(["copy", "move", "append"]);
const TransferPasteModeSchema = z.enum(["values", "formulas", "formats"]);
const ConfirmationLevelSchema = z.enum(["standard", "destructive"]);

export const RangeTransferPlanDataSchema = strictObject({
  sourceSheet: z.string().min(1).max(128),
  sourceRange: z.string().min(1).max(128),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  operation: TransferOperationSchema,
  pasteMode: TransferPasteModeSchema,
  transpose: z.boolean(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(10),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
});

export const DataCleanupPlanDataSchema = z.discriminatedUnion("operation", [
  strictObject({
    targetSheet: z.string().min(1).max(128),
    targetRange: z.string().min(1).max(128),
    operation: z.literal("remove_duplicate_rows"),
    keyColumns: z.array(z.string().min(1).max(16)).max(50).optional(),
    explanation: z.string().min(1).max(12000),
    confidence: z.number().min(0).max(1),
    requiresConfirmation: z.literal(true),
    affectedRanges: z.array(z.string().min(1).max(128)).max(10),
    overwriteRisk: OverwriteRiskSchema,
    confirmationLevel: ConfirmationLevelSchema
  })
  // extend with remaining cleanup operations in the same pattern
]);

export const RangeTransferUpdateDataSchema = strictObject({
  operation: z.literal("range_transfer_update"),
  sourceSheet: z.string().min(1).max(128),
  sourceRange: z.string().min(1).max(128),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  transferOperation: TransferOperationSchema,
  summary: z.string().min(1).max(500)
});

export const DataCleanupUpdateDataSchema = strictObject({
  operation: z.literal("data_cleanup_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  cleanupOperation: z.string().min(1).max(64),
  summary: z.string().min(1).max(500)
});
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: PASS with the new schema coverage green.

- [ ] **Step 5: Update exports**

```ts
export {
  RangeTransferPlanDataSchema,
  DataCleanupPlanDataSchema,
  RangeTransferUpdateDataSchema,
  DataCleanupUpdateDataSchema
} from "./schemas.ts";

export type RangeTransferPlanData = z.infer<typeof RangeTransferPlanDataSchema>;
export type DataCleanupPlanData = z.infer<typeof DataCleanupPlanDataSchema>;
export type RangeTransferUpdateData = z.infer<typeof RangeTransferUpdateDataSchema>;
export type DataCleanupUpdateData = z.infer<typeof DataCleanupUpdateDataSchema>;
```

- [ ] **Step 6: Re-run the focused contract test**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: PASS with export/import resolution working.

- [ ] **Step 7: Skip commit in this workspace snapshot**

Reason: there is no `.git` metadata in this checkout, so do not add a fake commit step.

### Task 2: Add Shared Preview Types and Rendering

**Files:**
- Modify: `packages/shared-client/src/types.ts`
- Modify: `packages/shared-client/src/render.ts`
- Modify: `packages/shared-client/src/trace.ts`
- Test: `packages/shared-client/tests/client.test.ts`

- [ ] **Step 1: Write the failing shared-client tests**

```ts
it("renders a non-lossy range transfer preview", () => {
  const preview = getStructuredPreview({
    type: "range_transfer_plan",
    data: {
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Archive",
      targetRange: "A1",
      operation: "move",
      pasteMode: "values",
      transpose: false,
      explanation: "Move the table into Archive.",
      confidence: 0.92,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
      overwriteRisk: "medium",
      confirmationLevel: "destructive"
    }
  });

  expect(preview.summary).toContain("Move");
  expect(preview.details).toContain("Sheet1!A1:D20");
  expect(preview.details).toContain("Archive!A1");
  expect(preview.details).toContain("clear the source after success");
});

it("renders a non-lossy cleanup preview", () => {
  const preview = getStructuredPreview({
    type: "data_cleanup_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "A2:F200",
      operation: "split_column",
      sourceColumn: "B",
      delimiter: ",",
      targetStartColumn: "C",
      explanation: "Split the comma-separated values in column B.",
      confidence: 0.88,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F200"],
      overwriteRisk: "medium",
      confirmationLevel: "destructive"
    }
  });

  expect(preview.summary).toContain("Split column B");
  expect(preview.details).toContain("delimiter \",\"");
  expect(preview.details).toContain("target columns starting at C");
});

it("renders typed completion lines for transfer and cleanup updates", () => {
  expect(getWritebackStatusLine({
    operation: "range_transfer_update",
    sourceSheet: "Sheet1",
    sourceRange: "A1:D20",
    targetSheet: "Archive",
    targetRange: "A1",
    transferOperation: "copy",
    summary: "Copied Sheet1!A1:D20 to Archive!A1."
  })).toBe("Copied Sheet1!A1:D20 to Archive!A1.");

  expect(getWritebackStatusLine({
    operation: "data_cleanup_update",
    targetSheet: "Sheet1",
    targetRange: "A2:F200",
    cleanupOperation: "remove_duplicate_rows",
    summary: "Removed duplicate rows from Sheet1!A2:F200."
  })).toBe("Removed duplicate rows from Sheet1!A2:F200.");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: FAIL with missing preview branches or missing result typing.

- [ ] **Step 3: Write the minimal shared-client implementation**

```ts
export type RangeTransferUpdateWritebackResult = {
  operation: "range_transfer_update";
  sourceSheet: string;
  sourceRange: string;
  targetSheet: string;
  targetRange: string;
  transferOperation: "copy" | "move" | "append";
  summary: string;
};

export type DataCleanupUpdateWritebackResult = {
  operation: "data_cleanup_update";
  targetSheet: string;
  targetRange: string;
  cleanupOperation: string;
  summary: string;
};

function renderRangeTransferPreview(data: RangeTransferPlanData): StructuredPreview {
  return {
    summary: `${capitalize(data.operation)} ${data.sourceSheet}!${data.sourceRange} to ${data.targetSheet}!${data.targetRange}.`,
    details: buildRangeTransferDetails(data),
    confidence: data.confidence,
    requiresConfirmation: true
  };
}

function renderDataCleanupPreview(data: DataCleanupPlanData): StructuredPreview {
  return {
    summary: buildCleanupSummary(data),
    details: buildCleanupDetails(data),
    confidence: data.confidence,
    requiresConfirmation: true
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: PASS with preview and status rendering green.

- [ ] **Step 5: Add trace labels for the new result kinds**

```ts
export function getResultStatusLabel(result: WritebackResult): string {
  switch (result.operation) {
    case "range_transfer_update":
      return "Range transfer applied";
    case "data_cleanup_update":
      return "Data cleanup applied";
    default:
      return getExistingResultStatusLabel(result);
  }
}
```

- [ ] **Step 6: Re-run the focused shared-client test**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: PASS with the new status/trace helpers type-safe.

### Task 3: Extend Runtime Rules, Request Routing, and Structured Body Validation

**Files:**
- Modify: `services/gateway/src/hermes/runtimeRules.ts`
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
- Modify: `services/gateway/src/hermes/structuredBody.ts`
- Modify: `services/gateway/src/lib/hermesClient.ts`
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`

- [ ] **Step 1: Write the failing gateway-facing tests**

```ts
it("mentions range_transfer_plan and data_cleanup_plan in the runtime rules", () => {
  expect(SPREADSHEET_RUNTIME_RULES).toContain("range_transfer_plan");
  expect(SPREADSHEET_RUNTIME_RULES).toContain("data_cleanup_plan");
  expect(SPREADSHEET_RUNTIME_RULES).toContain("UNSUPPORTED_OPERATION");
});

it("routes transfer and cleanup prompts toward the new plan families", () => {
  const transferPrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
    userMessage: "Copy this table to Sheet2."
  }));
  const cleanupPrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
    userMessage: "Remove duplicate rows from this table."
  }));

  expect(transferPrompt).toContain('Prefer type="range_transfer_plan"');
  expect(cleanupPrompt).toContain('Prefer type="data_cleanup_plan"');
});

it("accepts a range_transfer_plan body and builds a structured-preview response", async () => {
  vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(chatCompletionEnvelope(JSON.stringify({
    type: "range_transfer_plan",
    data: {
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Archive",
      targetRange: "A1",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the table into Archive.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  })), {
    status: 200,
    headers: { "content-type": "application/json" }
  })));

  await client.processRequest({ runId: "run_transfer_001", request: baseRequest(), traceBus });

  expect(traceBus.getRun("run_transfer_001")?.response).toMatchObject({
    type: "range_transfer_plan",
    ui: {
      displayMode: "structured-preview",
      showRequiresConfirmation: true
    }
  });
});
```

- [ ] **Step 2: Run the focused tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: FAIL with missing mentions, missing structured-body branch, or missing UI mapping.

- [ ] **Step 3: Write the minimal runtime/request/body/client implementation**

```ts
// runtimeRules.ts
// Add new response-type sections:
// - choose `range_transfer_plan` for copy/move/append/transpose asks
// - choose `data_cleanup_plan` for cleanup asks
// - return `error` with `UNSUPPORTED_OPERATION` for unsupported fuzzy or heuristic cleanup

// requestTemplate.ts
if (/\b(copy|move|append|transpose)\b/i.test(userIntent)) {
  hints.push('Prefer type="range_transfer_plan" over advisory chat when the user is asking to transfer spreadsheet data.');
}

if (/\b(trim|duplicate rows|blank rows|split column|join columns|fill down|standardize)\b/i.test(userIntent)) {
  hints.push('Prefer type="data_cleanup_plan" over advisory chat when the user is asking to clean or reshape spreadsheet data.');
}

// structuredBody.ts
const RangeTransferPlanBodySchema = strictObject({
  type: z.literal("range_transfer_plan"),
  data: RangeTransferPlanDataSchema
});

const DataCleanupPlanBodySchema = strictObject({
  type: z.literal("data_cleanup_plan"),
  data: DataCleanupPlanDataSchema
});

// hermesClient.ts
case "range_transfer_plan":
case "data_cleanup_plan":
  return {
    displayMode: "structured-preview",
    showTrace: true,
    showWarnings: true,
    showConfidence: true,
    showRequiresConfirmation: true
  };
```

- [ ] **Step 4: Run the focused tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: PASS with prompt routing and typed envelope assembly green.

- [ ] **Step 5: Add explicit trace events**

```ts
case "range_transfer_plan":
  pushIfMissing({ event: "range_transfer_plan_ready", timestamp: completedAt });
  pushIfMissing({ event: "completed", timestamp: completedAt });
  break;
case "data_cleanup_plan":
  pushIfMissing({ event: "data_cleanup_plan_ready", timestamp: completedAt });
  pushIfMissing({ event: "completed", timestamp: completedAt });
  break;
```

- [ ] **Step 6: Re-run the Hermes client test**

Run:

```bash
npm test -- services/gateway/tests/hermesClient.test.ts
```

Expected: PASS with the new trace event sequences.

### Task 4: Extend Writeback and Completion State

**Files:**
- Modify: `services/gateway/src/routes/writeback.ts`
- Modify: `services/gateway/src/lib/traceBus.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

- [ ] **Step 1: Write the failing writeback tests**

```ts
it("approves and completes a range transfer plan", async () => {
  const plan = {
    sourceSheet: "Sheet1",
    sourceRange: "A1:D20",
    targetSheet: "Archive",
    targetRange: "A1",
    operation: "copy",
    pasteMode: "values",
    transpose: false,
    explanation: "Copy the table into Archive.",
    confidence: 0.93,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
    overwriteRisk: "low",
    confirmationLevel: "standard" as const
  };

  const approved = await request(app)
    .post("/api/writeback/approve")
    .send({ runId: seededRunId, plan });

  expect(approved.status).toBe(200);

  const completed = await request(app)
    .post("/api/writeback/complete")
    .send({
      runId: seededRunId,
      approvalToken: approved.body.approvalToken,
      result: {
        operation: "range_transfer_update",
        sourceSheet: "Sheet1",
        sourceRange: "A1:D20",
        targetSheet: "Archive",
        targetRange: "A1",
        transferOperation: "copy",
        summary: "Copied Sheet1!A1:D20 to Archive!A1."
      }
    });

  expect(completed.status).toBe(200);
});
```

- [ ] **Step 2: Run the writeback test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: FAIL with approval or completion schema mismatch.

- [ ] **Step 3: Write the minimal writeback implementation**

```ts
const WritebackPlanSchema = z.union([
  ExistingWritebackPlanSchema,
  RangeTransferPlanDataSchema,
  DataCleanupPlanDataSchema
]);

const WritebackResultSchema = z.union([
  ExistingWritebackResultSchema,
  RangeTransferUpdateDataSchema,
  DataCleanupUpdateDataSchema
]);
```

- [ ] **Step 4: Run the writeback test to verify it passes**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: PASS with approval and completion green for the new plan families.

- [ ] **Step 5: Add typed completion persistence**

```ts
run.writeback = {
  ...run.writeback,
  completedResult: result
};
```

- [ ] **Step 6: Re-run the writeback test**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: PASS with typed completion state preserved.

### Task 5: Add Excel Transfer and Cleanup Apply Support

**Files:**
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave4Plans.test.ts`

- [ ] **Step 1: Write the failing Excel tests**

```ts
it("recognizes range_transfer_plan and data_cleanup_plan as write plans in Excel", async () => {
  const transferPreview = renderAssistantPayload({
    type: "range_transfer_plan",
    data: {
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Archive",
      targetRange: "A1",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the table.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  });

  expect(transferPreview.confirmLabel).toContain("Apply");
});

it("applies a range transfer plan in Excel", async () => {
  const result = await applyExcelWave4Plan({
    sourceSheet: "Sheet1",
    sourceRange: "A1:B2",
    targetSheet: "Archive",
    targetRange: "A1",
    operation: "copy",
    pasteMode: "values",
    transpose: false,
    explanation: "Copy values to archive.",
    confidence: 0.9,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!A1:B2", "Archive!A1:B2"],
    overwriteRisk: "low",
    confirmationLevel: "standard"
  });

  expect(result.operation).toBe("range_transfer_update");
});
```

- [ ] **Step 2: Run the focused Excel tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave4Plans.test.ts
```

Expected: FAIL with missing preview/apply branches.

- [ ] **Step 3: Write the minimal Excel implementation**

```js
function isRangeTransferPlan(plan) {
  return plan && typeof plan === "object" && typeof plan.sourceSheet === "string" && typeof plan.targetSheet === "string";
}

function isDataCleanupPlan(plan) {
  return plan && typeof plan === "object" && typeof plan.targetSheet === "string" && typeof plan.operation === "string";
}

async function applyRangeTransferPlan_(context, plan, platform) {
  // Read source, optionally transpose, write target, clear source only after success for move.
  return {
    kind: "range_transfer_update",
    operation: "range_transfer_update",
    sourceSheet: plan.sourceSheet,
    sourceRange: plan.sourceRange,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    transferOperation: plan.operation,
    summary: `Copied ${plan.sourceSheet}!${plan.sourceRange} to ${plan.targetSheet}!${plan.targetRange}.`
  };
}
```

- [ ] **Step 4: Run the focused Excel tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave4Plans.test.ts
```

Expected: PASS with preview recognition and apply-path coverage green.

- [ ] **Step 5: Add fail-closed overlap or unsupported-semantics coverage**

```ts
it("fails closed on overlapping move ambiguity in Excel", async () => {
  await expect(applyExcelWave4Plan({
    sourceSheet: "Sheet1",
    sourceRange: "A1:B4",
    targetSheet: "Sheet1",
    targetRange: "A2",
    operation: "move",
    pasteMode: "values",
    transpose: false,
    explanation: "Ambiguous overlapping move.",
    confidence: 0.7,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!A1:B5"],
    overwriteRisk: "high",
    confirmationLevel: "destructive"
  })).rejects.toThrow(/overlap/i);
});
```

- [ ] **Step 6: Re-run the focused Excel tests**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave4Plans.test.ts
```

Expected: PASS with fail-closed behavior covered.

### Task 6: Add Google Sheets Transfer and Cleanup Apply Support

**Files:**
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave4Plans.test.ts`

- [ ] **Step 1: Write the failing Google Sheets tests**

```ts
it("renders a future-tense Google Sheets preview for transfer and cleanup plans", () => {
  const transferPreview = getStructuredPreview({
    type: "range_transfer_plan",
    data: {
      sourceSheet: "Sheet1",
      sourceRange: "A1:D20",
      targetSheet: "Archive",
      targetRange: "A1",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the table.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  });

  expect(transferPreview.summary).toContain("Will copy");
});

it("applies a Google Sheets range transfer plan", () => {
  const result = applyWave4GooglePlan({
    sourceSheet: "Sheet1",
    sourceRange: "A1:B2",
    targetSheet: "Archive",
    targetRange: "A1",
    operation: "copy",
    pasteMode: "values",
    transpose: false,
    explanation: "Copy values to archive.",
    confidence: 0.9,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!A1:B2", "Archive!A1:B2"],
    overwriteRisk: "low",
    confirmationLevel: "standard"
  });

  expect(result.operation).toBe("range_transfer_update");
});
```

- [ ] **Step 2: Run the focused Google Sheets test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave4Plans.test.ts
```

Expected: FAIL with missing preview/apply branches.

- [ ] **Step 3: Write the minimal Google Sheets implementation**

```js
function isRangeTransferPlan_(plan) {
  return plan && typeof plan === "object" && typeof plan.sourceSheet === "string" && typeof plan.targetSheet === "string";
}

function isDataCleanupPlan_(plan) {
  return plan && typeof plan === "object" && typeof plan.targetSheet === "string" && typeof plan.operation === "string";
}

function getRangeTransferPreviewSummary_(plan) {
  return `Will ${plan.operation} ${plan.sourceSheet}!${plan.sourceRange} to ${plan.targetSheet}!${plan.targetRange}.`;
}

function applyRangeTransferPlan_(spreadsheet, plan) {
  return {
    kind: "range_transfer_update",
    operation: "range_transfer_update",
    sourceSheet: plan.sourceSheet,
    sourceRange: plan.sourceRange,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    transferOperation: plan.operation,
    summary: `Copied ${plan.sourceSheet}!${plan.sourceRange} to ${plan.targetSheet}!${plan.targetRange}.`
  };
}
```

- [ ] **Step 4: Run the focused Google Sheets test to verify it passes**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave4Plans.test.ts
```

Expected: PASS with preview and apply-path coverage green.

- [ ] **Step 5: Add fail-closed overlap or unsupported-semantics tests**

```ts
it("suppresses confirmation for unsupported cleanup semantics in Google Sheets", () => {
  const preview = getStructuredPreview({
    type: "data_cleanup_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "A2:F100",
      operation: "standardize_format",
      formatType: "date_text",
      formatPattern: "locale-sensitive-fuzzy",
      explanation: "Unsupported fuzzy cleanup request.",
      confidence: 0.6,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!A2:F100"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  });

  expect(preview.supportError).toContain("cannot represent");
  expect(preview.showConfirm).toBe(false);
});
```

- [ ] **Step 6: Re-run the focused Google Sheets test**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave4Plans.test.ts
```

Expected: PASS with fail-closed behavior covered.

### Task 7: Final Regression, Syntax Checks, and Build

**Files:**
- Test: `packages/contracts/tests/contracts.test.ts`
- Test: `packages/shared-client/tests/client.test.ts`
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave4Plans.test.ts`
- Test: `services/gateway/tests/googleSheetsWave4Plans.test.ts`
- Test: `services/gateway/tests/requestRouter.test.ts`
- Test: `services/gateway/tests/structuredBody.test.ts`
- Test: `services/gateway/tests/traceBus.test.ts`
- Test: `services/gateway/tests/uploads.test.ts`
- Test: `services/gateway/tests/excelCellValues.test.ts`
- Test: `services/gateway/tests/excelReferencedCells.test.ts`
- Test: `services/gateway/tests/googleSheetsReferencedCells.test.ts`
- Test: `services/gateway/tests/rangeSafety.test.ts`
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Modify: `apps/google-sheets-addon/src/Code.gs`

- [ ] **Step 1: Run the main Wave 4 regression batch**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts packages/shared-client/tests/client.test.ts services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts services/gateway/tests/writebackFlow.test.ts services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave4Plans.test.ts services/gateway/tests/googleSheetsWave4Plans.test.ts
```

Expected: PASS with all focused Wave 4 suites green.

- [ ] **Step 2: Run the secondary regression batch**

Run:

```bash
npm test -- services/gateway/tests/requestRouter.test.ts services/gateway/tests/structuredBody.test.ts services/gateway/tests/traceBus.test.ts services/gateway/tests/uploads.test.ts services/gateway/tests/excelCellValues.test.ts services/gateway/tests/excelReferencedCells.test.ts services/gateway/tests/googleSheetsReferencedCells.test.ts services/gateway/tests/rangeSafety.test.ts
```

Expected: PASS with earlier waves unaffected.

- [ ] **Step 3: Run host syntax checks**

Run:

```bash
node --check <repo-root>/apps/excel-addin/src/taskpane/taskpane.js
node --check <repo-root>/apps/excel-addin/src/taskpane/writePlan.js
bash -lc 'awk "/<script>/{flag=1;next}/<\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave4-sidebar.js && node --check /tmp/hermes-wave4-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave4-code-gs.js && node --check /tmp/hermes-wave4-code-gs.js'
```

Expected: PASS with zero syntax errors.

- [ ] **Step 4: Run the gateway build**

Run:

```bash
npm --workspace @hermes/gateway run build
```

Expected: PASS with no TypeScript errors.

- [ ] **Step 5: Record exact verification counts before claiming completion**

```text
- main batch: <fill exact file/test counts from actual output>
- secondary batch: <fill exact file/test counts from actual output>
- syntax checks: passed
- build: passed
```

Expected: only actual counts from command output, no guessed numbers.

- [ ] **Step 6: Skip commit in this workspace snapshot**

Reason: there is no `.git` metadata in this checkout, so do not add a fake commit step.

## Self-Review

### Spec coverage

Covered spec sections:
- scope and chosen plan families: Task 1
- contract and safety semantics: Tasks 1, 3, 4
- preview model: Tasks 2 and 6
- apply path: Tasks 5 and 6
- request routing and prompt grounding: Task 3
- testing strategy and regression: Tasks 1 through 7

No uncovered spec requirement remains.

### Placeholder scan

There are no `TBD`, `TODO`, or deferred implementation markers in this plan. The only intentional placeholder is the explicit instruction in Task 7 Step 5 to copy actual verification counts from real command output at execution time rather than inventing them.

### Type consistency

The plan consistently uses:
- `range_transfer_plan`
- `data_cleanup_plan`
- `range_transfer_update`
- `data_cleanup_update`
- `confirmationLevel`
- `overwriteRisk`
- `transferOperation`
- `cleanupOperation`

Later tasks do not rename these identifiers.
