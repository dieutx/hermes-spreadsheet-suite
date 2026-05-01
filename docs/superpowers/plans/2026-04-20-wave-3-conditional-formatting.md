# Wave 3 Conditional Formatting Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add contract-valid `conditional_format_plan` and `conditional_format_update` flows with strict preview, approval, and host apply support for Excel and Google Sheets.

**Architecture:** Extend the existing typed-plan pipeline with one new strict plan family instead of merging conditional formatting into `range_format_update`. Add contract schemas, runtime/request routing, structured-body parsing, writeback support, shared previews, then host-specific apply paths for Excel and Google Sheets with fail-closed semantics for unsupported exact behavior.

**Tech Stack:** TypeScript, Zod, Vitest, Express, Office.js, Google Apps Script, browser-safe shared client helpers.

---

## File Structure

### Contracts and shared types

- Modify: `packages/contracts/src/schemas.ts`
  - add `conditional_format_plan`, `conditional_format_update`, rule/style schemas, and result kinds
- Modify: `packages/contracts/src/index.ts`
  - export the new schemas and inferred types
- Test: `packages/contracts/tests/contracts.test.ts`
  - schema coverage for rule variants, management modes, comparators, color scales, and clear-mode invariants

### Shared preview and client types

- Modify: `packages/shared-client/src/types.ts`
  - extend `WritePlan`, `WritebackResult`, and structured preview unions
- Modify: `packages/shared-client/src/render.ts`
  - add non-lossy previews for conditional formatting
- Modify: `packages/shared-client/src/trace.ts`
  - add typed status/trace labels for the new result kind
- Test: `packages/shared-client/tests/client.test.ts`
  - preview coverage and result typing regression

### Gateway runtime and writeback

- Modify: `services/gateway/src/hermes/runtimeRules.ts`
  - teach Hermes the new response type and fail-closed behavior
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
  - route prompts toward `conditional_format_plan`
- Modify: `services/gateway/src/hermes/structuredBody.ts`
  - normalize and validate the new plan family
- Modify: `services/gateway/src/lib/hermesClient.ts`
  - map UI/trace behavior for the new plan family
- Modify: `services/gateway/src/routes/writeback.ts`
  - add approval/completion support for `conditional_format_update`
- Modify: `services/gateway/src/lib/traceBus.ts`
  - persist typed completion state for the new plan family
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

### Excel host

- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
  - integrate preview and apply paths for conditional formatting
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
  - extend status summaries for `conditional_format_update`
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave3Plans.test.ts`

### Google Sheets host

- Modify: `apps/google-sheets-addon/src/Code.gs`
  - apply conditional-format plans
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
  - preview and confirmation handling for conditional formatting
- Test: `services/gateway/tests/googleSheetsWave3Plans.test.ts`

### Final regression

- Re-run existing tests covering:
  - `sheet_update`
  - `sheet_import_plan`
  - `workbook_structure_update`
  - `range_format_update`
  - wave 1 plan families
  - wave 2 plan families

Note: this workspace snapshot currently has no `.git` metadata, so execution should skip commit steps locally unless the repo is re-initialized or restored.

---

### Task 1: Add Wave 3 Contract Schemas

**Files:**
- Modify: `packages/contracts/src/schemas.ts`
- Modify: `packages/contracts/src/index.ts`
- Test: `packages/contracts/tests/contracts.test.ts`

- [ ] **Step 1: Write the failing contract tests**

```ts
import {
  ConditionalFormatPlanDataSchema,
  ConditionalFormatUpdateDataSchema,
  HermesResponseSchema
} from "../src/index.ts";

it("accepts a replace-all single-color conditional format plan", () => {
  const parsed = ConditionalFormatPlanDataSchema.parse({
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    managementMode: "replace_all_on_target",
    ruleType: "number_compare",
    comparator: "greater_than",
    value: 10,
    style: {
      backgroundColor: "#ffdddd",
      textColor: "#990000",
      bold: true
    },
    explanation: "Highlight values above 10.",
    confidence: 0.96,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!B2:B20"],
    replacesExistingRules: true
  });

  expect(parsed.managementMode).toBe("replace_all_on_target");
  expect(parsed.ruleType).toBe("number_compare");
});

it("rejects clear_on_target plans that carry rule payload", () => {
  const parsed = ConditionalFormatPlanDataSchema.safeParse({
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    managementMode: "clear_on_target",
    ruleType: "text_contains",
    text: "overdue",
    style: {
      backgroundColor: "#ffeeaa"
    },
    explanation: "Invalid clear plan.",
    confidence: 0.4,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!B2:B20"],
    replacesExistingRules: true
  });

  expect(parsed.success).toBe(false);
});

it("accepts a 3-color scale plan", () => {
  const parsed = ConditionalFormatPlanDataSchema.parse({
    targetSheet: "Summary",
    targetRange: "A2:D20",
    managementMode: "add",
    ruleType: "color_scale",
    points: [
      { type: "min", color: "#f8696b" },
      { type: "percentile", value: 50, color: "#ffeb84" },
      { type: "max", color: "#63be7b" }
    ],
    explanation: "Apply a 3-color scale.",
    confidence: 0.91,
    requiresConfirmation: true,
    affectedRanges: ["Summary!A2:D20"],
    replacesExistingRules: false
  });

  expect(parsed.points).toHaveLength(3);
});

it("accepts a conditional_format_update response envelope", () => {
  const parsed = HermesResponseSchema.parse({
    schemaVersion: "1.0.0",
    type: "conditional_format_update",
    requestId: "req_cf_update_001",
    hermesRunId: "run_cf_update_001",
    processedBy: "host",
    serviceLabel: "hermes-gateway-local",
    environmentLabel: "local-dev",
    startedAt: "2026-04-20T10:00:00.000Z",
    completedAt: "2026-04-20T10:00:01.000Z",
    durationMs: 1000,
    trace: [{ event: "completed", timestamp: "2026-04-20T10:00:01.000Z" }],
    ui: {
      displayMode: "structured-preview",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: false
    },
    data: {
      operation: "conditional_format_update",
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      summary: "Added conditional formatting to Sheet1!B2:B20."
    }
  });

  expect(parsed.type).toBe("conditional_format_update");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: FAIL with missing exports or missing `conditional_format_plan` / `conditional_format_update` branches.

- [ ] **Step 3: Write the minimal contract implementation**

```ts
const ConditionalFormatManagementModeSchema = z.enum([
  "add",
  "replace_all_on_target",
  "clear_on_target"
]);

const ConditionalFormatComparatorSchema = z.enum([
  "between",
  "not_between",
  "equal_to",
  "not_equal_to",
  "greater_than",
  "greater_than_or_equal_to",
  "less_than",
  "less_than_or_equal_to"
]);

const ConditionalFormatStyleSchema = strictObject({
  backgroundColor: HexColorSchema.optional(),
  textColor: HexColorSchema.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  numberFormat: z.string().min(1).max(128).optional()
});

const ColorScalePointSchema = strictObject({
  type: z.enum(["min", "max", "number", "percent", "percentile"]),
  value: z.number().optional(),
  color: HexColorSchema
});

export const ConditionalFormatPlanDataSchema = z.discriminatedUnion("ruleType", [
  strictObject({
    targetSheet: z.string().min(1).max(128),
    targetRange: z.string().min(1).max(128),
    managementMode: ConditionalFormatManagementModeSchema,
    ruleType: z.literal("number_compare"),
    comparator: ConditionalFormatComparatorSchema,
    value: z.number(),
    value2: z.number().optional(),
    style: ConditionalFormatStyleSchema,
    explanation: z.string().min(1).max(12000),
    confidence: z.number().min(0).max(1),
    requiresConfirmation: z.literal(true),
    affectedRanges: z.array(z.string().min(1).max(128)).max(10),
    replacesExistingRules: z.boolean()
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    targetRange: z.string().min(1).max(128),
    managementMode: ConditionalFormatManagementModeSchema,
    ruleType: z.literal("color_scale"),
    points: z.array(ColorScalePointSchema).min(2).max(3),
    explanation: z.string().min(1).max(12000),
    confidence: z.number().min(0).max(1),
    requiresConfirmation: z.literal(true),
    affectedRanges: z.array(z.string().min(1).max(128)).max(10),
    replacesExistingRules: z.boolean()
  })
  // extend with the remaining rule families in the same pattern
]).superRefine((value, ctx) => {
  if (value.managementMode === "clear_on_target") {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "clear_on_target must not carry conditional-format rule payload."
    });
  }
});

export const ConditionalFormatUpdateDataSchema = strictObject({
  operation: z.literal("conditional_format_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  managementMode: ConditionalFormatManagementModeSchema,
  summary: z.string().min(1).max(500)
});
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: PASS with new schema coverage green.

- [ ] **Step 5: Update exports**

```ts
export {
  ConditionalFormatPlanDataSchema,
  ConditionalFormatUpdateDataSchema
} from "./schemas.ts";

export type ConditionalFormatPlanData = z.infer<typeof ConditionalFormatPlanDataSchema>;
export type ConditionalFormatUpdateData = z.infer<typeof ConditionalFormatUpdateDataSchema>;
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
it("renders a non-lossy conditional formatting preview", () => {
  const preview = getStructuredPreview({
    type: "conditional_format_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target",
      ruleType: "text_contains",
      text: "overdue",
      style: {
        backgroundColor: "#ffcccc",
        textColor: "#990000",
        bold: true
      },
      explanation: "Highlight overdue rows.",
      confidence: 0.94,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: true
    }
  });

  expect(preview.summary).toContain("Replace all conditional formatting");
  expect(preview.details).toContain("text contains \"overdue\"");
  expect(preview.details).toContain("background #ffcccc");
});

it("renders a typed completion line for conditional_format_update", () => {
  const line = getWritebackStatusLine({
    operation: "conditional_format_update",
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    managementMode: "add",
    summary: "Added conditional formatting to Sheet1!B2:B20."
  });

  expect(line).toBe("Added conditional formatting to Sheet1!B2:B20.");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: FAIL with missing preview branch or missing result typing.

- [ ] **Step 3: Write the minimal shared-client implementation**

```ts
export type WritePlan =
  | ExistingWritePlan
  | {
      type: "conditional_format_plan";
      summary: string;
      details: string[];
      requiresConfirmation: true;
    };

export type WritebackResult =
  | ExistingWritebackResult
  | {
      operation: "conditional_format_update";
      targetSheet: string;
      targetRange: string;
      managementMode: "add" | "replace_all_on_target" | "clear_on_target";
      summary: string;
    };

function renderConditionalFormatPreview(data: ConditionalFormatPlanData): StructuredPreview {
  const verb = data.managementMode === "add"
    ? "Add conditional formatting"
    : data.managementMode === "replace_all_on_target"
      ? "Replace all conditional formatting"
      : "Clear conditional formatting";

  return {
    summary: `${verb} on ${data.targetSheet}!${data.targetRange}.`,
    details: buildConditionalFormatDetails(data),
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

- [ ] **Step 5: Add trace labels for the new result kind**

```ts
export function getResultStatusLabel(result: WritebackResult): string {
  switch (result.operation) {
    case "conditional_format_update":
      return "Conditional formatting applied";
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

Expected: PASS with the trace/status helpers type-safe.

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
it("mentions conditional_format_plan in the runtime rules", () => {
  expect(SPREADSHEET_RUNTIME_RULES).toContain("conditional_format_plan");
  expect(SPREADSHEET_RUNTIME_RULES).toContain("UNSUPPORTED_OPERATION");
});

it("routes highlight prompts toward conditional formatting plans", () => {
  const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
    userMessage: "Highlight overdue dates in red."
  }));

  expect(prompt).toContain("conditional_format_plan");
  expect(prompt).toContain("range_format_update");
});

it("accepts a conditional_format_plan body and builds a structured-preview response", async () => {
  vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(chatCompletionEnvelope(JSON.stringify({
    type: "conditional_format_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "text_contains",
      text: "overdue",
      style: {
        backgroundColor: "#ffcccc",
        bold: true
      },
      explanation: "Highlight overdue text.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false
    }
  })), {
    status: 200,
    headers: { "content-type": "application/json" }
  })));

  await client.processRequest({ runId: "run_cf_001", request: baseRequest(), traceBus });

  expect(traceBus.getRun("run_cf_001")?.response).toMatchObject({
    type: "conditional_format_plan",
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
// Add a new response-type section:
// - choose `conditional_format_plan` for highlight / mark / color-scale / clear-conditional-format asks
// - return `error` with `UNSUPPORTED_OPERATION` for unsupported conditional-format features

// requestTemplate.ts
if (/\b(highlight|mark duplicates|conditional formatting|color scale|clear conditional formatting)\b/i.test(userIntent)) {
  hints.push("Prefer type=\"conditional_format_plan\" over advisory chat when the user is asking to apply conditional formatting.");
}

// structuredBody.ts
const ConditionalFormatPlanBodySchema = strictObject({
  type: z.literal("conditional_format_plan"),
  data: ConditionalFormatPlanDataSchema
});

// hermesClient.ts
case "conditional_format_plan":
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
case "conditional_format_plan":
  pushIfMissing({ event: "conditional_format_plan_ready", timestamp: completedAt });
  pushIfMissing({ event: "completed", timestamp: completedAt });
  break;
```

- [ ] **Step 6: Re-run the Hermes client test**

Run:

```bash
npm test -- services/gateway/tests/hermesClient.test.ts
```

Expected: PASS with the new trace event sequence.

### Task 4: Extend Writeback and Completion State

**Files:**
- Modify: `services/gateway/src/routes/writeback.ts`
- Modify: `services/gateway/src/lib/traceBus.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

- [ ] **Step 1: Write the failing writeback tests**

```ts
it("approves and completes a conditional-format plan", async () => {
  const plan = {
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    managementMode: "add",
    ruleType: "text_contains",
    text: "overdue",
    style: {
      backgroundColor: "#ffcccc"
    },
    explanation: "Highlight overdue text.",
    confidence: 0.92,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!B2:B20"],
    replacesExistingRules: false
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
        operation: "conditional_format_update",
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        managementMode: "add",
        summary: "Added conditional formatting to Sheet1!B2:B20."
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
  ConditionalFormatPlanDataSchema
]);

const WritebackResultSchema = z.union([
  ExistingWritebackResultSchema,
  ConditionalFormatUpdateDataSchema
]);

function isConditionalFormatPlan(plan: unknown): plan is ConditionalFormatPlanData {
  return ConditionalFormatPlanDataSchema.safeParse(plan).success;
}

function isConditionalFormatResult(result: unknown): result is ConditionalFormatUpdateData {
  return ConditionalFormatUpdateDataSchema.safeParse(result).success;
}
```

- [ ] **Step 4: Run the writeback test to verify it passes**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: PASS with approval and completion green for the new plan family.

- [ ] **Step 5: Persist typed completion in trace state**

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

### Task 5: Add Excel Conditional Formatting Apply Support

**Files:**
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave3Plans.test.ts`

- [ ] **Step 1: Write the failing Excel tests**

```ts
it("recognizes conditional_format_plan as a write plan in Excel", async () => {
  const preview = renderAssistantPayload({
    type: "conditional_format_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "number_compare",
      comparator: "greater_than",
      value: 10,
      style: { backgroundColor: "#ffdddd" },
      explanation: "Highlight values above 10.",
      confidence: 0.92,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false
    }
  });

  expect(preview.confirmLabel).toContain("Apply");
});

it("applies a text-contains conditional format in Excel", async () => {
  const calls = await applyExcelWave3Plan({
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    managementMode: "add",
    ruleType: "text_contains",
    text: "overdue",
    style: {
      backgroundColor: "#ffcccc",
      bold: true
    },
    explanation: "Highlight overdue text.",
    confidence: 0.9,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!B2:B20"],
    replacesExistingRules: false
  });

  expect(calls.addConditionalFormat).toHaveBeenCalled();
});
```

- [ ] **Step 2: Run the focused Excel tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave3Plans.test.ts
```

Expected: FAIL with missing preview/apply branches.

- [ ] **Step 3: Write the minimal Excel implementation**

```js
function isConditionalFormatPlan_(response) {
  return response?.type === "conditional_format_plan";
}

function getConditionalFormatStatusSummary_(result) {
  return result?.summary ?? "Conditional formatting applied.";
}

async function applyConditionalFormatPlan_(Excel, context, plan) {
  const sheet = context.workbook.worksheets.getItem(plan.targetSheet);
  const range = sheet.getRange(plan.targetRange);

  if (plan.managementMode === "clear_on_target") {
    range.conditionalFormats.clearAll();
    return {
      operation: "conditional_format_update",
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      managementMode: plan.managementMode,
      summary: `Cleared conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`
    };
  }

  if (plan.managementMode === "replace_all_on_target") {
    range.conditionalFormats.clearAll();
  }

  const format = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
  format.cellValue.rule = {
    operator: Excel.ConditionalCellValueOperator.greaterThan,
    formula1: String(plan.value)
  };
  format.cellValue.format.fill.color = plan.style.backgroundColor;

  return {
    operation: "conditional_format_update",
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    managementMode: plan.managementMode,
    summary: `Applied conditional formatting to ${plan.targetSheet}!${plan.targetRange}.`
  };
}
```

- [ ] **Step 4: Run the focused Excel tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave3Plans.test.ts
```

Expected: PASS with preview recognition and apply-path coverage green.

- [ ] **Step 5: Add fail-closed coverage**

```ts
it("fails closed when Excel cannot represent the requested conditional-format semantics exactly", async () => {
  await expect(applyExcelWave3Plan({
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    managementMode: "add",
    ruleType: "color_scale",
    points: [
      { type: "percentile", value: 10, color: "#f00" },
      { type: "percentile", value: 50, color: "#ff0" },
      { type: "percentile", value: 90, color: "#0f0" }
    ],
    explanation: "Unsupported exact mapping for this host branch.",
    confidence: 0.7,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!B2:B20"],
    replacesExistingRules: false
  })).rejects.toThrow(/cannot represent/i);
});
```

- [ ] **Step 6: Re-run the focused Excel tests**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave3Plans.test.ts
```

Expected: PASS with fail-closed semantics covered.

### Task 6: Add Google Sheets Conditional Formatting Apply Support

**Files:**
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave3Plans.test.ts`

- [ ] **Step 1: Write the failing Google Sheets tests**

```ts
it("renders a future-tense Google Sheets preview for conditional formatting", () => {
  const preview = getStructuredPreview({
    type: "conditional_format_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "replace_all_on_target",
      ruleType: "duplicate_values",
      style: {
        backgroundColor: "#ffeeaa",
        textColor: "#663300"
      },
      explanation: "Highlight duplicates.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: true
    }
  });

  expect(preview.summary).toContain("Will replace all conditional formatting");
});

it("applies a Google Sheets clear_on_target conditional format plan", () => {
  const result = applyWave3GooglePlan({
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    managementMode: "clear_on_target",
    explanation: "Clear existing rules.",
    confidence: 0.85,
    requiresConfirmation: true,
    affectedRanges: ["Sheet1!B2:B20"],
    replacesExistingRules: true
  });

  expect(result.operation).toBe("conditional_format_update");
  expect(result.summary).toContain("Cleared conditional formatting");
});
```

- [ ] **Step 2: Run the focused Google Sheets test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave3Plans.test.ts
```

Expected: FAIL with missing preview/apply branches.

- [ ] **Step 3: Write the minimal Google Sheets implementation**

```js
function isConditionalFormatPlan_(response) {
  return response && response.type === "conditional_format_plan";
}

function getConditionalFormatPreviewSummary_(plan) {
  if (plan.managementMode === "clear_on_target") {
    return `Will clear conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
  }

  if (plan.managementMode === "replace_all_on_target") {
    return `Will replace all conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
  }

  return `Will add conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
}

function applyConditionalFormatPlan_(plan) {
  if (plan.managementMode === "clear_on_target") {
    clearConditionalFormatRulesOnTarget_(plan.targetSheet, plan.targetRange);
    return {
      operation: "conditional_format_update",
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      managementMode: plan.managementMode,
      summary: `Cleared conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`
    };
  }

  const requests = buildConditionalFormatRequests_(plan);
  Sheets.Spreadsheets.batchUpdate({ requests }, spreadsheetId);

  return {
    operation: "conditional_format_update",
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    managementMode: plan.managementMode,
    summary: `Applied conditional formatting to ${plan.targetSheet}!${plan.targetRange}.`
  };
}
```

- [ ] **Step 4: Run the focused Google Sheets test to verify it passes**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave3Plans.test.ts
```

Expected: PASS with preview and apply-path coverage green.

- [ ] **Step 5: Add fail-closed tests for unsupported exact mappings**

```ts
it("suppresses confirmation when Google Sheets cannot represent the requested rule exactly", () => {
  const preview = getStructuredPreview({
    type: "conditional_format_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "color_scale",
      points: [
        { type: "number", value: 10, color: "#ff0000" },
        { type: "percentile", value: 50, color: "#ffff00" },
        { type: "number", value: 100, color: "#00ff00" }
      ],
      explanation: "Mixed scale point types unsupported for exact host parity.",
      confidence: 0.7,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false
    }
  });

  expect(preview.supportError).toContain("cannot represent");
  expect(preview.showConfirm).toBe(false);
});
```

- [ ] **Step 6: Re-run the focused Google Sheets test**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave3Plans.test.ts
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
- Test: `services/gateway/tests/excelWave3Plans.test.ts`
- Test: `services/gateway/tests/googleSheetsWave3Plans.test.ts`
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

- [ ] **Step 1: Run the main Wave 3 regression batch**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts packages/shared-client/tests/client.test.ts services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts services/gateway/tests/writebackFlow.test.ts services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave3Plans.test.ts services/gateway/tests/googleSheetsWave3Plans.test.ts
```

Expected: PASS with all focused Wave 3 suites green.

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
bash -lc 'awk "/<script>/{flag=1;next}/<\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave3-sidebar.js && node --check /tmp/hermes-wave3-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave3-code-gs.js && node --check /tmp/hermes-wave3-code-gs.js'
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
- scope and chosen plan family: Task 1
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
- `conditional_format_plan`
- `conditional_format_update`
- `managementMode`
- `ruleType`
- `replacesExistingRules`

Later tasks do not rename these identifiers.
