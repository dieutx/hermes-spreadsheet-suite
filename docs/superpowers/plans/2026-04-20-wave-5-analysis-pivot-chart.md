# Wave 5 Analysis Reports + Pivot Tables + Charts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add contract-valid `analysis_report_plan`, `pivot_table_plan`, and `chart_plan` flows with strict preview, exact-safe host execution, and writeback support for materialized artifacts.

**Architecture:** Extend the existing typed-plan pipeline with three new strict artifact families instead of folding them into `sheet_update` or a generic artifact bag. Keep `analysis_report_plan(chat_only)` outside writeback, while `analysis_report_plan(materialize_report)`, `pivot_table_plan`, and `chart_plan` follow the normal approval/writeback path with destructive second confirmation when artifact areas would be replaced or overwritten.

**Tech Stack:** TypeScript, Zod, Vitest, Express, Office.js, Google Apps Script, browser-safe shared client helpers.

---

## File Structure

### Contracts and shared types

- Modify: `packages/contracts/src/schemas.ts`
  - add Wave 5 plan schemas, response schemas, update schemas, trace events, and capability flags
- Modify: `packages/contracts/src/index.ts`
  - export the new schemas and inferred types
- Test: `packages/contracts/tests/contracts.test.ts`
  - schema coverage for report output modes, pivot invariants, chart invariants, and update envelopes

### Shared preview and client types

- Modify: `packages/shared-client/src/types.ts`
  - extend `WritePlan`, `WritebackResult`, and preview unions
- Modify: `packages/shared-client/src/render.ts`
  - add non-lossy previews and response-body text handling for all three plan families
- Modify: `packages/shared-client/src/trace.ts`
  - add typed trace labels for Wave 5 plan/update events
- Test: `packages/shared-client/tests/client.test.ts`
  - preview coverage and `chat_only` non-write behavior

### Gateway runtime and writeback

- Modify: `services/gateway/src/hermes/runtimeRules.ts`
  - teach Hermes the new plan families, exact-safe failure behavior, and `chat_only` vs `materialize_report`
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
  - route analysis / pivot / chart prompts toward Wave 5 plan types
- Modify: `services/gateway/src/hermes/structuredBody.ts`
  - normalize and validate the new plan families and update envelopes
- Modify: `services/gateway/src/lib/hermesClient.ts`
  - map preview and trace behavior for Wave 5 responses
- Modify: `services/gateway/src/routes/writeback.ts`
  - add approval/completion support for materialized report, pivot, and chart updates
- Modify: `services/gateway/src/lib/traceBus.ts`
  - persist typed completion state for Wave 5 updates
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

### Excel host

- Create: `apps/excel-addin/src/taskpane/analysisArtifactsPlan.js`
  - isolate Wave 5 status-summary and artifact-shape helpers
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
  - integrate preview/apply paths for report, pivot, and chart plans
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
  - extend status-line helpers for Wave 5 updates
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave5Plans.test.ts`

### Google Sheets host

- Modify: `apps/google-sheets-addon/src/Code.gs`
  - apply Wave 5 plans with exact-safe semantics
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
  - render previews and confirmation affordances for Wave 5 plans
- Test: `services/gateway/tests/googleSheetsWave5Plans.test.ts`

### Final regression

- Re-run existing tests covering Waves 1 through 4 and baseline request routing
- Re-run gateway build and host syntax checks

Note: this workspace snapshot currently has no `.git` metadata, so execution should skip commit steps locally unless the repo metadata is restored.

---

### Task 1: Add Wave 5 Contract Schemas

**Files:**
- Modify: `packages/contracts/src/schemas.ts`
- Modify: `packages/contracts/src/index.ts`
- Test: `packages/contracts/tests/contracts.test.ts`

- [ ] **Step 1: Write the failing contract tests**

```ts
import {
  AnalysisReportPlanDataSchema,
  ChartPlanDataSchema,
  HermesResponseSchema,
  PivotTablePlanDataSchema
} from "../src/index.ts";

it("accepts a chat-only analysis report plan", () => {
  const parsed = AnalysisReportPlanDataSchema.parse({
    sourceSheet: "Sales",
    sourceRange: "A1:F50",
    outputMode: "chat_only",
    sections: [
      {
        type: "summary_stats",
        title: "Revenue summary",
        summary: "Average revenue is 12,500.",
        sourceRanges: ["Sales!A1:F50"]
      }
    ],
    explanation: "Summarize the selected sales range.",
    confidence: 0.92,
    requiresConfirmation: false,
    affectedRanges: ["Sales!A1:F50"],
    overwriteRisk: "none",
    confirmationLevel: "standard"
  });

  expect(parsed.outputMode).toBe("chat_only");
  expect(parsed.requiresConfirmation).toBe(false);
});

it("rejects a materialized analysis report without a target sheet", () => {
  const parsed = AnalysisReportPlanDataSchema.safeParse({
    sourceSheet: "Sales",
    sourceRange: "A1:F50",
    outputMode: "materialize_report",
    sections: [
      {
        type: "summary_stats",
        title: "Revenue summary",
        summary: "Average revenue is 12,500.",
        sourceRanges: ["Sales!A1:F50"]
      }
    ],
    explanation: "Write a report sheet.",
    confidence: 0.91,
    requiresConfirmation: true,
    affectedRanges: ["Sales!A1:F50"],
    overwriteRisk: "low",
    confirmationLevel: "standard"
  });

  expect(parsed.success).toBe(false);
});

it("accepts a pivot table plan with multiple values and optional sort", () => {
  const parsed = PivotTablePlanDataSchema.parse({
    sourceSheet: "Sales",
    sourceRange: "A1:F50",
    targetSheet: "Sales Pivot",
    targetRange: "A1",
    rowGroups: ["Region", "Rep"],
    columnGroups: ["Quarter"],
    valueAggregations: [
      { field: "Revenue", aggregation: "sum" },
      { field: "Deals", aggregation: "count" }
    ],
    filters: [{ field: "Status", operator: "equal_to", value: "Closed Won" }],
    sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
    explanation: "Build a sales pivot by region and rep.",
    confidence: 0.9,
    requiresConfirmation: true,
    affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
    overwriteRisk: "low",
    confirmationLevel: "standard"
  });

  expect(parsed.valueAggregations).toHaveLength(2);
});

it("accepts a chart plan with explicit category and series mapping", () => {
  const parsed = ChartPlanDataSchema.parse({
    sourceSheet: "Sales",
    sourceRange: "A1:C20",
    targetSheet: "Sales Chart",
    targetRange: "A1",
    chartType: "line",
    categoryField: "Month",
    series: [
      { field: "Revenue", label: "Revenue" },
      { field: "Margin", label: "Margin" }
    ],
    title: "Revenue vs Margin",
    legendPosition: "bottom",
    explanation: "Chart monthly revenue and margin.",
    confidence: 0.93,
    requiresConfirmation: true,
    affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
    overwriteRisk: "low",
    confirmationLevel: "standard"
  });

  expect(parsed.chartType).toBe("line");
});

it("accepts a chart_update response envelope", () => {
  const parsed = HermesResponseSchema.parse({
    schemaVersion: "1.0.0",
    type: "chart_update",
    requestId: "req_chart_update_001",
    hermesRunId: "run_chart_update_001",
    processedBy: "host",
    serviceLabel: "hermes-gateway-local",
    environmentLabel: "local-dev",
    startedAt: "2026-04-20T12:00:00.000Z",
    completedAt: "2026-04-20T12:00:01.000Z",
    durationMs: 1000,
    trace: [{ event: "completed", timestamp: "2026-04-20T12:00:01.000Z" }],
    ui: {
      displayMode: "structured-preview",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: false
    },
    data: {
      operation: "chart_update",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line",
      summary: "Created line chart on Sales Chart!A1."
    }
  });

  expect(parsed.type).toBe("chart_update");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: FAIL with missing Wave 5 schemas, response branches, or exports.

- [ ] **Step 3: Write the minimal contract implementation**

```ts
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

export const AnalysisReportPlanDataSchema = strictObject({
  sourceSheet: z.string().min(1).max(128),
  sourceRange: z.string().min(1).max(128),
  outputMode: z.enum(["chat_only", "materialize_report"]),
  targetSheet: z.string().min(1).max(128).optional(),
  targetRange: z.string().min(1).max(128).optional(),
  sections: z.array(AnalysisReportSectionSchema).min(1).max(12),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.boolean(),
  affectedRanges: z.array(z.string().min(1).max(128)).max(12),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
}).superRefine((data, ctx) => {
  if (data.outputMode === "chat_only" && data.requiresConfirmation !== false) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "chat_only analysis reports must not require confirmation.",
      path: ["requiresConfirmation"]
    });
  }

  if (data.outputMode === "materialize_report") {
    if (!data.targetSheet) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "materialize_report requires targetSheet.",
        path: ["targetSheet"]
      });
    }
    if (data.requiresConfirmation !== true) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "materialize_report requires confirmation.",
        path: ["requiresConfirmation"]
      });
    }
  }
});

const PivotAggregationSchema = strictObject({
  field: z.string().min(1).max(128),
  aggregation: z.enum(["sum", "count", "average", "min", "max"])
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
  sourceRange: z.string().min(1).max(128),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
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
  sourceRange: z.string().min(1).max(128),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
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
  series: z.array(strictObject({
    field: z.string().min(1).max(128),
    label: z.string().min(1).max(128).optional()
  })).min(1).max(10),
  title: z.string().min(1).max(256).optional(),
  legendPosition: z.enum(["top", "bottom", "left", "right", "hidden"]).optional(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(12),
  overwriteRisk: OverwriteRiskSchema,
  confirmationLevel: ConfirmationLevelSchema
});

export const ChartUpdateDataSchema = strictObject({
  operation: z.literal("chart_update"),
  targetSheet: z.string().min(1).max(128),
  targetRange: z.string().min(1).max(128),
  chartType: ChartPlanDataSchema.shape.chartType,
  summary: z.string().min(1).max(500)
});
```

- [ ] **Step 4: Run the contract tests**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: PASS for the new Wave 5 schema coverage.

---

### Task 2: Add Shared Client Types and Previews

**Files:**
- Modify: `packages/shared-client/src/types.ts`
- Modify: `packages/shared-client/src/render.ts`
- Modify: `packages/shared-client/src/trace.ts`
- Test: `packages/shared-client/tests/client.test.ts`

- [ ] **Step 1: Write the failing shared-client tests**

```ts
import {
  getRequiresConfirmation,
  getResponseBodyText,
  getStructuredPreview,
  isWritePlanResponse
} from "../src/render.ts";

it("does not treat chat-only analysis reports as write plans", () => {
  const response = {
    type: "analysis_report_plan",
    data: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "chat_only",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Summarize the selected range.",
      confidence: 0.92,
      requiresConfirmation: false,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "none",
      confirmationLevel: "standard"
    }
  } as const;

  expect(isWritePlanResponse(response as never)).toBe(false);
  expect(getRequiresConfirmation(response as never)).toBe(false);
});

it("renders a pivot preview with explicit group and value metadata", () => {
  const response = {
    type: "pivot_table_plan",
    data: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      explanation: "Build a pivot.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  } as const;

  expect(getStructuredPreview(response as never)).toMatchObject({
    kind: "pivot_table_plan",
    sourceSheet: "Sales",
    targetSheet: "Sales Pivot",
    rowGroups: ["Region", "Rep"]
  });
  expect(getResponseBodyText(response as never)).toContain("pivot");
});

it("renders a chart update summary directly", () => {
  const response = {
    type: "chart_update",
    data: {
      operation: "chart_update",
      targetSheet: "Sales Chart",
      targetRange: "A1",
      chartType: "line",
      summary: "Created line chart on Sales Chart!A1."
    }
  } as const;

  expect(getResponseBodyText(response as never)).toBe("Created line chart on Sales Chart!A1.");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: FAIL with missing Wave 5 preview types or incorrect write-plan classification.

- [ ] **Step 3: Implement the shared-client branches**

```ts
export type AnalysisReportUpdateWritebackResult = {
  kind: "analysis_report_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: "analysis_report_update";
  targetSheet: string;
  targetRange: string;
  summary: string;
};

export type PivotTableUpdateWritebackResult = {
  kind: "pivot_table_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: "pivot_table_update";
  targetSheet: string;
  targetRange: string;
  summary: string;
};

export type ChartUpdateWritebackResult = {
  kind: "chart_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: "chart_update";
  targetSheet: string;
  targetRange: string;
  chartType: string;
  summary: string;
};
```

```ts
case "analysis_report_plan":
  return buildAnalysisReportPreview(response.data);
case "pivot_table_plan":
  return buildPivotTablePreview(response.data);
case "chart_plan":
  return buildChartPreview(response.data);
case "analysis_report_update":
case "pivot_table_update":
case "chart_update":
  return {
    kind: response.type,
    ...response.data
  };
```

```ts
if (response.type === "analysis_report_plan") {
  return response.data.outputMode === "materialize_report";
}

if (
  response.type === "pivot_table_plan" ||
  response.type === "chart_plan"
) {
  return true;
}
```

- [ ] **Step 4: Run the shared-client tests**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: PASS with explicit previews and correct `chat_only` non-write behavior.

---

### Task 3: Add Gateway Runtime, Request Routing, and Structured Response Handling

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
it("routes artifact prompts toward wave 5 plan families", () => {
  const template = buildRequestTemplate({
    prompt: "Create a pivot table of this data by region and rep.",
    host: { platform: "google_sheets" },
    capabilities: { canConfirmWriteBack: true, canRenderStructuredPreview: true, canRenderTrace: true }
  });

  expect(template).toContain("pivot_table_plan");
  expect(template).toContain("chart_plan");
  expect(template).toContain("analysis_report_plan");
});

it("accepts a chat-only analysis report response body", async () => {
  const response = await normalizeHermesStructuredBody({
    type: "analysis_report_plan",
    data: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "chat_only",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Summarize the selected range.",
      confidence: 0.92,
      requiresConfirmation: false,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "none",
      confirmationLevel: "standard"
    }
  });

  expect(response.type).toBe("analysis_report_plan");
  expect(response.data.outputMode).toBe("chat_only");
});

it("rejects a chart response with unsupported chartType", async () => {
  await expect(() =>
    normalizeHermesStructuredBody({
      type: "chart_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:C10",
        targetSheet: "Chart Sheet",
        targetRange: "A1",
        chartType: "combo",
        series: [{ field: "Revenue" }],
        explanation: "Invalid chart type.",
        confidence: 0.8,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C10", "Chart Sheet!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    })
  ).rejects.toThrow();
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: FAIL with missing Wave 5 routing, missing trace labels, or missing structured-body branches.

- [ ] **Step 3: Implement runtime rules and structured-body handling**

```ts
// runtimeRules.ts
Artifact plans you may return in this environment:
- analysis_report_plan
- pivot_table_plan
- chart_plan

Rules:
- analysis_report_plan with outputMode="chat_only" must set requiresConfirmation=false
- analysis_report_plan with outputMode="materialize_report" must set requiresConfirmation=true and include targetSheet
- if a pivot/chart/report request cannot be represented exactly by this host, return error code UNSUPPORTED_OPERATION
- do not invent unsupported chart types, pivot features, or report materialization semantics
```

```ts
// requestTemplate.ts
if (mentionsPivot(prompt)) {
  preferredResponseTypes.push("pivot_table_plan");
}

if (mentionsChart(prompt)) {
  preferredResponseTypes.push("chart_plan");
}

if (mentionsAnalysis(prompt)) {
  preferredResponseTypes.push("analysis_report_plan");
}
```

```ts
// hermesClient.ts / structuredBody.ts
case "analysis_report_plan":
  return HermesResponseSchema.parse({
    ...baseEnvelope,
    type: "analysis_report_plan",
    data: body.data
  });
case "pivot_table_plan":
  return HermesResponseSchema.parse({
    ...baseEnvelope,
    type: "pivot_table_plan",
    data: body.data
  });
case "chart_plan":
  return HermesResponseSchema.parse({
    ...baseEnvelope,
    type: "chart_plan",
    data: body.data
  });
```

- [ ] **Step 4: Run the gateway-facing tests**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: PASS with Wave 5 routing and response parsing.

---

### Task 4: Add Writeback and Typed Completion Support

**Files:**
- Modify: `services/gateway/src/routes/writeback.ts`
- Modify: `services/gateway/src/lib/traceBus.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

- [ ] **Step 1: Write the failing writeback tests**

```ts
it("does not send chat-only analysis reports through writeback approval", async () => {
  const response = {
    type: "analysis_report_plan",
    data: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "chat_only",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Summarize the selected range.",
      confidence: 0.92,
      requiresConfirmation: false,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "none",
      confirmationLevel: "standard"
    }
  };

  expect(isWritebackEligibleResponse(response as never)).toBe(false);
});

it("accepts a pivot_table_update completion result", async () => {
  const completed = await completeWriteback({
    requestId: "req_pivot_001",
    runId: "run_pivot_001",
    approvalToken: "token",
    planDigest: "digest",
    result: {
      kind: "pivot_table_update",
      operation: "pivot_table_update",
      hostPlatform: "google_sheets",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      summary: "Created pivot table on Sales Pivot!A1."
    }
  });

  expect(completed.status).toBe("completed");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: FAIL with missing update result kinds or incorrect writeback eligibility for `chat_only`.

- [ ] **Step 3: Implement writeback eligibility and completion branches**

```ts
function isWritebackEligibleResponse(response: HermesResponse): boolean {
  if (response.type === "analysis_report_plan") {
    return response.data.outputMode === "materialize_report";
  }

  return (
    response.type === "pivot_table_plan" ||
    response.type === "chart_plan" ||
    existingWritePlanTypes.has(response.type)
  );
}

const WritebackResultSchema = z.union([
  ExistingWritebackResultSchema,
  AnalysisReportUpdateResultSchema,
  PivotTableUpdateResultSchema,
  ChartUpdateResultSchema
]);
```

```ts
traceBus.record(runId, {
  event: "result_generated",
  timestamp: now,
  label: "Result generated"
});

traceBus.record(runId, {
  event: "completed",
  timestamp: now,
  label: "Completed"
});
```

- [ ] **Step 4: Run the writeback tests**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: PASS for Wave 5 writeback eligibility and completion state.

---

### Task 5: Add Excel Host Support

**Files:**
- Create: `apps/excel-addin/src/taskpane/analysisArtifactsPlan.js`
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave5Plans.test.ts`

- [ ] **Step 1: Write the failing Excel host tests**

```ts
it("treats chat-only analysis reports as non-write previews in Excel", async () => {
  const taskpane = await loadTaskpaneModule({ sync: vi.fn(async () => {}) });

  const response = {
    type: "analysis_report_plan",
    data: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "chat_only",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Summarize the selected range.",
      confidence: 0.92,
      requiresConfirmation: false,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "none",
      confirmationLevel: "standard"
    }
  };

  expect(taskpane.isWritePlanResponse(response)).toBe(false);
  expect(taskpane.renderStructuredPreview(response, {
    runId: "run_analysis_chat",
    requestId: "req_analysis_chat"
  })).not.toContain("Confirm");
});

it("applies a materialized analysis report in Excel", async () => {
  const targetRange = createRangeStub({
    address: "Report!A1:C4",
    rowCount: 4,
    columnCount: 3,
    values: [
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""]
    ]
  });

  const taskpane = await loadTaskpaneModule({
    sync: vi.fn(async () => {}),
    workbook: {
      worksheets: {
        getItem: vi.fn(() => ({
          getRange: vi.fn(() => targetRange)
        }))
      }
    }
  });

  await expect(taskpane.applyWritePlan({
    plan: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report",
      targetSheet: "Report",
      targetRange: "A1",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Write a report sheet.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50", "Report!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    },
    requestId: "req_excel_report_001",
    runId: "run_excel_report_001",
    approvalToken: "token"
  })).resolves.toMatchObject({
    kind: "analysis_report_update",
    operation: "analysis_report_update",
    targetSheet: "Report"
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave5Plans.test.ts
```

Expected: FAIL with missing Wave 5 write-plan helpers or missing Excel artifact apply branches.

- [ ] **Step 3: Implement the Excel artifact helpers and apply branches**

```js
// analysisArtifactsPlan.js
export function getAnalysisReportStatusSummary(plan) {
  return `Created analysis report on ${plan.targetSheet}!${plan.targetRange}.`;
}

export function getPivotTableStatusSummary(plan) {
  return `Created pivot table on ${plan.targetSheet}!${plan.targetRange}.`;
}

export function getChartStatusSummary(plan) {
  return `Created ${plan.chartType} chart on ${plan.targetSheet}!${plan.targetRange}.`;
}
```

```js
// taskpane.js
if (response.type === "analysis_report_plan") {
  return response.data.outputMode === "materialize_report";
}

if (response.type === "pivot_table_plan" || response.type === "chart_plan") {
  return true;
}
```

```js
if (isAnalysisReportPlan(plan) && plan.outputMode === "materialize_report") {
  const sheet = worksheets.getItem(plan.targetSheet);
  const target = sheet.getRange(plan.targetRange);
  target.values = buildAnalysisReportMatrix(plan.sections);
  await context.sync();
  return {
    kind: "analysis_report_update",
    hostPlatform: platform,
    operation: "analysis_report_update",
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    summary: getAnalysisReportStatusSummary(plan)
  };
}

if (isPivotTablePlan(plan)) {
  throw new Error("Excel host pivot implementation must use exact-safe Office.js pivot APIs before enabling this branch.");
}

if (isChartPlan(plan)) {
  throw new Error("Excel host chart implementation must use exact-safe Office.js chart APIs before enabling this branch.");
}
```

- [ ] **Step 4: Run the Excel tests**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave5Plans.test.ts
```

Expected: PASS for Wave 5 Excel preview/apply behavior, or fail closed for unsupported exact mappings until exact-safe branches are implemented.

---

### Task 6: Add Google Sheets Host Support

**Files:**
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave5Plans.test.ts`

- [ ] **Step 1: Write the failing Google Sheets host tests**

```ts
it("treats chat-only analysis reports as non-write previews in Google Sheets", () => {
  const sidebar = loadSidebarContext();

  const html = sidebar.renderStructuredPreview({
    type: "analysis_report_plan",
    data: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "chat_only",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Summarize the selected range.",
      confidence: 0.92,
      requiresConfirmation: false,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "none",
      confirmationLevel: "standard"
    }
  }, {
    runId: "run_analysis_chat",
    requestId: "req_analysis_chat"
  });

  expect(html).not.toContain("Confirm");
});

it("applies a materialized analysis report in Google Sheets", () => {
  const reportRange = createRangeStub({
    a1Notation: "A1:C4",
    row: 1,
    column: 1,
    numRows: 4,
    numColumns: 3,
    values: [
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
      ["", "", ""]
    ]
  });

  const sheet = {
    getRange: vi.fn(() => reportRange)
  };

  const spreadsheet = {
    getSheetByName: vi.fn(() => sheet)
  };

  const { applyWritePlan } = loadCodeModule({ spreadsheet });

  const result = applyWritePlan({
    plan: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report",
      targetSheet: "Report",
      targetRange: "A1",
      sections: [
        {
          type: "summary_stats",
          title: "Revenue summary",
          summary: "Average revenue is 12,500.",
          sourceRanges: ["Sales!A1:F50"]
        }
      ],
      explanation: "Write a report sheet.",
      confidence: 0.91,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50", "Report!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  });

  expect(result).toMatchObject({
    kind: "analysis_report_update",
    operation: "analysis_report_update",
    targetSheet: "Report"
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts
```

Expected: FAIL with missing Wave 5 preview or apply branches.

- [ ] **Step 3: Implement the Google Sheets preview and apply branches**

```js
// Sidebar.js.html
if (response.type === "analysis_report_plan") {
  return response.data.outputMode === "materialize_report";
}

if (response.type === "pivot_table_plan" || response.type === "chart_plan") {
  return true;
}
```

```js
// Code.gs
if (isAnalysisReportPlan_(plan) && plan.outputMode === "materialize_report") {
  const sheet = spreadsheet.getSheetByName(plan.targetSheet);
  const matrix = buildAnalysisReportMatrix_(plan.sections);
  const target = resolveReportWriteRange_(sheet, plan.targetRange, matrix);
  target.setValues(matrix);
  SpreadsheetApp.flush();
  return {
    kind: 'analysis_report_update',
    operation: 'analysis_report_update',
    hostPlatform: 'google_sheets',
    targetSheet: plan.targetSheet,
    targetRange: target.getA1Notation(),
    summary: `Created analysis report on ${plan.targetSheet}!${target.getA1Notation()}.`
  };
}

if (isPivotTablePlan_(plan)) {
  throw new Error('Google Sheets host must use exact-safe pivot semantics before enabling this branch.');
}

if (isChartPlan_(plan)) {
  throw new Error('Google Sheets host must use exact-safe chart semantics before enabling this branch.');
}
```

- [ ] **Step 4: Run the Google Sheets tests**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts
```

Expected: PASS for report materialization and fail-closed behavior for still-unsupported exact mappings.

---

### Task 7: Run Full Regression and Whole-Wave Review

**Files:**
- Verify only: existing Wave 1–4 and baseline tests
- Verify only: host syntax/build surfaces

- [ ] **Step 1: Run the focused Wave 5 and shared batches**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts packages/shared-client/tests/client.test.ts services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts services/gateway/tests/writebackFlow.test.ts services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave5Plans.test.ts services/gateway/tests/googleSheetsWave5Plans.test.ts
```

Expected: PASS with Wave 5 plan/update coverage green.

- [ ] **Step 2: Run the baseline regression batch**

Run:

```bash
npm test -- services/gateway/tests/requestRouter.test.ts services/gateway/tests/structuredBody.test.ts services/gateway/tests/traceBus.test.ts services/gateway/tests/uploads.test.ts services/gateway/tests/excelCellValues.test.ts services/gateway/tests/excelReferencedCells.test.ts services/gateway/tests/googleSheetsReferencedCells.test.ts services/gateway/tests/rangeSafety.test.ts
```

Expected: PASS with no regressions outside Wave 5.

- [ ] **Step 3: Run syntax checks and gateway build**

Run:

```bash
node --check <repo-root>/apps/excel-addin/src/taskpane/taskpane.js
bash -lc 'awk "/<script>/{flag=1;next}/<\\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave5-sidebar.js && node --check /tmp/hermes-wave5-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave5-code-gs.js && node --check /tmp/hermes-wave5-code-gs.js'
npm --workspace @hermes/gateway run build
```

Expected: all commands PASS.

- [ ] **Step 4: Request whole-wave review before claiming completion**

Use a review agent to inspect only the final Wave 5 surfaces and respond with either concrete findings or `APPROVED with no findings`.

Expected: APPROVED, then mark Wave 5 complete.

---

## Self-Review

Spec coverage check:
- `analysis_report_plan` tasks: covered in Tasks 1–6
- `pivot_table_plan` tasks: covered in Tasks 1–5, with fail-closed host behavior required until exact-safe branches are implemented
- `chart_plan` tasks: covered in Tasks 1–5, with fail-closed host behavior required until exact-safe branches are implemented
- `chat_only` vs `materialize_report`: covered in Tasks 1–4 and host tests
- destructive confirmation / overwrite handling: covered in Tasks 1 and 4, plus host expectations
- regression requirement: covered in Task 7

Placeholder scan:
- no `TODO`
- no `TBD`
- no “write tests for above” placeholders without concrete examples

Type consistency check:
- plan families use:
  - `analysis_report_plan`
  - `pivot_table_plan`
  - `chart_plan`
- update kinds use:
  - `analysis_report_update`
  - `pivot_table_update`
  - `chart_update`
- `analysis_report_plan(chat_only)` remains non-write throughout the plan
