# Wave 1 Sheet Structure + Sort/Filter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add contract-valid `sheet_structure_update`, `range_sort_plan`, and `range_filter_plan` flows with strict preview, approval, and host apply support for Excel and Google Sheets.

**Architecture:** Extend the existing typed-plan pipeline instead of introducing a generic operation bag. Add three new strict plan families in contracts, teach Hermes runtime/request routing to emit them, extend gateway writeback with destructive second-confirm support, then add typed preview and host adapters in Excel and Google Sheets.

**Tech Stack:** TypeScript, Zod, Vitest, Express, Office.js, Google Apps Script, browser-safe shared client helpers.

---

## File Structure

### Contracts and shared types

- Modify: `packages/contracts/src/schemas.ts`
  - add plan schemas, trace events, capabilities, and response unions for wave 1
- Modify: `packages/contracts/src/index.ts`
  - export the new plan and result types
- Test: `packages/contracts/tests/contracts.test.ts`
  - schema coverage for new plan families and destructive confirmation metadata

### Shared preview and client types

- Modify: `packages/shared-client/src/types.ts`
  - extend `WritePlan` and `WritebackResult`
- Modify: `packages/shared-client/src/render.ts`
  - add typed previews for structure, sort, and filter
- Modify: `packages/shared-client/src/index.ts`
  - export any new preview helpers/types
- Test: `packages/shared-client/tests/client.test.ts`
  - preview and client typing regression

### Gateway runtime and writeback

- Modify: `services/gateway/src/hermes/runtimeRules.ts`
  - teach Hermes the new response types and safety rules
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
  - route prompts toward the right wave 1 type
- Modify: `services/gateway/src/hermes/structuredBody.ts`
  - normalize and validate the new plan families
- Modify: `services/gateway/src/routes/writeback.ts`
  - add approval/completion support for wave 1
- Modify: `services/gateway/src/lib/traceBus.ts`
  - store destructive confirmation state if needed by the approval lifecycle
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`

### Excel host

- Create: `apps/excel-addin/src/taskpane/structurePlan.js`
  - structure-plan detection, summary, and apply helpers
- Create: `apps/excel-addin/src/taskpane/sortFilterPlan.js`
  - sort/filter helper compilation and summaries
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
  - re-export or consume new summary helpers as needed
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
  - integrate new plan preview/apply paths and second-confirm UX plumbing
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave1Plans.test.ts`

### Google Sheets host

- Create: `apps/google-sheets-addon/src/Wave1Plans.js`
  - pure helper functions for structure/sort/filter compilation with `module.exports` guard for tests
- Modify: `apps/google-sheets-addon/src/Code.gs`
  - integrate wave 1 apply paths
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
  - integrate preview/apply/confirmation UI handling for new plan types
- Test: `services/gateway/tests/googleSheetsWave1Plans.test.ts`

### Final regression

- Re-run existing tests covering:
  - `sheet_update`
  - `sheet_import_plan`
  - `workbook_structure_update`
  - `range_format_update`

---

### Task 1: Add Wave 1 Contract Schemas

**Files:**
- Modify: `packages/contracts/src/schemas.ts`
- Modify: `packages/contracts/src/index.ts`
- Test: `packages/contracts/tests/contracts.test.ts`

- [ ] **Step 1: Write the failing contract tests**

```ts
import {
  HermesResponseSchema,
  SheetStructureUpdateDataSchema,
  RangeSortPlanDataSchema,
  RangeFilterPlanDataSchema
} from "../src/index.ts";

it("accepts a destructive sheet structure delete-row plan", () => {
  const parsed = SheetStructureUpdateDataSchema.parse({
    targetSheet: "Sheet1",
    operation: "delete_rows",
    startIndex: 7,
    count: 3,
    explanation: "Delete three empty rows below the table.",
    confidence: 0.92,
    requiresConfirmation: true,
    confirmationLevel: "destructive",
    affectedRanges: ["Sheet1!8:10"]
  });

  expect(parsed.operation).toBe("delete_rows");
  expect(parsed.confirmationLevel).toBe("destructive");
});

it("accepts a multi-key range sort response envelope", () => {
  const parsed = HermesResponseSchema.parse({
    schemaVersion: "1.0.0",
    type: "range_sort_plan",
    requestId: "req_sort_001",
    hermesRunId: "run_sort_001",
    processedBy: "hermes",
    serviceLabel: "hermes-gateway-local",
    environmentLabel: "local-dev",
    startedAt: "2026-04-19T12:00:00.000Z",
    completedAt: "2026-04-19T12:00:01.000Z",
    durationMs: 1000,
    trace: [
      { event: "range_sort_plan_ready", timestamp: "2026-04-19T12:00:01.000Z" }
    ],
    ui: {
      displayMode: "structured-preview",
      showTrace: true,
      showWarnings: true,
      showConfidence: true,
      showRequiresConfirmation: true
    },
    data: {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      keys: [
        { columnRef: "Status", direction: "asc" },
        { columnRef: "Due Date", direction: "desc" }
      ],
      explanation: "Sort open items first, then latest due date.",
      confidence: 0.94,
      requiresConfirmation: true
    }
  });

  expect(parsed.type).toBe("range_sort_plan");
  expect(parsed.data.keys).toHaveLength(2);
});

it("rejects a filter plan without conditions", () => {
  const parsed = RangeFilterPlanDataSchema.safeParse({
    targetSheet: "Sheet1",
    targetRange: "A1:F25",
    hasHeader: true,
    conditions: [],
    combiner: "and",
    clearExistingFilters: true,
    explanation: "Invalid filter plan.",
    confidence: 0.8,
    requiresConfirmation: true
  });

  expect(parsed.success).toBe(false);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts -t "accepts a destructive sheet structure delete-row plan"
```

Expected: FAIL with missing exports or invalid discriminated union branches for `sheet_structure_update`, `range_sort_plan`, or `range_filter_plan`.

- [ ] **Step 3: Write the minimal contract implementation**

```ts
export const ConfirmationLevelSchema = z.enum(["standard", "destructive"]);

export const SheetStructureUpdateDataSchema = z.discriminatedUnion("operation", [
  strictObject({
    targetSheet: z.string().min(1).max(128),
    operation: z.literal("delete_rows"),
    startIndex: z.number().int().min(0),
    count: z.number().int().min(1),
    explanation: z.string().min(1).max(12000),
    confidence: z.number().min(0).max(1),
    requiresConfirmation: z.literal(true),
    confirmationLevel: ConfirmationLevelSchema,
    affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional(),
    overwriteRisk: OverwriteRiskSchema.optional()
  }),
  // add remaining row/column, range-layout, freeze, and tab-color operations here
]);

export const RangeSortKeySchema = strictObject({
  columnRef: z.union([z.string().min(1).max(128), z.number().int().min(1)]),
  direction: z.enum(["asc", "desc"]),
  sortOn: z.enum(["values"]).optional()
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: PASS with new schemas, new response types, and updated trace event coverage.

- [ ] **Step 5: Commit**

```bash
git add packages/contracts/src/schemas.ts packages/contracts/src/index.ts packages/contracts/tests/contracts.test.ts
git commit -m "feat: add wave 1 sheet structure and sort filter contracts"
```

### Task 2: Extend Shared Client Types And Previews

**Files:**
- Modify: `packages/shared-client/src/types.ts`
- Modify: `packages/shared-client/src/render.ts`
- Modify: `packages/shared-client/src/index.ts`
- Test: `packages/shared-client/tests/client.test.ts`

- [ ] **Step 1: Write the failing preview tests**

```ts
import { describe, expect, it } from "vitest";
import { getStructuredPreview } from "../src/render.ts";

describe("wave 1 structured previews", () => {
  it("builds a sheet structure summary preview", () => {
    const preview = getStructuredPreview({
      schemaVersion: "1.0.0",
      type: "sheet_structure_update",
      requestId: "req_struct_001",
      hermesRunId: "run_struct_001",
      processedBy: "hermes",
      serviceLabel: "gateway",
      environmentLabel: "test",
      startedAt: "2026-04-19T12:00:00.000Z",
      completedAt: "2026-04-19T12:00:01.000Z",
      durationMs: 1000,
      trace: [{ event: "sheet_structure_update_ready", timestamp: "2026-04-19T12:00:01.000Z" }],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        operation: "insert_rows",
        startIndex: 7,
        count: 3,
        explanation: "Insert three rows above the totals block.",
        confidence: 0.91,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });

    expect(preview).toMatchObject({
      kind: "sheet_structure_update",
      operation: "insert_rows",
      targetSheet: "Sheet1"
    });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts -t "builds a sheet structure summary preview"
```

Expected: FAIL because `getStructuredPreview()` and `WritePlan` do not know the new types yet.

- [ ] **Step 3: Write the minimal shared-client implementation**

```ts
export type WritePlan =
  | SheetImportPlanData
  | SheetUpdateData
  | WorkbookStructureUpdateData
  | RangeFormatUpdateData
  | SheetStructureUpdateData
  | RangeSortPlanData
  | RangeFilterPlanData;

export type SheetStructureWritebackResult = {
  kind: "sheet_structure_update";
  hostPlatform: HermesRequest["host"]["platform"];
  operation: SheetStructureUpdateData["operation"];
  targetSheet: string;
  summary: string;
};

export type RangeSortWritebackResult = {
  kind: "range_sort";
  hostPlatform: HermesRequest["host"]["platform"];
  targetSheet: string;
  targetRange: string;
  summary: string;
};

export type RangeFilterWritebackResult = {
  kind: "range_filter";
  hostPlatform: HermesRequest["host"]["platform"];
  targetSheet: string;
  targetRange: string;
  summary: string;
};
```

```ts
case "sheet_structure_update":
  return {
    kind: "sheet_structure_update",
    targetSheet: response.data.targetSheet,
    operation: response.data.operation,
    summary: response.data.explanation,
    confirmationLevel: response.data.confirmationLevel
  };

case "range_sort_plan":
  return {
    kind: "range_sort_plan",
    targetSheet: response.data.targetSheet,
    targetRange: response.data.targetRange,
    hasHeader: response.data.hasHeader,
    keys: response.data.keys,
    explanation: response.data.explanation
  };

case "range_filter_plan":
  return {
    kind: "range_filter_plan",
    targetSheet: response.data.targetSheet,
    targetRange: response.data.targetRange,
    hasHeader: response.data.hasHeader,
    conditions: response.data.conditions,
    combiner: response.data.combiner,
    clearExistingFilters: response.data.clearExistingFilters,
    explanation: response.data.explanation
  };
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: PASS with preview branches and writeback result unions for wave 1.

- [ ] **Step 5: Commit**

```bash
git add packages/shared-client/src/types.ts packages/shared-client/src/render.ts packages/shared-client/src/index.ts packages/shared-client/tests/client.test.ts
git commit -m "feat: add shared client preview support for wave 1 plans"
```

### Task 3: Teach Hermes Runtime And Request Routing About Wave 1

**Files:**
- Modify: `services/gateway/src/hermes/runtimeRules.ts`
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
- Modify: `services/gateway/src/hermes/structuredBody.ts`
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`

- [ ] **Step 1: Write the failing routing and structured-body tests**

```ts
it("documents wave 1 response types in runtime rules", () => {
  expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="sheet_structure_update"');
  expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="range_sort_plan"');
  expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="range_filter_plan"');
});

it("routes sort prompts toward range_sort_plan", () => {
  const prompt = buildHermesSpreadsheetRequestPrompt({
    ...baseRequest,
    userMessage: "Sort this table by Status asc and Due Date desc"
  });

  expect(prompt).toContain('Prefer type="range_sort_plan"');
});

it("rejects extra keys inside a range_filter_plan data object", async () => {
  const response = normalizeHermesStructuredBodyInput({
    type: "range_filter_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      conditions: [{ columnRef: "Status", operator: "equals", value: "Open" }],
      combiner: "and",
      clearExistingFilters: true,
      explanation: "Filter open rows.",
      confidence: 0.9,
      requiresConfirmation: true,
      unexpected: "drop-me"
    }
  });

  expect(() => HermesStructuredBodySchema.parse(response)).toThrow();
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: FAIL because wave 1 types and routing hints are not recognized yet.

- [ ] **Step 3: Write the minimal runtime and validation implementation**

```ts
const STRUCTURED_BODY_TYPES = [
  "chat",
  "formula",
  "workbook_structure_update",
  "range_format_update",
  "sheet_structure_update",
  "range_sort_plan",
  "range_filter_plan",
  "sheet_update",
  "sheet_import_plan",
  "error",
  "attachment_analysis",
  "extracted_table",
  "document_summary"
] as const;
```

```ts
if (/\bsort\b/.test(userMessage) && /\btable\b|\brange\b|[A-Z]+\d+\:[A-Z]+\d+/i.test(userMessage)) {
  return "range_sort_plan";
}

if (/\bfilter\b/.test(userMessage)) {
  return "range_filter_plan";
}

if (/\binsert\b|\bdelete\b|\bhide\b|\bunhide\b|\bmerge\b|\bfreeze\b|\bgroup\b|\bautofit\b/.test(userMessage)) {
  return "sheet_structure_update";
}
```

```ts
createStructuredBodySchema("sheet_structure_update", SheetStructureUpdateDataSchema),
createStructuredBodySchema("range_sort_plan", RangeSortPlanDataSchema),
createStructuredBodySchema("range_filter_plan", RangeFilterPlanDataSchema),
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: PASS with new response type guidance and no silent coercion into old plan families.

- [ ] **Step 5: Commit**

```bash
git add services/gateway/src/hermes/runtimeRules.ts services/gateway/src/hermes/requestTemplate.ts services/gateway/src/hermes/structuredBody.ts services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
git commit -m "feat: route Hermes runtime to wave 1 structure sort filter plans"
```

### Task 4: Add Destructive Approval And New Result Types In Gateway Writeback

**Files:**
- Modify: `services/gateway/src/routes/writeback.ts`
- Modify: `services/gateway/src/lib/traceBus.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

- [ ] **Step 1: Write the failing gateway approval tests**

```ts
it("rejects destructive row deletion approval without a second confirmation payload", () => {
  const plan = {
    targetSheet: "Sheet1",
    operation: "delete_rows",
    startIndex: 7,
    count: 2,
    explanation: "Delete two empty rows.",
    confidence: 0.91,
    requiresConfirmation: true,
    confirmationLevel: "destructive"
  };

  expect(() => approveWriteback({
    requestId: "req_delete_rows",
    runId: "run_delete_rows",
    plan,
    traceBus,
    config
  })).toThrow("Destructive confirmation required.");
});

it("records a range sort completion with a typed result", () => {
  const completion = completeWriteback({
    requestId: "req_sort",
    runId: "run_sort",
    approvalToken,
    planDigest,
    result: {
      kind: "range_sort",
      hostPlatform: "google_sheets",
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      summary: "Sorted Sheet1!A1:F25 by Status asc, Due Date desc."
    },
    traceBus,
    config
  });

  expect(completion.ok).toBe(true);
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: FAIL because the approval schema and completion result union do not yet support the new plan/result families or destructive second-confirm checks.

- [ ] **Step 3: Write the minimal writeback implementation**

```ts
const ApprovalRequestSchema = z.object({
  requestId: z.string().min(1),
  runId: z.string().min(1),
  destructiveConfirmation: z.object({
    confirmed: z.literal(true)
  }).optional(),
  plan: z.union([
    SheetUpdateDataSchema,
    SheetImportPlanDataSchema,
    WorkbookStructureUpdateDataSchema,
    RangeFormatUpdateDataSchema,
    SheetStructureUpdateDataSchema,
    RangeSortPlanDataSchema,
    RangeFilterPlanDataSchema
  ])
});

function requiresDestructiveConfirmation(plan: z.infer<typeof ApprovalRequestSchema>["plan"]) {
  return "confirmationLevel" in plan && plan.confirmationLevel === "destructive";
}

if (requiresDestructiveConfirmation(input.plan) && !input.destructiveConfirmation?.confirmed) {
  throw new Error("Destructive confirmation required.");
}
```

```ts
const CompletionRequestSchema = z.object({
  requestId: z.string().min(1),
  runId: z.string().min(1),
  approvalToken: z.string().min(1),
  planDigest: z.string().min(1),
  result: z.union([
    RangeWritebackResultSchema,
    WorkbookStructureWritebackResultSchema,
    SheetStructureWritebackResultSchema,
    RangeSortWritebackResultSchema,
    RangeFilterWritebackResultSchema
  ])
});
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: PASS with explicit destructive gating and typed completion results.

- [ ] **Step 5: Commit**

```bash
git add services/gateway/src/routes/writeback.ts services/gateway/src/lib/traceBus.ts services/gateway/tests/writebackFlow.test.ts
git commit -m "feat: add wave 1 writeback approval and result handling"
```

### Task 5: Add Excel Helper Modules For Structure, Sort, And Filter

**Files:**
- Create: `apps/excel-addin/src/taskpane/structurePlan.js`
- Create: `apps/excel-addin/src/taskpane/sortFilterPlan.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Test: `services/gateway/tests/excelWave1Plans.test.ts`
- Test: `services/gateway/tests/excelWritePlan.test.ts`

- [ ] **Step 1: Write the failing Excel helper tests**

```ts
import { describe, expect, it } from "vitest";
import {
  isSheetStructurePlan,
  getSheetStructureStatusSummary
} from "../../../apps/excel-addin/src/taskpane/structurePlan.js";
import {
  isRangeSortPlan,
  isRangeFilterPlan,
  buildExcelSortFields,
  getRangeSortStatusSummary,
  getRangeFilterStatusSummary
} from "../../../apps/excel-addin/src/taskpane/sortFilterPlan.js";

describe("Excel wave 1 plan helpers", () => {
  it("detects structure, sort, and filter plans", () => {
    expect(isSheetStructurePlan({ operation: "insert_rows", targetSheet: "Sheet1" })).toBe(true);
    expect(isRangeSortPlan({ targetSheet: "Sheet1", targetRange: "A1:F25", keys: [{ columnRef: "Status", direction: "asc" }] })).toBe(true);
    expect(isRangeFilterPlan({ targetSheet: "Sheet1", targetRange: "A1:F25", conditions: [{ columnRef: "Status", operator: "equals", value: "Open" }], combiner: "and" })).toBe(true);
  });

  it("builds readable Excel status summaries", () => {
    expect(getSheetStructureStatusSummary({
      operation: "insert_rows",
      targetSheet: "Sheet1",
      startIndex: 7,
      count: 3
    })).toBe("Inserted 3 rows at Sheet1 row 8.");

    expect(getRangeFilterStatusSummary({
      targetSheet: "Sheet1",
      targetRange: "A1:F25"
    })).toBe("Applied filter to Sheet1!A1:F25");
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/excelWave1Plans.test.ts services/gateway/tests/excelWritePlan.test.ts
```

Expected: FAIL because the helper modules do not exist yet.

- [ ] **Step 3: Write the minimal Excel helper implementation**

```js
export function isSheetStructurePlan(plan) {
  return Boolean(plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.operation === "string" &&
    [
      "insert_rows",
      "delete_rows",
      "hide_rows",
      "unhide_rows",
      "group_rows",
      "ungroup_rows",
      "insert_columns",
      "delete_columns",
      "hide_columns",
      "unhide_columns",
      "group_columns",
      "ungroup_columns",
      "merge_cells",
      "unmerge_cells",
      "freeze_panes",
      "unfreeze_panes",
      "autofit_rows",
      "autofit_columns",
      "set_sheet_tab_color"
    ].includes(plan.operation));
}

export function getSheetStructureStatusSummary(plan) {
  if (plan.operation === "insert_rows") {
    return `Inserted ${plan.count} rows at ${plan.targetSheet} row ${plan.startIndex + 1}.`;
  }
  return "Applied sheet structure update.";
}
```

```js
export function isRangeSortPlan(plan) {
  return Boolean(plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    Array.isArray(plan.keys));
}

export function isRangeFilterPlan(plan) {
  return Boolean(plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    Array.isArray(plan.conditions));
}

export function buildExcelSortFields(plan) {
  return plan.keys.map(function(key) {
    return {
      key: key.columnRef,
      ascending: key.direction !== "desc"
    };
  });
}

export function getRangeSortStatusSummary(plan) {
  return `Sorted ${plan.targetSheet}!${plan.targetRange}`;
}

export function getRangeFilterStatusSummary(plan) {
  return `Applied filter to ${plan.targetSheet}!${plan.targetRange}`;
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/excelWave1Plans.test.ts services/gateway/tests/excelWritePlan.test.ts
```

Expected: PASS with helper detection and summary text coverage.

- [ ] **Step 5: Commit**

```bash
git add apps/excel-addin/src/taskpane/structurePlan.js apps/excel-addin/src/taskpane/sortFilterPlan.js apps/excel-addin/src/taskpane/writePlan.js services/gateway/tests/excelWave1Plans.test.ts services/gateway/tests/excelWritePlan.test.ts
git commit -m "feat: add Excel helper modules for wave 1 plans"
```

### Task 6: Integrate Wave 1 Apply Paths In Excel

**Files:**
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Test: `services/gateway/tests/excelWave1Plans.test.ts`

- [ ] **Step 1: Write the failing Excel integration tests**

```ts
it("builds a structured preview message for a sort plan", () => {
  expect(getResponseBodyText({
    type: "range_sort_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      keys: [
        { columnRef: "Status", direction: "asc" },
        { columnRef: "Due Date", direction: "desc" }
      ],
      explanation: "Sort by status then due date.",
      confidence: 0.93,
      requiresConfirmation: true
    }
  })).toContain("Prepared a sort plan");
});

it("returns a typed writeback result for an applied filter plan", async () => {
  const result = await applyWritePlan({
    plan: {
      targetSheet: "Sheet1",
      targetRange: "A1:F25",
      hasHeader: true,
      conditions: [{ columnRef: "Status", operator: "equals", value: "Open" }],
      combiner: "and",
      clearExistingFilters: true,
      explanation: "Filter open rows.",
      confidence: 0.9,
      requiresConfirmation: true
    },
    requestId: "req_filter_001",
    runId: "run_filter_001",
    approvalToken: "token"
  });

  expect(result.kind).toBe("range_filter");
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/excelWave1Plans.test.ts
```

Expected: FAIL because `taskpane.js` does not yet render or apply wave 1 plans.

- [ ] **Step 3: Write the minimal Excel integration**

```js
import {
  getSheetStructureStatusSummary,
  isSheetStructurePlan
} from "./structurePlan.js";
import {
  buildExcelSortFields,
  getRangeFilterStatusSummary,
  getRangeSortStatusSummary,
  isRangeFilterPlan,
  isRangeSortPlan
} from "./sortFilterPlan.js";
```

```js
case "sheet_structure_update":
  return `Prepared a sheet structure update for ${response.data.targetSheet}.`;
case "range_sort_plan":
  return `Prepared a sort plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
case "range_filter_plan":
  return `Prepared a filter plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
```

```js
if (isSheetStructurePlan(plan)) {
  return Excel.run(async (context) => {
    // resolve worksheet and apply structure mutation
    return {
      kind: "sheet_structure_update",
      hostPlatform: detectExcelPlatform(),
      operation: plan.operation,
      targetSheet: plan.targetSheet,
      summary: getSheetStructureStatusSummary(plan)
    };
  });
}

if (isRangeSortPlan(plan)) {
  return Excel.run(async (context) => {
    // worksheet.getRange(plan.targetRange).sort.apply(...)
    return {
      kind: "range_sort",
      hostPlatform: detectExcelPlatform(),
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      summary: getRangeSortStatusSummary(plan)
    };
  });
}

if (isRangeFilterPlan(plan)) {
  return Excel.run(async (context) => {
    // worksheet.autoFilter.apply(...)
    return {
      kind: "range_filter",
      hostPlatform: detectExcelPlatform(),
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      summary: getRangeFilterStatusSummary(plan)
    };
  });
}
```

- [ ] **Step 4: Run tests and syntax checks**

Run:

```bash
npm test -- services/gateway/tests/excelWave1Plans.test.ts services/gateway/tests/excelWritePlan.test.ts
node --check <repo-root>/apps/excel-addin/src/taskpane/structurePlan.js
node --check <repo-root>/apps/excel-addin/src/taskpane/sortFilterPlan.js
node --check <repo-root>/apps/excel-addin/src/taskpane/taskpane.js
```

Expected: PASS with valid syntax and typed result coverage.

- [ ] **Step 5: Commit**

```bash
git add apps/excel-addin/src/taskpane/taskpane.js apps/excel-addin/src/taskpane/writePlan.js services/gateway/tests/excelWave1Plans.test.ts
git commit -m "feat: integrate wave 1 plan execution in Excel host"
```

### Task 7: Add Google Sheets Helper Compilation For Wave 1

**Files:**
- Create: `apps/google-sheets-addon/src/Wave1Plans.js`
- Test: `services/gateway/tests/googleSheetsWave1Plans.test.ts`

- [ ] **Step 1: Write the failing Google helper tests**

```ts
import { describe, expect, it } from "vitest";
// eslint-disable-next-line @typescript-eslint/no-require-imports
const {
  isSheetStructurePlan_,
  isRangeSortPlan_,
  isRangeFilterPlan_,
  buildSortSpec_,
  getSheetStructureStatusSummary_,
  getRangeSortStatusSummary_,
  getRangeFilterStatusSummary_
} = require("../../../apps/google-sheets-addon/src/Wave1Plans.js");

describe("Google Sheets wave 1 helper compilation", () => {
  it("detects wave 1 plan types", () => {
    expect(isSheetStructurePlan_({ operation: "merge_cells", targetSheet: "Sheet1", targetRange: "A1:C1" })).toBe(true);
    expect(isRangeSortPlan_({ targetSheet: "Sheet1", targetRange: "A1:F25", keys: [{ columnRef: "Status", direction: "asc" }] })).toBe(true);
    expect(isRangeFilterPlan_({ targetSheet: "Sheet1", targetRange: "A1:F25", conditions: [{ columnRef: "Status", operator: "equals", value: "Open" }], combiner: "and" })).toBe(true);
  });

  it("compiles a simple sort spec", () => {
    expect(buildSortSpec_({
      hasHeader: true,
      keys: [{ columnRef: 2, direction: "desc" }]
    })).toEqual([{ dimensionIndex: 1, sortOrder: "DESCENDING" }]);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave1Plans.test.ts
```

Expected: FAIL because the helper file does not exist yet.

- [ ] **Step 3: Write the minimal Google helper implementation**

```js
function isSheetStructurePlan_(plan) {
  return Boolean(plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.operation === "string");
}

function isRangeSortPlan_(plan) {
  return Boolean(plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    Array.isArray(plan.keys));
}

function isRangeFilterPlan_(plan) {
  return Boolean(plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    Array.isArray(plan.conditions));
}

function getSheetStructureStatusSummary_(plan) {
  if (plan.operation === "merge_cells") {
    return "Merged " + plan.targetSheet + "!" + plan.targetRange + ".";
  }

  return "Applied sheet structure update.";
}

function buildSortSpec_(plan) {
  return plan.keys.map(function(key) {
    const zeroBased = typeof key.columnRef === "number" ? key.columnRef - 1 : 0;
    return {
      dimensionIndex: zeroBased,
      sortOrder: key.direction === "desc" ? "DESCENDING" : "ASCENDING"
    };
  });
}

function getRangeSortStatusSummary_(plan) {
  return "Sorted " + plan.targetSheet + "!" + plan.targetRange + ".";
}

function getRangeFilterStatusSummary_(plan) {
  return "Applied filter to " + plan.targetSheet + "!" + plan.targetRange;
}

if (typeof module !== "undefined") {
  module.exports = {
    isSheetStructurePlan_,
    isRangeSortPlan_,
    isRangeFilterPlan_,
    buildSortSpec_,
    getSheetStructureStatusSummary_,
    getRangeSortStatusSummary_,
    getRangeFilterStatusSummary_
  };
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave1Plans.test.ts
```

Expected: PASS with helper detection and compilation coverage.

- [ ] **Step 5: Commit**

```bash
git add apps/google-sheets-addon/src/Wave1Plans.js services/gateway/tests/googleSheetsWave1Plans.test.ts
git commit -m "feat: add Google Sheets helper compilation for wave 1 plans"
```

### Task 8: Integrate Wave 1 Apply Paths In Google Sheets

**Files:**
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave1Plans.test.ts`

- [ ] **Step 1: Write the failing Google integration tests**

```ts
it("builds a filter status summary for Google Sheets", () => {
  expect(getRangeFilterStatusSummary_({
    targetSheet: "Sheet1",
    targetRange: "A1:F25"
  })).toBe("Applied filter to Sheet1!A1:F25");
});

it("renders a wave 1 preview in the sidebar client bundle", () => {
  expect(sidebarScript).toContain("range_sort_plan");
  expect(sidebarScript).toContain("range_filter_plan");
  expect(sidebarScript).toContain("sheet_structure_update");
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave1Plans.test.ts
```

Expected: FAIL because `Code.gs` and `Sidebar.js.html` do not yet branch on the new plan types.

- [ ] **Step 3: Write the minimal Google Sheets integration**

```js
function applyWave1Plan_(spreadsheet, plan) {
  if (isSheetStructurePlan_(plan)) {
    return {
      kind: "sheet_structure_update",
      hostPlatform: "google_sheets",
      operation: plan.operation,
      targetSheet: plan.targetSheet,
      summary: getSheetStructureStatusSummary_(plan)
    };
  }

  if (isRangeSortPlan_(plan)) {
    return {
      kind: "range_sort",
      hostPlatform: "google_sheets",
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      summary: getRangeSortStatusSummary_(plan)
    };
  }

  if (isRangeFilterPlan_(plan)) {
    return {
      kind: "range_filter",
      hostPlatform: "google_sheets",
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      summary: getRangeFilterStatusSummary_(plan)
    };
  }

  return null;
}
```

```js
case "sheet_structure_update":
  return "Prepared a sheet structure update for " + response.data.targetSheet + ".";
case "range_sort_plan":
  return "Prepared a sort plan for " + response.data.targetSheet + "!" + response.data.targetRange + ".";
case "range_filter_plan":
  return "Prepared a filter plan for " + response.data.targetSheet + "!" + response.data.targetRange + ".";
```

- [ ] **Step 4: Run tests and syntax checks**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave1Plans.test.ts
bash -lc 'awk "/<script>/{flag=1;next}/<\\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave1-sidebar.js && node --check /tmp/hermes-wave1-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave1-code-gs.js && node --check /tmp/hermes-wave1-code-gs.js'
```

Expected: PASS with valid Apps Script syntax and sidebar branching.

- [ ] **Step 5: Commit**

```bash
git add apps/google-sheets-addon/src/Code.gs apps/google-sheets-addon/html/Sidebar.js.html
git commit -m "feat: integrate wave 1 plan execution in Google Sheets host"
```

### Task 9: Run Full Wave 1 Regression And Build Verification

**Files:**
- Modify: `services/gateway/tests/requestTemplate.test.ts`
- Modify: `services/gateway/tests/runtimeRules.test.ts`
- Modify: `services/gateway/tests/writebackFlow.test.ts`
- Modify: `services/gateway/tests/hermesClient.test.ts`
- Modify: `services/gateway/tests/excelWritePlan.test.ts`
- Modify: `services/gateway/tests/googleSheetsWave1Plans.test.ts`

- [ ] **Step 1: Add the final regression assertions**

```ts
it("keeps existing plan families valid alongside wave 1", () => {
  expect(HermesResponseSchema.parse(existingSheetUpdateResponse).type).toBe("sheet_update");
  expect(HermesResponseSchema.parse(existingSheetImportResponse).type).toBe("sheet_import_plan");
  expect(HermesResponseSchema.parse(existingWorkbookStructureResponse).type).toBe("workbook_structure_update");
  expect(HermesResponseSchema.parse(existingRangeFormatResponse).type).toBe("range_format_update");
});
```

- [ ] **Step 2: Run the focused wave 1 suite**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts packages/shared-client/tests/client.test.ts services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts services/gateway/tests/writebackFlow.test.ts services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave1Plans.test.ts services/gateway/tests/googleSheetsWave1Plans.test.ts
```

Expected: PASS with wave 1 and existing write-plan behaviors green.

- [ ] **Step 3: Run the broader regression suite**

Run:

```bash
npm test -- services/gateway/tests/requestRouter.test.ts services/gateway/tests/structuredBody.test.ts services/gateway/tests/traceBus.test.ts services/gateway/tests/uploads.test.ts services/gateway/tests/excelCellValues.test.ts services/gateway/tests/excelReferencedCells.test.ts services/gateway/tests/googleSheetsReferencedCells.test.ts services/gateway/tests/rangeSafety.test.ts
```

Expected: PASS with no regressions to existing request validation, upload, and host context logic.

- [ ] **Step 4: Run syntax and build verification**

Run:

```bash
node --check <repo-root>/apps/excel-addin/src/taskpane/structurePlan.js
node --check <repo-root>/apps/excel-addin/src/taskpane/sortFilterPlan.js
node --check <repo-root>/apps/excel-addin/src/taskpane/taskpane.js
bash -lc 'awk "/<script>/{flag=1;next}/<\\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave1-sidebar.js && node --check /tmp/hermes-wave1-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave1-code-gs.js && node --check /tmp/hermes-wave1-code-gs.js'
npm --workspace @hermes/gateway run build
```

Expected: PASS with clean syntax and gateway build success.

- [ ] **Step 5: Commit**

```bash
git add services/gateway/tests/requestTemplate.test.ts services/gateway/tests/runtimeRules.test.ts services/gateway/tests/writebackFlow.test.ts services/gateway/tests/hermesClient.test.ts services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave1Plans.test.ts services/gateway/tests/googleSheetsWave1Plans.test.ts
git commit -m "test: verify wave 1 sheet structure sort filter regression coverage"
```
