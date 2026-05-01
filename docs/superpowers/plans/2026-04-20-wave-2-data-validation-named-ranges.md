# Wave 2 Data Validation + Named Ranges Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add contract-valid `data_validation_plan` and `named_range_update` flows with strict preview, approval, and host apply support for Excel and Google Sheets.

**Architecture:** Extend the existing typed-plan pipeline with two new strict plan families instead of introducing a generic metadata plan. Add contract schemas, runtime/request routing, structured-body parsing, writeback support, shared previews, then host-specific apply paths for Excel and Google Sheets with fail-closed semantics for unsupported exact behavior.

**Tech Stack:** TypeScript, Zod, Vitest, Express, Office.js, Google Apps Script, browser-safe shared client helpers.

---

## File Structure

### Contracts and shared types

- Modify: `packages/contracts/src/schemas.ts`
  - add `data_validation_plan`, `named_range_update`, and their result kinds
- Modify: `packages/contracts/src/index.ts`
  - export the new schemas and inferred types
- Test: `packages/contracts/tests/contracts.test.ts`
  - schema coverage for validation-rule variants, comparators, and named-range operation invariants

### Shared preview and client types

- Modify: `packages/shared-client/src/types.ts`
  - extend `WritePlan`, `WritebackResult`, and structured preview unions
- Modify: `packages/shared-client/src/render.ts`
  - add non-lossy previews for validation and named ranges
- Modify: `packages/shared-client/src/index.ts`
  - export any new preview helpers/types
- Test: `packages/shared-client/tests/client.test.ts`
  - preview coverage and result typing regression

### Gateway runtime and writeback

- Modify: `services/gateway/src/hermes/runtimeRules.ts`
  - teach Hermes the new response types and fail-closed behavior
- Modify: `services/gateway/src/hermes/requestTemplate.ts`
  - route prompts toward validation or named-range plans
- Modify: `services/gateway/src/hermes/structuredBody.ts`
  - normalize and validate the new plan families
- Modify: `services/gateway/src/lib/hermesClient.ts`
  - map UI/trace behavior for the new plan families
- Modify: `services/gateway/src/routes/writeback.ts`
  - add approval/completion support for the new result kinds
- Modify: `services/gateway/src/lib/traceBus.ts`
  - persist typed completion state for the new plans
- Test: `services/gateway/tests/runtimeRules.test.ts`
- Test: `services/gateway/tests/requestTemplate.test.ts`
- Test: `services/gateway/tests/hermesClient.test.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

### Excel host

- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
  - integrate preview and apply paths for validation and named ranges
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
  - extend status summaries for the new result kinds if needed
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave2Plans.test.ts`

### Google Sheets host

- Modify: `apps/google-sheets-addon/src/Code.gs`
  - apply validation and named-range plans
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
  - preview and confirmation handling for both plan families
- Test: `services/gateway/tests/googleSheetsWave2Plans.test.ts`

### Final regression

- Re-run existing tests covering:
  - `sheet_update`
  - `sheet_import_plan`
  - `workbook_structure_update`
  - `range_format_update`
  - wave 1 plan families

Note: this workspace currently has no active `.git` metadata, so execution should skip commit steps locally unless the repo is re-initialized or restored.

---

### Task 1: Add Wave 2 Contract Schemas

**Files:**
- Modify: `packages/contracts/src/schemas.ts`
- Modify: `packages/contracts/src/index.ts`
- Test: `packages/contracts/tests/contracts.test.ts`

- [ ] **Step 1: Write the failing contract tests**

```ts
import {
  DataValidationPlanDataSchema,
  HermesResponseSchema,
  NamedRangeUpdateDataSchema
} from "../src/index.ts";

it("accepts a list validation plan backed by a named range", () => {
  const parsed = DataValidationPlanDataSchema.parse({
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    ruleType: "list",
    namedRangeName: "StatusOptions",
    showDropdown: true,
    allowBlank: false,
    invalidDataBehavior: "reject",
    helpText: "Choose a valid status.",
    explanation: "Restrict the status column to approved options.",
    confidence: 0.95,
    requiresConfirmation: true,
    replacesExistingValidation: true
  });

  expect(parsed.ruleType).toBe("list");
  expect(parsed.namedRangeName).toBe("StatusOptions");
});

it("rejects list validation plans that define both values and sourceRange", () => {
  const parsed = DataValidationPlanDataSchema.safeParse({
    targetSheet: "Sheet1",
    targetRange: "B2:B20",
    ruleType: "list",
    values: ["Open", "Closed"],
    sourceRange: "Lookup!A1:A2",
    allowBlank: false,
    invalidDataBehavior: "reject",
    explanation: "Invalid double-defined list source.",
    confidence: 0.6,
    requiresConfirmation: true
  });

  expect(parsed.success).toBe(false);
});

it("accepts a sheet-scoped named range retarget update", () => {
  const parsed = NamedRangeUpdateDataSchema.parse({
    operation: "retarget",
    name: "InputRange",
    scope: "sheet",
    sheetName: "Sheet1",
    targetSheet: "Sheet1",
    targetRange: "B2:D20",
    explanation: "Retarget the named input block.",
    confidence: 0.91,
    requiresConfirmation: true
  });

  expect(parsed.operation).toBe("retarget");
  expect(parsed.scope).toBe("sheet");
});

it("accepts a data_validation_plan response envelope", () => {
  const parsed = HermesResponseSchema.parse({
    schemaVersion: "1.0.0",
    type: "data_validation_plan",
    requestId: "req_validation_001",
    hermesRunId: "run_validation_001",
    processedBy: "hermes",
    serviceLabel: "hermes-gateway-local",
    environmentLabel: "local-dev",
    startedAt: "2026-04-20T09:00:00.000Z",
    completedAt: "2026-04-20T09:00:01.000Z",
    durationMs: 1000,
    trace: [
      { event: "data_validation_plan_ready", timestamp: "2026-04-20T09:00:01.000Z" }
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
      targetRange: "B2:B20",
      ruleType: "whole_number",
      comparator: "between",
      value: 1,
      value2: 10,
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Restrict values to integers between 1 and 10.",
      confidence: 0.93,
      requiresConfirmation: true
    }
  });

  expect(parsed.type).toBe("data_validation_plan");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: FAIL with missing exports or missing `data_validation_plan` / `named_range_update` response branches.

- [ ] **Step 3: Write the minimal contract implementation**

```ts
const ValidationComparatorSchema = z.enum([
  "between",
  "not_between",
  "equal_to",
  "not_equal_to",
  "greater_than",
  "greater_than_or_equal_to",
  "less_than",
  "less_than_or_equal_to"
]);

const InvalidDataBehaviorSchema = z.enum(["warn", "reject"]);

const ValidationListSourceSchema = z.union([
  strictObject({
    ruleType: z.literal("list"),
    values: z.array(z.string().min(1).max(256)).min(1).max(500),
    showDropdown: z.boolean().optional()
  }),
  strictObject({
    ruleType: z.literal("list"),
    sourceRange: z.string().min(1).max(128),
    showDropdown: z.boolean().optional()
  }),
  strictObject({
    ruleType: z.literal("list"),
    namedRangeName: z.string().min(1).max(255),
    showDropdown: z.boolean().optional()
  })
]);

export const DataValidationPlanDataSchema = z.discriminatedUnion("ruleType", [
  ValidationListSourceSchema.extend({
    targetSheet: z.string().min(1).max(128),
    targetRange: z.string().min(1).max(128),
    allowBlank: z.boolean(),
    invalidDataBehavior: InvalidDataBehaviorSchema,
    helpText: z.string().min(1).max(500).optional(),
    explanation: z.string().min(1).max(12000),
    confidence: z.number().min(0).max(1),
    requiresConfirmation: z.literal(true),
    affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional(),
    replacesExistingValidation: z.boolean().optional()
  }),
  strictObject({
    targetSheet: z.string().min(1).max(128),
    targetRange: z.string().min(1).max(128),
    ruleType: z.literal("whole_number"),
    comparator: ValidationComparatorSchema,
    value: z.number(),
    value2: z.number().optional(),
    allowBlank: z.boolean(),
    invalidDataBehavior: InvalidDataBehaviorSchema,
    helpText: z.string().min(1).max(500).optional(),
    explanation: z.string().min(1).max(12000),
    confidence: z.number().min(0).max(1),
    requiresConfirmation: z.literal(true),
    affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional(),
    replacesExistingValidation: z.boolean().optional()
  }),
  // repeat for checkbox, decimal, date, text_length, custom_formula
]);

export const NamedRangeUpdateDataSchema = strictObject({
  operation: z.enum(["create", "rename", "delete", "retarget"]),
  name: z.string().min(1).max(255),
  scope: z.enum(["workbook", "sheet"]),
  sheetName: z.string().min(1).max(128).optional(),
  targetSheet: z.string().min(1).max(128).optional(),
  targetRange: z.string().min(1).max(128).optional(),
  newName: z.string().min(1).max(255).optional(),
  explanation: z.string().min(1).max(12000),
  confidence: z.number().min(0).max(1),
  requiresConfirmation: z.literal(true),
  affectedRanges: z.array(z.string().min(1).max(128)).max(10).optional(),
  overwriteRisk: OverwriteRiskSchema.optional()
}).superRefine((data, ctx) => {
  if (data.scope === "sheet" && !data.sheetName) {
    ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["sheetName"], message: "sheet-scoped names require sheetName." });
  }
  if (data.scope === "workbook" && data.sheetName) {
    ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["sheetName"], message: "workbook-scoped names must omit sheetName." });
  }
  if ((data.operation === "create" || data.operation === "retarget") && (!data.targetSheet || !data.targetRange)) {
    ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["targetRange"], message: `${data.operation} requires targetSheet and targetRange.` });
  }
  if (data.operation === "rename" && !data.newName) {
    ctx.addIssue({ code: z.ZodIssueCode.custom, path: ["newName"], message: "rename requires newName." });
  }
});
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts
```

Expected: PASS with new schema branches accepted and invalid combinations rejected.

---

### Task 2: Add Shared Preview And Client Types

**Files:**
- Modify: `packages/shared-client/src/types.ts`
- Modify: `packages/shared-client/src/render.ts`
- Modify: `packages/shared-client/src/index.ts`
- Test: `packages/shared-client/tests/client.test.ts`

- [ ] **Step 1: Write the failing shared-client tests**

```ts
import { buildStructuredPreview } from "../src/render.ts";

it("renders a non-lossy validation preview", () => {
  const preview = buildStructuredPreview({
    type: "data_validation_plan",
    data: {
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      ruleType: "list",
      values: ["Open", "Closed"],
      allowBlank: false,
      invalidDataBehavior: "reject",
      helpText: "Choose a valid status.",
      explanation: "Restrict the status column.",
      confidence: 0.95,
      requiresConfirmation: true,
      replacesExistingValidation: true
    }
  } as any);

  expect(preview?.kind).toBe("data_validation_plan");
  expect(preview?.targetRange).toBe("B2:B20");
  expect(preview?.invalidDataBehavior).toBe("reject");
  expect(preview?.replacesExistingValidation).toBe(true);
});

it("renders a non-lossy named range preview", () => {
  const preview = buildStructuredPreview({
    type: "named_range_update",
    data: {
      operation: "retarget",
      name: "InputRange",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      explanation: "Retarget the named range.",
      confidence: 0.9,
      requiresConfirmation: true
    }
  } as any);

  expect(preview?.kind).toBe("named_range_update");
  expect(preview?.operation).toBe("retarget");
  expect(preview?.scope).toBe("sheet");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
```

Expected: FAIL because the preview union does not yet know the new plan families.

- [ ] **Step 3: Write the minimal shared preview implementation**

```ts
export type WritePlan =
  | WorkbookStructureUpdateResponse["data"]
  | RangeFormatUpdateResponse["data"]
  | SheetStructureUpdateResponse["data"]
  | RangeSortPlanResponse["data"]
  | RangeFilterPlanResponse["data"]
  | DataValidationPlanResponse["data"]
  | NamedRangeUpdateResponse["data"]
  | SheetUpdateResponse["data"]
  | SheetImportPlanResponse["data"];

export type WritebackResult =
  | { kind: "data_validation_update"; hostPlatform: HostPlatform; targetSheet: string; targetRange: string; summary: string }
  | { kind: "named_range_update"; hostPlatform: HostPlatform; operation: string; name: string; summary: string }
  | ExistingWritebackResult;

export function buildDataValidationPreview(plan: DataValidationPlanData) {
  return {
    kind: "data_validation_plan" as const,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    ruleType: plan.ruleType,
    invalidDataBehavior: plan.invalidDataBehavior,
    allowBlank: plan.allowBlank,
    helpText: plan.helpText,
    replacesExistingValidation: Boolean(plan.replacesExistingValidation),
    summary: plan.explanation
  };
}

export function buildNamedRangeUpdatePreview(plan: NamedRangeUpdateData) {
  return {
    kind: "named_range_update" as const,
    operation: plan.operation,
    name: plan.name,
    scope: plan.scope,
    sheetName: plan.sheetName,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    newName: plan.newName,
    overwriteRisk: plan.overwriteRisk,
    summary: plan.explanation
  };
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- packages/shared-client/tests/client.test.ts
npm run build --workspace @hermes/shared-client
```

Expected: PASS with typed previews for both new plan families.

---

### Task 3: Add Gateway Runtime, Request Routing, And Structured-Body Support

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
it("documents validation and named-range response types", () => {
  expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="data_validation_plan"');
  expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="named_range_update"');
});

it("routes validation prompts toward data_validation_plan", () => {
  const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
    userMessage: "Add a dropdown in B2:B20 using the StatusOptions named range."
  }));

  expect(prompt).toContain('Prefer type="data_validation_plan"');
});

it("routes named-range prompts toward named_range_update", () => {
  const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
    userMessage: "Rename the named range SalesData to SalesData2026."
  }));

  expect(prompt).toContain('Prefer type="named_range_update"');
});

it("assembles a final data_validation_plan response with typed ui and trace metadata", async () => {
  // mirror existing hermesClient tests with type=data_validation_plan
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
```

Expected: FAIL because the runtime rules, request template, and Hermes client do not yet recognize the new plan families.

- [ ] **Step 3: Write the minimal gateway-facing implementation**

```ts
// runtimeRules.ts
- For type="data_validation_plan":
  - data.targetSheet is required
  - data.targetRange is required
  - data.ruleType is required
  - data.invalidDataBehavior is required
  - data.allowBlank is required
  - data.explanation is required
  - data.confidence is required
  - data.requiresConfirmation must be true
  - if exact host semantics are unavailable, return type="error" with data.code="UNSUPPORTED_OPERATION"

- For type="named_range_update":
  - data.operation is required and must be create, rename, delete, or retarget
  - data.name is required
  - data.scope is required and must be workbook or sheet
  - data.explanation is required
  - data.confidence is required
  - data.requiresConfirmation must be true

// requestTemplate.ts
- For explicit dropdown, checkbox, validation, allow only, reject invalid, or named-range-backed validation requests, prefer type="data_validation_plan".
- For explicit create, rename, delete, or retarget named range requests, prefer type="named_range_update".

// structuredBody.ts
case "data_validation_plan":
  return normalizeDataValidationPlanData(raw);
case "named_range_update":
  return normalizeNamedRangeUpdateData(raw);

// hermesClient.ts
case "data_validation_plan":
  return {
    displayMode: "structured-preview",
    showTrace: true,
    showWarnings: true,
    showConfidence: true,
    showRequiresConfirmation: true
  };
case "named_range_update":
  return {
    displayMode: "structured-preview",
    showTrace: true,
    showWarnings: true,
    showConfidence: true,
    showRequiresConfirmation: true
  };
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts
npx tsc -p services/gateway/tsconfig.json --noEmit
```

Expected: PASS with routing and final response assembly for both new types.

---

### Task 4: Extend Writeback And Trace Handling

**Files:**
- Modify: `services/gateway/src/routes/writeback.ts`
- Modify: `services/gateway/src/lib/traceBus.ts`
- Test: `services/gateway/tests/writebackFlow.test.ts`

- [ ] **Step 1: Write the failing writeback tests**

```ts
it("approves and completes a data validation plan", async () => {
  // follow existing writebackFlow pattern using type=data_validation_plan
});

it("approves and completes a named range update", async () => {
  // follow existing writebackFlow pattern using type=named_range_update
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: FAIL because approval/completion schema and result-kind handling do not yet include the new plan families.

- [ ] **Step 3: Write the minimal writeback implementation**

```ts
const ApprovablePlanSchema = z.union([
  WorkbookStructureUpdateDataSchema,
  RangeFormatUpdateDataSchema,
  SheetStructureUpdateDataSchema,
  RangeSortPlanDataSchema,
  RangeFilterPlanDataSchema,
  DataValidationPlanDataSchema,
  NamedRangeUpdateDataSchema,
  SheetUpdateDataSchema,
  SheetImportPlanDataSchema
]);

const CompletionResultSchema = z.union([
  RangeWriteResultSchema,
  WorkbookStructureResultSchema,
  SheetStructureResultSchema,
  RangeSortResultSchema,
  RangeFilterResultSchema,
  DataValidationUpdateResultSchema,
  NamedRangeUpdateResultSchema
]);
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/writebackFlow.test.ts
```

Expected: PASS with new writeback kinds flowing through approval/completion.

---

### Task 5: Add Excel Validation And Named Range Apply Paths

**Files:**
- Modify: `apps/excel-addin/src/taskpane/taskpane.js`
- Modify: `apps/excel-addin/src/taskpane/writePlan.js`
- Test: `services/gateway/tests/excelWritePlan.test.ts`
- Test: `services/gateway/tests/excelWave2Plans.test.ts`

- [ ] **Step 1: Write the failing Excel host tests**

```ts
it("applies a whole-number validation rule in Excel", async () => {
  // mock Office.js range.dataValidation and assert the exact rule shape
});

it("creates a workbook-scoped named range in Excel", async () => {
  // mock workbook.names.add and assert it is called with name and range
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave2Plans.test.ts
```

Expected: FAIL because Excel host apply paths and preview/status handling do not yet include the new plan families.

- [ ] **Step 3: Write the minimal Excel implementation**

```js
if (response.type === "data_validation_plan") {
  return `Prepared a validation plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
}

if (response.type === "named_range_update") {
  return `Prepared a named range update for ${response.data.name}.`;
}

if (preview.kind === "data_validation_plan") {
  // render target, rule summary, allowBlank, invalidDataBehavior, helpText
}

if (preview.kind === "named_range_update") {
  // render operation, scope, name, old/new target
}

if (isDataValidationPlan(plan)) {
  const target = sheet.getRange(plan.targetRange);
  target.dataValidation.rule = buildExcelValidationRule(plan);
  return {
    kind: "data_validation_update",
    hostPlatform: platform,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    summary: getDataValidationStatusSummary(plan)
  };
}

if (isNamedRangeUpdate(plan)) {
  applyExcelNamedRangeUpdate(workbook, worksheets, plan);
  return {
    kind: "named_range_update",
    hostPlatform: platform,
    operation: plan.operation,
    name: plan.name,
    summary: getNamedRangeStatusSummary(plan)
  };
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave2Plans.test.ts
node --check <repo-root>/apps/excel-addin/src/taskpane/taskpane.js
```

Expected: PASS with typed Excel apply paths and status summaries.

---

### Task 6: Add Google Sheets Validation And Named Range Apply Paths

**Files:**
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave2Plans.test.ts`

- [ ] **Step 1: Write the failing Google host tests**

```ts
it("applies a list validation rule in Google Sheets", () => {
  // vm-load Code.gs, mock SpreadsheetApp.newDataValidation(), assert criteria wiring
});

it("retargets a workbook named range in Google Sheets", () => {
  // vm-load Code.gs, mock spreadsheet named range APIs, assert target replacement
});

it("renders detailed validation and named-range previews in the sidebar", () => {
  // vm-load Sidebar.js.html script and assert preview strings
});
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave2Plans.test.ts
```

Expected: FAIL because Google host apply and preview logic do not yet include validation or named ranges.

- [ ] **Step 3: Write the minimal Google implementation**

```js
if (isDataValidationPlan_(plan)) {
  const target = sheet.getRange(plan.targetRange);
  target.setDataValidation(buildGoogleSheetsValidationRule_(plan));
  SpreadsheetApp.flush();
  return {
    kind: "data_validation_update",
    hostPlatform: "google_sheets",
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    summary: getDataValidationStatusSummary_(plan)
  };
}

if (isNamedRangeUpdate_(plan)) {
  applyGoogleNamedRangeUpdate_(spreadsheet, plan);
  SpreadsheetApp.flush();
  return {
    kind: "named_range_update",
    hostPlatform: "google_sheets",
    operation: plan.operation,
    name: plan.name,
    summary: getNamedRangeStatusSummary_(plan)
  };
}
```

```js
case "data_validation_plan":
  return `Prepared a validation plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
case "named_range_update":
  return `Prepared a named range update for ${response.data.name}.`;
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave2Plans.test.ts
bash -lc 'awk "/<script>/{flag=1;next}/<\\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave2-sidebar.js && node --check /tmp/hermes-wave2-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave2-code-gs.js && node --check /tmp/hermes-wave2-code-gs.js'
```

Expected: PASS with typed Google apply paths and non-lossy previews.

---

### Task 7: Run Final Wave 2 Regression And Build Verification

**Files:**
- Modify: `packages/contracts/tests/contracts.test.ts`
- Modify: `packages/shared-client/tests/client.test.ts`
- Modify: `services/gateway/tests/runtimeRules.test.ts`
- Modify: `services/gateway/tests/requestTemplate.test.ts`
- Modify: `services/gateway/tests/hermesClient.test.ts`
- Modify: `services/gateway/tests/writebackFlow.test.ts`
- Modify: `services/gateway/tests/excelWritePlan.test.ts`
- Modify: `services/gateway/tests/googleSheetsWave2Plans.test.ts`

- [ ] **Step 1: Add final coexistence assertions**

```ts
it("keeps existing plan families valid alongside wave 2", () => {
  expect(HermesResponseSchema.parse(existingSheetUpdateResponse).type).toBe("sheet_update");
  expect(HermesResponseSchema.parse(existingSheetImportResponse).type).toBe("sheet_import_plan");
  expect(HermesResponseSchema.parse(existingWorkbookStructureResponse).type).toBe("workbook_structure_update");
  expect(HermesResponseSchema.parse(existingRangeFormatResponse).type).toBe("range_format_update");
});
```

- [ ] **Step 2: Run the focused wave 2 suite**

Run:

```bash
npm test -- packages/contracts/tests/contracts.test.ts packages/shared-client/tests/client.test.ts services/gateway/tests/runtimeRules.test.ts services/gateway/tests/requestTemplate.test.ts services/gateway/tests/hermesClient.test.ts services/gateway/tests/writebackFlow.test.ts services/gateway/tests/excelWritePlan.test.ts services/gateway/tests/excelWave2Plans.test.ts services/gateway/tests/googleSheetsWave2Plans.test.ts
```

Expected: PASS with wave 2 and existing write-plan behaviors green.

- [ ] **Step 3: Run the broader regression suite**

Run:

```bash
npm test -- services/gateway/tests/requestRouter.test.ts services/gateway/tests/structuredBody.test.ts services/gateway/tests/traceBus.test.ts services/gateway/tests/uploads.test.ts services/gateway/tests/excelCellValues.test.ts services/gateway/tests/excelReferencedCells.test.ts services/gateway/tests/googleSheetsReferencedCells.test.ts services/gateway/tests/rangeSafety.test.ts
```

Expected: PASS with no regressions to existing request validation, upload, and host context logic.

- [ ] **Step 4: Run syntax and build verification**

Run:

```bash
node --check <repo-root>/apps/excel-addin/src/taskpane/taskpane.js
bash -lc 'awk "/<script>/{flag=1;next}/<\\/script>/{flag=0}flag" <repo-root>/apps/google-sheets-addon/html/Sidebar.js.html > /tmp/hermes-wave2-sidebar.js && node --check /tmp/hermes-wave2-sidebar.js'
bash -lc 'cp <repo-root>/apps/google-sheets-addon/src/Code.gs /tmp/hermes-wave2-code-gs.js && node --check /tmp/hermes-wave2-code-gs.js'
npm --workspace @hermes/gateway run build
```

Expected: PASS with clean syntax and gateway build success.

---

## Self-Review

### Spec Coverage

Covered:
- `data_validation_plan` contract, preview, runtime, writeback, Excel apply, Google apply
- `named_range_update` contract, preview, runtime, writeback, Excel apply, Google apply
- fail-closed unsupported semantics
- workbook vs sheet scope handling
- final coexistence regression

No uncovered spec sections remain for wave 2.

### Placeholder Scan

Checked for:
- `TBD`
- `TODO`
- “implement later”
- vague “add tests” without test content

None remain.

### Type Consistency

Names used consistently across tasks:
- `data_validation_plan`
- `named_range_update`
- `data_validation_update`
- `named_range_update` result kind
- `invalidDataBehavior`
- `replacesExistingValidation`

No internal naming contradictions remain.
