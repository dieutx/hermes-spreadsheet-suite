# Google Sheets Live Demo Reliability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the provided Google Sheets demo workbook reliably support live Hermes read/chat plus confirm-before-write flows, including exact-safe Google Sheets pivot and chart creation.

**Architecture:** Keep the existing host/gateway split intact and treat the local repository as the source of truth. Patch the Google Sheets host where it still fails closed for `pivot_table_plan` and `chart_plan`, align sidebar support messaging with the real host capabilities, then redeploy the bound Apps Script project and verify the full demo matrix on disposable demo tabs in the target workbook.

**Tech Stack:** Google Apps Script, JavaScript, Vitest, Express, Node.js, existing Hermes gateway/contracts/shared-client packages.

---

## File Structure

### Google Sheets host apply layer

- Modify: `apps/google-sheets-addon/src/Code.gs`
  - add exact-safe pivot-table helpers
  - add exact-safe chart helpers
  - keep explicit fail-closed checks for unsupported host semantics

### Google Sheets sidebar behavior

- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
  - treat supported pivot/chart plans as confirmable write previews
  - keep support detection honest for unsupported subcases
  - keep composite previews aligned with the host support matrix

### Regression tests

- Modify: `services/gateway/tests/googleSheetsWave5Plans.test.ts`
  - flip current pivot/chart expectations from unsupported to supported for the demo-safe subset
  - add apply-path coverage for pivot/chart updates
- Modify: `services/gateway/tests/googleSheetsWave6Plans.test.ts`
  - verify composite previews remain confirmable when child pivot/chart steps are supported

### Demo docs and operator checklist

- Modify: `docs/demo-runbook.md`
  - add exact redeploy, property, and prompt guidance for the Google Sheets demo workbook
- Create: `docs/review/google-sheets-live-demo-checklist.md`
  - list demo tabs, prompts, expected previews, and expected final mutations

### Verification outputs

- Reuse: `services/gateway/tests/writebackFlow.test.ts`
  - confirm gateway approval/completion still accepts pivot/chart updates
- Reuse: root `package.json`
  - run existing `npm test` / `npm run build` entry points

Note: this workspace snapshot currently has no `.git` metadata, so execution should skip local commit steps unless repo metadata is restored.

---

### Task 1: Add Exact-Safe Google Sheets Pivot Table Support

**Files:**
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave5Plans.test.ts`

- [ ] **Step 1: Write the failing pivot tests**

Add these expectations to `services/gateway/tests/googleSheetsWave5Plans.test.ts` near the existing pivot/chart Google Sheets preview block:

```ts
it("treats a demo-safe pivot plan as a confirmable Google Sheets write preview", () => {
  const sidebar = loadSidebarContext();
  const response = {
    type: "pivot_table_plan",
    data: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      columnGroups: ["Quarter"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      filters: [{ field: "Status", operator: "equal_to", value: "Closed Won" }],
      sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
      explanation: "Build a pivot table by region and rep.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  };

  expect(sidebar.isWritePlanResponse(response)).toBe(true);
  expect(sidebar.getRequiresConfirmation(response)).toBe(true);

  const html = sidebar.renderStructuredPreview(response, {
    runId: "run_pivot_preview",
    requestId: "req_pivot_preview"
  });

  expect(html).toContain("Will create a pivot table on Sales Pivot!A1.");
  expect(html).toContain("Confirm Pivot Table");
  expect(html).not.toContain("does not support exact-safe pivot table creation yet");
});

it("applies a demo-safe pivot plan through Code.gs", () => {
  const anchorRange = createRangeStub({
    a1Notation: "A1",
    row: 1,
    column: 1,
    numRows: 1,
    numColumns: 1,
    values: [[""]]
  });
  const sourceRange = createRangeStub({
    a1Notation: "A1:F50",
    row: 1,
    column: 1,
    numRows: 50,
    numColumns: 6,
    displayValues: [
      ["Region", "Rep", "Quarter", "Revenue", "Status", "Deals"]
    ]
  });
  const pivotTable = {
    addRowGroup: vi.fn(),
    addColumnGroup: vi.fn(),
    addPivotValue: vi.fn(),
    addFilter: vi.fn()
  };
  anchorRange.createPivotTable = vi.fn(() => pivotTable);

  const spreadsheet = {
    getSheetByName: vi.fn((sheetName: string) => {
      if (sheetName === "Sales") {
        return { getRange: vi.fn(() => sourceRange) };
      }
      if (sheetName === "Sales Pivot") {
        return { getRange: vi.fn(() => anchorRange) };
      }
      return null;
    })
  };

  const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });
  const result = applyWritePlan({
    plan: {
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      targetSheet: "Sales Pivot",
      targetRange: "A1",
      rowGroups: ["Region", "Rep"],
      columnGroups: ["Quarter"],
      valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
      filters: [{ field: "Status", operator: "equal_to", value: "Closed Won" }],
      sort: { field: "Revenue", direction: "desc", sortOn: "aggregated_value" },
      explanation: "Build a pivot table by region and rep.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    }
  });

  expect(anchorRange.createPivotTable).toHaveBeenCalledWith(sourceRange);
  expect(result).toEqual({
    kind: "pivot_table_update",
    operation: "pivot_table_update",
    hostPlatform: "google_sheets",
    targetSheet: "Sales Pivot",
    targetRange: "A1",
    summary: "Created pivot table on Sales Pivot!A1."
  });
  expect(flush).toHaveBeenCalledTimes(1);
});
```

- [ ] **Step 2: Run the targeted pivot tests and verify they fail for current unsupported behavior**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts
```

Expected:

- FAIL because the sidebar still marks pivot plans as unsupported
- FAIL because `applyWritePlan()` still throws `Google Sheets host does not support exact-safe pivot table creation yet.`

- [ ] **Step 3: Implement the minimal exact-safe pivot path in `Code.gs` and sidebar support checks**

Patch `apps/google-sheets-addon/src/Code.gs` by replacing the current hard failure branch with a focused helper set like this:

```js
function buildHeaderMap_(sourceRange) {
  const headerRow = sourceRange.getDisplayValues()[0] || [];
  return headerRow.reduce(function(map, value, index) {
    const key = String(value || "").trim();
    if (key) {
      map[key] = index + 1;
    }
    return map;
  }, {});
}

function requireSingleCellAnchor_(range, kind) {
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    throw new Error("Google Sheets host requires a single-cell target anchor for " + kind + ".");
  }
}

function getPivotSummarizeFunction_(aggregation) {
  switch (aggregation) {
    case "sum":
      return SpreadsheetApp.PivotTableSummarizeFunction.SUM;
    case "count":
      return SpreadsheetApp.PivotTableSummarizeFunction.COUNTA;
    case "average":
      return SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE;
    case "min":
      return SpreadsheetApp.PivotTableSummarizeFunction.MIN;
    case "max":
      return SpreadsheetApp.PivotTableSummarizeFunction.MAX;
    default:
      throw new Error("Unsupported pivot aggregation: " + aggregation);
  }
}

function buildPivotFilterCriteria_(filter) {
  const builder = SpreadsheetApp.newFilterCriteria();
  if (filter.operator === "equal_to") {
    builder.whenTextEqualTo(String(filter.value));
    return builder.build();
  }

  throw new Error("Google Sheets host only supports equal_to pivot filters in the live demo subset.");
}

function applyPivotTablePlan_(spreadsheet, plan) {
  const sourceSheet = spreadsheet.getSheetByName(plan.sourceSheet);
  const targetSheet = spreadsheet.getSheetByName(plan.targetSheet);
  if (!sourceSheet) {
    throw new Error("Source sheet not found: " + plan.sourceSheet);
  }
  if (!targetSheet) {
    throw new Error("Target sheet not found: " + plan.targetSheet);
  }

  const sourceRange = sourceSheet.getRange(plan.sourceRange);
  const anchorRange = targetSheet.getRange(plan.targetRange);
  requireSingleCellAnchor_(anchorRange, "pivot tables");

  if (typeof anchorRange.createPivotTable !== "function") {
    throw new Error("Google Sheets host does not expose pivot creation on this range.");
  }

  const headerMap = buildHeaderMap_(sourceRange);
  const pivotTable = anchorRange.createPivotTable(sourceRange);

  plan.rowGroups.forEach(function(field) {
    pivotTable.addRowGroup(headerMap[field]);
  });
  (plan.columnGroups || []).forEach(function(field) {
    pivotTable.addColumnGroup(headerMap[field]);
  });
  plan.valueAggregations.forEach(function(aggregation) {
    pivotTable.addPivotValue(
      headerMap[aggregation.field],
      getPivotSummarizeFunction_(aggregation.aggregation)
    );
  });
  (plan.filters || []).forEach(function(filter) {
    pivotTable.addFilter(
      headerMap[filter.field],
      buildPivotFilterCriteria_(filter)
    );
  });

  SpreadsheetApp.flush();
  return {
    kind: "pivot_table_update",
    operation: "pivot_table_update",
    hostPlatform: "google_sheets",
    targetSheet: plan.targetSheet,
    targetRange: normalizeA1_(plan.targetRange),
    summary: "Created pivot table on " + plan.targetSheet + "!" + normalizeA1_(plan.targetRange) + "."
  };
}
```

Update the pivot support branch in `apps/google-sheets-addon/html/Sidebar.js.html` so that demo-safe pivot previews stop returning the unsupported string:

```js
if (preview.kind === "pivot_table_plan") {
  if (typeof preview.targetRange !== "string" || preview.targetRange.includes(":")) {
    return "Google Sheets host requires a single-cell pivot target anchor.";
  }

  return "";
}
```

Then swap the current pivot hard-stop in `applyWritePlan()` to:

```js
if (isPivotTablePlan_(plan)) {
  return applyPivotTablePlan_(spreadsheet, plan);
}
```

- [ ] **Step 4: Re-run pivot tests and host syntax checks**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts
bash -lc "cp apps/google-sheets-addon/src/Code.gs /tmp/hermes-codegs-pivot.js && node --check /tmp/hermes-codegs-pivot.js"
bash -lc "tail -n +2 apps/google-sheets-addon/html/Sidebar.js.html | head -n -1 > /tmp/hermes-sidebar-pivot.js && node --check /tmp/hermes-sidebar-pivot.js"
```

Expected:

- the Wave 5 Google Sheets test file passes
- both syntax checks pass with no parse errors

---

### Task 2: Add Exact-Safe Google Sheets Chart Support

**Files:**
- Modify: `apps/google-sheets-addon/src/Code.gs`
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave5Plans.test.ts`

- [ ] **Step 1: Write the failing chart tests**

Add chart coverage beside the pivot tests in `services/gateway/tests/googleSheetsWave5Plans.test.ts`:

```ts
it("treats a demo-safe chart plan as a confirmable Google Sheets write preview", () => {
  const sidebar = loadSidebarContext();
  const response = {
    type: "chart_plan",
    data: {
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
    }
  };

  expect(sidebar.isWritePlanResponse(response)).toBe(true);
  expect(sidebar.getRequiresConfirmation(response)).toBe(true);

  const html = sidebar.renderStructuredPreview(response, {
    runId: "run_chart_preview",
    requestId: "req_chart_preview"
  });

  expect(html).toContain("Will create a line chart on Sales Chart!A1.");
  expect(html).toContain("Confirm Chart");
  expect(html).not.toContain("does not support exact-safe chart creation yet");
});

it("applies a demo-safe chart plan through Code.gs", () => {
  const sourceRange = createRangeStub({
    a1Notation: "A1:C20",
    row: 1,
    column: 1,
    numRows: 20,
    numColumns: 3,
    values: Array.from({ length: 20 }, () => ["Jan", 10, 5])
  });
  const anchorRange = createRangeStub({
    a1Notation: "A1",
    row: 1,
    column: 1,
    numRows: 1,
    numColumns: 1,
    values: [[""]]
  });
  const builtChart = { id: "chart-1" };
  const chartBuilder = {
    addRange: vi.fn().mockReturnThis(),
    setChartType: vi.fn().mockReturnThis(),
    setPosition: vi.fn().mockReturnThis(),
    setOption: vi.fn().mockReturnThis(),
    build: vi.fn(() => builtChart)
  };
  const chartSheet = {
    getRange: vi.fn(() => anchorRange),
    newChart: vi.fn(() => chartBuilder),
    insertChart: vi.fn()
  };
  const spreadsheet = {
    getSheetByName: vi.fn((sheetName: string) => {
      if (sheetName === "Sales") {
        return { getRange: vi.fn(() => sourceRange) };
      }
      if (sheetName === "Sales Chart") {
        return chartSheet;
      }
      return null;
    })
  };

  const { applyWritePlan, flush } = loadCodeModule({ spreadsheet });
  const result = applyWritePlan({
    plan: {
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
    }
  });

  expect(chartSheet.newChart).toHaveBeenCalledTimes(1);
  expect(chartSheet.insertChart).toHaveBeenCalledWith(builtChart);
  expect(result).toEqual({
    kind: "chart_update",
    operation: "chart_update",
    hostPlatform: "google_sheets",
    targetSheet: "Sales Chart",
    targetRange: "A1",
    chartType: "line",
    summary: "Created line chart on Sales Chart!A1."
  });
  expect(flush).toHaveBeenCalledTimes(1);
});
```

- [ ] **Step 2: Run the targeted chart tests and verify they fail for current unsupported behavior**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts
```

Expected:

- FAIL because the sidebar still marks chart plans as unsupported
- FAIL because `applyWritePlan()` still throws `Google Sheets host does not support exact-safe chart creation yet.`

- [ ] **Step 3: Implement the minimal exact-safe chart path in `Code.gs` and sidebar support checks**

Add a chart helper in `apps/google-sheets-addon/src/Code.gs`:

```js
function getGoogleSheetsChartType_(chartType) {
  switch (chartType) {
    case "bar":
      return Charts.ChartType.BAR;
    case "column":
      return Charts.ChartType.COLUMN;
    case "stacked_bar":
      return Charts.ChartType.BAR;
    case "stacked_column":
      return Charts.ChartType.COLUMN;
    case "line":
      return Charts.ChartType.LINE;
    case "area":
      return Charts.ChartType.AREA;
    case "pie":
      return Charts.ChartType.PIE;
    case "scatter":
      return Charts.ChartType.SCATTER;
    default:
      throw new Error("Unsupported chart type: " + chartType);
  }
}

function applyChartPlan_(spreadsheet, plan) {
  const sourceSheet = spreadsheet.getSheetByName(plan.sourceSheet);
  const targetSheet = spreadsheet.getSheetByName(plan.targetSheet);
  if (!sourceSheet) {
    throw new Error("Source sheet not found: " + plan.sourceSheet);
  }
  if (!targetSheet) {
    throw new Error("Target sheet not found: " + plan.targetSheet);
  }

  const sourceRange = sourceSheet.getRange(plan.sourceRange);
  const anchorRange = targetSheet.getRange(plan.targetRange);
  requireSingleCellAnchor_(anchorRange, "charts");

  if (typeof targetSheet.newChart !== "function") {
    throw new Error("Google Sheets host does not expose chart creation on this sheet.");
  }

  const builder = targetSheet.newChart()
    .addRange(sourceRange)
    .setChartType(getGoogleSheetsChartType_(plan.chartType))
    .setPosition(anchorRange.getRow(), anchorRange.getColumn(), 0, 0);

  if (plan.title) {
    builder.setOption("title", plan.title);
  }
  if (plan.legendPosition) {
    builder.setOption(
      "legend.position",
      plan.legendPosition === "hidden" ? "none" : plan.legendPosition
    );
  }
  if (plan.chartType === "stacked_bar" || plan.chartType === "stacked_column") {
    builder.setOption("isStacked", true);
  }

  const chart = builder.build();
  targetSheet.insertChart(chart);
  SpreadsheetApp.flush();

  return {
    kind: "chart_update",
    operation: "chart_update",
    hostPlatform: "google_sheets",
    targetSheet: plan.targetSheet,
    targetRange: normalizeA1_(plan.targetRange),
    chartType: plan.chartType,
    summary: "Created " + plan.chartType + " chart on " + plan.targetSheet + "!" + normalizeA1_(plan.targetRange) + "."
  };
}
```

Update the chart support branch in `apps/google-sheets-addon/html/Sidebar.js.html`:

```js
if (preview.kind === "chart_plan") {
  if (typeof preview.targetRange !== "string" || preview.targetRange.includes(":")) {
    return "Google Sheets host requires a single-cell chart target anchor.";
  }

  return "";
}
```

Then swap the current chart hard-stop in `applyWritePlan()` to:

```js
if (isChartPlan_(plan)) {
  return applyChartPlan_(spreadsheet, plan);
}
```

- [ ] **Step 4: Re-run chart tests and host syntax checks**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts
bash -lc "cp apps/google-sheets-addon/src/Code.gs /tmp/hermes-codegs-chart.js && node --check /tmp/hermes-codegs-chart.js"
bash -lc "tail -n +2 apps/google-sheets-addon/html/Sidebar.js.html | head -n -1 > /tmp/hermes-sidebar-chart.js && node --check /tmp/hermes-sidebar-chart.js"
```

Expected:

- the Wave 5 Google Sheets test file passes with chart support enabled
- both syntax checks pass with no parse errors

---

### Task 3: Align Google Sheets Support Detection And Composite Preview Behavior

**Files:**
- Modify: `apps/google-sheets-addon/html/Sidebar.js.html`
- Test: `services/gateway/tests/googleSheetsWave5Plans.test.ts`
- Test: `services/gateway/tests/googleSheetsWave6Plans.test.ts`

- [ ] **Step 1: Write the failing support-matrix and composite tests**

Add a focused composite expectation in `services/gateway/tests/googleSheetsWave6Plans.test.ts`:

```ts
it("keeps a composite workflow confirmable when its pivot/chart steps are demo-safe on Google Sheets", () => {
  const sidebar = loadSidebarContext();
  const response = {
    type: "composite_plan",
    data: {
      steps: [
        {
          stepId: "step_pivot",
          dependsOn: [],
          continueOnError: false,
          plan: {
            sourceSheet: "Sales",
            sourceRange: "A1:F50",
            targetSheet: "Sales Pivot",
            targetRange: "A1",
            rowGroups: ["Region"],
            valueAggregations: [{ field: "Revenue", aggregation: "sum" }],
            explanation: "Build a pivot table.",
            confidence: 0.91,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50", "Sales Pivot!A1"],
            overwriteRisk: "low",
            confirmationLevel: "standard"
          }
        },
        {
          stepId: "step_chart",
          dependsOn: ["step_pivot"],
          continueOnError: false,
          plan: {
            sourceSheet: "Sales",
            sourceRange: "A1:C20",
            targetSheet: "Sales Chart",
            targetRange: "A1",
            chartType: "column",
            series: [{ field: "Revenue", label: "Revenue" }],
            explanation: "Chart revenue.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
            overwriteRisk: "low",
            confirmationLevel: "standard"
          }
        }
      ],
      explanation: "Create a pivot and then chart it.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales Pivot!A1", "Sales Chart!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: false,
      dryRunRecommended: false,
      dryRunRequired: false
    }
  };

  const html = sidebar.renderStructuredPreview(response, {
    runId: "run_composite_wave5_supported",
    requestId: "req_composite_wave5_supported"
  });

  expect(html).toContain("Confirm Workflow");
  expect(html).not.toContain("does not support exact-safe pivot table creation yet");
  expect(html).not.toContain("does not support exact-safe chart creation yet");
});
```

- [ ] **Step 2: Run the support-matrix tests and verify they fail before the sidebar logic is aligned**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts services/gateway/tests/googleSheetsWave6Plans.test.ts
```

Expected:

- FAIL because the sidebar still reports pivot/chart unsupported for Google Sheets composite previews

- [ ] **Step 3: Tighten `getGoogleSheetsPlanSupportError()` and keep composite previews honest**

Update `apps/google-sheets-addon/html/Sidebar.js.html` so support detection matches the patched host behavior. Insert these new early-return branches near the top of the existing `getGoogleSheetsPlanSupportError()` function and leave the current validation, named-range, conditional-format, transfer, and cleanup branches intact below them:

```js
if (preview.kind === "pivot_table_plan") {
  if (typeof preview.targetRange !== "string" || preview.targetRange.includes(":")) {
    return "Google Sheets host requires a single-cell pivot target anchor.";
  }
  return "";
}

if (preview.kind === "chart_plan") {
  if (typeof preview.targetRange !== "string" || preview.targetRange.includes(":")) {
    return "Google Sheets host requires a single-cell chart target anchor.";
  }
  return "";
}
```

Do not change the Wave 6 reversibility policy in the sidebar. Keep pivot/chart composite steps non-reversible unless an exact inverse is actually implemented.

- [ ] **Step 4: Re-run support-matrix tests**

Run:

```bash
npm test -- services/gateway/tests/googleSheetsWave5Plans.test.ts services/gateway/tests/googleSheetsWave6Plans.test.ts services/gateway/tests/writebackFlow.test.ts
```

Expected:

- all three test files pass
- gateway writeback flow remains green for pivot/chart approval and completion

---

### Task 4: Update The Demo Runbook And Operator Checklist

**Files:**
- Modify: `docs/demo-runbook.md`
- Create: `docs/review/google-sheets-live-demo-checklist.md`

- [ ] **Step 1: Update `docs/demo-runbook.md` for the real Google Sheets rollout**

Add a Google Sheets section like this to `docs/demo-runbook.md`:

```md
## Google Sheets live-demo deploy

1. Open the provided demo workbook.
2. Open the bound Apps Script project from Extensions -> Apps Script.
3. Replace the bound files with the local source-of-truth files from:
   - `apps/google-sheets-addon/appsscript.json`
   - `apps/google-sheets-addon/src/Code.gs`
   - `apps/google-sheets-addon/src/ReferencedCells.js`
   - `apps/google-sheets-addon/src/Wave1Plans.js`
   - `apps/google-sheets-addon/html/Sidebar.html`
   - `apps/google-sheets-addon/html/Sidebar.css.html`
   - `apps/google-sheets-addon/html/Sidebar.js.html`
4. Set script properties:
   - `HERMES_GATEWAY_URL`
   - `HERMES_CLIENT_VERSION=google-sheets-addon-live-demo`
   - optionally `HERMES_REVIEWER_SAFE_MODE=false`
   - optionally `HERMES_FORCE_EXTRACTION_MODE=real`
5. Reload the workbook and verify the `Hermes` menu still appears.
```

- [ ] **Step 2: Create the live checklist document with tabs, prompts, and expected outcomes**

Create `docs/review/google-sheets-live-demo-checklist.md` with concrete entries like:

```md
# Google Sheets Live Demo Checklist

## Demo tabs

- `Demo_ReadChat`
- `Demo_SheetUpdate`
- `Demo_Structure`
- `Demo_SortFilterFormat`
- `Demo_ValidationNamedRanges`
- `Demo_TransferCleanup`
- `Demo_Reports`
- `Demo_Pivot`
- `Demo_Charts`
- `Demo_Composite`
- `Demo_ImageImport`

## Prompt checklist

- Read/chat: `Explain the current selection.`
- Formula: `Suggest a formula that calculates total revenue in G2.`
- Sort/filter: `Sort this table by Revenue descending and then filter Status = Closed Won.`
- Validation: `Add a dropdown on D2:D20 for Open, Blocked, Done.`
- Cleanup: `Trim whitespace in A2:C20.`
- Report: `Create an analysis report for this table on Demo_Reports starting at A1.`
- Pivot: `Create a pivot table from this data on Demo_Pivot starting at A1 grouped by Region with Revenue summed.`
- Chart: `Create a column chart on Demo_Charts at A1 using Month as category and Revenue as the series.`
- Composite: `Create a pivot on Demo_Pivot and then create a chart on Demo_Charts.`
```

- [ ] **Step 3: Verify the docs mention the exact workbook rollout path**

Run:

```bash
rg -n "Google Sheets live-demo deploy|Demo_Pivot|Demo_Charts|HERMES_GATEWAY_URL" docs/demo-runbook.md docs/review/google-sheets-live-demo-checklist.md
```

Expected:

- both docs contain the exact deploy path, target tabs, and live prompts

---

### Task 5: Redeploy The Bound Apps Script Project And Run The Live Workbook Smoke Test

**Files:**
- Use: `apps/google-sheets-addon/appsscript.json`
- Use: `apps/google-sheets-addon/src/Code.gs`
- Use: `apps/google-sheets-addon/src/ReferencedCells.js`
- Use: `apps/google-sheets-addon/src/Wave1Plans.js`
- Use: `apps/google-sheets-addon/html/Sidebar.html`
- Use: `apps/google-sheets-addon/html/Sidebar.css.html`
- Use: `apps/google-sheets-addon/html/Sidebar.js.html`
- Use: `docs/review/google-sheets-live-demo-checklist.md`

- [ ] **Step 1: Start the gateway and verify health**

Run:

```bash
cd <repo-root>
npm run dev:gateway
```

In a second shell:

```bash
curl http://127.0.0.1:8787/health
```

Expected:

- `ok: true`
- the configured `service` and `environment` values are present

- [ ] **Step 2: Redeploy the bound Apps Script files from local source**

Manual action in the bound Apps Script editor:

```text
Replace appsscript.json, Code.gs, ReferencedCells.js, Wave1Plans.js, Sidebar.html, Sidebar.css.html, and Sidebar.js.html with the patched local versions from this repo.
```

Expected:

- the script saves cleanly
- no Apps Script syntax errors are reported

- [ ] **Step 3: Set the live script properties**

Manual action in Apps Script Project Settings -> Script properties:

```text
HERMES_GATEWAY_URL=http://<reachable-gateway-host>:8787
HERMES_CLIENT_VERSION=google-sheets-addon-live-demo
HERMES_REVIEWER_SAFE_MODE=false
HERMES_FORCE_EXTRACTION_MODE=real
```

Expected:

- the sidebar points at the gateway reachable on the shared VPN/network

- [ ] **Step 4: Seed the disposable demo tabs in the workbook**

Create these tabs and seed obvious sample data:

```text
Demo_ReadChat
Demo_SheetUpdate
Demo_Structure
Demo_SortFilterFormat
Demo_ValidationNamedRanges
Demo_TransferCleanup
Demo_Reports
Demo_Pivot
Demo_Charts
Demo_Composite
Demo_ImageImport
```

Expected:

- every destructive demo operates only on disposable demo tabs

- [ ] **Step 5: Run the live smoke matrix from the checklist**

Execute the prompts in `docs/review/google-sheets-live-demo-checklist.md` and confirm:

```text
Read/chat -> proof line visible, no mutation
Write flows -> preview first, explicit confirm required, success status after apply
Pivot -> created on Demo_Pivot
Chart -> created on Demo_Charts
Composite -> preview + confirm + child steps complete in order
```

Expected:

- the live workbook matches the checklist outcomes
- any flaky flow is fixed or removed from the final live-demo script before signoff
