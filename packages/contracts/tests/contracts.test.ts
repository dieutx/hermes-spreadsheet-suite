import { describe, expect, it } from "vitest";
import {
  AttachmentSchema,
  AnalysisReportPlanDataSchema,
  ChartPlanDataSchema,
  CompositePlanDataSchema,
  DryRunResultSchema,
  DataValidationPlanDataSchema,
  ConditionalFormatPlanDataSchema,
  DataCleanupPlanDataSchema,
  DataCleanupPlanResponseSchema,
  DataCleanupUpdateDataSchema,
  DataCleanupUpdateResponseSchema,
  HermesRequestSchema,
  HermesResponseSchema,
  HermesTraceEventSchema,
  PlanHistoryEntrySchema,
  PlanHistoryPageSchema,
  NamedRangeUpdateDataSchema,
  RangeFilterConditionSchema,
  RangeTransferPlanDataSchema,
  RangeTransferPlanResponseSchema,
  RangeFilterPlanDataSchema,
  RangeSortPlanDataSchema,
  RedoRequestSchema,
  PivotTablePlanDataSchema,
  SheetImportPlanDataSchema,
  SheetStructureUpdateDataSchema,
  SheetUpdateDataSchema,
  UndoRequestSchema
} from "../src/index.ts";

describe("Hermes spreadsheet contracts", () => {
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
            affectedRanges: ["Sales!A1:F50"]
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
            conditions: [{ columnRef: "Status", operator: "equals", value: "Open" }],
            combiner: "and",
            clearExistingFilters: true,
            explanation: "Filter open rows.",
            confidence: 0.88,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F50"]
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

  it("rejects composite steps with duplicate ids", () => {
    const parsed = CompositePlanDataSchema.safeParse({
      steps: [
        {
          stepId: "step_one",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:A10",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "asc" }],
            explanation: "Sort rows.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:A10"]
          }
        },
        {
          stepId: "step_one",
          dependsOn: [],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:A10",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "desc" }],
            explanation: "Sort rows again.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:A10"]
          }
        }
      ],
      explanation: "Duplicate ids are invalid.",
      confidence: 0.8,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:A10"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: true,
      dryRunRecommended: false,
      dryRunRequired: false
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects composite dependencies that point at missing steps", () => {
    const parsed = CompositePlanDataSchema.safeParse({
      steps: [
        {
          stepId: "step_filter",
          dependsOn: ["missing_step"],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:F10",
            hasHeader: true,
            conditions: [{ columnRef: "Status", operator: "equal_to", value: "Open" }],
            combiner: "and",
            clearExistingFilters: true,
            explanation: "Filter rows.",
            confidence: 0.86,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:F10"]
          }
        }
      ],
      explanation: "Bad dependency.",
      confidence: 0.8,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F10"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: true,
      dryRunRecommended: false,
      dryRunRequired: false
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects composite dependency cycles", () => {
    const parsed = CompositePlanDataSchema.safeParse({
      steps: [
        {
          stepId: "step_a",
          dependsOn: ["step_b"],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:A10",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "asc" }],
            explanation: "Sort rows.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:A10"]
          }
        },
        {
          stepId: "step_b",
          dependsOn: ["step_a"],
          continueOnError: false,
          plan: {
            targetSheet: "Sales",
            targetRange: "A1:A10",
            hasHeader: true,
            keys: [{ columnRef: "Revenue", direction: "desc" }],
            explanation: "Sort rows again.",
            confidence: 0.9,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:A10"]
          }
        }
      ],
      explanation: "Cyclic dependency.",
      confidence: 0.8,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:A10"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: true,
      dryRunRecommended: false,
      dryRunRequired: false
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects composite plans that fail to escalate destructive confirmation", () => {
    const parsed = CompositePlanDataSchema.safeParse({
      steps: [
        {
          stepId: "step_cleanup",
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
      explanation: "Cleanup workflow.",
      confidence: 0.9,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:F50"],
      overwriteRisk: "medium",
      confirmationLevel: "standard",
      reversible: false,
      dryRunRecommended: true,
      dryRunRequired: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects composite plans that mark known non-reversible steps as reversible", () => {
    const parsed = CompositePlanDataSchema.safeParse({
      steps: [
        {
          stepId: "step_chart",
          dependsOn: [],
          continueOnError: false,
          plan: {
            sourceSheet: "Sales",
            sourceRange: "A1:C20",
            targetSheet: "Sales Chart",
            targetRange: "A1",
            chartType: "line",
            categoryField: "Month",
            series: [
              { field: "Revenue", label: "Revenue" }
            ],
            explanation: "Create a chart.",
            confidence: 0.92,
            requiresConfirmation: true,
            affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
            overwriteRisk: "low",
            confirmationLevel: "standard"
          }
        }
      ],
      explanation: "Chart workflow.",
      confidence: 0.92,
      requiresConfirmation: true,
      affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard",
      reversible: true,
      dryRunRecommended: true,
      dryRunRequired: false
    });

    expect(parsed.success).toBe(false);
  });

  it("accepts undo and redo request envelopes", () => {
    expect(
      UndoRequestSchema.parse({
        executionId: "exec_001",
        requestId: "req_undo_001",
        workbookSessionKey: "excel_windows::workbook-123"
      }).executionId
    ).toBe("exec_001");

    expect(
      RedoRequestSchema.parse({
        executionId: "exec_undo_001",
        requestId: "req_redo_001",
        workbookSessionKey: "excel_windows::workbook-123"
      }).executionId
    ).toBe("exec_undo_001");
  });

  it("accepts dry-run results and history pages", () => {
    expect(
      DryRunResultSchema.parse({
        planDigest: "digest_001",
        workbookSessionKey: "excel_windows::workbook-123",
        simulated: true,
        steps: [
          {
            stepId: "step_sort",
            status: "simulated",
            summary: "Will sort rows."
          }
        ],
        predictedAffectedRanges: ["Sales!A1:F50"],
        predictedSummaries: ["Sort Sales rows, then filter open rows."],
        overwriteRisk: "low",
        reversible: true,
        expiresAt: "2026-04-20T10:00:00.000Z"
      }).planDigest
    ).toBe("digest_001");

    expect(
      PlanHistoryEntrySchema.parse({
        executionId: "exec_001",
        requestId: "req_001",
        runId: "run_001",
        planType: "composite_plan",
        planDigest: "digest_001",
        status: "completed",
        timestamp: "2026-04-20T09:00:00.000Z",
        reversible: true,
        undoEligible: true,
        redoEligible: false,
        summary: "Completed the composite plan.",
        stepEntries: [
          {
            stepId: "step_sort",
            planType: "range_sort_plan",
            status: "completed",
            summary: "Sorted Sales rows."
          }
        ]
      }).status
    ).toBe("completed");

    expect(
      PlanHistoryPageSchema.parse({
        entries: [
          {
            executionId: "exec_001",
            requestId: "req_001",
            runId: "run_001",
            planType: "composite_plan",
            planDigest: "digest_001",
            status: "completed",
            timestamp: "2026-04-20T09:00:00.000Z",
            reversible: true,
            undoEligible: true,
            redoEligible: false,
            summary: "Completed the composite plan."
          }
        ],
        nextCursor: "2"
      }).entries
    ).toHaveLength(1);
  });

  it("accepts composite Hermes response envelopes", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "composite_plan",
      requestId: "req_comp_001",
      hermesRunId: "run_comp_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      startedAt: "2026-04-20T12:00:00.000Z",
      completedAt: "2026-04-20T12:00:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "completed",
          timestamp: "2026-04-20T12:00:01.000Z"
        }
      ],
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
            affectedRanges: ["Sales!A1:F50"]
          }
        }
      ],
        explanation: "Sort the current table.",
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
    expect(parsed.data.steps).toHaveLength(1);
  });

  it("accepts composite update Hermes response envelopes", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "composite_update",
      requestId: "req_comp_update_001",
      hermesRunId: "run_comp_update_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      startedAt: "2026-04-20T12:10:00.000Z",
      completedAt: "2026-04-20T12:10:01.000Z",
      durationMs: 1000,
      trace: [
        {
          event: "completed",
          timestamp: "2026-04-20T12:10:01.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "composite_update",
        executionId: "exec_001",
        stepResults: [
          {
            stepId: "step_sort",
            status: "completed",
            summary: "Sorted Sales rows."
          }
        ],
        summary: "Completed the composite update."
      }
    });

    expect(parsed.type).toBe("composite_update");
    expect(parsed.data.stepResults).toHaveLength(1);
  });

  it("accepts the Step 2 backend request envelope exactly", () => {
    const parsed = HermesRequestSchema.parse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0",
        sessionId: "sess_123"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2",
        selectedRange: "A1:B2",
        locale: "en-US",
        timeZone: "America/Los_Angeles"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          range: "A1:B2",
          headers: ["Region", "Revenue"],
          values: [
            ["APAC", 1200],
            ["EMEA", 1800]
          ],
          formulas: [
            [null, null],
            [null, "=SUM(B2:B2)"]
          ]
        },
        currentRegion: {
          range: "A1:B8",
          headers: ["Region", "Revenue"]
        },
        currentRegionArtifactTarget: "A10",
        currentRegionAppendTarget: "A9:B9",
        attachments: [
          {
            id: "att_001",
            type: "image",
            mimeType: "image/png",
            source: "clipboard",
            fileName: "capture.png",
            size: 2048,
            uploadToken: "upl_123"
          }
        ]
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true,
        supportsStructureEdits: true,
        supportsAutofit: true,
        supportsSortFilter: true,
        supportsImageInputs: true,
        supportsWriteBackExecution: true
      },
      reviewer: {
        reviewerSafeMode: false,
        forceExtractionMode: null
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.source.channel).toBe("google_sheets");
    expect(parsed.host.platform).toBe("google_sheets");
    expect(parsed.context.selection?.range).toBe("A1:B2");
    expect(parsed.context.currentRegion?.range).toBe("A1:B8");
    expect(parsed.context.currentRegionArtifactTarget).toBe("A10");
    expect(parsed.context.currentRegionAppendTarget).toBe("A9:B9");
    expect(parsed.context.attachments?.[0]?.source).toBe("clipboard");
    expect(parsed.capabilities.supportsStructureEdits).toBe(true);
    expect(parsed.capabilities.supportsAutofit).toBe(true);
    expect(parsed.capabilities.supportsSortFilter).toBe(true);
  });

  it("rejects malformed selection ranges in requests", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_bad_a1_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2"
      },
      userMessage: "Explain the current selection.",
      conversation: [{ role: "user", content: "Explain the current selection." }],
      context: {
        selection: {
          range: "foo"
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects malformed host selectedRange values in requests", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_host_bad_a1_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2",
        selectedRange: "foo"
      },
      userMessage: "Explain the current selection.",
      conversation: [{ role: "user", content: "Explain the current selection." }],
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects ragged selection values in requests", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_ragged_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          range: "A1:B3",
          headers: ["Region", "Revenue"],
          values: [
            ["APAC", 1200],
            ["EMEA"]
          ]
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects selection formulas that do not match header width", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_formula_width_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          range: "A1:B3",
          headers: ["Region", "Revenue"],
          values: [
            ["APAC", 1200],
            ["EMEA", 1800]
          ],
          formulas: [
            [null, null, null],
            [null, "=SUM(B2:B2)", null]
          ]
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects selection headers that do not match the selection range width", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_headers_width_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          range: "A1:B2",
          headers: ["Region", "Revenue", "Extra"]
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects selection values that do not match the selection range", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_range_values_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2",
        selectedRange: "A1:B3"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          range: "A1:B3",
          values: [
            ["APAC", 1200, "extra"],
            ["EMEA", 1800, "extra"]
          ]
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects selection formulas that do not match the selection range", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_range_formulas_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2",
        selectedRange: "A1:B3"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          range: "A1:B3",
          formulas: [
            [null, null, null],
            [null, "=SUM(B2:B2)", null]
          ]
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects selection values and formulas with mismatched shapes", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_shape_mismatch_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          values: [
            ["APAC", 1200],
            ["EMEA", 1800]
          ],
          formulas: [
            [null, null, null],
            [null, "=SUM(B2:B2)", null]
          ]
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects mismatched host and selection ranges in requests", () => {
    const parsed = HermesRequestSchema.safeParse({
      schemaVersion: "1.0.0",
      requestId: "req_sel_host_range_001",
      source: {
        channel: "google_sheets",
        clientVersion: "0.1.0"
      },
      host: {
        platform: "google_sheets",
        workbookTitle: "Revenue Tracker",
        activeSheet: "Q2",
        selectedRange: "A1:C3"
      },
      userMessage: "Explain the current selection.",
      conversation: [
        { role: "user", content: "Explain the current selection." }
      ],
      context: {
        selection: {
          range: "A1:B3",
          values: [
            ["APAC", 1200],
            ["EMEA", 1800]
          ]
        }
      },
      capabilities: {
        canRenderTrace: true,
        canRenderStructuredPreview: true,
        canConfirmWriteBack: true
      },
      reviewer: {
        reviewerSafeMode: false
      },
      confirmation: {
        state: "none"
      }
    });

    expect(parsed.success).toBe(false);
  });

  it("accepts public trace events only in the Step 1 shape", () => {
    const parsed = HermesTraceEventSchema.parse({
      event: "ocr_started",
      timestamp: "2026-04-19T09:00:00.000Z",
      label: "OCR started",
      skillName: "ocr-and-documents",
      details: {
        attachmentId: "att_001",
        mode: "real"
      }
    });

    expect(parsed.event).toBe("ocr_started");
    expect(parsed.details?.mode).toBe("real");
  });

  it("accepts wave-4 plan-ready trace events in the Step 1 shape", () => {
    const parsed = [
      HermesTraceEventSchema.parse({
        event: "range_transfer_plan_ready",
        timestamp: "2026-04-20T09:00:00.000Z"
      }),
      HermesTraceEventSchema.parse({
        event: "data_cleanup_plan_ready",
        timestamp: "2026-04-20T09:00:01.000Z"
      })
    ];

    expect(parsed.map((event) => event.event)).toEqual([
      "range_transfer_plan_ready",
      "data_cleanup_plan_ready"
    ]);
  });

  it("enforces full targetRange and matrix shape semantics for sheet updates", () => {
    const parsed = SheetUpdateDataSchema.parse({
      targetSheet: "Sheet3",
      targetRange: "B4:C7",
      operation: "replace_range",
      values: [
        ["North", 1200],
        ["South", 1400],
        ["East", 1600],
        ["West", 1800]
      ],
      explanation: "Proposed normalized values.",
      confidence: 0.93,
      requiresConfirmation: true,
      overwriteRisk: "medium",
      shape: {
        rows: 4,
        columns: 2
      }
    });

    expect(parsed.shape.rows).toBe(4);
    expect(parsed.shape.columns).toBe(2);
    expect(parsed.targetRange).toBe("B4:C7");
  });

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

  it("accepts a non-delete sheet structure merge-cells plan", () => {
    const parsed = SheetStructureUpdateDataSchema.parse({
      targetSheet: "Sheet1",
      operation: "merge_cells",
      targetRange: "A1:C1",
      explanation: "Merge the title row across the table.",
      confidence: 0.89,
      requiresConfirmation: true,
      confirmationLevel: "standard"
    });

    expect(parsed.operation).toBe("merge_cells");
    expect(parsed.targetRange).toBe("A1:C1");
  });

  it("rejects numeric formula cells in set_formulas updates", () => {
    const parsed = SheetUpdateDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "A1:A1",
      operation: "set_formulas",
      formulas: [[1]],
      explanation: "Set the formula.",
      confidence: 0.91,
      requiresConfirmation: true,
      shape: { rows: 1, columns: 1 }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects boolean note cells in set_notes updates", () => {
    const parsed = SheetUpdateDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "A1:A1",
      operation: "set_notes",
      notes: [[true]],
      explanation: "Set the note.",
      confidence: 0.91,
      requiresConfirmation: true,
      shape: { rows: 1, columns: 1 }
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects wrong operation payload combinations for set_formulas and set_notes", () => {
    const formulasOnly = SheetUpdateDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "A1:A1",
      operation: "set_formulas",
      values: [["=A1+1"]],
      explanation: "Set the formula.",
      confidence: 0.91,
      requiresConfirmation: true,
      shape: { rows: 1, columns: 1 }
    });
    const notesOnly = SheetUpdateDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "A1:A1",
      operation: "set_notes",
      formulas: [["Note text"]],
      explanation: "Set the note.",
      confidence: 0.91,
      requiresConfirmation: true,
      shape: { rows: 1, columns: 1 }
    });

    expect(formulasOnly.success).toBe(false);
    expect(notesOnly.success).toBe(false);
  });

  it("rejects empty mixed_update payloads", () => {
    const parsed = SheetUpdateDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "A1:A1",
      operation: "mixed_update",
      explanation: "Update multiple cell types.",
      confidence: 0.91,
      requiresConfirmation: true,
      shape: { rows: 1, columns: 1 }
    });

    expect(parsed.success).toBe(false);
  });

  it("accepts mixed_update when at least one matrix is present", () => {
    const parsed = SheetUpdateDataSchema.parse({
      targetSheet: "Sheet1",
      targetRange: "A1:A1",
      operation: "mixed_update",
      values: [[1]],
      explanation: "Update multiple cell types.",
      confidence: 0.91,
      requiresConfirmation: true,
      shape: { rows: 1, columns: 1 }
    });

    expect(parsed.operation).toBe("mixed_update");
    expect(parsed.values).toHaveLength(1);
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
        {
          event: "range_sort_plan_ready",
          timestamp: "2026-04-19T12:00:01.000Z"
        }
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

  it("rejects a materialized analysis report without a target range", () => {
    const parsed = AnalysisReportPlanDataSchema.safeParse({
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report",
      targetSheet: "Sales Report",
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
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    });

    expect(parsed.success).toBe(false);
  });

  it("accepts a materialized analysis report with a target sheet and range", () => {
    const parsed = AnalysisReportPlanDataSchema.parse({
      sourceSheet: "Sales",
      sourceRange: "A1:F50",
      outputMode: "materialize_report",
      targetSheet: "Sales Report",
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
      affectedRanges: ["Sales!A1:F50", "Sales Report!A1"],
      overwriteRisk: "low",
      confirmationLevel: "standard"
    });

    expect(parsed.outputMode).toBe("materialize_report");
    expect(parsed.targetSheet).toBe("Sales Report");
    expect(parsed.targetRange).toBe("A1");
    expect(parsed.requiresConfirmation).toBe(true);
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
      processedBy: "hermes",
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

  it("rejects filter conditions that violate operator-specific rules", () => {
    const emptyWithValue = RangeFilterConditionSchema.safeParse({
      columnRef: "Status",
      operator: "isEmpty",
      value: "x"
    });
    const topNWithoutValue = RangeFilterConditionSchema.safeParse({
      columnRef: "Status",
      operator: "topN"
    });
    const containsWithoutValue = RangeFilterConditionSchema.safeParse({
      columnRef: "Status",
      operator: "contains"
    });
    const topNWithValue2 = RangeFilterConditionSchema.safeParse({
      columnRef: "Status",
      operator: "topN",
      value: 5,
      value2: 10
    });
    const containsWithValue2 = RangeFilterConditionSchema.safeParse({
      columnRef: "Status",
      operator: "contains",
      value: "open",
      value2: "closed"
    });
    const containsWithBoolean = RangeFilterConditionSchema.safeParse({
      columnRef: "Status",
      operator: "contains",
      value: true
    });
    const greaterThanWithNull = RangeFilterConditionSchema.safeParse({
      columnRef: "Revenue",
      operator: "greaterThan",
      value: null
    });

    expect(emptyWithValue.success).toBe(false);
    expect(topNWithoutValue.success).toBe(false);
    expect(containsWithoutValue.success).toBe(false);
    expect(topNWithValue2.success).toBe(false);
    expect(containsWithValue2.success).toBe(false);
    expect(containsWithBoolean.success).toBe(false);
    expect(greaterThanWithNull.success).toBe(false);
  });

  it("accepts a valid topN filter condition", () => {
    const parsed = RangeFilterConditionSchema.parse({
      columnRef: "Revenue",
      operator: "topN",
      value: 5
    });

    expect(parsed.operator).toBe("topN");
    expect(parsed.value).toBe(5);
  });

  it("enforces header, values, shape, and targetRange semantics for sheet import plans", () => {
    const parsed = SheetImportPlanDataSchema.parse({
      sourceAttachmentId: "att_002",
      targetSheet: "Sheet3",
      targetRange: "B4:C6",
      headers: ["Name", "Qty"],
      values: [
        ["Widget", 10],
        ["Cable", 12]
      ],
      confidence: 0.88,
      warnings: [
        {
          code: "DEMO_MODE",
          message: "Demo-only extraction preview.",
          severity: "warning"
        }
      ],
      requiresConfirmation: true,
      extractionMode: "demo",
      shape: {
        rows: 3,
        columns: 2
      }
    });

    expect(parsed.headers).toEqual(["Name", "Qty"]);
    expect(parsed.values).toHaveLength(2);
    expect(parsed.shape.rows).toBe(3);
    expect(parsed.shape.columns).toBe(2);
    expect(parsed.targetRange).toBe("B4:C6");
  });

  it("accepts the exact Step 1/Step 2 response envelope", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "sheet_import_plan",
      requestId: "req_img_001",
      hermesRunId: "run_img_demo_001",
      processedBy: "hermes",
      serviceLabel: "spreadsheet-gateway",
      environmentLabel: "demo-review",
      startedAt: "2026-04-19T09:00:00.000Z",
      completedAt: "2026-04-19T09:00:02.000Z",
      durationMs: 2000,
      skillsUsed: [],
      downstreamProvider: null,
      warnings: [
        {
          code: "DEMO_MODE",
          message: "This import plan is demo-only and not derived from real extraction.",
          severity: "warning"
        }
      ],
      trace: [
        {
          event: "request_received",
          timestamp: "2026-04-19T09:00:00.000Z"
        },
        {
          event: "attachment_received",
          timestamp: "2026-04-19T09:00:00.500Z",
          details: {
            attachmentId: "att_101"
          }
        },
        {
          event: "table_extraction_started",
          timestamp: "2026-04-19T09:00:01.000Z",
          details: {
            mode: "demo"
          }
        },
        {
          event: "sheet_import_plan_ready",
          timestamp: "2026-04-19T09:00:01.500Z"
        },
        {
          event: "completed",
          timestamp: "2026-04-19T09:00:02.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceAttachmentId: "att_101",
        targetSheet: "Imported Table",
        targetRange: "A1:B3",
        headers: ["Item", "Qty"],
        values: [
          ["Cable", 2],
          ["Adapter", 3]
        ],
        confidence: 0.77,
        warnings: [
          {
            code: "DEMO_MODE",
            message: "This import plan is demo-only and not derived from real extraction.",
            severity: "warning"
          }
        ],
        requiresConfirmation: true,
        extractionMode: "demo",
        shape: {
          rows: 3,
          columns: 2
        }
      }
    });

    expect(parsed.requestId).toBe("req_img_001");
    expect(parsed.hermesRunId).toBe("run_img_demo_001");
    expect(parsed.serviceLabel).toBe("spreadsheet-gateway");
    expect(parsed.environmentLabel).toBe("demo-review");
    expect(parsed.data.extractionMode).toBe("demo");
  });

  it("accepts workbook structure update responses for confirmed sheet creation", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "workbook_structure_update",
      requestId: "req_sheet_create_001",
      hermesRunId: "run_sheet_create_001",
      processedBy: "hermes",
      serviceLabel: "spreadsheet-gateway",
      environmentLabel: "demo-review",
      startedAt: "2026-04-19T09:00:00.000Z",
      completedAt: "2026-04-19T09:00:02.000Z",
      durationMs: 2000,
      skillsUsed: [],
      downstreamProvider: null,
      warnings: [],
      trace: [
        {
          event: "request_received",
          timestamp: "2026-04-19T09:00:00.000Z"
        },
        {
          event: "result_generated",
          timestamp: "2026-04-19T09:00:02.000Z"
        },
        {
          event: "completed",
          timestamp: "2026-04-19T09:00:02.000Z"
        }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "create_sheet",
        sheetName: "New Sheet",
        position: "end",
        explanation: "Create a new sheet at the end of the workbook.",
        confidence: 0.94,
        requiresConfirmation: true,
        overwriteRisk: "none"
      }
    } as any);

    expect(parsed.type).toBe("workbook_structure_update");
    expect((parsed.data as any).operation).toBe("create_sheet");
    expect((parsed.data as any).sheetName).toBe("New Sheet");
  });

  it("accepts range format update responses for confirmed formatting changes", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "range_format_update",
      requestId: "req_format_001",
      hermesRunId: "run_format_001",
      processedBy: "hermes",
      serviceLabel: "spreadsheet-gateway",
      environmentLabel: "demo-review",
      startedAt: "2026-04-19T09:00:00.000Z",
      completedAt: "2026-04-19T09:00:02.000Z",
      durationMs: 2000,
      skillsUsed: [],
      downstreamProvider: null,
      warnings: [],
      trace: [
        {
          event: "request_received",
          timestamp: "2026-04-19T09:00:00.000Z"
        },
        {
          event: "result_generated",
          timestamp: "2026-04-19T09:00:02.000Z"
        },
        {
          event: "completed",
          timestamp: "2026-04-19T09:00:02.000Z"
        }
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
        targetRange: "A1:J10",
        format: {
          backgroundColor: "#fff2cc",
          textColor: "#1f1f1f",
          bold: true,
          horizontalAlignment: "center",
          verticalAlignment: "middle",
          wrapStrategy: "wrap",
          numberFormat: "0.00",
          columnWidth: 96,
          rowHeight: 24
        },
        explanation: "Apply the requested square-table formatting.",
        confidence: 0.9,
        requiresConfirmation: true,
        overwriteRisk: "low"
      }
    } as any);

    expect(parsed.type).toBe("range_format_update");
    expect((parsed.data as any).targetRange).toBe("A1:J10");
    expect((parsed.data as any).format.bold).toBe(true);
  });

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

  it("rejects non-finite numeric conditional format comparator values", () => {
    const parsed = ConditionalFormatPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "add",
      ruleType: "number_compare",
      comparator: "greater_than",
      value: Number.POSITIVE_INFINITY,
      style: {
        backgroundColor: "#ffdddd"
      },
      explanation: "Invalid non-finite comparator value.",
      confidence: 0.92,
      requiresConfirmation: true,
      affectedRanges: ["Sheet1!B2:B20"],
      replacesExistingRules: false
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects clear_on_target plans that carry rule payload", () => {
    const parsed = ConditionalFormatPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      managementMode: "clear_on_target",
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

  it("accepts a conditional_format_plan response envelope", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "conditional_format_plan",
      requestId: "req_cf_plan_001",
      hermesRunId: "run_cf_plan_001",
      processedBy: "hermes",
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
        showRequiresConfirmation: true
      },
      data: {
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
      }
    });

    expect(parsed.type).toBe("conditional_format_plan");
  });

  it("accepts a conditional_format_update response envelope", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "conditional_format_update",
      requestId: "req_cf_update_001",
      hermesRunId: "run_cf_update_001",
      processedBy: "hermes",
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

  it("accepts the other wave 2 validation rule families", () => {
    const cases = [
      {
        targetSheet: "Sheet1",
        targetRange: "C2:C20",
        ruleType: "checkbox",
        checkedValue: "Done",
        uncheckedValue: "Pending",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Use a checkbox for task completion.",
        confidence: 0.82,
        requiresConfirmation: true
      },
      {
        targetSheet: "Sheet1",
        targetRange: "D2:D20",
        ruleType: "decimal",
        comparator: "greater_than",
        value: 0,
        allowBlank: false,
        invalidDataBehavior: "warn",
        explanation: "Only allow positive decimal values.",
        confidence: 0.83,
        requiresConfirmation: true
      },
      {
        targetSheet: "Sheet1",
        targetRange: "E2:E20",
        ruleType: "date",
        comparator: "between",
        value: "2026-04-01",
        value2: "2026-04-30",
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Only allow dates in April 2026.",
        confidence: 0.84,
        requiresConfirmation: true
      },
      {
        targetSheet: "Sheet1",
        targetRange: "F2:F20",
        ruleType: "text_length",
        comparator: "less_than_or_equal_to",
        value: 12,
        allowBlank: true,
        invalidDataBehavior: "warn",
        explanation: "Keep input under 12 characters.",
        confidence: 0.85,
        requiresConfirmation: true
      },
      {
        targetSheet: "Sheet1",
        targetRange: "G2:G20",
        ruleType: "custom_formula",
        formula: '=COUNTIF($B$2:$B$20,G2)<=1',
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Reject duplicate values in the input range.",
        confidence: 0.86,
        requiresConfirmation: true
      }
    ] as const;

    for (const candidate of cases) {
      const parsed = DataValidationPlanDataSchema.safeParse(candidate);
      expect(parsed.success).toBe(true);
    }
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

  it("rejects range comparators without value2 in data validation plans", () => {
    const betweenPlan = DataValidationPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      ruleType: "whole_number",
      comparator: "between",
      value: 1,
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Missing upper bound.",
      confidence: 0.73,
      requiresConfirmation: true
    });
    const notBetweenPlan = DataValidationPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      ruleType: "decimal",
      comparator: "not_between",
      value: 1.5,
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Missing second bound.",
      confidence: 0.74,
      requiresConfirmation: true
    });

    expect(betweenPlan.success).toBe(false);
    expect(notBetweenPlan.success).toBe(false);
  });

  it("rejects non-range comparators that include value2 in data validation plans", () => {
    const parsed = DataValidationPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "B2:B20",
      ruleType: "text_length",
      comparator: "less_than_or_equal_to",
      value: 12,
      value2: 15,
      allowBlank: true,
      invalidDataBehavior: "warn",
      explanation: "Upper-bound comparators should not accept value2.",
      confidence: 0.75,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects data validation plans with invalid targetRange", () => {
    const parsed = DataValidationPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "not-a-range",
      ruleType: "list",
      values: ["Open", "Closed"],
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Invalid target range string.",
      confidence: 0.61,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects date validation plans with invalid date literals", () => {
    const parsed = DataValidationPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "E2:E20",
      ruleType: "date",
      comparator: "equal_to",
      value: "2026-02-30",
      allowBlank: false,
      invalidDataBehavior: "reject",
      explanation: "Impossible calendar date.",
      confidence: 0.76,
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

  it("rejects sheet-scoped named ranges without sheetName", () => {
    const parsed = NamedRangeUpdateDataSchema.safeParse({
      operation: "delete",
      name: "InputRange",
      scope: "sheet",
      explanation: "Sheet-scoped names require sheet metadata.",
      confidence: 0.72,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects workbook-scoped named ranges that include sheetName", () => {
    const parsed = NamedRangeUpdateDataSchema.safeParse({
      operation: "delete",
      name: "InputRange",
      scope: "workbook",
      sheetName: "Sheet1",
      explanation: "Workbook names should not carry sheet scope metadata.",
      confidence: 0.7,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects create named range updates without targetSheet and targetRange", () => {
    const parsed = NamedRangeUpdateDataSchema.safeParse({
      operation: "create",
      name: "InputRange",
      scope: "workbook",
      explanation: "Target location is required.",
      confidence: 0.77,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects rename named range updates without newName", () => {
    const parsed = NamedRangeUpdateDataSchema.safeParse({
      operation: "rename",
      name: "InputRange",
      scope: "sheet",
      sheetName: "Sheet1",
      explanation: "Missing replacement name.",
      confidence: 0.71,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects named range updates with invalid targetRange", () => {
    const parsed = NamedRangeUpdateDataSchema.safeParse({
      operation: "retarget",
      name: "InputRange",
      scope: "sheet",
      sheetName: "Sheet1",
      targetSheet: "Sheet1",
      targetRange: "column-b",
      explanation: "Invalid target range string.",
      confidence: 0.78,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects named range updates with operation-irrelevant fields", () => {
    const parsed = NamedRangeUpdateDataSchema.safeParse({
      operation: "rename",
      name: "InputRange",
      scope: "sheet",
      sheetName: "Sheet1",
      newName: "RenamedInputRange",
      targetSheet: "Sheet1",
      targetRange: "B2:D20",
      explanation: "Rename should not silently accept retarget fields.",
      confidence: 0.79,
      requiresConfirmation: true
    });

    expect(parsed.success).toBe(false);
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

  it("accepts a named_range_update response envelope", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "named_range_update",
      requestId: "req_named_range_001",
      hermesRunId: "run_named_range_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      startedAt: "2026-04-20T09:00:00.000Z",
      completedAt: "2026-04-20T09:00:01.000Z",
      durationMs: 1000,
      trace: [
        { event: "named_range_update_ready", timestamp: "2026-04-20T09:00:01.000Z" }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "create",
        name: "StatusOptions",
        scope: "workbook",
        targetSheet: "Lookup",
        targetRange: "A1:A5",
        explanation: "Create a workbook-scoped named range for status values.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low"
      }
    });

    expect(parsed.type).toBe("named_range_update");
  });

  it("accepts a destructive move transfer plan", () => {
    const parsed = RangeTransferPlanDataSchema.parse({
      sourceSheet: "Sheet1",
      sourceRange: "B2:C10",
      targetSheet: "Archive",
      targetRange: "A1:B9",
      operation: "move",
      pasteMode: "values",
      transpose: false,
      explanation: "Move the staged rows into the archive sheet.",
      confidence: 0.91,
      requiresConfirmation: true,
      confirmationLevel: "destructive",
      affectedRanges: ["Sheet1!B2:C10", "Archive!A1:B9"],
      overwriteRisk: "high"
    });

    expect(parsed.operation).toBe("move");
    expect(parsed.confirmationLevel).toBe("destructive");
  });

  it("rejects a move transfer plan with standard confirmation", () => {
    const parsed = RangeTransferPlanDataSchema.safeParse({
      sourceSheet: "Sheet1",
      sourceRange: "B2:C10",
      targetSheet: "Archive",
      targetRange: "A1:B9",
      operation: "move",
      pasteMode: "values",
      transpose: false,
      explanation: "Move the staged rows into the archive sheet.",
      confidence: 0.91,
      requiresConfirmation: true,
      confirmationLevel: "standard",
      affectedRanges: ["Sheet1!B2:C10", "Archive!A1:B9"],
      overwriteRisk: "high"
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects a copy transfer plan with destructive confirmation", () => {
    const parsed = RangeTransferPlanDataSchema.safeParse({
      sourceSheet: "Sheet1",
      sourceRange: "B2:C10",
      targetSheet: "Archive",
      targetRange: "A1:B9",
      operation: "copy",
      pasteMode: "values",
      transpose: false,
      explanation: "Copy the staged rows into the archive sheet.",
      confidence: 0.91,
      requiresConfirmation: true,
      confirmationLevel: "destructive",
      affectedRanges: ["Sheet1!B2:C10", "Archive!A1:B9"],
      overwriteRisk: "low"
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects an append transfer plan without a target range", () => {
    const parsed = RangeTransferPlanDataSchema.safeParse({
      sourceSheet: "Sheet1",
      sourceRange: "B2:C10",
      targetSheet: "Archive",
      operation: "append",
      pasteMode: "values",
      transpose: false,
      explanation: "Append the staged rows to the archive sheet.",
      confidence: 0.91,
      requiresConfirmation: true,
      confirmationLevel: "standard"
    });

    expect(parsed.success).toBe(false);
  });

  it("accepts a destructive remove_duplicate_rows cleanup plan", () => {
    const parsed = DataCleanupPlanDataSchema.parse({
      targetSheet: "Sheet1",
      targetRange: "A1:D50",
      operation: "remove_duplicate_rows",
      explanation: "Remove duplicate records before publishing.",
      confidence: 0.87,
      requiresConfirmation: true,
      confirmationLevel: "destructive",
      affectedRanges: ["Sheet1!A1:D50"],
      overwriteRisk: "medium"
    });

    expect(parsed.operation).toBe("remove_duplicate_rows");
    expect(parsed.confirmationLevel).toBe("destructive");
  });

  it("rejects a destructive cleanup branch with standard confirmation", () => {
    const parsed = DataCleanupPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "A1:D50",
      operation: "remove_duplicate_rows",
      explanation: "Remove duplicate records before publishing.",
      confidence: 0.87,
      requiresConfirmation: true,
      confirmationLevel: "standard",
      affectedRanges: ["Sheet1!A1:D50"],
      overwriteRisk: "medium"
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects a non-destructive cleanup branch with destructive confirmation", () => {
    const parsed = DataCleanupPlanDataSchema.safeParse({
      targetSheet: "Sheet1",
      targetRange: "A1:D50",
      operation: "trim_whitespace",
      explanation: "Trim whitespace before publishing.",
      confidence: 0.87,
      requiresConfirmation: true,
      confirmationLevel: "destructive",
      affectedRanges: ["Sheet1!A1:D50"],
      overwriteRisk: "low"
    });

    expect(parsed.success).toBe(false);
  });

  it("rejects a data cleanup update with an unsupported cleanup operation", () => {
    const parsed = DataCleanupUpdateDataSchema.safeParse({
      operation: "data_cleanup_update",
      targetSheet: "Sheet1",
      targetRange: "A1:D50",
      cleanupOperation: "definitely_not_supported",
      summary: "Unsupported cleanup result."
    });

    expect(parsed.success).toBe(false);
  });

  it("accepts a data_cleanup_update response envelope", () => {
    const parsed = DataCleanupUpdateResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "data_cleanup_update",
      requestId: "req_cleanup_001",
      hermesRunId: "run_cleanup_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      startedAt: "2026-04-20T09:00:00.000Z",
      completedAt: "2026-04-20T09:00:01.000Z",
      durationMs: 1000,
      trace: [
        { event: "data_cleanup_update_ready", timestamp: "2026-04-20T09:00:01.000Z" }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "data_cleanup_update",
        targetSheet: "Sheet1",
        targetRange: "A1:D50",
        cleanupOperation: "remove_duplicate_rows",
        summary: "Remove duplicate rows from the target range."
      }
    });

    expect(parsed.type).toBe("data_cleanup_update");
  });

  it("accepts a range_transfer_plan response envelope", () => {
    const parsed = RangeTransferPlanResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "range_transfer_plan",
      requestId: "req_transfer_plan_001",
      hermesRunId: "run_transfer_plan_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      startedAt: "2026-04-20T09:00:00.000Z",
      completedAt: "2026-04-20T09:00:01.000Z",
      durationMs: 1000,
      trace: [
        { event: "range_transfer_plan_ready", timestamp: "2026-04-20T09:00:01.000Z" }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceSheet: "Sheet1",
        sourceRange: "B2:C10",
        targetSheet: "Archive",
        targetRange: "A1:B9",
        operation: "copy",
        pasteMode: "values",
        transpose: false,
        explanation: "Copy the selected rows into the archive sheet.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:C10", "Archive!A1:B9"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(parsed.type).toBe("range_transfer_plan");
  });

  it("accepts a data_cleanup_plan response envelope", () => {
    const parsed = DataCleanupPlanResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "data_cleanup_plan",
      requestId: "req_cleanup_plan_001",
      hermesRunId: "run_cleanup_plan_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      startedAt: "2026-04-20T09:00:00.000Z",
      completedAt: "2026-04-20T09:00:01.000Z",
      durationMs: 1000,
      trace: [
        { event: "data_cleanup_plan_ready", timestamp: "2026-04-20T09:00:01.000Z" }
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
        targetRange: "A1:D50",
        operation: "remove_blank_rows",
        explanation: "Remove blank rows from the current table.",
        confidence: 0.88,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!A1:D50"],
        overwriteRisk: "medium",
        confirmationLevel: "destructive"
      }
    });

    expect(parsed.type).toBe("data_cleanup_plan");
  });

  it("accepts a range_transfer_update response envelope", () => {
    const parsed = HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "range_transfer_update",
      requestId: "req_transfer_001",
      hermesRunId: "run_transfer_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      startedAt: "2026-04-20T09:00:00.000Z",
      completedAt: "2026-04-20T09:00:01.000Z",
      durationMs: 1000,
      trace: [
        { event: "range_transfer_update_ready", timestamp: "2026-04-20T09:00:01.000Z" }
      ],
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "range_transfer_update",
        sourceSheet: "Sheet1",
        sourceRange: "B2:C10",
        targetSheet: "Archive",
        targetRange: "A1:B9",
        transferOperation: "copy",
        pasteMode: "values",
        transpose: false,
        summary: "Copy the selected rows into the archive sheet."
      }
    });

    expect(parsed.type).toBe("range_transfer_update");
  });

  it("restricts attachments to the MVP image set", () => {
    expect(() => {
      AttachmentSchema.parse({
        id: "att_bad",
        type: "image",
        mimeType: "application/pdf",
        source: "upload"
      });
    }).toThrow(/mimeType/i);
  });
});
