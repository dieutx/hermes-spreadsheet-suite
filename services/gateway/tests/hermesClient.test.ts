import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesResponseSchema, type HermesRequest } from "@hermes/contracts";
import { getConfig } from "../src/lib/config.ts";
import { SPREADSHEET_RUNTIME_RULES } from "../src/hermes/runtimeRules.ts";
import { buildHermesSpreadsheetRequestPrompt } from "../src/hermes/requestTemplate.ts";
import {
  HermesStructuredBodySchema,
  normalizeHermesStructuredBodyInput
} from "../src/hermes/structuredBody.ts";
import { HermesAgentClient } from "../src/lib/hermesClient.ts";
import { TraceBus } from "../src/lib/traceBus.ts";

beforeEach(() => {
  process.env.APPROVAL_SECRET = "test-approval-secret";
});

afterEach(() => {
  vi.restoreAllMocks();
  delete process.env.APPROVAL_SECRET;
  delete process.env.HERMES_DEBUG_INVALID_RESPONSES;
  delete process.env.HERMES_AGENT_BASE_URL;
  delete process.env.HERMES_API_SERVER_KEY;
  delete process.env.HERMES_AGENT_API_KEY;
  delete process.env.HERMES_AGENT_MODEL;
  delete process.env.HERMES_AGENT_ID;
  delete process.env.HERMES_AGENT_TIMEOUT_MS;
  delete process.env.HERMES_BASE_URL;
  delete process.env.HERMES_SERVICE_LABEL;
  delete process.env.HERMES_ENVIRONMENT_LABEL;
});

function baseRequest(overrides?: Partial<HermesRequest>): HermesRequest {
  return {
    schemaVersion: "1.0.0",
    requestId: "req_123",
    source: {
      channel: "google_sheets",
      clientVersion: "0.1.0",
      sessionId: "sess_123"
    },
    host: {
      platform: "google_sheets",
      workbookTitle: "Revenue Demo",
      activeSheet: "Sheet1",
      selectedRange: "A1:F6",
      locale: "en-US",
      timeZone: "America/Los_Angeles"
    },
    userMessage: "Explain this selection",
    conversation: [{ role: "user", content: "Explain this selection" }],
    context: {
      selection: {
        range: "A1:F6",
        headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"],
        values: [["2026-04-01", "Audio", "Cable", "North", 1, 15.5]]
      }
    },
    capabilities: {
      canRenderTrace: true,
      canRenderStructuredPreview: true,
      canConfirmWriteBack: true,
      supportsImageInputs: true,
      supportsWriteBackExecution: true
    },
    reviewer: {
      reviewerSafeMode: false,
      forceExtractionMode: null
    },
    confirmation: {
      state: "none"
    },
    ...overrides
  };
}

function chatCompletionEnvelope(content: string) {
  return {
    id: "chatcmpl_001",
    object: "chat.completion",
    created: 1_776_522_000,
    model: "hermes-agent",
    choices: [
      {
        index: 0,
        finish_reason: "stop",
        message: {
          role: "assistant",
          content
        }
      }
    ]
  };
}

describe("HermesAgentClient", () => {
  it("accepts wave-1 structured plan families", () => {
    const bodies = [
      {
        type: "sheet_structure_update",
        data: {
          targetSheet: "Sheet1",
          operation: "freeze_panes",
          frozenRows: 1,
          frozenColumns: 0,
          explanation: "Freeze the header row.",
          confidence: 0.9,
          requiresConfirmation: true,
          confirmationLevel: "standard"
        }
      },
      {
        type: "range_sort_plan",
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1:F25",
          hasHeader: true,
          keys: [
            { columnRef: "Status", direction: "asc" },
            { columnRef: "Due Date", direction: "desc" }
          ],
          explanation: "Sort by status and due date.",
          confidence: 0.91,
          requiresConfirmation: true
        }
      },
      {
        type: "range_filter_plan",
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1:F25",
          hasHeader: true,
          conditions: [
            { columnRef: "Status", operator: "equals", value: "Open" }
          ],
          combiner: "and",
          clearExistingFilters: true,
          explanation: "Filter to open rows.",
          confidence: 0.92,
          requiresConfirmation: true
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          targetSheet: "Sheet1",
          targetRange: "B2:B20",
          explanation: "Highlight values above 10 in red.",
          confidence: 0.93,
          requiresConfirmation: true,
          affectedRanges: ["B2:B20"],
          replacesExistingRules: true,
          managementMode: "replace_all_on_target",
          ruleType: "number_compare",
          comparator: "greater_than",
          value: 10,
          style: {
            backgroundColor: "#FEE2E2",
            textColor: "#B91C1C"
          }
        }
      }
    ] as const;

    for (const body of bodies) {
      expect(() =>
        HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput(body))
      ).not.toThrow();
    }
  });

  it("accepts wave-4 transfer and cleanup plan families", () => {
    const bodies = [
      {
        type: "range_transfer_plan",
        data: {
          sourceSheet: "Sheet1",
          sourceRange: "A1:D20",
          targetSheet: "Archive",
          targetRange: "A1:D20",
          operation: "move",
          pasteMode: "values",
          transpose: false,
          explanation: "Move the source block into the archive sheet.",
          confidence: 0.94,
          requiresConfirmation: true,
          affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
          overwriteRisk: "high",
          confirmationLevel: "destructive"
        }
      },
      {
        type: "data_cleanup_plan",
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1:D50",
          operation: "remove_duplicate_rows",
          keyColumns: ["A", "C"],
          explanation: "Remove duplicate rows using ID and email as keys.",
          confidence: 0.9,
          requiresConfirmation: true,
          affectedRanges: ["Sheet1!A1:D50"],
          overwriteRisk: "medium",
          confirmationLevel: "destructive"
        }
      }
    ] as const;

    for (const body of bodies) {
      expect(() =>
        HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput(body))
      ).not.toThrow();
    }
  });

  it("normalizes common report, chart, and cleanup aliases into contract-valid plan data", () => {
    const bodies = [
      {
        type: "analysis_report_plan",
        data: {
          sourceSheet: "Support",
          sourceRange: "A1:H10",
          outputMode: "materialize_report",
          targetSheet: "Summary",
          targetRange: "A1:F12",
          sections: ["ticket_counts_by_priority", "sla_risk_summary"],
          explanation: "Create a support summary report.",
          confidence: 0.86,
          requiresConfirmation: true
        }
      },
      {
        type: "chart_plan",
        data: {
          sourceSheet: "Marketing",
          sourceRange: "A1:D10",
          targetSheet: "Marketing",
          targetRange: "F2",
          chartType: "column",
          categoryField: "Channel",
          series: [
            { name: "Spend", range: "B2:B6" },
            { name: "Revenue", range: "C2:C6" }
          ],
          explanation: "Compare spend and revenue by channel.",
          confidence: 0.9,
          requiresConfirmation: true
        }
      },
      {
        type: "data_cleanup_plan",
        data: {
          targetSheet: "Sales",
          targetRange: "A1:F20",
          operation: "remove_duplicates",
          explanation: "Remove duplicate rows in the imported sales table.",
          confidence: 0.84,
          requiresConfirmation: true
        }
      },
      {
        type: "data_cleanup_plan",
        data: {
          targetSheet: "Sales",
          targetRange: "D2:D20",
          operation: "standardize_case",
          explanation: "Standardize region values to title case.",
          confidence: 0.85,
          requiresConfirmation: true
        }
      },
      {
        type: "range_sort_plan",
        data: {
          targetSheet: "Marketing",
          targetRange: "A15:D20",
          hasHeader: true,
          keys: [
            { field: "ROAS", direction: "descending" }
          ],
          explanation: "Sort the summary from best to worst ROAS.",
          confidence: 0.9,
          requiresConfirmation: true
        }
      },
      {
        type: "conditional_format_plan",
        data: {
          targetSheet: "Support",
          targetRange: "A2:H10",
          explanation: "Highlight rows where the SLA breached flag is Yes.",
          confidence: 0.92,
          requiresConfirmation: true,
          affectedRanges: ["Support!A2:H10"],
          replacesExistingRules: false,
          managementMode: "add",
          rule: {
            ruleType: "custom_formula",
            formula: "=$G2=\"Yes\"",
            format: {
              backgroundColor: "#FDECEC",
              textColor: "#9C0006",
              bold: true
            }
          }
        }
      }
    ] as const;

    for (const body of bodies) {
      expect(() =>
        HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput(body))
      ).not.toThrow();
    }
  });

  it("normalizes composite-plan reversible and confirmation defaults when chart steps are present", () => {
    expect(() =>
      HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
        type: "composite_plan",
        data: {
          steps: [
            {
              stepId: "write_summary",
              dependsOn: [],
              continueOnError: false,
              plan: {
                targetSheet: "Marketing",
                targetRange: "A15:D20",
                operation: "replace_range",
                explanation: "Write a summary table.",
                confidence: 0.9,
                requiresConfirmation: true,
                shape: { rows: 6, columns: 4 },
                values: [
                  ["Channel", "Spend", "Revenue", "ROAS"],
                  ["Email", 1200, 11880, 9.9],
                  ["Google Search", 15600, 71970, 4.6],
                  ["Meta", 16000, 62340, 3.9],
                  ["LinkedIn", 4200, 12600, 3],
                  ["TikTok", 3100, 6900, 2.2]
                ],
                overwriteRisk: "low"
              }
            },
            {
              stepId: "create_chart",
              dependsOn: ["write_summary"],
              continueOnError: false,
              plan: {
                sourceSheet: "Marketing",
                sourceRange: "A15:C20",
                targetSheet: "Marketing",
                targetRange: "F15:M30",
                chartType: "column",
                series: [{ name: "Spend", range: "B16:B20" }],
                categoryField: "Channel",
                explanation: "Create a spend chart.",
                confidence: 0.9,
                requiresConfirmation: true
              }
            }
          ],
          explanation: "Build the summary and chart.",
          confidence: 0.91,
          requiresConfirmation: true,
          affectedRanges: ["Marketing!A15:D20", "Marketing!F15:M30"],
          overwriteRisk: "low",
          confirmationLevel: "destructive",
          reversible: false,
          dryRunRecommended: true,
          dryRunRequired: false
        }
      }))
    ).not.toThrow();
  });

  it("normalizes sheet_update formula writes that also include header values into mixed_update", () => {
    expect(() =>
      HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
        type: "sheet_update",
        data: {
          targetSheet: "Budget",
          targetRange: "F1:G10",
          operation: "set_formulas",
          explanation: "Write headers and formulas together.",
          confidence: 0.9,
          requiresConfirmation: true,
          shape: { rows: 10, columns: 2 },
          values: [["Variance", "Variance %"], ...Array.from({ length: 9 }, () => [null, null])],
          formulas: [[null, null], ...Array.from({ length: 9 }, (_, index) => [`=D${index + 2}-C${index + 2}`, `=IFERROR(F${index + 2}/C${index + 2},\"\")`])],
          overwriteRisk: "low"
        }
      }))
    ).not.toThrow();
  });

  it("accepts wave-5 analysis, pivot, and chart plan families and rejects unsupported chart types", () => {
    const validBodies = [
      {
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
          explanation: "Summarize the selected sales range.",
          confidence: 0.92,
          requiresConfirmation: false,
          affectedRanges: ["Sales!A1:F50"],
          overwriteRisk: "none",
          confirmationLevel: "standard"
        }
      },
      {
        type: "pivot_table_plan",
        data: {
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
        }
      },
      {
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
      }
    ] as const;

    for (const body of validBodies) {
      expect(() =>
        HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput(body))
      ).not.toThrow();
    }

    expect(() => HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "chart_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "combo",
        series: [{ field: "Revenue", label: "Revenue" }],
        explanation: "Attempt an unsupported combo chart.",
        confidence: 0.5,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    }))).toThrow();
  });

  it("accepts wave-5 update families", () => {
    const bodies = [
      {
        type: "analysis_report_update",
        data: {
          operation: "analysis_report_update",
          targetSheet: "Analysis Report",
          targetRange: "A1:F20",
          summary: "Created analysis report on Analysis Report!A1:F20."
        }
      },
      {
        type: "pivot_table_update",
        data: {
          operation: "pivot_table_update",
          targetSheet: "Sales Pivot",
          targetRange: "A1:E12",
          summary: "Created pivot table on Sales Pivot!A1:E12."
        }
      },
      {
        type: "chart_update",
        data: {
          operation: "chart_update",
          targetSheet: "Sales Chart",
          targetRange: "A1",
          chartType: "line",
          summary: "Created line chart on Sales Chart!A1."
        }
      }
    ] as const;

    for (const body of bodies) {
      expect(() =>
        HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput(body))
      ).not.toThrow();
    }
  });

  it("accepts and renders composite_plan responses with strict preview metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
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
          reversible: false,
          dryRunRecommended: true,
          dryRunRequired: false
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_composite_plan_001",
      request: baseRequest({
        requestId: "req_composite_plan_001",
        userMessage: "Sort this table by revenue, then filter status to Open."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_composite_plan_001")?.response;
    expect(response).toMatchObject({
      type: "composite_plan",
      requestId: "req_composite_plan_001",
      hermesRunId: "run_composite_plan_001",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false,
        confirmationLevel: "standard"
      }
    });
    expect(() => HermesResponseSchema.parse(response)).not.toThrow();
    expect(response?.trace.some((event) => event.event === "composite_plan_ready")).toBe(true);
  });

  it("normalizes wrapped composite step plans that arrive as {type,data} envelopes", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "composite_plan",
        data: {
          steps: [
            {
              stepId: "step_create_sheet",
              dependsOn: [],
              continueOnError: false,
              plan: {
                type: "workbook_structure_update",
                data: {
                  operation: "create_sheet",
                  sheetName: "demo_sales_data",
                  explanation: "Create a new sheet named demo_sales_data.",
                  confidence: 0.99,
                  requiresConfirmation: true
                }
              }
            },
          {
            stepId: "step_seed_sales",
            dependsOn: ["step_create_sheet"],
            continueOnError: false,
            plan: {
                type: "sheet_update",
                data: {
                  targetSheet: "demo_sales_data",
                  targetRange: "A1:B3",
                  operation: "set_values",
                  explanation: "Populate sample sales data on demo_sales_data.",
                  confidence: 0.97,
                  requiresConfirmation: true,
                  shape: { rows: 3, columns: 2 },
                  values: [
                    ["Date", "Revenue"],
                    ["2026-04-01", 1250000],
                    ["2026-04-02", 970000]
                  ],
                  overwriteRisk: "low"
                }
              }
            }
          ],
          explanation: "Create a new sales-data sheet and populate it with sample rows.",
          confidence: 0.98,
          requiresConfirmation: true,
          affectedRanges: ["demo_sales_data!A1:B3"],
          overwriteRisk: "low",
          confirmationLevel: "standard",
          reversible: false,
          dryRunRecommended: false,
          dryRunRequired: false
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_composite_plan_wrapped_001",
        request: baseRequest({
          requestId: "req_composite_plan_wrapped_001",
          userMessage: "Create a new sheet named demo_sales_data and fill it with random sales data."
        }),
      traceBus
    });

    const response = traceBus.getRun("run_composite_plan_wrapped_001")?.response;
    expect(response).toMatchObject({
      type: "composite_plan",
      requestId: "req_composite_plan_wrapped_001",
      data: {
        reversible: false,
        steps: [
          {
            stepId: "step_create_sheet",
            plan: {
              operation: "create_sheet",
              sheetName: "demo_sales_data"
            }
          },
          {
            stepId: "step_seed_sales",
            plan: {
              targetSheet: "demo_sales_data",
              targetRange: "A1:B3",
              operation: "replace_range"
            }
          }
        ]
      }
    });
    expect(() => HermesResponseSchema.parse(response)).not.toThrow();
  });

  it("caps gateway traces while preserving the newest completion and failure events", () => {
    const client = new HermesAgentClient(getConfig()) as any;
    const longTrace = Array.from({ length: 260 }, (_, index) => ({
      event: "tool_selected",
      timestamp: `2026-04-20T12:${String(Math.floor(index / 60)).padStart(2, "0")}:${String(index % 60).padStart(2, "0")}.000Z`,
      label: `trace_${index}`
    }));

    const completedTrace = client.buildResponseTrace(
      "chat",
      longTrace,
      "2026-04-20T13:00:00.000Z"
    );
    expect(completedTrace).toHaveLength(200);
    expect(completedTrace.at(-2)).toMatchObject({ event: "result_generated" });
    expect(completedTrace.at(-1)).toMatchObject({ event: "completed" });
    expect(completedTrace[0]).toMatchObject({ label: "trace_62" });

    const failedTrace = client.withFailedTrace(longTrace);
    expect(failedTrace).toHaveLength(200);
    expect(failedTrace.at(-1)).toMatchObject({ event: "failed" });
    expect(failedTrace[0]).toMatchObject({ label: "trace_61" });
  });

  it("rejects extra keys inside wave-1 plan data objects", () => {
    const body = normalizeHermesStructuredBodyInput({
      type: "range_filter_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:F25",
        hasHeader: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" }
        ],
        combiner: "and",
        clearExistingFilters: true,
        explanation: "Filter to open rows.",
        confidence: 0.92,
        requiresConfirmation: true,
        unexpected: "drop-me"
      }
    });

    expect(() => HermesStructuredBodySchema.parse(body)).toThrow();
  });

  it("normalizes wave-1 plan data into fresh objects without dropping strictness", () => {
    const bodies = [
      {
        type: "sheet_structure_update" as const,
        data: {
          targetSheet: "Sheet1",
          operation: "freeze_panes",
          frozenRows: 1,
          frozenColumns: 0,
          explanation: "Freeze the header row.",
          confidence: 0.9,
          requiresConfirmation: true,
          confirmationLevel: "standard",
          unexpected: "keep-me"
        },
        assertNested(normalizedData: Record<string, unknown>, originalData: Record<string, unknown>) {
          expect(normalizedData).not.toBe(originalData);
        }
      },
      {
        type: "range_sort_plan" as const,
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1:F25",
          hasHeader: true,
          keys: [
            { columnRef: "Status", direction: "asc", unexpected: "keep-me" }
          ],
          explanation: "Sort by status.",
          confidence: 0.91,
          requiresConfirmation: true,
          unexpected: "keep-me"
        },
        assertNested(normalizedData: Record<string, unknown>, originalData: Record<string, unknown>) {
          expect(normalizedData).not.toBe(originalData);
          expect(normalizedData.keys).not.toBe(originalData.keys);
          expect((normalizedData.keys as Array<Record<string, unknown>>)[0]).not.toBe(
            (originalData.keys as Array<Record<string, unknown>>)[0]
          );
        }
      },
      {
        type: "range_filter_plan" as const,
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1:F25",
          hasHeader: true,
          conditions: [
            { columnRef: "Status", operator: "equals", value: "Open", unexpected: "keep-me" }
          ],
          combiner: "and",
          clearExistingFilters: true,
          explanation: "Filter to open rows.",
          confidence: 0.92,
          requiresConfirmation: true,
          unexpected: "keep-me"
        },
        assertNested(normalizedData: Record<string, unknown>, originalData: Record<string, unknown>) {
          expect(normalizedData).not.toBe(originalData);
          expect(normalizedData.conditions).not.toBe(originalData.conditions);
          expect((normalizedData.conditions as Array<Record<string, unknown>>)[0]).not.toBe(
            (originalData.conditions as Array<Record<string, unknown>>)[0]
          );
        }
      }
    ];

    for (const body of bodies) {
      const normalized = normalizeHermesStructuredBodyInput(body) as {
        type: string;
        data: Record<string, unknown>;
      };

      expect(normalized).not.toBe(body);
      expect(normalized.type).toBe(body.type);
      body.assertNested(normalized.data, body.data);
      expect(normalized.data.unexpected).toBe("keep-me");
      expect(() => HermesStructuredBodySchema.parse(normalized)).toThrow();
    }
  });

  it("normalizes data validation and named range bodies into fresh objects without dropping strictness", () => {
    const bodies = [
      {
        type: "data_validation_plan" as const,
        data: {
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
          replacesExistingValidation: true,
          unexpected: "keep-me"
        },
        assertNested(normalizedData: Record<string, unknown>, originalData: Record<string, unknown>) {
          expect(normalizedData).not.toBe(originalData);
        }
      },
      {
        type: "named_range_update" as const,
        data: {
          operation: "retarget",
          scope: "sheet",
          name: "InputRange",
          sheetName: "Sheet1",
          targetSheet: "Sheet1",
          targetRange: "B2:D20",
          explanation: "Retarget the named input block.",
          confidence: 0.91,
          requiresConfirmation: true,
          unexpected: "keep-me"
        },
        assertNested(normalizedData: Record<string, unknown>, originalData: Record<string, unknown>) {
          expect(normalizedData).not.toBe(originalData);
        }
      }
    ];

    for (const body of bodies) {
      const normalized = normalizeHermesStructuredBodyInput(body) as {
        type: string;
        data: Record<string, unknown>;
      };

      expect(normalized).not.toBe(body);
      expect(normalized.type).toBe(body.type);
      body.assertNested(normalized.data, body.data);
      expect(normalized.data.unexpected).toBeUndefined();
      expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
    }
  });

  it("assembles a final conditional_format_plan response with structured-preview ui and typed trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "conditional_format_plan",
        data: {
          targetSheet: "Sheet1",
          targetRange: "B2:B20",
          explanation: "Replace the target rules with a red highlight for values above 10.",
          confidence: 0.93,
          requiresConfirmation: true,
          affectedRanges: ["B2:B20"],
          replacesExistingRules: true,
          managementMode: "replace_all_on_target",
          ruleType: "number_compare",
          comparator: "greater_than",
          value: 10,
          style: {
            backgroundColor: "#FEE2E2",
            textColor: "#B91C1C"
          }
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_conditional_format_001",
      request: baseRequest({
        requestId: "req_conditional_format_001",
        userMessage: "Color values above 10 in red."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_conditional_format_001")?.response;
    expect(response).toMatchObject({
      type: "conditional_format_plan",
      requestId: "req_conditional_format_001",
      hermesRunId: "run_conditional_format_001",
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
        requiresConfirmation: true,
        replacesExistingRules: true
      }
    });
    expect(response?.trace.some((event) => event.event === "conditional_format_plan_ready")).toBe(true);
  });

  it("assembles a final sheet_structure_update response with wave-1 ui and trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "sheet_structure_update",
        data: {
          targetSheet: "Sheet1",
          operation: "unfreeze_panes",
          frozenRows: 0,
          frozenColumns: 0,
          explanation: "Unfreeze the current sheet.",
          confidence: 0.9,
          requiresConfirmation: true,
          confirmationLevel: "standard"
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_sheet_structure_001",
      request: baseRequest({
        requestId: "req_sheet_structure_001",
        userMessage: "Unfreeze the panes on this sheet."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_sheet_structure_001")?.response;
    expect(response).toMatchObject({
      type: "sheet_structure_update",
      requestId: "req_sheet_structure_001",
      hermesRunId: "run_sheet_structure_001",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        operation: "unfreeze_panes",
        frozenRows: 0,
        frozenColumns: 0,
        requiresConfirmation: true,
        confirmationLevel: "standard"
      }
    });
    expect(response?.trace.some((event) => event.event === "sheet_structure_update_ready")).toBe(true);
  });

  it("assembles a final range_sort_plan response with wave-1 ui and trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "range_sort_plan",
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1:F25",
          hasHeader: true,
          keys: [
            { columnRef: "Status", direction: "asc" }
          ],
          explanation: "Sort by status.",
          confidence: 0.91,
          requiresConfirmation: true
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_range_sort_001",
      request: baseRequest({
        requestId: "req_range_sort_001",
        userMessage: "Sort this table by Status."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_range_sort_001")?.response;
    expect(response).toMatchObject({
      type: "range_sort_plan",
      requestId: "req_range_sort_001",
      hermesRunId: "run_range_sort_001",
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
        requiresConfirmation: true
      }
    });
    expect(response?.trace.some((event) => event.event === "range_sort_plan_ready")).toBe(true);
  });

  it("assembles a final range_filter_plan response with wave-1 ui and trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "range_filter_plan",
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1:F25",
          hasHeader: true,
          conditions: [
            { columnRef: "Status", operator: "equals", value: "Open" }
          ],
          combiner: "and",
          clearExistingFilters: true,
          explanation: "Filter to open rows.",
          confidence: 0.92,
          requiresConfirmation: true
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_range_filter_001",
      request: baseRequest({
        requestId: "req_range_filter_001",
        userMessage: "Filter this range to open rows."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_range_filter_001")?.response;
    expect(response).toMatchObject({
      type: "range_filter_plan",
      requestId: "req_range_filter_001",
      hermesRunId: "run_range_filter_001",
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
        combiner: "and",
        clearExistingFilters: true,
        requiresConfirmation: true
      }
    });
    expect(response?.trace.some((event) => event.event === "range_filter_plan_ready")).toBe(true);
  });

  it("assembles a final data_validation_plan response with typed ui and trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "data_validation_plan",
        data: {
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
          requiresConfirmation: true
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_data_validation_001",
      request: baseRequest({
        requestId: "req_data_validation_001",
        userMessage: "Add a dropdown in B2:B20 using the StatusOptions named range."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_data_validation_001")?.response;
    expect(response).toMatchObject({
      type: "data_validation_plan",
      requestId: "req_data_validation_001",
      hermesRunId: "run_data_validation_001",
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
        ruleType: "list",
        namedRangeName: "StatusOptions",
        requiresConfirmation: true
      }
    });
    expect(response?.trace.some((event) => event.event === "data_validation_plan_ready")).toBe(true);
  });

  it("assembles a final named_range_update retarget response with typed ui and trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "named_range_update",
        data: {
          operation: "retarget",
          scope: "sheet",
          name: "InputRange",
          sheetName: "Sheet1",
          targetSheet: "Sheet1",
          targetRange: "B2:D20",
          explanation: "Retarget the named input block.",
          confidence: 0.91,
          requiresConfirmation: true
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_named_range_001",
      request: baseRequest({
        requestId: "req_named_range_001",
        userMessage: "Retarget the named range InputRange to B2:D20."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_named_range_001")?.response;
    expect(response).toMatchObject({
      type: "named_range_update",
      requestId: "req_named_range_001",
      hermesRunId: "run_named_range_001",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        operation: "retarget",
        scope: "sheet",
        name: "InputRange",
        sheetName: "Sheet1",
        targetSheet: "Sheet1",
        targetRange: "B2:D20",
        requiresConfirmation: true
      }
    });
    expect(response?.trace.some((event) => event.event === "named_range_update_ready")).toBe(true);
  });

  it("assembles a final range_transfer_plan response with structured-preview ui and typed trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "range_transfer_plan",
        data: {
          sourceSheet: "Sheet1",
          sourceRange: "A1:D20",
          targetSheet: "Archive",
          targetRange: "A1:D20",
          operation: "move",
          pasteMode: "values",
          transpose: false,
          explanation: "Move the source block into the archive sheet.",
          confidence: 0.94,
          requiresConfirmation: true,
          affectedRanges: ["Sheet1!A1:D20", "Archive!A1:D20"],
          overwriteRisk: "high",
          confirmationLevel: "destructive"
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_range_transfer_001",
      request: baseRequest({
        requestId: "req_range_transfer_001",
        userMessage: "Move A1:D20 to Archive!A1:D20."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_range_transfer_001")?.response;
    expect(response).toMatchObject({
      type: "range_transfer_plan",
      requestId: "req_range_transfer_001",
      hermesRunId: "run_range_transfer_001",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        sourceSheet: "Sheet1",
        sourceRange: "A1:D20",
        targetSheet: "Archive",
        targetRange: "A1:D20",
        operation: "move",
        pasteMode: "values",
        requiresConfirmation: true,
        confirmationLevel: "destructive"
      }
    });
    expect(() => HermesResponseSchema.parse(response)).not.toThrow();
    expect(response?.trace.some((event) => event.event === "range_transfer_plan_ready")).toBe(true);
  });

  it("assembles a final data_cleanup_plan response with structured-preview ui and typed trace metadata", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
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
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_data_cleanup_001",
      request: baseRequest({
        requestId: "req_data_cleanup_001",
        userMessage: "Split column B on commas into columns C:E."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_data_cleanup_001")?.response;
    expect(response).toMatchObject({
      type: "data_cleanup_plan",
      requestId: "req_data_cleanup_001",
      hermesRunId: "run_data_cleanup_001",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        targetSheet: "Sheet1",
        targetRange: "A2:F200",
        operation: "split_column",
        sourceColumn: "B",
        delimiter: ",",
        targetStartColumn: "C",
        requiresConfirmation: true,
        confirmationLevel: "destructive"
      }
    });
    expect(() => HermesResponseSchema.parse(response)).not.toThrow();
    expect(response?.trace.some((event) => event.event === "data_cleanup_plan_ready")).toBe(true);
  });

  it("normalizes chat bodies with extra data keys and warnings string[]", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    process.env.HERMES_API_SERVER_KEY = "agent-secret";
    process.env.HERMES_AGENT_MODEL = "hermes-agent";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();
    const request = baseRequest({
      requestId: "req_chat_001"
    });

    vi.stubGlobal("fetch", vi.fn(async (_url: string, init?: RequestInit) => {
      const body = JSON.parse(String(init?.body)) as {
        model: string;
        messages: Array<{ role: string; content: string }>;
      };

      expect(body.model).toBe("hermes-agent");
      expect(body.messages).toEqual([
        {
          role: "system",
          content: SPREADSHEET_RUNTIME_RULES
        },
        {
          role: "user",
          content: buildHermesSpreadsheetRequestPrompt(request)
        }
      ]);

      return new Response(JSON.stringify(chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "The current selection is a small sales table grouped by date, category, product, region, units, and revenue.",
          confidence: 0.93,
          selection: {
            sheet: "Sheet1",
            range: "A1:F6"
          },
          highlights: [
            "Structured tabular dataset with headers in row 1"
          ]
        },
        warnings: [
          "Selection range A1:F6 suggests more rows may exist than were included in the request payload."
        ],
        skillsUsed: ["spreadsheet-expert"]
      }))), {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    }));

    await client.processRequest({
      runId: "run_chat_001",
      request,
      traceBus
    });

    const response = traceBus.getRun("run_chat_001")?.response;
    expect(response).toMatchObject({
      schemaVersion: "1.0.0",
      type: "chat",
      requestId: "req_chat_001",
      hermesRunId: "run_chat_001",
      processedBy: "hermes",
      serviceLabel: "hermes-gateway-local",
      environmentLabel: "local-dev",
      skillsUsed: ["spreadsheet-expert"],
      warnings: [
        {
          code: "MODEL_WARNING",
          message: "Selection range A1:F6 suggests more rows may exist than were included in the request payload.",
          severity: "warning"
        }
      ],
      ui: {
        displayMode: "chat-first",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        message: "The current selection is a small sales table grouped by date, category, product, region, units, and revenue.",
        confidence: 0.93
      }
    });
    expect(response?.trace.at(-1)?.event).toBe("completed");
  });

  it("assembles a final chat_only analysis_report_plan response without confirmation affordances", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
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
          explanation: "Summarize the selected sales range.",
          confidence: 0.92,
          requiresConfirmation: false,
          affectedRanges: ["Sales!A1:F50"],
          overwriteRisk: "none",
          confirmationLevel: "standard"
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_analysis_report_plan_001",
      request: baseRequest({
        requestId: "req_analysis_report_plan_001",
        userMessage: "Analyze this range and give me a report summary."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_analysis_report_plan_001")?.response;
    expect(response).toMatchObject({
      type: "analysis_report_plan",
      requestId: "req_analysis_report_plan_001",
      hermesRunId: "run_analysis_report_plan_001",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: false
      },
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        outputMode: "chat_only",
        requiresConfirmation: false
      }
    });
    expect(() => HermesResponseSchema.parse(response)).not.toThrow();
    expect(response?.trace.some((event) => event.event === "analysis_report_plan_ready")).toBe(true);
  });

  it("assembles final wave-5 update responses with typed trace metadata and non-confirming ui", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";

    const cases = [
      {
        runId: "run_analysis_report_update_001",
        requestId: "req_analysis_report_update_001",
        userMessage: "Apply the report writeback.",
        body: {
          type: "analysis_report_update" as const,
          data: {
            operation: "analysis_report_update" as const,
            targetSheet: "Analysis Report",
            targetRange: "A1:F20",
            summary: "Created analysis report on Analysis Report!A1:F20."
          }
        },
        traceEvent: "analysis_report_update_ready",
        expectedData: {
          operation: "analysis_report_update",
          targetSheet: "Analysis Report",
          targetRange: "A1:F20",
          summary: "Created analysis report on Analysis Report!A1:F20."
        }
      },
      {
        runId: "run_pivot_table_update_001",
        requestId: "req_pivot_table_update_001",
        userMessage: "Apply the pivot writeback.",
        body: {
          type: "pivot_table_update" as const,
          data: {
            operation: "pivot_table_update" as const,
            targetSheet: "Sales Pivot",
            targetRange: "A1:E12",
            summary: "Created pivot table on Sales Pivot!A1:E12."
          }
        },
        traceEvent: "pivot_table_update_ready",
        expectedData: {
          operation: "pivot_table_update",
          targetSheet: "Sales Pivot",
          targetRange: "A1:E12",
          summary: "Created pivot table on Sales Pivot!A1:E12."
        }
      },
      {
        runId: "run_chart_update_001",
        requestId: "req_chart_update_001",
        userMessage: "Apply the chart writeback.",
        body: {
          type: "chart_update" as const,
          data: {
            operation: "chart_update" as const,
            targetSheet: "Sales Chart",
            targetRange: "A1",
            chartType: "line" as const,
            summary: "Created line chart on Sales Chart!A1."
          }
        },
        traceEvent: "chart_update_ready",
        expectedData: {
          operation: "chart_update",
          targetSheet: "Sales Chart",
          targetRange: "A1",
          chartType: "line",
          summary: "Created line chart on Sales Chart!A1."
        }
      }
    ] as const;

    for (const testCase of cases) {
      const client = new HermesAgentClient(getConfig());
      const traceBus = new TraceBus();

      vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
        chatCompletionEnvelope(JSON.stringify(testCase.body))
      ), {
        status: 200,
        headers: { "content-type": "application/json" }
      })));

      await client.processRequest({
        runId: testCase.runId,
        request: baseRequest({
          requestId: testCase.requestId,
          userMessage: testCase.userMessage
        }),
        traceBus
      });

      const response = traceBus.getRun(testCase.runId)?.response;
      expect(response).toMatchObject({
        type: testCase.body.type,
        requestId: testCase.requestId,
        hermesRunId: testCase.runId,
        ui: {
          displayMode: "structured-preview",
          showTrace: true,
          showWarnings: true,
          showConfidence: false,
          showRequiresConfirmation: false
        },
        data: testCase.expectedData
      });
      expect(() => HermesResponseSchema.parse(response)).not.toThrow();
      expect(response?.trace.some((event) => event.event === testCase.traceEvent)).toBe(true);
    }
  });

  it("accepts chat bodies with only allowed keys unchanged", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "The current selection is a revenue table.",
          confidence: 0.9,
          followUpSuggestions: [
            "Summarize revenue by region",
            "Show the average revenue per unit"
          ]
        },
        warnings: [
          {
            code: "PARTIAL_CONTEXT",
            message: "Only a subset of selected rows was present in the request payload.",
            severity: "warning",
            field: "context.selection.values"
          }
        ]
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_warning_objects_001",
      request: baseRequest({
        requestId: "req_warning_objects_001"
      }),
      traceBus
    });

    expect(traceBus.getRun("run_warning_objects_001")?.response).toMatchObject({
      type: "chat",
      warnings: [
        {
          code: "PARTIAL_CONTEXT",
          message: "Only a subset of selected rows was present in the request payload.",
          severity: "warning",
          field: "context.selection.values"
        }
      ],
      data: {
        message: "The current selection is a revenue table.",
        confidence: 0.9,
        followUpSuggestions: [
          "Summarize revenue by region",
          "Show the average revenue per unit"
        ]
      }
    });
  });

  it("accepts demo extraction responses when the demo warning is top-level only", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "sheet_import_plan",
        data: {
          sourceAttachmentId: "att_demo_001",
          targetSheet: "Sheet3",
          targetRange: "B4:D6",
          headers: ["Task", "Owner", "Status"],
          values: [
            ["Set up cron job", "Alex", "Done"],
            ["Review import flow", "Sam", "In Progress"]
          ],
          confidence: 0.2,
          requiresConfirmation: true,
          extractionMode: "demo",
          shape: {
            rows: 3,
            columns: 3
          }
        },
        warnings: [
          {
            code: "demo_output",
            message: "This is demo output only.",
            severity: "warning"
          }
        ]
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_demo_top_level_warning_001",
      request: baseRequest({
        requestId: "req_demo_top_level_warning_001",
        context: {
          attachments: [
            {
              id: "att_demo_001",
              type: "image",
              mimeType: "image/png",
              source: "upload",
              storageRef: "blob://att_demo_001"
            }
          ]
        },
        reviewer: {
          reviewerSafeMode: true,
          forceExtractionMode: "demo"
        }
      }),
      traceBus
    });

    const response = traceBus.getRun("run_demo_top_level_warning_001")?.response;
    expect(response?.type).toBe("sheet_import_plan");
    expect(response?.data.extractionMode).toBe("demo");
    expect(response?.warnings?.[0]?.code).toBe("demo_output");
    expect(response?.data.code).not.toBe("INTERNAL_ERROR");
  });

  it("assembles a valid Step 1 formula response from a minimal Hermes formula body", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "formula",
        data: {
          intent: "suggest",
          targetCell: "F12",
          formula: "=SUMIFS(F2:F11, D2:D11, \"North\")",
          formulaLanguage: "google_sheets",
          explanation: "This sums Revenue in column F where Region equals North.",
          confidence: 0.95,
          requiresConfirmation: true
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_formula_001",
      request: baseRequest({
        requestId: "req_formula_001",
        userMessage: "Suggest a formula for North revenue"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_formula_001")?.response;
    expect(response).toMatchObject({
      schemaVersion: "1.0.0",
      type: "formula",
      requestId: "req_formula_001",
      hermesRunId: "run_formula_001",
      processedBy: "hermes",
      ui: {
        displayMode: "structured-preview",
        showTrace: true,
        showWarnings: true,
        showConfidence: true,
        showRequiresConfirmation: true
      },
      data: {
        intent: "suggest",
        targetCell: "F12",
        formula: "=SUMIFS(F2:F11, D2:D11, \"North\")",
        formulaLanguage: "google_sheets",
        explanation: "This sums Revenue in column F where Region equals North.",
        confidence: 0.95,
        requiresConfirmation: true
      }
    });
    expect(response?.trace.at(-2)?.event).toBe("result_generated");
    expect(response?.trace.at(-1)?.event).toBe("completed");
  });

  it("coerces string alternateFormulas from Hermes into contract-valid formula alternatives", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "formula",
        data: {
          intent: "suggest",
          formula: "=SUMIF($B$2:$B$7,B8,$F$2:$F$7)",
          formulaLanguage: "google_sheets",
          explanation: "This sums Revenue for the category in B8.",
          confidence: 0.74,
          alternateFormulas: [
            "=SUMIFS($F$2:$F$7,$B$2:$B$7,B8)",
            "=SUMIF(B2:B7,B8,F2:F7)"
          ]
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_formula_alt_strings_001",
      request: baseRequest({
        requestId: "req_formula_alt_strings_001",
        userMessage: "sumif at the row 8"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_formula_alt_strings_001")?.response;
    expect(response?.type).toBe("formula");
    expect(response?.data.alternateFormulas).toEqual([
      {
        formula: "=SUMIFS($F$2:$F$7,$B$2:$B$7,B8)",
        explanation: "Alternative formulation."
      },
      {
        formula: "=SUMIF(B2:B7,B8,F2:F7)",
        explanation: "Alternative formulation."
      }
    ]);
    expect(response?.data.code).not.toBe("INTERNAL_ERROR");
  });

  it("maps Hermes missing-context error codes into contract-safe spreadsheet context errors", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "error",
        data: {
          code: "MISSING_REQUIRED_CONTEXT",
          message: "Tell me which cell and condition to use.",
          retryable: true,
          userAction: "Specify the target cell and the SUMIF condition."
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_error_missing_context_001",
      request: baseRequest({
        requestId: "req_error_missing_context_001",
        userMessage: "sumif at the row 8"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_error_missing_context_001")?.response;
    expect(response?.type).toBe("error");
    expect(response?.data).toMatchObject({
      code: "SPREADSHEET_CONTEXT_MISSING",
      message: "Tell me which cell and condition to use.",
      retryable: true,
      userAction: "Specify the target cell and the SUMIF condition."
    });
  });

  it("sanitizes internal wording from otherwise valid Hermes error bodies", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "error",
        data: {
          code: "UNSUPPORTED_OPERATION",
          message: "This resize request cannot be represented by the available contract schema.",
          retryable: false,
          userAction: "Review the schema and choose a supported contract operation."
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_error_sanitized_001",
      request: baseRequest({
        requestId: "req_error_sanitized_001",
        userMessage: "Resize this sheet to 100x100."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_error_sanitized_001")?.response;
    expect(response?.type).toBe("error");
    expect(response?.data).toMatchObject({
      code: "UNSUPPORTED_OPERATION",
      message: "I can't do that exact spreadsheet action here.",
      retryable: false,
      userAction: "Tell me the target sheet, range, cell, or output location you want me to use, or ask for the closest supported alternative."
    });
  });

  it("bounds provider error text before returning gateway error responses", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();
    const oversizedProviderMessage = "A".repeat(12001);

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify({
      error: {
        message: oversizedProviderMessage
      }
    }), {
      status: 503,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_provider_error_bound_001",
      request: baseRequest({
        requestId: "req_provider_error_bound_001",
        userMessage: "Explain this workbook."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_provider_error_bound_001")?.response;
    expect(response?.type).toBe("error");
    expect(response?.data).toMatchObject({
      code: "PROVIDER_ERROR",
      message: "The Hermes service couldn't complete that request right now.",
      retryable: true,
      userAction: "Retry the request after the remote Hermes Agent service recovers."
    });
    expect(JSON.stringify(response)).not.toContain(oversizedProviderMessage);
  });

  it("sanitizes internal warning text before returning gateway envelopes", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "I checked the workbook context.",
          confidence: 0.91
        },
        warnings: [
          {
            code: "HERMES_API_SERVER_KEY",
            message: "ReferenceError at /srv/hermes/services/gateway/src/app.ts:99 HERMES_API_SERVER_KEY=secret_123",
            severity: "warning",
            field: "/srv/hermes/services/gateway/src/app.ts"
          }
        ]
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_warning_sanitized_001",
      request: baseRequest({
        requestId: "req_warning_sanitized_001",
        userMessage: "Check this workbook."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_warning_sanitized_001")?.response;
    expect(response?.type).toBe("chat");
    expect(response?.warnings?.[0]).toEqual({
      code: "INTERNAL_WARNING",
      message: "A gateway warning was hidden because it contained internal diagnostic details.",
      severity: "warning"
    });
    expect(JSON.stringify(response?.warnings)).not.toContain("HERMES_API_SERVER_KEY");
    expect(JSON.stringify(response?.warnings)).not.toContain("/srv/hermes");
  });

  it("sanitizes nested data warnings before returning import previews", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "sheet_import_plan",
        data: {
          sourceAttachmentId: "att_warning_001",
          targetSheet: "Sheet1",
          targetRange: "A1:B2",
          headers: ["A", "B"],
          values: [["1", "2"]],
          confidence: 0.81,
          warnings: [
            {
              code: "EXTRACTION_NOTE",
              message: "Traceback at /root/hermes/extractor.py:88 with APPROVAL_SECRET=secret_123",
              severity: "warning",
              field: "/root/hermes/extractor.py"
            }
          ],
          requiresConfirmation: true,
          extractionMode: "real",
          shape: { rows: 2, columns: 2 }
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_nested_warning_sanitized_001",
      request: baseRequest({
        requestId: "req_nested_warning_sanitized_001",
        userMessage: "Import this image table."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_nested_warning_sanitized_001")?.response;
    expect(response?.type).toBe("sheet_import_plan");
    expect((response?.data as any).warnings?.[0]).toEqual({
      code: "EXTRACTION_NOTE",
      message: "A gateway warning was hidden because it contained internal diagnostic details.",
      severity: "warning"
    });
    expect(JSON.stringify(response?.data)).not.toContain("APPROVAL_SECRET");
    expect(JSON.stringify(response?.data)).not.toContain("/root/hermes");
  });

  it("sanitizes client-facing response metadata before returning gateway envelopes", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "I checked the workbook context.",
          confidence: 0.88
        },
        skillsUsed: [
          "SelectionExplainerSkill",
          "/srv/hermes/private_tool.ts",
          "HERMES_API_SERVER_KEY=secret_123"
        ],
        downstreamProvider: {
          label: "https://internal.example/provider",
          model: "gpt-5 HERMES_API_SERVER_KEY=secret_123"
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_metadata_sanitized_001",
      request: baseRequest({
        requestId: "req_metadata_sanitized_001",
        userMessage: "Check this workbook."
      }),
      traceBus
    });

    const response = traceBus.getRun("run_metadata_sanitized_001")?.response;
    expect(response?.type).toBe("chat");
    expect(response?.skillsUsed).toEqual(["SelectionExplainerSkill"]);
    expect(response?.downstreamProvider).toBeNull();
    expect(JSON.stringify(response)).not.toContain("HERMES_API_SERVER_KEY");
    expect(JSON.stringify(response)).not.toContain("/srv/hermes");
    expect(JSON.stringify(response)).not.toContain("internal.example");
  });

  it("maps descriptive overwriteRisk text into a contract-safe risk level for sheet_update responses", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "sheet_update",
        data: {
          targetSheet: "Sheet1",
          targetRange: "H11",
          operation: "set_formulas",
          explanation: "Set H11 to the corrected SUMIF formula.",
          confidence: 0.93,
          requiresConfirmation: true,
          shape: { rows: 1, columns: 1 },
          formulas: [["=SUMIF(B:B,\"north\",F:F)"]],
          overwriteRisk: "Replaces the existing formula in H11."
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_sheet_update_overwrite_risk_text_001",
      request: baseRequest({
        requestId: "req_sheet_update_overwrite_risk_text_001",
        userMessage: "insert to cell h11"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_sheet_update_overwrite_risk_text_001")?.response;
    expect(response?.type).toBe("sheet_update");
    expect(response?.data).toMatchObject({
      targetSheet: "Sheet1",
      targetRange: "H11",
      operation: "set_formulas",
      overwriteRisk: "low"
    });
  });

  it("accepts a single fenced JSON body and preserves requestId and hermesRunId", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope([
        "```json",
        JSON.stringify({
          type: "chat",
          data: {
            message: "Fenced JSON normalized correctly.",
            confidence: 0.88
          }
        }, null, 2),
        "```"
      ].join("\n"))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_fenced_001",
      request: baseRequest({
        requestId: "req_fenced_001"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_fenced_001")?.response;
    expect(response?.requestId).toBe("req_fenced_001");
    expect(response?.hermesRunId).toBe("run_fenced_001");
    expect(response?.type).toBe("chat");
  });

  it("rejects invalid prose output without writing debug files unless explicitly enabled", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();
    const debugPrefix = "hermes-spreadsheet-invalid-req_mixed_001";
    const debugDir = tmpdir();
    const rawAssistantContent = [
      "Here is the JSON body you asked for:",
      JSON.stringify({
        type: "chat",
        data: {
          message: "This should be rejected because prose surrounds the JSON."
        }
      })
    ].join("\n");
    const existingDebugFiles = new Set(
      (await fs.readdir(debugDir)).filter((fileName) => fileName.startsWith(debugPrefix))
    );

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(rawAssistantContent)
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_mixed_001",
      request: baseRequest({
        requestId: "req_mixed_001"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_mixed_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_mixed_001",
      hermesRunId: "run_mixed_001",
      processedBy: "hermes",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare a valid spreadsheet response for that request.",
        retryable: true,
        userAction: "Try again with the target sheet, range, cell, or attachment, or split the request into smaller steps."
      }
    });

    const newDebugFiles = (await fs.readdir(debugDir))
      .filter((fileName) => fileName.startsWith(debugPrefix))
      .filter((fileName) => !existingDebugFiles.has(fileName))
      .sort();

    expect(newDebugFiles.length).toBe(0);
  });

  it("writes invalid assistant debug files only when explicitly enabled", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    process.env.HERMES_DEBUG_INVALID_RESPONSES = "true";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();
    const debugPrefix = "hermes-spreadsheet-invalid-req_mixed_001";
    const debugDir = tmpdir();
    const rawAssistantContent = [
      "Here is the JSON body you asked for:",
      JSON.stringify({
        type: "chat",
        data: {
          message: "This should be rejected because prose surrounds the JSON."
        }
      })
    ].join("\n");
    const existingDebugFiles = new Set(
      (await fs.readdir(debugDir)).filter((fileName) => fileName.startsWith(debugPrefix))
    );

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(rawAssistantContent)
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_mixed_001",
      request: baseRequest({
        requestId: "req_mixed_001"
      }),
      traceBus
    });

    const newDebugFiles = (await fs.readdir(debugDir))
      .filter((fileName) => fileName.startsWith(debugPrefix))
      .filter((fileName) => !existingDebugFiles.has(fileName))
      .sort();

    expect(newDebugFiles.length).toBe(1);

    const debugContents = await fs.readFile(path.join(debugDir, newDebugFiles[0]), "utf8");
    expect(debugContents).toContain("reason: assistant_content_not_single_json_object");
    expect(debugContents).toContain(rawAssistantContent);
  });

  it("rejects invalid warnings shapes beyond string[] or warning objects", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "This payload has invalid warnings."
        },
        warnings: [123]
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_invalid_warnings_001",
      request: baseRequest({
        requestId: "req_invalid_warnings_001"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_invalid_warnings_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_invalid_warnings_001",
      hermesRunId: "run_invalid_warnings_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare a valid spreadsheet response for that request.",
        retryable: true,
        userAction: "Try again with the target sheet, range, cell, or attachment, or split the request into smaller steps."
      }
    });
  });

  it("normalizes model warning severities like medium to contract-safe warning severities", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "This selection shows one revenue row.",
          confidence: 0.92
        },
        warnings: [
          {
            code: "WRITEBACK_NOT_EMITTED",
            message: "The response stayed conversational instead of returning a richer structured plan.",
            severity: "medium"
          }
        ]
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_warning_severity_001",
      request: baseRequest({
        requestId: "req_warning_severity_001",
        userMessage: "Explain this selection",
        conversation: [{ role: "user", content: "Explain this selection" }]
      }),
      traceBus
    });

    const response = traceBus.getRun("run_warning_severity_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_warning_severity_001",
      hermesRunId: "run_warning_severity_001",
      type: "chat",
      data: {
        message: "This selection shows one revenue row.",
        confidence: 0.92
      },
      warnings: [
        {
          code: "WRITEBACK_NOT_EMITTED",
          severity: "warning"
        }
      ]
    });
  });

  it("fails closed when a writeback request incorrectly comes back as chat-only", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "Confirmed. Please create a new sheet named demo_sales_data.",
          confidence: 0.92
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_missing_writeback_001",
      request: baseRequest({
        requestId: "req_missing_writeback_001",
        userMessage: "Confirm create sheet demo_sales_data",
        conversation: [{ role: "user", content: "Confirm create sheet demo_sales_data" }]
      }),
      traceBus
    });

    const response = traceBus.getRun("run_missing_writeback_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_missing_writeback_001",
      hermesRunId: "run_missing_writeback_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare an actionable spreadsheet plan for that request.",
        retryable: true,
        userAction: "Retry the sheet action explicitly, or tell me the sheet/tab name you want me to create, rename, move, hide, or delete."
      }
    });
    expect(response?.trace.some((event) => event.event === "failed")).toBe(true);
  });

  it("fails closed when a targeted formula write request incorrectly comes back as formula-only advice", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "formula",
        data: {
          intent: "fix",
          formula: "=SUMIF(B:B,\"north\",F:F)",
          formulaLanguage: "google_sheets",
          explanation: "Use SUMIF with the revenue column.",
          confidence: 0.91
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_missing_formula_writeback_001",
      request: baseRequest({
        requestId: "req_missing_formula_writeback_001",
        userMessage: "Fix the formula in H11 and apply it.",
        conversation: [{ role: "user", content: "Fix the formula in H11 and apply it." }]
      }),
      traceBus
    });

    const response = traceBus.getRun("run_missing_formula_writeback_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_missing_formula_writeback_001",
      hermesRunId: "run_missing_formula_writeback_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare an actionable spreadsheet plan for that request.",
        retryable: true,
        userAction: "Tell me the exact cell or range you want me to change, or retry the write request more explicitly."
      }
    });
  });

  it("fails closed when a mixed advisory-plus-write request comes back as chat-only", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "The selection looks fine, then you can sort it.",
          confidence: 0.84
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_mixed_writeback_001",
      request: baseRequest({
        requestId: "req_mixed_writeback_001",
        userMessage: "Explain this selection and then sort it by revenue descending.",
        conversation: [{ role: "user", content: "Explain this selection and then sort it by revenue descending." }]
      }),
      traceBus
    });

    const response = traceBus.getRun("run_mixed_writeback_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_mixed_writeback_001",
      hermesRunId: "run_mixed_writeback_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare an actionable spreadsheet plan for that request.",
        retryable: true,
        userAction: "Split the analysis and writeback into separate steps, or tell me the target sheet, range, or output location you want me to change now."
      }
    });
  });

  it("fails closed when a materialized analysis request incorrectly comes back as chat_only", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "analysis_report_plan",
        data: {
          sourceSheet: "Sheet1",
          sourceRange: "A1:F20",
          outputMode: "chat_only",
          sections: [
            {
              type: "summary_stats",
              title: "Summary"
              ,
              summary: "Summarize the current selection before writing it anywhere.",
              sourceRanges: ["Sheet1!A1:F20"]
            }
          ],
          explanation: "Summarize the selection.",
          confidence: 0.88,
          affectedRanges: ["Sheet1!A1:F20"],
          overwriteRisk: "low",
          confirmationLevel: "standard",
          requiresConfirmation: false
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_materialized_analysis_001",
      request: baseRequest({
        requestId: "req_materialized_analysis_001",
        userMessage: "Analyze this range and put the report on a new sheet named Summary.",
        conversation: [{ role: "user", content: "Analyze this range and put the report on a new sheet named Summary." }]
      }),
      traceBus
    });

    const response = traceBus.getRun("run_materialized_analysis_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_materialized_analysis_001",
      hermesRunId: "run_materialized_analysis_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare an actionable spreadsheet plan for that request.",
        retryable: true,
        userAction: "Tell me the source table or range and the output sheet or anchor you want me to use, or retry the request more explicitly."
      }
    });
  });

  it("keeps current-table guidance when a pivot write request degrades to chat-only", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "You can build a pivot table from the current data.",
          confidence: 0.8
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_missing_pivot_writeback_001",
      request: baseRequest({
        requestId: "req_missing_pivot_writeback_001",
        userMessage: "Create a pivot table from the current table.",
        conversation: [{ role: "user", content: "Create a pivot table from the current table." }],
        context: {
          ...baseRequest().context,
          selection: {
            range: "J6",
            values: [[123]]
          },
          currentRegion: {
            range: "A1:F11",
            headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"]
          },
          currentRegionArtifactTarget: "A13"
        }
      }),
      traceBus
    });

    const response = traceBus.getRun("run_missing_pivot_writeback_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_missing_pivot_writeback_001",
      hermesRunId: "run_missing_pivot_writeback_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare an actionable spreadsheet plan for that request.",
        retryable: true,
        userAction: "Retry the request for the current table, or tell me a different output sheet, range, or anchor if you want the result somewhere else."
      }
    });
  });

  it("uses write-aware malformed-payload guidance for workbook actions instead of asking for a range", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope("not-json")
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_malformed_workbook_write_001",
      request: baseRequest({
        requestId: "req_malformed_workbook_write_001",
        userMessage: "Create a new sheet named Summary.",
        conversation: [{ role: "user", content: "Create a new sheet named Summary." }]
      }),
      traceBus
    });

    const response = traceBus.getRun("run_malformed_workbook_write_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_malformed_workbook_write_001",
      hermesRunId: "run_malformed_workbook_write_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare a valid spreadsheet response for that request.",
        retryable: true,
        userAction: "Retry the sheet action explicitly, or tell me the sheet/tab name you want me to create, rename, move, hide, or delete."
      }
    });
  });

  it("suggests a helper sheet when input cells overlap the current source table", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "Use a lookup formula.",
          confidence: 0.82
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_input_layout_conflict_001",
      request: baseRequest({
        requestId: "req_input_layout_conflict_001",
        host: {
          ...baseRequest().host,
          activeSheet: "Sheet3",
          selectedRange: "J6"
        },
        userMessage: "I have an employee dataset and want to build a dynamic search tool. With the EEID in cell A1 and the target attribute in cell B1, I need a formula that looks up the specific value.",
        conversation: [{
          role: "user",
          content: "I have an employee dataset and want to build a dynamic search tool. With the EEID in cell A1 and the target attribute in cell B1, I need a formula that looks up the specific value."
        }],
        context: {
          selection: {
            range: "J6",
            values: [[""]]
          },
          currentRegion: {
            range: "A1:N6",
            headers: [
              "EEID",
              "Full Name",
              "Job Title",
              "Department",
              "Business Unit",
              "Gender",
              "Ethnicity",
              "Age",
              "Hire Date",
              "Annual Salary",
              "Bonus %",
              "Country",
              "City",
              "Exit Date"
            ]
          }
        }
      }),
      traceBus
    });

    const response = traceBus.getRun("run_input_layout_conflict_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_input_layout_conflict_001",
      hermesRunId: "run_input_layout_conflict_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare an actionable spreadsheet plan for that request.",
        retryable: true,
        userAction: "Your requested input or result cells overlap the current source table. I can create a separate helper sheet for the inputs and output, or you can tell me a different target sheet or cells."
      }
    });
  });

  it("augments helper-sheet composite plans with a visible guidance block when the model omits scaffold content", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "composite_plan",
        data: {
          steps: [
            {
              stepId: "create_lookup_sheet",
              dependsOn: [],
              continueOnError: false,
              plan: {
                operation: "create_sheet",
                sheetName: "Lookup_Demo",
                explanation: "Create a helper sheet for the lookup tool.",
                confidence: 0.9,
                requiresConfirmation: true
              }
            },
            {
              stepId: "set_lookup_formula",
              dependsOn: ["create_lookup_sheet"],
              continueOnError: false,
              plan: {
                targetSheet: "Lookup_Demo",
                targetRange: "C1",
                operation: "set_formulas",
                formulas: [[
                  '=IF(VLOOKUP(A1,Sheet3!A:N,MATCH("Exit Date",Sheet3!1:1,0),FALSE)<>"","Terminated",VLOOKUP(A1,Sheet3!A:N,MATCH(B1,Sheet3!1:1,0),FALSE))'
                ]],
                explanation: "Set the lookup result formula.",
                confidence: 0.9,
                requiresConfirmation: true,
                overwriteRisk: "low",
                shape: {
                  rows: 1,
                  columns: 1
                }
              }
            }
          ],
          explanation: "Create a lookup helper sheet.",
          confidence: 0.9,
          requiresConfirmation: true,
          affectedRanges: ["Lookup_Demo!C1"],
          overwriteRisk: "low",
          confirmationLevel: "standard",
          reversible: false,
          dryRunRecommended: false,
          dryRunRequired: false
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_helper_sheet_scaffold_001",
      request: baseRequest({
        requestId: "req_helper_sheet_scaffold_001",
        host: {
          ...baseRequest().host,
          activeSheet: "Sheet3",
          selectedRange: "J6"
        },
        userMessage: "I have an employee dataset and want to build a dynamic search tool. With the EEID in cell A1 and the target attribute in cell B1, I need a formula that looks up the specific value.",
        conversation: [{
          role: "user",
          content: "I have an employee dataset and want to build a dynamic search tool. With the EEID in cell A1 and the target attribute in cell B1, I need a formula that looks up the specific value."
        }],
        context: {
          selection: {
            range: "J6",
            values: [[""]]
          },
          currentRegion: {
            range: "A1:N6",
            headers: [
              "EEID",
              "Full Name",
              "Job Title",
              "Department",
              "Business Unit",
              "Gender",
              "Ethnicity",
              "Age",
              "Hire Date",
              "Annual Salary",
              "Bonus %",
              "Country",
              "City",
              "Exit Date"
            ]
          }
        }
      }),
      traceBus
    });

    const response = traceBus.getRun("run_helper_sheet_scaffold_001")?.response;
    expect(response?.type).toBe("composite_plan");
    expect(response?.data.steps).toHaveLength(3);
    expect(response?.data.steps[1]).toMatchObject({
      stepId: "seed_helper_sheet_guidance",
      dependsOn: ["create_lookup_sheet"],
      continueOnError: false,
      plan: {
        targetSheet: "Lookup_Demo",
        operation: "replace_range",
        targetRange: "A3:B7",
        explanation: "Seed helper-sheet guidance for the input and output cells.",
        values: [
          ["Hermes helper sheet", ""],
          ["How to use", "Edit the input cells, then read the result cell."],
          ["A1", "Input value"],
          ["B1", "Parameter or field"],
          ["C1", "Result cell"]
        ],
        shape: {
          rows: 5,
          columns: 2
        }
      }
    });
    expect(response?.data.steps[2]).toMatchObject({
      stepId: "set_lookup_formula",
      dependsOn: ["create_lookup_sheet", "seed_helper_sheet_guidance"]
    });
  });

  it("rejects invalid chat bodies without data.message", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          selection: {
            sheet: "Sheet1",
            range: "A1:F6"
          }
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_invalid_chat_001",
      request: baseRequest({
        requestId: "req_invalid_chat_001"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_invalid_chat_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_invalid_chat_001",
      hermesRunId: "run_invalid_chat_001",
      type: "error",
      data: {
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare a valid spreadsheet response for that request.",
        retryable: true,
        userAction: "Try again with the target sheet, range, cell, or attachment, or split the request into smaller steps."
      }
    });
  });

  it("rejects chat bodies with invalid followUpSuggestions instead of silently dropping them", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(
      chatCompletionEnvelope(JSON.stringify({
        type: "chat",
        data: {
          message: "This should fail.",
          followUpSuggestions: [
            "one",
            "two",
            "three",
            "four",
            "five",
            "six"
          ]
        }
      }))
    ), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_invalid_followups_001",
      request: baseRequest({
        requestId: "req_invalid_followups_001"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_invalid_followups_001")?.response;
    expect(response?.type).toBe("error");
    expect(response?.data.code).toBe("INTERNAL_ERROR");
    expect(response?.data.message).toBe("I couldn't prepare a valid spreadsheet response for that request.");
    expect(response?.data.userAction).toBe(
      "Try again with the target sheet, range, cell, or attachment, or split the request into smaller steps."
    );
  });

  it("keeps reviewer-safe no-fabrication behavior unchanged", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://hermes.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(chatCompletionEnvelope(
      JSON.stringify({
        type: "sheet_import_plan",
        data: {
          sourceAttachmentId: "att_102",
          targetSheet: "Sheet1",
          targetRange: "A1:B3",
          headers: ["A", "B"],
          values: [
            ["fake", "preview"],
            ["still", "fake"]
          ],
          confidence: 0.6,
          warnings: [],
          requiresConfirmation: true,
          extractionMode: "real",
          shape: { rows: 3, columns: 2 }
        }
      })
    )), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_unavailable_001",
      request: baseRequest({
        requestId: "req_img_unavail_001",
        context: {
          attachments: [
            {
              id: "att_102",
              type: "image",
              mimeType: "image/png",
              source: "upload",
              storageRef: "blob://att_102"
            }
          ]
        },
        reviewer: {
          reviewerSafeMode: true,
          forceExtractionMode: "unavailable"
        }
      }),
      traceBus
    });

    const response = traceBus.getRun("run_unavailable_001")?.response;
    expect(response?.type).toBe("error");
    expect(response?.data.code).toBe("EXTRACTION_UNAVAILABLE");
    expect(response?.data.message).toContain("unavailable");
    expect(response?.type).not.toBe("sheet_import_plan");
    expect(response?.type).not.toBe("extracted_table");
    expect(response?.data.code).not.toBe("INTERNAL_ERROR");
  });

  it("returns a contract-valid unavailable error when reviewer-safe mode forces unavailable and Hermes replies with chat", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://hermes.test/v1";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async () => new Response(JSON.stringify(chatCompletionEnvelope(
      JSON.stringify({
        type: "chat",
        data: {
          message: "Image analysis is unavailable in the current reviewer-safe mode.",
          confidence: 0.99
        },
        warnings: [
          {
            code: "EXTRACTION_UNAVAILABLE",
            message: "Reviewer-safe mode forced extractionMode=\"unavailable\".",
            severity: "info"
          }
        ]
      })
    )), {
      status: 200,
      headers: { "content-type": "application/json" }
    })));

    await client.processRequest({
      runId: "run_unavailable_chat_001",
      request: baseRequest({
        requestId: "req_img_unavail_chat_001",
        context: {
          attachments: [
            {
              id: "att_103",
              type: "image",
              mimeType: "image/png",
              source: "upload",
              storageRef: "blob://att_103"
            }
          ]
        },
        reviewer: {
          reviewerSafeMode: true,
          forceExtractionMode: "unavailable"
        }
      }),
      traceBus
    });

    const response = traceBus.getRun("run_unavailable_chat_001")?.response;
    expect(response?.type).toBe("error");
    expect(response?.data.code).toBe("EXTRACTION_UNAVAILABLE");
    expect(response?.type).not.toBe("sheet_import_plan");
    expect(response?.type).not.toBe("extracted_table");
    expect(response?.data.code).not.toBe("INTERNAL_ERROR");
  });

  it("returns a terminal TIMEOUT error when the Hermes provider does not answer before the deadline", async () => {
    vi.useFakeTimers();
    process.env.HERMES_AGENT_BASE_URL = "http://hermes.test/v1";
    process.env.HERMES_AGENT_TIMEOUT_MS = "25";
    const client = new HermesAgentClient(getConfig());
    const traceBus = new TraceBus();

    vi.stubGlobal("fetch", vi.fn(async (_url: string, init?: RequestInit) => (
      await new Promise<Response>((_resolve, reject) => {
        const signal = init?.signal;
        if (signal && typeof signal.addEventListener === "function") {
          signal.addEventListener("abort", () => reject(new Error("aborted")), { once: true });
        }
      })
    )));

    const requestPromise = client.processRequest({
      runId: "run_timeout_001",
      request: baseRequest({
        requestId: "req_timeout_001"
      }),
      traceBus
    });

    await vi.advanceTimersByTimeAsync(25);
    await requestPromise;

    const response = traceBus.getRun("run_timeout_001")?.response;
    expect(response?.type).toBe("error");
    expect(response?.data).toMatchObject({
      code: "TIMEOUT",
      retryable: true
    });
    expect(response?.trace.at(-1)).toMatchObject({
      event: "failed"
    });
  });

  it("uses the injected TraceBus clock for response timestamps", async () => {
    process.env.HERMES_AGENT_BASE_URL = "http://agent.test/v1";
    const client = new HermesAgentClient(getConfig());
    let nowMs = Date.UTC(2026, 3, 23, 2, 0, 0);
    const traceBus = new TraceBus({
      now: () => nowMs
    });
    traceBus.ensureRun("run_clock_001", "req_clock_001");

    vi.stubGlobal("fetch", vi.fn(async () => {
      nowMs += 2_000;
      return new Response(JSON.stringify(
        chatCompletionEnvelope(JSON.stringify({
          type: "chat",
          data: {
            message: "Clocked response"
          }
        }))
      ), {
        status: 200,
        headers: { "content-type": "application/json" }
      });
    }));

    await client.processRequest({
      runId: "run_clock_001",
      request: baseRequest({
        requestId: "req_clock_001"
      }),
      traceBus
    });

    const response = traceBus.getRun("run_clock_001")?.response;
    expect(response).toMatchObject({
      requestId: "req_clock_001",
      startedAt: "2026-04-23T02:00:00.000Z",
      completedAt: "2026-04-23T02:00:02.000Z"
    });
    expect(response?.trace.at(-1)).toMatchObject({
      event: "completed",
      timestamp: "2026-04-23T02:00:02.000Z"
    });
  });
});
