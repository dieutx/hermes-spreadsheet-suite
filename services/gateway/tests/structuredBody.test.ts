import { describe, expect, it } from "vitest";
import {
  extractSingleJsonObjectText,
  HermesStructuredBodySchema,
  normalizeHermesStructuredBodyInput
} from "../src/hermes/structuredBody.ts";

describe("structured body normalization", () => {
  it.each([
    [
      "chat",
      {
        type: "chat",
        data: {
          message: "The selection is a revenue table.",
          confidence: 0.93,
          followUpSuggestions: ["Summarize revenue by region"],
          selection: { sheet: "Sheet1", range: "A1:F6" },
          highlights: ["Revenue is present in column F"]
        },
        warnings: ["Partial selection payload."],
        skillsUsed: ["spreadsheet-expert"],
        requestId: "drop-me"
      },
      {
        type: "chat",
        data: {
          message: "The selection is a revenue table.",
          confidence: 0.93,
          followUpSuggestions: ["Summarize revenue by region"]
        },
        warnings: [
          {
            code: "MODEL_WARNING",
            message: "Partial selection payload.",
            severity: "warning"
          }
        ],
        skillsUsed: ["spreadsheet-expert"]
      }
    ],
    [
      "formula",
      {
        type: "formula",
        data: {
          intent: "suggest",
          targetCell: "F12",
          formula: "=SUMIFS(F2:F11, D2:D11, \"North\")",
          formulaLanguage: "google_sheets",
          explanation: "This sums revenue for the North region.",
          alternateFormulas: [
            {
              formula: "=SUM(FILTER(F2:F11, D2:D11=\"North\"))",
              explanation: "Alternative formulation.",
              note: "drop-me"
            }
          ],
          confidence: 0.95,
          requiresConfirmation: true,
          highlights: ["drop-me"]
        },
        downstreamProvider: "openai"
      },
      {
        type: "formula",
        data: {
          intent: "suggest",
          targetCell: "F12",
          formula: "=SUMIFS(F2:F11, D2:D11, \"North\")",
          formulaLanguage: "google_sheets",
          explanation: "This sums revenue for the North region.",
          alternateFormulas: [
            {
              formula: "=SUM(FILTER(F2:F11, D2:D11=\"North\"))",
              explanation: "Alternative formulation."
            }
          ],
          confidence: 0.95,
          requiresConfirmation: true
        },
        downstreamProvider: {
          label: "openai"
        }
      }
    ],
    [
      "sheet_update",
      {
        type: "sheet_update",
        data: {
          targetSheet: "Sheet1",
          targetRange: "F12",
          operation: "set_formulas",
          formulas: [["=SUMIFS(F2:F11, D2:D11, \"North\")"]],
          explanation: "Set the North revenue formula.",
          confidence: 0.97,
          requiresConfirmation: true,
          overwriteRisk: "low",
          shape: { rows: 1, columns: 1, previewRows: 1 },
          preview: { anchor: "F12" }
        },
        skillsUsed: ["sheet-writer"],
        debug: true
      },
      {
        type: "sheet_update",
        data: {
          targetSheet: "Sheet1",
          targetRange: "F12",
          operation: "set_formulas",
          formulas: [["=SUMIFS(F2:F11, D2:D11, \"North\")"]],
          explanation: "Set the North revenue formula.",
          confidence: 0.97,
          requiresConfirmation: true,
          overwriteRisk: "low",
          shape: { rows: 1, columns: 1 }
        },
        skillsUsed: ["sheet-writer"]
      }
    ],
    [
      "sheet_import_plan",
      {
        type: "sheet_import_plan",
        data: {
          sourceAttachmentId: "att_100",
          targetSheet: "Imported Table",
          targetRange: "A1:B2",
          headers: ["Date", "Amount"],
          values: [["2026-04-01", 15.5]],
          confidence: 0.89,
          warnings: ["OCR confidence is moderate."],
          requiresConfirmation: true,
          extractionMode: "real",
          shape: { rows: 2, columns: 2, preview: true },
          previewRows: 5
        },
        extra: "drop-me"
      },
      {
        type: "sheet_import_plan",
        data: {
          sourceAttachmentId: "att_100",
          targetSheet: "Imported Table",
          targetRange: "A1:B2",
          headers: ["Date", "Amount"],
          values: [["2026-04-01", 15.5]],
          confidence: 0.89,
          warnings: [
            {
              code: "MODEL_WARNING",
              message: "OCR confidence is moderate.",
              severity: "warning"
            }
          ],
          requiresConfirmation: true,
          extractionMode: "real",
          shape: { rows: 2, columns: 2 }
        }
      }
    ],
    [
      "external_data_plan",
      {
        type: "external_data_plan",
        data: {
          sourceType: "market_data",
          provider: "googlefinance",
          query: {
            symbol: "NASDAQ:GOOG",
            attribute: "price",
            dropMe: true
          },
          targetSheet: "Market Data",
          targetRange: "B2",
          formula: '=GOOGLEFINANCE("NASDAQ:GOOG","price")',
          explanation: "Anchor the latest GOOG price in B2.",
          confidence: 0.92,
          requiresConfirmation: true,
          overwriteRisk: "Anchors a live formula in B2."
        },
        extra: "drop-me"
      },
      {
        type: "external_data_plan",
        data: {
          sourceType: "market_data",
          provider: "googlefinance",
          query: {
            symbol: "NASDAQ:GOOG",
            attribute: "price"
          },
          targetSheet: "Market Data",
          targetRange: "B2",
          formula: '=GOOGLEFINANCE("NASDAQ:GOOG","price")',
          explanation: "Anchor the latest GOOG price in B2.",
          confidence: 0.92,
          requiresConfirmation: true,
          affectedRanges: ["Market Data!B2"],
          overwriteRisk: "low",
          confirmationLevel: "standard"
        }
      }
    ],
    [
      "error",
      {
        type: "error",
        data: {
          code: "INTERNAL_ERROR",
          message: "Something failed.",
          retryable: true,
          userAction: "Retry later.",
          details: "drop-me"
        },
        warnings: [
          {
            code: "SAFE",
            message: "Already structured.",
            severity: "warning"
          }
        ],
        metadata: {}
      },
      {
        type: "error",
        data: {
          code: "INTERNAL_ERROR",
          message: "Something failed.",
          retryable: true,
          userAction: "Retry later."
        },
        warnings: [
          {
            code: "SAFE",
            message: "Already structured.",
            severity: "warning"
          }
        ]
      }
    ],
    [
      "attachment_analysis",
      {
        type: "attachment_analysis",
        data: {
          sourceAttachmentId: "att_101",
          contentKind: "table",
          summary: "The image contains a small pricing table.",
          confidence: 0.8,
          warnings: ["The image is slightly blurred."],
          extractionMode: "real",
          blocks: []
        }
      },
      {
        type: "attachment_analysis",
        data: {
          sourceAttachmentId: "att_101",
          contentKind: "table",
          summary: "The image contains a small pricing table.",
          confidence: 0.8,
          warnings: [
            {
              code: "MODEL_WARNING",
              message: "The image is slightly blurred.",
              severity: "warning"
            }
          ],
          extractionMode: "real"
        }
      }
    ],
    [
      "extracted_table",
      {
        type: "extracted_table",
        data: {
          sourceAttachmentId: "att_102",
          headers: ["Item"],
          rows: [["Cable"]],
          confidence: 0.86,
          warnings: ["Column alignment may be approximate."],
          extractionMode: "real",
          shape: { rows: 1, columns: 1, preview: false },
          highlights: ["drop-me"]
        }
      },
      {
        type: "extracted_table",
        data: {
          sourceAttachmentId: "att_102",
          headers: ["Item"],
          rows: [["Cable"]],
          confidence: 0.86,
          warnings: [
            {
              code: "MODEL_WARNING",
              message: "Column alignment may be approximate.",
              severity: "warning"
            }
          ],
          extractionMode: "real",
          shape: { rows: 1, columns: 1 }
        }
      }
    ],
    [
      "document_summary",
      {
        type: "document_summary",
        data: {
          sourceAttachmentId: "att_103",
          summary: "The attachment describes quarterly revenue results.",
          contentKind: "plain_text",
          keyPoints: ["Revenue increased quarter over quarter."],
          confidence: 0.91,
          warnings: ["Only the first page was fully legible."],
          extractionMode: "real",
          sections: []
        }
      },
      {
        type: "document_summary",
        data: {
          sourceAttachmentId: "att_103",
          summary: "The attachment describes quarterly revenue results.",
          contentKind: "plain_text",
          keyPoints: ["Revenue increased quarter over quarter."],
          confidence: 0.91,
          warnings: [
            {
              code: "MODEL_WARNING",
              message: "Only the first page was fully legible.",
              severity: "warning"
            }
          ],
          extractionMode: "real"
        }
      }
    ]
  ])("normalizes %s bodies with extra unsupported keys before validation", (
    _label,
    rawBody,
    expectedBody
  ) => {
    const normalized = normalizeHermesStructuredBodyInput(rawBody);
    const parsed = HermesStructuredBodySchema.parse(normalized);

    expect(parsed).toEqual(expectedBody);
  });

  it("normalizes range format target and style aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "range_format_update",
      data: {
        sheet: "Sheet1",
        range: "A1:B2",
        background: "#ffeeaa",
        fontColor: "#112233",
        bold: true,
        wrapText: true,
        explanation: "Format the summary header.",
        confidence: 0.89,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "range_format_update",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        affectedRanges: ["Sheet1!A1:B2"],
        confirmationLevel: "standard",
        format: {
          backgroundColor: "#ffeeaa",
          textColor: "#112233",
          bold: true,
          wrapStrategy: "wrap"
        }
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("preserves range format affected range and confirmation metadata during normalization", () => {
    const parsed = HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "range_format_update",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:D8",
        affectedRanges: ["Sheet1!B2:D8"],
        confirmationLevel: "standard",
        explanation: "Apply static formatting.",
        confidence: 0.92,
        requiresConfirmation: true,
        format: {
          backgroundColor: "#eeeeee"
        }
      }
    }));

    expect(parsed.data.affectedRanges).toEqual(["Sheet1!B2:D8"]);
    expect(parsed.data.confirmationLevel).toBe("standard");
  });

  it("normalizes workbook structure action and sheet-name aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "workbook_structure_update",
      data: {
        action: "add",
        name: "Summary",
        position: "end",
        explanation: "Create a summary worksheet.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "workbook_structure_update",
      data: {
        operation: "create_sheet",
        sheetName: "Summary",
        position: "end"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes sheet structure range aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "sheet_structure_update",
      data: {
        sheet: "Sheet1",
        range: "A1:B2",
        action: "merge",
        explanation: "Merge the title cells.",
        confidence: 0.88,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "sheet_structure_update",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        operation: "merge_cells",
        confirmationLevel: "standard",
        affectedRanges: ["Sheet1!A1:B2"]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes sheet import row-matrix aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "sheet_import_plan",
      data: {
        attachmentId: "att_200",
        sheet: "Imported",
        range: "A1:C3",
        rows: [
          ["Date", "Amount", "Region"],
          ["2026-04-01", 15.5, "North"],
          ["2026-04-02", 9, "South"]
        ],
        confidence: 0.88,
        warnings: ["OCR column alignment is approximate."],
        requiresConfirmation: true,
        mode: "real"
      }
    });

    expect(normalized).toEqual({
      type: "sheet_import_plan",
      data: {
        sourceAttachmentId: "att_200",
        targetSheet: "Imported",
        targetRange: "A1:C3",
        headers: ["Date", "Amount", "Region"],
        values: [
          ["2026-04-01", 15.5, "North"],
          ["2026-04-02", 9, "South"]
        ],
        confidence: 0.88,
        warnings: [
          {
            code: "MODEL_WARNING",
            message: "OCR column alignment is approximate.",
            severity: "warning"
          }
        ],
        requiresConfirmation: true,
        extractionMode: "real",
        shape: { rows: 3, columns: 3 }
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("synthesizes external data formulas from provider fields before validation", () => {
    const marketData = normalizeHermesStructuredBodyInput({
      type: "external_data_plan",
      data: {
        provider: "GOOGLEFINANCE",
        query: {
          symbol: "NASDAQ:GOOG",
          attribute: "price",
          startDate: "2026-01-01",
          endDate: "2026-04-01",
          interval: "DAILY"
        },
        targetSheet: "Market Data",
        targetRange: "B2",
        explanation: "Anchor GOOG history in B2.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    });

    expect(marketData).toMatchObject({
      type: "external_data_plan",
      data: {
        sourceType: "market_data",
        provider: "googlefinance",
        formula: '=GOOGLEFINANCE("NASDAQ:GOOG","price","2026-01-01","2026-04-01","DAILY")',
        affectedRanges: ["Market Data!B2"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(marketData)).not.toThrow();

    const webImport = normalizeHermesStructuredBodyInput({
      type: "external_data_plan",
      data: {
        provider: "IMPORTHTML",
        sourceUrl: "https://example.com/markets",
        selectorType: "table",
        selector: 1,
        targetSheet: "Imports",
        targetRange: "A1",
        explanation: "Import the first market table.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(webImport).toMatchObject({
      type: "external_data_plan",
      data: {
        sourceType: "web_table_import",
        provider: "importhtml",
        formula: '=IMPORTHTML("https://example.com/markets","table",1)',
        affectedRanges: ["Imports!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(webImport)).not.toThrow();
  });

  it("defaults provider-specific web import selector types before validation", () => {
    const xmlImport = normalizeHermesStructuredBodyInput({
      type: "external_data_plan",
      data: {
        provider: "IMPORTXML",
        sourceUrl: "https://example.com/feed",
        selector: "//item/title",
        targetSheet: "Imports",
        targetRange: "A1",
        explanation: "Import feed titles.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(xmlImport).toMatchObject({
      type: "external_data_plan",
      data: {
        sourceType: "web_table_import",
        provider: "importxml",
        selectorType: "xpath",
        formula: '=IMPORTXML("https://example.com/feed","//item/title")'
      }
    });
    expect(() => HermesStructuredBodySchema.parse(xmlImport)).not.toThrow();

    const dataImport = normalizeHermesStructuredBodyInput({
      type: "external_data_plan",
      data: {
        provider: "IMPORTDATA",
        sourceUrl: "https://example.com/data.csv",
        targetSheet: "Imports",
        targetRange: "C1",
        explanation: "Import a CSV feed.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(dataImport).toMatchObject({
      type: "external_data_plan",
      data: {
        sourceType: "web_table_import",
        provider: "importdata",
        selectorType: "direct",
        formula: '=IMPORTDATA("https://example.com/data.csv")'
      }
    });
    expect(() => HermesStructuredBodySchema.parse(dataImport)).not.toThrow();
  });

  it("normalizes external data enum aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "external_data_plan",
      data: {
        sourceType: "WEB_TABLE_IMPORT",
        provider: "IMPORTXML",
        sourceUrl: "https://example.com/feed.xml",
        selectorType: "XPATH",
        selector: "//item/title",
        targetSheet: "Imports",
        targetRange: "A1",
        explanation: "Import feed titles.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "external_data_plan",
      data: {
        sourceType: "web_table_import",
        provider: "importxml",
        selectorType: "xpath",
        formula: '=IMPORTXML("https://example.com/feed.xml","//item/title")'
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("coerces importhtml selector indexes before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "external_data_plan",
      data: {
        provider: "IMPORTHTML",
        sourceUrl: "https://example.com/markets",
        selectorType: "table",
        selector: "1",
        targetSheet: "Imports",
        targetRange: "A1",
        explanation: "Import the first market table.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "external_data_plan",
      data: {
        sourceType: "web_table_import",
        provider: "importhtml",
        selector: 1,
        formula: '=IMPORTHTML("https://example.com/markets","table",1)'
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes sheet update aliases and infers shape before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "sheet_update",
      data: {
        sheet: "Sheet1",
        range: "A1:B2",
        operation: "set_values",
        data: [
          ["Region", "Revenue"],
          ["North", 1200]
        ],
        explanation: "Write the summarized revenue table.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "sheet_update",
      data: {
        targetSheet: "Sheet1",
        targetRange: "A1:B2",
        operation: "replace_range",
        values: [
          ["Region", "Revenue"],
          ["North", 1200]
        ],
        shape: { rows: 2, columns: 2 }
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes chart legendPosition none to the contract-safe hidden alias", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "chart_plan",
      data: {
        sourceSheet: "Crypto Prices",
        sourceRange: "A1:B11",
        targetSheet: "Crypto Prices",
        targetRange: "D1",
        chartType: "column",
        categoryField: "Asset",
        series: [{ field: "Price (USD)", label: "Price (USD)" }],
        legendPosition: "none",
        explanation: "Compare prices by asset.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Crypto Prices!A1:B11", "Crypto Prices!D1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(normalized).toEqual({
      type: "chart_plan",
      data: {
        sourceSheet: "Crypto Prices",
        sourceRange: "A1:B11",
        targetSheet: "Crypto Prices",
        targetRange: "D1",
        chartType: "column",
        categoryField: "Asset",
        series: [{ field: "Price (USD)", label: "Price (USD)" }],
        legendPosition: "hidden",
        explanation: "Compare prices by asset.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Crypto Prices!A1:B11", "Crypto Prices!D1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes chart axis title aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "chart_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "A1",
        chartType: "line",
        categoryField: "Month",
        series: [{ field: "Revenue", label: "Revenue" }],
        title: "Revenue",
        xAxisTitle: "Month",
        yAxisTitle: "USD",
        explanation: "Chart monthly revenue.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C20", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(normalized).toMatchObject({
      type: "chart_plan",
      data: {
        horizontalAxisTitle: "Month",
        verticalAxisTitle: "USD"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes chart creation aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "chart_plan",
      data: {
        dataRange: "Sales!A1:C20",
        targetSheet: "Sales Chart",
        insertAt: "D2",
        chartType: "column",
        categoryField: "Month",
        series: [{ field: "Revenue", label: "Revenue" }],
        chartTitle: "Quarterly Revenue",
        explanation: "Chart quarterly revenue.",
        confidence: 0.92,
        requiresConfirmation: true,
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(normalized).toMatchObject({
      type: "chart_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:C20",
        targetSheet: "Sales Chart",
        targetRange: "D2",
        title: "Quarterly Revenue",
        affectedRanges: ["Sales!A1:C20", "Sales Chart!D2"]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes pivot rows columns and values aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "pivot_table_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:F50",
        targetSheet: "Pivot",
        targetRange: "A1",
        rows: ["Region"],
        columns: ["Quarter"],
        values: ["Revenue", "Deals"],
        aggregation: "sum",
        explanation: "Build a sales pivot.",
        confidence: 0.91,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(normalized).toMatchObject({
      type: "pivot_table_plan",
      data: {
        rowGroups: ["Region"],
        columnGroups: ["Quarter"],
        valueAggregations: [
          { field: "Revenue", aggregation: "sum" },
          { field: "Deals", aggregation: "sum" }
        ]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes range sort aliases and affected ranges before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "range_sort_plan",
      data: {
        sheet: "Sales",
        range: "A1:D50",
        header: true,
        sortKeys: [
          { column: "Revenue", order: "descending" },
          { field: "Region", order: "ascending" }
        ],
        explanation: "Sort sales by revenue and region.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "range_sort_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:D50",
        hasHeader: true,
        keys: [
          { columnRef: "Revenue", direction: "desc" },
          { columnRef: "Region", direction: "asc" }
        ],
        affectedRanges: ["Sales!A1:D50"]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("fails closed for unsupported range sort modes instead of dropping them", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "range_sort_plan",
      data: {
        sheet: "Sales",
        range: "A1:D50",
        header: true,
        sortKeys: [
          { column: "Revenue", order: "descending", sortOn: "cellColor" }
        ],
        explanation: "Sort sales by revenue cell color.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(() => HermesStructuredBodySchema.parse(normalized)).toThrow();
  });

  it("normalizes analysis report source and output aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "analysis_report_plan",
      data: {
        dataRange: "Support!A1:H50",
        output: "sheet",
        outputRange: "Summary!A1",
        reportSections: ["sla risk summary", "next actions"],
        explanation: "Create a materialized support analysis report.",
        confidence: 0.88
      }
    });

    expect(normalized).toMatchObject({
      type: "analysis_report_plan",
      data: {
        sourceSheet: "Support",
        sourceRange: "A1:H50",
        targetSheet: "Summary",
        targetRange: "A1:D6",
        outputMode: "materialize_report",
        requiresConfirmation: true,
        sections: [
          {
            type: "anomalies",
            title: "Sla Risk Summary",
            sourceRanges: ["Support!A1:H50"]
          },
          {
            type: "next_actions",
            title: "Next Actions",
            sourceRanges: ["Support!A1:H50"]
          }
        ],
        affectedRanges: ["Support!A1:H50", "Summary!A1:D6"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes conditional format aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "conditional_format_plan",
      data: {
        sheet: "Support",
        range: "G2:G50",
        mode: "replace",
        rule: {
          type: "formula",
          customFormula: '=$G2="Breached"',
          format: {
            background: "#FDECEC",
            fontColor: "#9C0006",
            bold: true
          }
        },
        explanation: "Highlight breached SLA rows.",
        confidence: 0.92,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "conditional_format_plan",
      data: {
        targetSheet: "Support",
        targetRange: "G2:G50",
        managementMode: "replace_all_on_target",
        replacesExistingRules: true,
        ruleType: "custom_formula",
        formula: '=$G2="Breached"',
        style: {
          backgroundColor: "#FDECEC",
          textColor: "#9C0006",
          bold: true
        },
        affectedRanges: ["Support!G2:G50"]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes data validation prompt and error-message aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "list",
        values: ["Open", "Closed"],
        showDropdown: true,
        allowBlank: false,
        invalidDataBehavior: "reject",
        promptTitle: "Status",
        promptMessage: "Choose a valid status.",
        errorAlertTitle: "Invalid status",
        errorAlertMessage: "Pick a value from the dropdown.",
        explanation: "Restrict status values.",
        confidence: 0.95,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "data_validation_plan",
      data: {
        inputTitle: "Status",
        inputMessage: "Choose a valid status.",
        errorTitle: "Invalid status",
        errorMessage: "Pick a value from the dropdown."
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes comparison aliases for validation conditional formatting and pivot filters", () => {
    const validation = normalizeHermesStructuredBodyInput({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        ruleType: "number",
        comparator: "greaterThanOrEqual",
        value: 100,
        allowBlank: false,
        invalidDataBehavior: "reject",
        explanation: "Require values greater than or equal to 100.",
        confidence: 0.95,
        requiresConfirmation: true
      }
    });

    expect(validation).toMatchObject({
      type: "data_validation_plan",
      data: {
        ruleType: "decimal",
        comparator: "greater_than_or_equal_to"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(validation)).not.toThrow();

    const conditionalFormat = normalizeHermesStructuredBodyInput({
      type: "conditional_format_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "C2:C20",
        managementMode: "add",
        ruleType: "number_compare",
        comparator: "lessThanOrEqual",
        value: 10,
        style: {
          backgroundColor: "#FDECEC"
        },
        explanation: "Highlight values at or below 10.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(conditionalFormat).toMatchObject({
      type: "conditional_format_plan",
      data: {
        comparator: "less_than_or_equal_to"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(conditionalFormat)).not.toThrow();

    const pivot = normalizeHermesStructuredBodyInput({
      type: "pivot_table_plan",
      data: {
        sourceSheet: "Sales",
        sourceRange: "A1:E50",
        targetSheet: "Pivot",
        targetRange: "A1",
        rowGroups: ["Region"],
        valueAggregations: [
          {
            field: "Revenue",
            aggregation: "sum"
          }
        ],
        filters: [
          {
            field: "Revenue",
            operator: "notEquals",
            value: 0
          }
        ],
        explanation: "Summarize non-zero revenue by region.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:E50", "Pivot!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });

    expect(pivot).toMatchObject({
      type: "pivot_table_plan",
      data: {
        filters: [
          {
            field: "Revenue",
            operator: "not_equal_to",
            value: 0
          }
        ]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(pivot)).not.toThrow();
  });

  it("normalizes range transfer aliases and required defaults before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "range_transfer_plan",
      data: {
        sourceSheet: "Raw",
        sourceRange: "A1:B3",
        targetSheet: "Archive",
        targetRange: "D1:E3",
        transferOperation: "copy",
        pasteMode: "values",
        explanation: "Copy the raw values into the archive.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "range_transfer_plan",
      data: {
        operation: "copy",
        transpose: false,
        affectedRanges: ["Raw!A1:B3", "Archive!D1:E3"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes range filter condition aliases and defaults before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "range_filter_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        hasHeader: true,
        conditions: [
          { field: "Status", operator: "equal_to", value: "Open" },
          { column: "Revenue", operator: "greater_than_or_equal_to", value: 1000 },
          { column: "Units", operator: "top_n", value: "5" }
        ],
        explanation: "Filter to open high-value rows.",
        confidence: 0.9,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "range_filter_plan",
      data: {
        combiner: "and",
        clearExistingFilters: true,
        conditions: [
          { columnRef: "Status", operator: "equals", value: "Open" },
          { columnRef: "Revenue", operator: "greaterThanOrEqual", value: 1000 },
          { columnRef: "Units", operator: "topN", value: 5 }
        ],
        affectedRanges: ["Sales!A1:F50"]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes dropdown validation aliases and defaults before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "data_validation_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "C2:C20",
        ruleType: "dropdown",
        options: ["Open", "Closed", "Paused"],
        promptMessage: "Choose a status.",
        explanation: "Restrict status values.",
        confidence: 0.94,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "data_validation_plan",
      data: {
        ruleType: "list",
        values: ["Open", "Closed", "Paused"],
        showDropdown: true,
        allowBlank: true,
        invalidDataBehavior: "reject",
        inputMessage: "Choose a status.",
        affectedRanges: ["Sheet1!C2:C20"]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes named range define aliases and qualified references before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "named_range_update",
      data: {
        operation: "define",
        rangeName: "SalesData",
        refersTo: "Sales!A1:D20",
        explanation: "Define a named range for the sales table.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "named_range_update",
      data: {
        operation: "create",
        scope: "workbook",
        name: "SalesData",
        targetSheet: "Sales",
        targetRange: "A1:D20",
        affectedRanges: ["Sales!A1:D20"]
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes table plans before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "table_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        styleName: "TableStyleMedium2",
        showBandedRows: true,
        showBandedColumns: false,
        showFilterButton: true,
        showTotalsRow: false,
        explanation: "Convert the selected sales range into a native table.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "This creates table metadata over the selected cells.",
        confirmationLevel: "standard",
        previewOnly: "drop-me"
      },
      extra: "drop-me"
    });

    expect(normalized).toEqual({
      type: "table_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        styleName: "TableStyleMedium2",
        showBandedRows: true,
        showBandedColumns: false,
        showFilterButton: true,
        showTotalsRow: false,
        explanation: "Convert the selected sales range into a native table.",
        confidence: 0.92,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes table plan aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "table_plan",
      data: {
        sheet: "Sales",
        range: "A1:F50",
        tableName: "SalesTable",
        hasHeader: true,
        tableStyle: "TableStyleMedium2",
        bandedRows: true,
        bandedColumns: false,
        filterButton: true,
        totalsRow: true,
        explanation: "Create a native table for the sales range.",
        confidence: 0.91,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "table_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "A1:F50",
        name: "SalesTable",
        hasHeaders: true,
        styleName: "TableStyleMedium2",
        showBandedRows: true,
        showBandedColumns: false,
        showFilterButton: true,
        showTotalsRow: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("normalizes cleanup split-column aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "data_cleanup_plan",
      data: {
        sheet: "Sales",
        range: "C2:D20",
        action: "split",
        column: "C",
        separator: ",",
        startColumn: "D",
        explanation: "Split customer name values into separate columns.",
        confidence: 0.87,
        requiresConfirmation: true
      }
    });

    expect(normalized).toMatchObject({
      type: "data_cleanup_plan",
      data: {
        targetSheet: "Sales",
        targetRange: "C2:D20",
        operation: "split_column",
        sourceColumn: "C",
        delimiter: ",",
        targetStartColumn: "D",
        affectedRanges: ["Sales!C2:D20"],
        overwriteRisk: "high",
        confirmationLevel: "destructive"
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it("coerces formula alternateFormulas string entries into contract-valid objects", () => {
    const rawBody = {
      type: "formula",
      data: {
        intent: "suggest",
        formula: "=SUMIF(B2:B7,B8,F2:F7)",
        formulaLanguage: "google_sheets",
        explanation: "Sum matching values for the category in B8.",
        alternateFormulas: [
          "=SUMIFS(F2:F7,B2:B7,B8)",
          "=SUMIF($B$2:$B$7,B8,$F$2:$F$7)"
        ],
        confidence: 0.74
      }
    };

    const normalized = normalizeHermesStructuredBodyInput(rawBody);
    const parsed = HermesStructuredBodySchema.parse(normalized);

    expect(parsed).toEqual({
      type: "formula",
      data: {
        intent: "suggest",
        formula: "=SUMIF(B2:B7,B8,F2:F7)",
        formulaLanguage: "google_sheets",
        explanation: "Sum matching values for the category in B8.",
        alternateFormulas: [
          {
            formula: "=SUMIFS(F2:F7,B2:B7,B8)",
            explanation: "Alternative formulation."
          },
          {
            formula: "=SUMIF($B$2:$B$7,B8,$F$2:$F$7)",
            explanation: "Alternative formulation."
          }
        ],
        confidence: 0.74
      }
    });
  });

  it("maps model-specific missing-context errors into contract-safe spreadsheet context errors", () => {
    const rawBody = {
      type: "error",
      data: {
        code: "MISSING_REQUIRED_CONTEXT",
        message: "Tell me which cell and condition to use.",
        retryable: true,
        userAction: "Specify the target cell and the SUMIF condition."
      }
    };

    const normalized = normalizeHermesStructuredBodyInput(rawBody);
    const parsed = HermesStructuredBodySchema.parse(normalized);

    expect(parsed).toEqual({
      type: "error",
      data: {
        code: "SPREADSHEET_CONTEXT_MISSING",
        message: "Tell me which cell and condition to use.",
        retryable: true,
        userAction: "Specify the target cell and the SUMIF condition."
      }
    });
  });

  it("maps descriptive overwriteRisk text into a contract-safe risk level for sheet updates", () => {
    const rawBody = {
      type: "sheet_update",
      data: {
        targetSheet: "Sheet1",
        targetRange: "H11",
        operation: "set_formulas",
        formulas: [["=SUMIF(B:B,\"north\",F:F)"]],
        explanation: "Write the corrected formula into H11.",
        confidence: 0.93,
        requiresConfirmation: true,
        shape: { rows: 1, columns: 1 },
        overwriteRisk: "Replaces the existing formula in H11."
      }
    };

    const normalized = normalizeHermesStructuredBodyInput(rawBody);
    const parsed = HermesStructuredBodySchema.parse(normalized);

    expect(parsed).toEqual({
      type: "sheet_update",
      data: {
        targetSheet: "Sheet1",
        targetRange: "H11",
        operation: "set_formulas",
        formulas: [["=SUMIF(B:B,\"north\",F:F)"]],
        explanation: "Write the corrected formula into H11.",
        confidence: 0.93,
        requiresConfirmation: true,
        shape: { rows: 1, columns: 1 },
        overwriteRisk: "low"
      }
    });
  });

  it("normalizes composite_plan bodies with nested executable steps before validation", () => {
    const rawBody = {
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
              keys: [
                { columnRef: "Revenue", direction: "desc", unexpected: "drop-me" }
              ],
              explanation: "Sort by revenue.",
              confidence: 0.91,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50"],
              unexpected: "drop-me"
            },
            unexpected: "drop-me"
          }
        ],
        explanation: "Run the workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false,
        unexpected: "drop-me"
      },
      extra: "drop-me"
    };

    const normalized = normalizeHermesStructuredBodyInput(rawBody);
    const parsed = HermesStructuredBodySchema.parse(normalized);

    expect(parsed).toEqual({
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
              keys: [
                { columnRef: "Revenue", direction: "desc" }
              ],
              explanation: "Sort by revenue.",
              confidence: 0.91,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50"]
            }
          }
        ],
        explanation: "Run the workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });
  });

  it("preserves explicit reversible composite plans when every child step can be snapshot-backed", () => {
    const parsed = HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "write_headers",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:B1",
              operation: "replace_range",
              values: [["Region", "Revenue"]],
              explanation: "Write report headers.",
              confidence: 0.92,
              requiresConfirmation: true,
              shape: { rows: 1, columns: 2 }
            }
          }
        ],
        explanation: "Write the report headers.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:B1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    }));

    expect(parsed.data.reversible).toBe(true);
  });

  it("preserves reversible analytic composite plans during normalization", () => {
    const parsed = HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_pivot",
            dependsOn: [],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:C20",
              targetSheet: "Sales Pivot",
              targetRange: "A1",
              rowGroups: ["Region"],
              valueAggregations: [
                { field: "Revenue", aggregation: "sum" }
              ],
              explanation: "Create a pivot table.",
              confidence: 0.92,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:C20", "Sales Pivot!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "step_chart",
            dependsOn: ["step_pivot"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales Pivot",
              sourceRange: "A1:B5",
              targetSheet: "Sales Chart",
              targetRange: "A1",
              chartType: "line",
              categoryField: "Region",
              series: [{ field: "Revenue", label: "Revenue" }],
              explanation: "Create a chart from the pivot table.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales Pivot!A1:B5", "Sales Chart!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Create a pivot table and chart.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales Pivot!A1", "Sales Chart!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    }));

    expect(parsed.data.reversible).toBe(true);
  });

  it("preserves reversible composite plans with table steps during normalization", () => {
    const parsed = HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_table",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:F50",
              name: "SalesTable",
              hasHeaders: true,
              showBandedRows: true,
              showBandedColumns: false,
              showFilterButton: true,
              showTotalsRow: false,
              explanation: "Convert the sales range into a table.",
              confidence: 0.92,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ],
        explanation: "Create a table.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: true,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    }));

    expect(parsed.data.reversible).toBe(true);
  });

  it("normalizes composite action aliases before validation", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "composite_plan",
      data: {
        actions: [
          {
            id: "sort_sales",
            type: "range_sort_plan",
            data: {
              targetSheet: "Sales",
              targetRange: "A1:C10",
              hasHeader: true,
              keys: [{ columnRef: "Revenue", direction: "desc" }],
              explanation: "Sort sales by revenue.",
              confidence: 0.91,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:C10"]
            }
          }
        ],
        explanation: "Run a one-step workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C10"],
        overwriteRisk: "low"
      }
    });

    expect(normalized).toEqual({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "sort_sales",
            dependsOn: [],
            continueOnError: false,
            plan: {
              targetSheet: "Sales",
              targetRange: "A1:C10",
              hasHeader: true,
              keys: [{ columnRef: "Revenue", direction: "desc" }],
              explanation: "Sort sales by revenue.",
              confidence: 0.91,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:C10"]
            }
          }
        ],
        explanation: "Run a one-step workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:C10"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });
    expect(() => HermesStructuredBodySchema.parse(normalized)).not.toThrow();
  });

  it.each([
    [
      "chat missing message",
      {
        type: "chat",
        data: {
          selection: { sheet: "Sheet1" }
        }
      },
      ["data", "message"]
    ],
    [
      "formula missing formula",
      {
        type: "formula",
        data: {
          intent: "suggest",
          formulaLanguage: "excel",
          explanation: "Missing formula should fail.",
          confidence: 0.8
        }
      },
      ["data", "formula"]
    ],
    [
      "sheet_update missing shape",
      {
        type: "sheet_update",
        data: {
          targetSheet: "Sheet1",
          targetRange: "A1",
          operation: "append_rows",
          values: [[1]],
          explanation: "Shape is required.",
          confidence: 0.8,
          requiresConfirmation: true
        }
      },
      ["data", "shape"]
    ]
  ])("rejects %s after normalization", (_label, rawBody, expectedPath) => {
    const normalized = normalizeHermesStructuredBodyInput(rawBody);
    const parsed = HermesStructuredBodySchema.safeParse(normalized);

    expect(parsed.success).toBe(false);
    if (parsed.success) {
      return;
    }

    expect(parsed.error.issues.some((issue) =>
      JSON.stringify(issue.path) === JSON.stringify(expectedPath)
    )).toBe(true);
  });

  it("rejects prose mixed with JSON", () => {
    expect(extractSingleJsonObjectText(
      "Here is the payload:\n{\"type\":\"chat\",\"data\":{\"message\":\"hello\"}}"
    )).toBeNull();
  });

  it("preserves raw composite step families without misclassifying them as sheet structure updates", () => {
    const normalized = normalizeHermesStructuredBodyInput({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "analysis",
            dependsOn: [],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:F50",
              outputMode: "materialize_report",
              targetSheet: "Report",
              targetRange: "A1",
              sections: [
                {
                  type: "summary_stats",
                  title: "Summary",
                  summary: "Revenue summary.",
                  sourceRanges: ["Sales!A1:F50"]
                }
              ],
              explanation: "Build a report.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50", "Report!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "pivot",
            dependsOn: ["analysis"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:F50",
              targetSheet: "Pivot",
              targetRange: "A1",
              rowGroups: ["Region"],
              valueAggregations: [
                {
                  field: "Revenue",
                  aggregation: "sum"
                }
              ],
              explanation: "Build a pivot.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50", "Pivot!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "chart",
            dependsOn: ["pivot"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A1:F50",
              targetSheet: "Chart",
              targetRange: "A1",
              chartType: "column",
              series: [
                {
                  field: "Revenue",
                  label: "Revenue"
                }
              ],
              categoryField: "Month",
              title: "Revenue by Month",
              explanation: "Build a chart.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A1:F50", "Chart!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "transfer",
            dependsOn: ["chart"],
            continueOnError: false,
            plan: {
              sourceSheet: "Sales",
              sourceRange: "A2:B10",
              targetSheet: "Archive",
              targetRange: "D5:E13",
              operation: "copy",
              pasteMode: "values",
              transpose: false,
              explanation: "Copy values.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Sales!A2:B10", "Archive!D5:E13"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          },
          {
            stepId: "cleanup",
            dependsOn: ["transfer"],
            continueOnError: false,
            plan: {
              targetSheet: "Archive",
              targetRange: "D5:E13",
              operation: "split_column",
              sourceColumn: "D",
              targetStartColumn: "D",
              delimiter: "-",
              explanation: "Split codes.",
              confidence: 0.9,
              requiresConfirmation: true,
              affectedRanges: ["Archive!D5:E13"],
              overwriteRisk: "high",
              confirmationLevel: "destructive"
            }
          }
        ],
        explanation: "Run the workflow.",
        confidence: 0.9,
        requiresConfirmation: true,
        affectedRanges: ["Sales!A1:F50", "Report!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    });

    const parsed = HermesStructuredBodySchema.parse(normalized);
    expect(parsed.type).toBe("composite_plan");
    expect(parsed.data.reversible).toBe(false);
    expect(parsed.data.steps[0].plan.outputMode).toBe("materialize_report");
    expect(parsed.data.steps[1].plan.rowGroups).toEqual(["Region"]);
    expect(parsed.data.steps[2].plan.series[0]).toMatchObject({ field: "Revenue" });
    expect(parsed.data.steps[3].plan.pasteMode).toBe("values");
    expect(parsed.data.steps[4].plan.delimiter).toBe("-");
  });

  it("preserves all supported conditional-format style fields during normalization", () => {
    const parsed = HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "conditional_format_plan",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:B20",
        explanation: "Highlight revenue values.",
        confidence: 0.94,
        requiresConfirmation: true,
        affectedRanges: ["Sheet1!B2:B20"],
        replacesExistingRules: false,
        managementMode: "add",
        ruleType: "custom_formula",
        formula: '=B2>1000',
        style: {
          underline: true,
          strikethrough: false,
          numberFormat: "#,##0.00",
          textColor: "#111111"
        }
      }
    }));

    expect(parsed.data.style).toEqual({
      underline: true,
      strikethrough: false,
      numberFormat: "#,##0.00",
      textColor: "#111111"
    });
  });

  it("preserves all supported static range format fields during normalization", () => {
    const parsed = HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "range_format_update",
      data: {
        targetSheet: "Sheet1",
        targetRange: "B2:D8",
        explanation: "Apply detailed static formatting.",
        confidence: 0.92,
        requiresConfirmation: true,
        format: {
          fontFamily: "Aptos",
          fontSize: 12,
          underline: true,
          strikethrough: false,
          border: {
            outer: {
              style: "solid",
              color: "#222222"
            },
            innerHorizontal: {
              style: "dotted"
            }
          }
        }
      }
    }));

    expect(parsed.data.format).toEqual({
      fontFamily: "Aptos",
      fontSize: 12,
      underline: true,
      strikethrough: false,
      border: {
        outer: {
          style: "solid",
          color: "#222222"
        },
        innerHorizontal: {
          style: "dotted"
        }
      }
    });
  });

  it("normalizes raw external_data_plan steps inside composite plans before validation", () => {
    const parsed = HermesStructuredBodySchema.parse(normalizeHermesStructuredBodyInput({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_import_market",
            dependsOn: [],
            continueOnError: false,
            plan: {
              sourceType: "web_table_import",
              provider: "importhtml",
              sourceUrl: "https://example.com/markets",
              selectorType: "table",
              selector: 1,
              targetSheet: "Imported Data",
              targetRange: "A1",
              formula: '=IMPORTHTML("https://example.com/markets","table",1)',
              explanation: "Import the first public table.",
              confidence: 0.89,
              requiresConfirmation: true
            }
          }
        ],
        explanation: "Import one external table.",
        confidence: 0.89,
        requiresConfirmation: true,
        affectedRanges: ["Imported Data!A1"],
        overwriteRisk: "low",
        confirmationLevel: "standard",
        reversible: false,
        dryRunRecommended: true,
        dryRunRequired: false
      }
    }));

    expect(parsed).toMatchObject({
      type: "composite_plan",
      data: {
        steps: [
          {
            stepId: "step_import_market",
            plan: {
              sourceType: "web_table_import",
              provider: "importhtml",
              sourceUrl: "https://example.com/markets",
              selectorType: "table",
              selector: 1,
              targetSheet: "Imported Data",
              targetRange: "A1",
              formula: '=IMPORTHTML("https://example.com/markets","table",1)',
              affectedRanges: ["Imported Data!A1"],
              overwriteRisk: "low",
              confirmationLevel: "standard"
            }
          }
        ]
      }
    });
  });
});
