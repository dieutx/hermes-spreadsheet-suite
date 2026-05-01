import { describe, expect, it } from "vitest";
import {
  evaluateCapabilityFixtureDirectory,
  evaluateCapabilityFixturePack,
  type CapabilityFixturePack
} from "../../../scripts/eval_fixtures.ts";

describe("capability fixture eval runner", () => {
  it("normalizes and validates structured-body fixtures with path expectations", () => {
    const pack: CapabilityFixturePack = {
      version: 1,
      fixtures: [
        {
          id: "gateway.sheet-import.rows-normalization",
          host: "gateway",
          family: "sheet_import_plan",
          structuredBody: {
            type: "sheet_import_plan",
            data: {
              sourceAttachmentId: "att_fixture_001",
              targetSheet: "Import",
              targetRange: "A1:B2",
              rows: [
                ["Name", "Revenue"],
                ["Cable", 15.5]
              ],
              confidence: 0.91,
              requiresConfirmation: true,
              extractionMode: "demo"
            }
          },
          expect: {
            valid: true,
            paths: {
              "type": "sheet_import_plan",
              "data.headers.0": "Name",
              "data.values.0.1": 15.5,
              "data.shape.rows": 2,
              "data.shape.columns": 2
            }
          }
        }
      ]
    };

    const result = evaluateCapabilityFixturePack(pack, { source: "inline" });

    expect(result).toMatchObject({
      source: "inline",
      total: 1,
      passed: 1,
      failed: 0
    });
    expect(result.results[0]).toMatchObject({
      id: "gateway.sheet-import.rows-normalization",
      status: "passed"
    });
  });

  it("treats reviewer-safe unavailable extracted rows as an expected negative fixture", () => {
    const pack: CapabilityFixturePack = {
      version: 1,
      fixtures: [
        {
          id: "gateway.reviewer-safe.unavailable-extracted-table",
          host: "gateway",
          family: "extracted_table",
          structuredBody: {
            type: "extracted_table",
            data: {
              sourceAttachmentId: "att_fixture_002",
              headers: ["Name"],
              rows: [["Cable"]],
              confidence: 0.4,
              extractionMode: "unavailable",
              shape: {
                rows: 1,
                columns: 1
              }
            }
          },
          expect: {
            valid: false,
            errorIncludes: "reviewer-safe unavailable mode"
          }
        }
      ]
    };

    const result = evaluateCapabilityFixturePack(pack, { source: "inline" });

    expect(result).toMatchObject({
      total: 1,
      passed: 1,
      failed: 0
    });
    expect(result.results[0]).toMatchObject({
      status: "passed",
      actualValid: false
    });
  });

  it("validates gateway request fixtures with the contract schema", () => {
    const pack: CapabilityFixturePack = {
      version: 1,
      fixtures: [
        {
          id: "gateway.request.selection-explain",
          host: "gateway",
          family: "request",
          request: {
            schemaVersion: "1.0.0",
            requestId: "req_fixture_001",
            source: {
              channel: "excel_windows",
              clientVersion: "0.1.0",
              sessionId: "sess_fixture_001"
            },
            host: {
              platform: "excel_windows",
              workbookTitle: "Revenue Demo",
              workbookId: "workbook_fixture_001",
              activeSheet: "Sales",
              selectedRange: "A1:B1"
            },
            userMessage: "Explain this selected range",
            conversation: [
              { role: "user", content: "Explain this selected range" }
            ],
            context: {
              selection: {
                range: "A1:B1",
                headers: ["Name", "Revenue"],
                values: [["Cable", 15.5]]
              }
            },
            capabilities: {
              canRenderTrace: true,
              canRenderStructuredPreview: true,
              canConfirmWriteBack: true,
              supportsImageInputs: false,
              supportsWriteBackExecution: true
            },
            reviewer: {
              reviewerSafeMode: true,
              forceExtractionMode: "demo"
            },
            confirmation: {
              state: "none"
            }
          },
          expect: {
            valid: true,
            paths: {
              "host.platform": "excel_windows",
              "context.selection.range": "A1:B1"
            }
          }
        }
      ]
    };

    const result = evaluateCapabilityFixturePack(pack, { source: "inline" });

    expect(result).toMatchObject({
      total: 1,
      passed: 1,
      failed: 0
    });
  });

  it("reports fixture directory sources without exposing the local checkout path", () => {
    const result = evaluateCapabilityFixtureDirectory(`${process.cwd()}/fixtures/capability-eval`);

    expect(result.source).toBe("fixtures/capability-eval");
    expect(JSON.stringify(result)).not.toContain(process.cwd());
  });
});
