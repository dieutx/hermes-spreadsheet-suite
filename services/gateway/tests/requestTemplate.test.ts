import { describe, expect, it } from "vitest";
import type { HermesRequest } from "@hermes/contracts";
import { buildHermesSpreadsheetRequestPrompt } from "../src/hermes/requestTemplate.ts";

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

describe("Hermes spreadsheet request prompt", () => {
  it("includes type-specific contract guidance beyond chat", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest());

    expect(prompt).toContain("formulaLanguage");
    expect(prompt).toContain("requiresConfirmation must be true");
    expect(prompt).toContain("targetRange must match");
    expect(prompt).toContain("supportsNoteWrites");
  });

  it("includes a host capability matrix for Google Sheets planning", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest());

    expect(prompt).toContain("Host capability matrix for google_sheets");
    expect(prompt).toContain("pivot_table_plan: limited");
    expect(prompt).toContain("Supported pivot sorting: group_field on an existing row or column group");
    expect(prompt).toContain("Supported pivot filters use existing pivot fields");
    expect(prompt).toContain("between, or not_between");
    expect(prompt).toContain("Series labels may use explicit custom legend text when provided.");
    expect(prompt).toContain("Supported legend positions: bottom, left, right, top, hidden.");
    expect(prompt).toContain("If you would naturally say none, emit hidden.");
    expect(prompt).toContain("range_transfer_plan: limited");
    expect(prompt).toContain("Supported pasteMode values: values, formulas, formats.");
    expect(prompt).toContain("normalize_case only supports upper, lower, title, and sentence.");
    expect(prompt).toContain("named_range_update: limited");
    expect(prompt).toContain("external_data_plan: limited");
    expect(prompt).toContain("market_data/googlefinance");
    expect(prompt).toContain("web_table_import/importhtml, importxml, or importdata");
  });

  it("treats missing note-write capability as unsupported", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest());

    expect(prompt).toContain("note_writes: unsupported");
    expect(prompt).toContain("Do not propose sheet_update operations that depend on notes.");
  });

  it("includes a host capability matrix for Excel planning", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      source: {
        ...baseRequest().source,
        channel: "excel_windows"
      },
      host: {
        ...baseRequest().host,
        platform: "excel_windows"
      },
      capabilities: {
        ...baseRequest().capabilities,
        supportsNoteWrites: false
      }
    }));

    expect(prompt).toContain("Host capability matrix for excel_windows");
    expect(prompt).toContain("pivot_table_plan: limited");
    expect(prompt).toContain("Supported pivot sorting: group_field on an existing row or column group");
    expect(prompt).toContain("Supported pivot filters use existing pivot fields");
    expect(prompt).toContain("between, or not_between");
    expect(prompt).toContain("chart_plan: limited");
    expect(prompt).toContain("Series labels may use explicit custom legend text when provided.");
    expect(prompt).toContain("note_writes: unsupported");
    expect(prompt).toContain("external_data_plan: unsupported");
    expect(prompt).toContain("Do not emit external_data_plan on Excel hosts.");
    expect(prompt).toContain("Repeated conditions on the same column are exact-safe only when exactly two custom criteria can be combined with AND.");
  });

  it("includes reviewer-safe unavailable error guidance", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      reviewer: {
        reviewerSafeMode: true,
        forceExtractionMode: "unavailable"
      }
    }));

    expect(prompt).toContain("EXTRACTION_UNAVAILABLE");
    expect(prompt).toContain('type="error"');
  });

  it("redacts attachment access fields from the serialized Hermes prompt", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Import the uploaded receipt image.",
      context: {
        selection: {
          range: "A1:B2",
          headers: ["A", "B"],
          values: [["", ""]]
        },
        attachments: [
          {
            id: "att_prompt_001",
            type: "image",
            mimeType: "image/png",
            fileName: "receipt.png",
            size: 2048,
            source: "upload",
            previewUrl: "https://gateway.example/api/uploads/att_prompt_001/content?uploadToken=upl_secret_123&sessionId=sess_123",
            uploadToken: "upl_secret_123",
            storageRef: "blob://att_prompt_001",
            extractedText: "raw extracted text that should not be forwarded",
            metadata: {
              internalUrl: "https://internal.example/receipt.png"
            }
          }
        ]
      }
    }));

    expect(prompt).toContain('"id": "att_prompt_001"');
    expect(prompt).toContain('"mimeType": "image/png"');
    expect(prompt).toContain('"fileName": "receipt.png"');
    expect(prompt).not.toContain("uploadToken");
    expect(prompt).not.toContain("upl_secret_123");
    expect(prompt).not.toContain("previewUrl");
    expect(prompt).not.toContain("storageRef");
    expect(prompt).not.toContain("raw extracted text");
    expect(prompt).not.toContain("internal.example");
  });

  it("guides targeted formula-fix prompts toward executable sheet updates", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Fix the formula in cell F6 and apply it."
    }));

    expect(prompt).toContain('Prefer type="sheet_update"');
    expect(prompt).toContain('operation="set_formulas"');
    expect(prompt).toContain("targetCell");
  });

  it("routes current-cell formula apply prompts toward executable sheet updates", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Fix this formula and apply it to the current cell.",
      context: {
        selection: {
          range: "F6",
          headers: ["Revenue"],
          values: [[null]],
          formulas: [["=SUMIF(B:B,,F:F)"]]
        },
        activeCell: {
          a1Notation: "F6",
          displayValue: "#N/A",
          value: "#N/A",
          formula: "=SUMIF(B:B,,F:F)"
        }
      }
    }));

    expect(prompt).toContain('Prefer type="sheet_update"');
    expect(prompt).toContain('operation="set_formulas"');
    expect(prompt).toContain("context.activeCell.formula");
  });

  it("routes advisory formula-debug prompts toward formula responses when the user is not applying a change yet", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Why is this formula broken?",
      context: {
        selection: {
          range: "F6",
          headers: ["Revenue"],
          values: [[null]],
          formulas: [["=SUMIF(B:B,,F:F)"]]
        },
        activeCell: {
          a1Notation: "F6",
          displayValue: "#N/A",
          value: "#N/A",
          formula: "=SUMIF(B:B,,F:F)"
        }
      }
    }));

    expect(prompt).toContain('Prefer type="formula"');
    expect(prompt).toContain('intent="explain"');
    expect(prompt).toContain("context.activeCell.formula");
    expect(prompt).toContain("context.selection.formulas");
  });

  it("routes formula-debug correction prompts without a target apply action toward advisory formula fixes", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Debug this #REF! formula and tell me the fix."
    }));

    expect(prompt).toContain('Prefer type="formula"');
    expect(prompt).toContain('intent="fix"');
    expect(prompt).not.toContain('Prefer type="sheet_update"');
  });

  it("spells out the unsupported-operation error path for workbook actions with no contract support", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Resize this sheet to a square 10x10."
    }));

    expect(prompt).toContain("UNSUPPORTED_OPERATION");
    expect(prompt).toContain("unsupported workbook or formatting action");
    expect(prompt).toContain("Do not mention internal contracts, schema names, or validation failures");
    expect(prompt).toContain("suggest up to three closest supported alternatives");
  });

  it("routes create-sheet plus sample-data asks toward composite_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create a new sheet named demo_sales_data and fill it with random sales data."
    }));

    expect(prompt).toContain('Prefer type="composite_plan"');
    expect(prompt).toContain('create_sheet step followed by a sheet_update step that uses operation="replace_range"');
  });

  it("routes Vietnamese create-sheet plus generated-data asks toward composite_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Tao ngau nhien 1 vung du lieu ve doanh thu ban hang trong 1 sheet moi ten la demo_sales_data."
    }));

    expect(prompt).toContain('Prefer type="composite_plan"');
    expect(prompt).toContain('create_sheet step followed by a sheet_update step that uses operation="replace_range"');
  });

  it("routes Vietnamese generated-data asks on the current sheet toward sheet_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Tao ngau nhien du lieu vao sheet nay."
    }));

    expect(prompt).toContain('Prefer type="sheet_update"');
    expect(prompt).toContain("explicit data-population flow");
    expect(prompt).toContain('operation="replace_range"');
  });

  it("routes current-table fill-from-lookup asks toward sheet_update formula application", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Fill in the Item Name column based on the first left character of Item Code and look it up in the Lookup Table.",
      context: {
        selection: {
          range: "A1:K33",
          headers: ["Item Code", "Item Name", "Qty"],
          values: [["A100", "", 1]]
        },
        currentRegion: {
          range: "A1:K33",
          headers: ["Item Code", "Item Name", "Qty"]
        }
      }
    }));

    expect(prompt).toContain('Prefer type="sheet_update"');
    expect(prompt).toContain("fills or populates an existing column in the current table");
    expect(prompt).toContain('operation="set_formulas"');
    expect(prompt).toContain("using currentRegion headers to infer the source and target columns");
  });

  it("routes totals-row asks on the current table toward composite_plan instead of asking for reselection", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Thêm hàng tổng cho bảng doanh thu hiện tại.",
      context: {
        selection: {
          range: "J6",
          values: [["North"]]
        },
        currentRegion: {
          range: "A1:F11",
          headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"],
          values: [["2026-04-01", "Audio", "Cable", "North", 1, 15.5]]
        },
        currentRegionAppendTarget: "A12:F12"
      }
    }));

    expect(prompt).toContain('Prefer type="composite_plan"');
    expect(prompt).toContain("full values/formulas matrices may be omitted to keep the payload bounded");
    expect(prompt).toContain("Do not ask the user to select the whole table again when context.currentRegion already identifies it.");
    expect(prompt).toContain('insert_rows step followed by a sheet_update step that uses operation="set_formulas"');
  });

  it("routes control-cell requests with source-table overlap toward a helper-sheet composite plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      host: {
        ...baseRequest().host,
        activeSheet: "Sheet3",
        selectedRange: "J6"
      },
      userMessage: "I have an employee dataset and want to build a dynamic search tool. With the EEID in cell A1 and the target attribute in cell B1, I need a formula that looks up the specific value.",
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
    }));

    expect(prompt).toContain('Prefer type="composite_plan"');
    expect(prompt).toContain("The request names explicit input/output control cells that appear to overlap the current source table or header region.");
    expect(prompt).toContain("Do not reject the request solely because those control cells overlap the current source table.");
    expect(prompt).toContain("creates a helper sheet (for example Lookup_Demo)");
  });

  it("treats explicit tool scaffolding asks as helper-sheet composite flows even without cell overlap", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create a lookup tool on a helper sheet for the current employee table with one input cell for EEID, one output cell for Department, and a formula that returns the result.",
      context: {
        selection: {
          range: "A1:N20",
          headers: [
            "EEID",
            "Full Name",
            "Department",
            "Country"
          ],
          values: [["EMP-001", "Ada Lovelace", "Engineering", "UK"]]
        },
        currentRegion: {
          range: "A1:N20",
          headers: [
            "EEID",
            "Full Name",
            "Department",
            "Country"
          ]
        },
        currentRegionArtifactTarget: "P1"
      }
    }));

    expect(prompt).toContain('Prefer type="composite_plan"');
    expect(prompt).toContain("This request is a tool-like spreadsheet flow with user-facing inputs and outputs.");
    expect(prompt).toContain("creates or reuses a dedicated helper sheet");
    expect(prompt).toContain("seed clear labels for inputs and outputs");
    expect(prompt).toContain("short guidance row or section");
  });

  it("adds mixed advisory-and-write guidance for prompts that combine explanation with a write action", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Explain this selection and then sort it by revenue descending."
    }));

    expect(prompt).toContain("This request explicitly asks for a spreadsheet change. Do not return type=\"chat\"");
    expect(prompt).toContain("This request mixes an advisory explanation/debug step with a spreadsheet write action.");
    expect(prompt).toContain("split the analysis and writeback into separate steps");
  });

  it("keeps targeted formula apply requests on a write path even when they also ask for debugging context", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Why is this formula broken and then apply the fix to H11?"
    }));

    expect(prompt).toContain('Prefer type="sheet_update"');
    expect(prompt).toContain("This request mixes an advisory explanation/debug step with a spreadsheet write action.");
    expect(prompt).toContain("If one write-capable plan can satisfy the write and the explanation fits naturally in data.explanation");
  });

  it("routes sort prompts toward range_sort_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Sort this table by Status asc and Due Date desc."
    }));

    expect(prompt).toContain('Prefer type="range_sort_plan"');
  });

  it("routes filter prompts toward range_filter_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Filter this range to only show rows where Status equals Open."
    }));

    expect(prompt).toContain('Prefer type="range_filter_plan"');
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

  it("routes sheet structure prompts toward sheet_structure_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Freeze the top row and autofit the columns on this sheet."
    }));

    expect(prompt).toContain('Prefer type="sheet_structure_update"');
  });

  it("routes explicit move-sheet phrasing toward workbook_structure_update instead of range_transfer_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Move sheet Sheet1 to the end."
    }));

    expect(prompt).toContain('Prefer type="workbook_structure_update"');
    expect(prompt).not.toContain('Prefer type="range_transfer_plan"');
  });

  it("keeps explicit workbook confirmations on the workbook_structure_update path instead of chat acknowledgements", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Confirm create sheet demo_sales_data"
    }));

    expect(prompt).toContain('Prefer type="workbook_structure_update"');
    expect(prompt).toContain("Do not answer explicit confirmation phrasing with a chat acknowledgement");
  });

  it("routes ungroup prompts toward sheet_structure_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Ungroup these rows in the current sheet."
    }));

    expect(prompt).toContain('Prefer type="sheet_structure_update"');
  });

  it("routes unmerge prompts toward sheet_structure_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Unmerge the cells in range B2:D2."
    }));

    expect(prompt).toContain('Prefer type="sheet_structure_update"');
  });

  it("routes unfreeze prompts toward sheet_structure_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Unfreeze the panes on this sheet."
    }));

    expect(prompt).toContain('Prefer type="sheet_structure_update"');
  });

  it("routes tab color prompts toward sheet_structure_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Set the sheet tab color to red."
    }));

    expect(prompt).toContain('Prefer type="sheet_structure_update"');
  });

  it("routes natural named-range rename prompts toward named_range_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Rename SalesData to SalesData2026."
    }));

    expect(prompt).toContain('Prefer type="named_range_update"');
  });

  it("routes ordinary named-range identifiers toward named_range_update", () => {
    const renamePrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Rename Sales to Revenue."
    }));
    const createPrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create Revenue for B2:D20."
    }));
    const retargetPrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Retarget Sales to B2:D20."
    }));
    const deletePrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Delete Revenue."
    }));

    expect(renamePrompt).toContain('Prefer type="named_range_update"');
    expect(createPrompt).toContain('Prefer type="named_range_update"');
    expect(retargetPrompt).toContain('Prefer type="named_range_update"');
    expect(deletePrompt).toContain('Prefer type="named_range_update"');
  });

  it("routes natural named-range retarget prompts toward named_range_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Retarget SalesData to B2:D20."
    }));

    expect(prompt).toContain('Prefer type="named_range_update"');
  });

  it("routes natural named-range create prompts toward named_range_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create SalesData for B2:D20."
    }));

    expect(prompt).toContain('Prefer type="named_range_update"');
  });

  it("routes natural named-range delete prompts toward named_range_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Delete SalesData."
    }));

    expect(prompt).toContain('Prefer type="named_range_update"');
  });

  it("routes sheet-scoped named-range create prompts toward named_range_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create SalesData on Sheet1 for B2:D20."
    }));

    expect(prompt).toContain('Prefer type="named_range_update"');
  });

  it("routes sheet-scoped named-range prompts with spaced sheet names toward named_range_update", () => {
    const createPrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create SalesData on Sheet 1 for B2:D20."
    }));
    const retargetPrompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Retarget InputRange on Sheet 1 to B2:D20."
    }));

    expect(createPrompt).toContain('Prefer type="named_range_update"');
    expect(retargetPrompt).toContain('Prefer type="named_range_update"');
  });

  it("routes sheet-scoped named-range retarget prompts toward named_range_update", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Retarget InputRange on Sheet1 to B2:D20."
    }));

    expect(prompt).toContain('Prefer type="named_range_update"');
  });

  it("keeps named-range mutations ahead of validation wording when both appear", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Rename the named range used by this dropdown to SalesData2026."
    }));

    expect(prompt).toContain('Prefer type="named_range_update"');
  });

  it("routes whole-number validation prompts toward data_validation_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Only allow whole numbers between 1 and 10 in B2:B20."
    }));

    expect(prompt).toContain('Prefer type="data_validation_plan"');
  });

  it("routes custom-formula validation prompts toward data_validation_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Apply a custom formula validation to G2:G20."
    }));

    expect(prompt).toContain('Prefer type="data_validation_plan"');
  });

  it("routes analysis prompts toward analysis_report_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Analyze this range and give me a short report with trends and anomalies."
    }));

    expect(prompt).toContain('Prefer type="analysis_report_plan"');
  });

  it("routes explicit multi-step workflows toward composite_plan instead of a single write plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Sort this table by revenue, then filter status to Open, then add a summary report."
    }));

    expect(prompt).toContain('Prefer type="composite_plan"');
    expect(prompt).not.toContain('Prefer type="analysis_report_plan" for structured analysis reports, with chat_only for non-write summaries and materialize_report for confirmable report artifacts.');
  });

  it("keeps single-step analysis summaries on the analysis_report_plan path", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Summarize this range with trends and anomalies."
    }));

    expect(prompt).toContain('Prefer type="analysis_report_plan"');
    expect(prompt).not.toContain('Prefer type="composite_plan"');
  });

  it("treats analysis requests that explicitly place a report onto a sheet as write-capable", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Analyze this range and put the report on a new sheet named Summary."
    }));

    expect(prompt).toContain('Prefer type="analysis_report_plan"');
    expect(prompt).toContain("This request explicitly asks for a spreadsheet change. Do not return type=\"chat\"");
  });

  it("routes pivot prompts toward pivot_table_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create a pivot table showing total revenue by region and quarter."
    }));

    expect(prompt).toContain('Prefer type="pivot_table_plan"');
  });

  it("routes Vietnamese pivot aliases toward pivot_table_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Tao bang tong hop doanh thu theo khu vuc."
    }));

    expect(prompt).toContain('Prefer type="pivot_table_plan"');
  });

  it("keeps advisory pivot how-to prompts on the chat path", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Explain how to create a pivot table for this data."
    }));

    expect(prompt).toContain('Prefer type="chat"');
    expect(prompt).not.toContain('Prefer type="pivot_table_plan"');
    expect(prompt).not.toContain("This request explicitly asks for a spreadsheet change.");
  });

  it("uses currentRegion guidance for pivot asks over the current data region", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "tao pivot cho vung data hien tai",
      context: {
        selection: {
          range: "J6",
          values: [["North"]]
        },
        currentRegion: {
          range: "A1:F11",
          headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"]
        },
        currentRegionArtifactTarget: "A13"
      }
    }));

    expect(prompt).toContain('Prefer type="pivot_table_plan"');
    expect(prompt).toContain("use context.currentRegion as the implicit table/range");
    expect(prompt).toContain("use context.currentRegionArtifactTarget as the default targetRange");
  });

  it("guides Hermes to infer a conservative default layout for under-specified pivot requests", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Do some pivot here for data in range A1:F9. I have no idea the shape of the pivot, so please help me.",
      context: {
        selection: {
          range: "B2",
          values: [["North"]]
        },
        currentRegion: {
          range: "A1:F9",
          headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"]
        },
        currentRegionArtifactTarget: "A11"
      }
    }));

    expect(prompt).toContain("do not degrade to chat-only just because the user did not name every pivot field");
    expect(prompt).toContain("infer a conservative default pivot");
    expect(prompt).toContain("Category, Region, Department, Type, or Status");
    expect(prompt).toContain("use context.currentRegionArtifactTarget when it is present");
  });

  it("keeps explicit pivot requests ahead of broader analysis wording", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create a pivot table report with revenue trends by quarter."
    }));

    expect(prompt).toContain('Prefer type="pivot_table_plan"');
    expect(prompt).not.toContain('Prefer type="analysis_report_plan"');
  });

  it("routes chart prompts toward chart_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create a line chart of revenue by month on a new sheet."
    }));

    expect(prompt).toContain('Prefer type="chart_plan"');
  });

  it("keeps advisory chart how-to prompts on the chat path", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Explain how to create a line chart for this data."
    }));

    expect(prompt).toContain('Prefer type="chat"');
    expect(prompt).not.toContain('Prefer type="chart_plan"');
    expect(prompt).not.toContain("This request explicitly asks for a spreadsheet change.");
  });

  it("uses currentRegion guidance for chart asks over the current data region", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Tạo biểu đồ doanh thu từ bảng hiện tại.",
      context: {
        selection: {
          range: "J6",
          values: [["North"]]
        },
        currentRegion: {
          range: "A1:F11",
          headers: ["Date", "Category", "Product", "Region", "Units", "Revenue"]
        },
        currentRegionArtifactTarget: "A13"
      }
    }));

    expect(prompt).toContain('Prefer type="chart_plan"');
    expect(prompt).toContain("use context.currentRegion as the implicit table/range");
    expect(prompt).toContain("use context.currentRegionArtifactTarget as the default targetRange");
  });

  it("keeps explicit chart requests ahead of broader analysis wording", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create a chart report showing revenue trends by month."
    }));

    expect(prompt).toContain('Prefer type="chart_plan"');
    expect(prompt).not.toContain('Prefer type="analysis_report_plan"');
  });

  it("does not treat build as a chart-intent substring inside larger words", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Analyze this graph and include a rebuild note in the report."
    }));

    expect(prompt).toContain('Prefer type="analysis_report_plan"');
    expect(prompt).not.toContain('Prefer type="chart_plan"');
  });

  it("routes highlight-style prompts toward conditional_format_plan", () => {
    const prompts = [
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Highlight overdue dates."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Color values above 10 in red."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Mark duplicates in column B."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Apply a 3-color scale to this table."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Clear conditional formatting from A:A."
      }))
    ];

    for (const prompt of prompts) {
      expect(prompt).toContain('Prefer type="conditional_format_plan"');
    }
  });

  it("does not route plain static formatting wording to conditional_format_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Format column B in red."
    }));

    expect(prompt).not.toContain('Prefer type="conditional_format_plan"');
  });

  it("routes static formatting prompts toward range_format_update", () => {
    const prompts = [
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Make A1:F1 bold with blue fill and a bottom border."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Set C2:C20 to currency number format."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Wrap text and center align this selected range."
      }))
    ];

    for (const prompt of prompts) {
      expect(prompt).toContain('Prefer type="range_format_update"');
      expect(prompt).not.toContain('Prefer type="conditional_format_plan"');
      expect(prompt).not.toContain('Prefer type="table_plan"');
    }
  });

  it("keeps advisory static formatting how-to prompts on the chat path", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Explain how to make A1:F1 bold in Excel."
    }));

    expect(prompt).toContain('Prefer type="chat"');
    expect(prompt).not.toContain('Prefer type="range_format_update"');
  });

  it("includes wave-1 sheet structure confirmation invariants in the request guidance", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Delete rows 8 to 10."
    }));

    expect(prompt).toContain('delete_rows and delete_columns require data.confirmationLevel="destructive"');
    expect(prompt).toContain('all other sheet_structure_update operations require data.confirmationLevel="standard"');
    expect(prompt).toContain("unfreeze_panes must resolve to data.frozenRows=0 and data.frozenColumns=0");
  });

  it("routes transfer prompts toward range_transfer_plan", () => {
    const prompts = [
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Copy this table to Sheet2."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Move A1:D20 to Archive!A1."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Append these values to the end of Sheet3."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Transpose this list into row 1 on Summary."
      }))
    ];

    for (const prompt of prompts) {
      expect(prompt).toContain('Prefer type="range_transfer_plan"');
      expect(prompt).toContain("data.targetRange is required and must be the full destination rectangle");
      expect(prompt).toContain("do not default data.targetRange to A1");
      expect(prompt).not.toContain("full destination rectangle or A1 anchor");
      expect(prompt).not.toContain(
        "If the target sheet is known but no target range is specified by the user, default data.targetRange to A1"
      );
    }
  });

  it("routes cleanup prompts toward data_cleanup_plan", () => {
    const prompts = [
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Trim whitespace in this table."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Remove duplicate rows from this range."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Split column B on commas into columns C:E."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Join columns A and B into column C with a dash."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Fill down missing values in column D."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Standardize these date strings to YYYY-MM-DD."
      })),
      buildHermesSpreadsheetRequestPrompt(baseRequest({
        userMessage: "Cleanup and reshape this imported table."
      }))
    ];

    for (const prompt of prompts) {
      expect(prompt).toContain('Prefer type="data_cleanup_plan"');
    }
  });

  it("routes market-data prompts toward external_data_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Get the latest BTC price into Sheet2 starting at B2 using GOOGLEFINANCE."
    }));

    expect(prompt).toContain('Prefer type="external_data_plan"');
    expect(prompt).toContain("stock/crypto market data");
  });

  it("routes Excel external-data prompts toward explicit unsupported errors", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      source: {
        ...baseRequest().source,
        channel: "excel_windows"
      },
      host: {
        ...baseRequest().host,
        platform: "excel_windows"
      },
      userMessage: "Get the latest BTC price into Sheet2 starting at B2 using GOOGLEFINANCE."
    }));

    expect(prompt).toContain('Prefer type="error"');
    expect(prompt).toContain('data.code="UNSUPPORTED_OPERATION"');
    expect(prompt).not.toContain('Prefer type="external_data_plan"');
  });

  it("routes public website-table import prompts toward external_data_plan", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Import table 1 from a public website into Sheet2 starting at A1 using IMPORTHTML."
    }));

    expect(prompt).toContain('Prefer type="external_data_plan"');
    expect(prompt).toContain("public website-table imports into Google Sheets");
  });

  it("keeps advisory external-data how-to prompts on the chat path", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Explain how to use IMPORTHTML to import a website table."
    }));

    expect(prompt).toContain('Prefer type="chat"');
    expect(prompt).not.toContain('Prefer type="external_data_plan"');
    expect(prompt).not.toContain("This request explicitly asks for a spreadsheet change.");
  });

  it("includes exact chart, report, conditional-format, and cleanup contract hints for local-brain planning", () => {
    const prompt = buildHermesSpreadsheetRequestPrompt(baseRequest({
      userMessage: "Create a management summary, highlight SLA risks, and chart spend versus revenue."
    }));

    expect(prompt).toContain("data.sections must be an array of objects with type, title, summary, and sourceRanges");
    expect(prompt).toContain("Each series item must use field to reference a source header name");
    expect(prompt).toContain("For row-highlighting logic driven by a status/breach/overdue column or by comparisons between columns, prefer ruleType=\"custom_formula\"");
    expect(prompt).toContain("Do not compress multiple cleanup transforms into one broad cleanup step");
  });
});
