import { describe, expect, it } from "vitest";
import { SPREADSHEET_RUNTIME_RULES } from "../src/hermes/runtimeRules.ts";

describe("spreadsheet runtime rules", () => {
  it("documents all structured response types with contract-shaping guidance", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="formula"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="sheet_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="sheet_import_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="data_validation_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="named_range_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="workbook_structure_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="range_format_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="sheet_structure_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="range_sort_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="range_filter_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="error"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="attachment_analysis"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="extracted_table"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="document_summary"');
  });

  it("spells out formula language and write-plan confirmation requirements", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain("formulaLanguage");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("requiresConfirmation");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("targetRange must match");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("supportsNoteWrites");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("host capability matrix");
  });

  it("documents wave-1 sheet structure confirmation-level invariants", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('delete_rows and delete_columns require data.confirmationLevel="destructive"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('all other sheet_structure_update operations require data.confirmationLevel="standard"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain("unfreeze_panes must resolve to data.frozenRows=0 and data.frozenColumns=0");
  });

  it("makes reviewer-safe unavailable map to a concrete error response", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain("EXTRACTION_UNAVAILABLE");
    expect(SPREADSHEET_RUNTIME_RULES).toContain('prefer type="error"');
  });

  it("documents fail-closed handling for explicit write requests and mixed advisory-plus-write asks", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('If the user explicitly asks for a spreadsheet change, do not return type="chat"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain("If the user mixes explanation/debug/summarization with a write request");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("split the analysis and writeback into separate steps");
  });

  it("requires unsupported workbook actions to return a clean unsupported-operation error", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain("UNSUPPORTED_OPERATION");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("unsupported workbook or formatting action");
  });

  it("documents advanced static range formatting fields", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain("fontFamily");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("fontSize");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("underline");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("strikethrough");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("border");
  });

  it("documents conditional formatting as a distinct write-plan family with fail-closed semantics", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain("conditional_format_plan");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("distinct from range_format_update");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("clear_on_target");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("replace_all_on_target");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("host-exact");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("unsupported mappings must fail closed");
  });

  it("documents wave-4 transfer and cleanup plan families with destructive confirmation rules", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="range_transfer_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="data_cleanup_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain(
      "data.targetRange is required and must be the full destination rectangle"
    );
    expect(SPREADSHEET_RUNTIME_RULES).toContain("do not default data.targetRange to A1");
    expect(SPREADSHEET_RUNTIME_RULES).not.toContain("full destination rectangle or A1 anchor");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("distinct from sheet_update");
    expect(SPREADSHEET_RUNTIME_RULES).toContain('move transfer plans require data.confirmationLevel="destructive"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('destructive cleanup operations require data.confirmationLevel="destructive"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('data.clearExistingFilters=true requires data.confirmationLevel="destructive"');
    expect(SPREADSHEET_RUNTIME_RULES).not.toContain('target-sheet-only transfer defaults to data.targetRange="A1"');
  });

  it("documents unsupported fuzzy cleanup and overlap ambiguity as fail-closed unsupported operations", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain("fuzzy");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("heuristic");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("overlap ambiguity");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("UNSUPPORTED_OPERATION");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("fail closed");
  });

  it("documents wave-5 analysis, pivot, and chart families with exact-safe semantics", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="analysis_report_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="pivot_table_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="chart_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="analysis_report_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="pivot_table_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="chart_update"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('outputMode is required and must be chat_only or materialize_report');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('chat_only requires data.requiresConfirmation=false');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('materialize_report requires data.requiresConfirmation=true');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('cannot be represented exactly on the current host');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('unsupported pivot or chart mappings must fail closed');
    expect(SPREADSHEET_RUNTIME_RULES).toContain('every pivot filter requires value');
  });

  it("documents current-region grounding so current-table requests do not force a reselection round-trip", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain("context.currentRegion");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("context.currentRegionArtifactTarget");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("context.currentRegionAppendTarget");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("do not degrade to chat-only just because the user did not name every pivot field");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("infer a conservative default pivot");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("full values/formulas matrices may be omitted to keep the payload bounded");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("Do not ask the user to select the whole table again when context.currentRegion already identifies it.");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("lookup tools, trackers, helper sheets, input sheets, or output sheets");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("move that scaffold to a helper sheet instead of rejecting the task");
  });

  it("documents wave-6 composite plans as a strict Hermes-emission boundary", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="composite_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain("must not contain another composite_plan");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("analysis_report_plan(chat_only) is not allowed in composite plans");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("composite_plan always requires confirmation");
  });

  it("documents public-url requirements for external data web imports", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('For type="external_data_plan"');
    expect(SPREADSHEET_RUNTIME_RULES).toContain("data.sourceUrl must be a public HTTP(S) URL");
    expect(SPREADSHEET_RUNTIME_RULES).toContain("external-data formulas must not reference private or internal URLs");
  });

  it("keeps explicit confirmation phrasing on supported plan types instead of chat acknowledgements", () => {
    expect(SPREADSHEET_RUNTIME_RULES).toContain('confirm create sheet X');
    expect(SPREADSHEET_RUNTIME_RULES).toContain("return that plan type as a proposal instead of chat");
  });
});
