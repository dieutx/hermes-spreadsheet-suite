# Opus 4.7 Feature Expansion Roadmap

Date: 2026-04-19

Model review source:
- `/tmp/hermes_feature_expansion_opus_review.md`
- `/tmp/hermes_feature_expansion_opus_review_cont2.md`

Scope:
- Spreadsheet assistant feature expansion beyond MVP
- Grounded in current repo boundaries
- No assumed Hermes core changes

## Current Supported Surface

Source-evidenced from:
- [packages/contracts/src/schemas.ts](<repo-root>/packages/contracts/src/schemas.ts)
- [services/gateway/src/routes/writeback.ts](<repo-root>/services/gateway/src/routes/writeback.ts)
- [apps/excel-addin/src/taskpane/taskpane.js](<repo-root>/apps/excel-addin/src/taskpane/taskpane.js)
- [apps/google-sheets-addon/src/Code.gs](<repo-root>/apps/google-sheets-addon/src/Code.gs)

Already supported:
- `chat`
- `formula`
- `sheet_update`
- `sheet_import_plan`
- `workbook_structure_update`
- `range_format_update`
- `attachment_analysis`
- `extracted_table`
- `document_summary`
- `error`

Current writeback-able plans:
- `sheet_update`
- `sheet_import_plan`
- `workbook_structure_update`
- `range_format_update`

Current workbook structure ops:
- `create_sheet`
- `delete_sheet`
- `rename_sheet`
- `duplicate_sheet`
- `move_sheet`
- `hide_sheet`
- `unhide_sheet`

Current formatting surface:
- `numberFormat`
- `backgroundColor`
- `textColor`
- `bold`
- `italic`
- `horizontalAlignment`
- `verticalAlignment`
- `wrapStrategy`
- `columnWidth`
- `rowHeight`

## Highest-Value Missing Capability Families

Opus 4.7 identified these as the biggest expansion gaps for a practical spreadsheet assistant:

1. Sheet structure operations
- insert/delete rows
- insert/delete columns
- merge/unmerge cells
- freeze/unfreeze panes
- group/ungroup rows and columns
- hide/unhide rows and columns
- autofit rows and columns
- set sheet tab color

2. Range transfer and cleanup operations
- copy/cut/paste with paste modes
- fill down/right/series
- find and replace
- clear contents / clear formatting / clear all
- trim whitespace
- normalize case
- split text to columns
- text-to-number/date cleanup
- dedupe rows
- remove empty rows/columns
- regex extract
- reshape wide/long

3. Sorting, filtering, validation, and names
- sort plans
- filter plans
- data validation plans
- named range updates
- protected range / sheet plans

4. Rich formatting and visual features
- font family / size / underline / strikethrough
- borders
- text rotation
- indentation
- conditional formatting plans
- formatting clear mode

5. Analysis and reporting
- structured analysis report plans
- pivot table plans
- chart plans
- sparklines
- subtotal / total-row helpers

6. Richer import flows
- PDF attachments
- CSV/TSV/XLSX import
- URL fetch imports
- multi-table / multi-region import plans

7. Safety and workflow
- explicit dry-run mode
- reversible / undo-aware plans
- secondary confirmation for destructive ops
- plan expiry / stale-plan detection
- composite multi-step plans

## Recommended Contract Additions

In [packages/contracts/src/schemas.ts](<repo-root>/packages/contracts/src/schemas.ts):

1. Expand `CapabilitiesSchema`
- `supportsConditionalFormat`
- `supportsDataValidation`
- `supportsNamedRanges`
- `supportsProtectedRanges`
- `supportsStructureEdits`
- `supportsAutofit`
- `supportsFontFamily`
- `supportsBorders`
- `supportsPivotTables`
- `supportsCharts`
- `supportsSparklines`
- `supportsTablesListObjects`
- `supportsSortFilter`
- `supportsFindReplace`
- `supportsHyperlinks`
- `supportsImageInsert`
- `supportsPdfAttachments`
- `supportsCsvAttachments`
- `supportsXlsxAttachments`
- `supportsCompositePlans`
- `supportsDryRun`
- `supportsUndo`
- `hostApiTier`

2. Expand `RangeFormatSchema`
- `fontFamily`
- `fontSize`
- `underline`
- `strikethrough`
- `borders`
- `textRotationDegrees`
- `indent`
- `clear`
- `perColumnWidths`
- `perRowHeights`

3. Add new response/data types
- `sheet_structure_update`
- `range_transfer_plan`
- `range_sort_plan`
- `range_filter_plan`
- `formula_audit`
- `pivot_table_plan`
- `chart_plan`
- `data_validation_plan`
- `named_range_update`
- `conditional_format_plan`
- `data_cleanup_plan`
- `data_reshape_plan`
- `analysis_report_plan`
- `composite_plan`

4. Add envelope metadata
- `planId`
- `requiresConfirm`
- `affectedRanges`
- `reversible`
- `costEstimate`

## Gateway / Runtime / Render / Writeback Work

Primary source files:
- [services/gateway/src/hermes/runtimeRules.ts](<repo-root>/services/gateway/src/hermes/runtimeRules.ts)
- [services/gateway/src/hermes/requestTemplate.ts](<repo-root>/services/gateway/src/hermes/requestTemplate.ts)
- [services/gateway/src/routes/writeback.ts](<repo-root>/services/gateway/src/routes/writeback.ts)
- [packages/shared-client/src/render.ts](<repo-root>/packages/shared-client/src/render.ts)
- [packages/shared-client/src/types.ts](<repo-root>/packages/shared-client/src/types.ts)

Recommended changes:
- Teach runtime rules and request template every new response type explicitly.
- Add unsupported-action routing for every not-yet-supported category.
- Add a plan cache keyed by `planId`.
- Add stale-plan detection before apply.
- Extend writeback result types beyond `range_write` and `workbook_structure_update`.
- Add server-side preview generation for:
  - sort preview
  - filter preview
  - pivot preview
  - chart preview metadata
- Break response handling into a registry instead of an ever-growing switch.
- Add composite plan orchestration for multi-step flows.
- Add undo stack support for non-cell operations.

## Host Adapter Work

Excel:
- Add handler modules for structure, format, sort/filter, validation, charts, pivots.
- Batch Office.js mutations under one `Excel.run` with one final `context.sync()`.
- Add translation helpers for:
  - borders
  - conditional formatting
  - data validation
  - chart types

Google Sheets:
- Compile new plans into `batchUpdate` requests.
- Maintain `sheetName <-> sheetId` cache safely after rename/add/delete.
- Centralize `fields` mask generation for every update request.
- Add GridRange/A1 conversion helpers for all new op families.

Cross-host:
- Move feature support decisions behind typed host capability declarations.
- Gracefully degrade unsupported host features instead of emitting invalid plans.

## Test Matrix

Opus 4.7 recommends three main layers plus property-based tests.

1. Unit tests
- schema validation
- request compilation
- A1/GridRange conversion
- field-mask generation
- conditional-format translation
- capability gating

2. Adapter tests with mocked host APIs
- exact Office.js call sequence
- exact Google Sheets `batchUpdate` payloads
- correct cache refresh behavior
- correct batching/chunking behavior
- proper unsupported-operation rejection

3. Integration tests with live hosts
- build a workbook from empty state
- destructive op + undo
- concurrent op serialization
- large batch stress
- cross-host parity scenario

4. Property-based tests
- A1 round-trip
- commutative op reorderings
- format merge semantics

## Debug Playbook

Opus 4.7 highlighted these debugging anchors:
- log compiled host requests before apply
- log plan validation, plan preview, and apply separately
- log stale-plan / snapshot-hash mismatches
- add per-host debug namespaces
- add preflight shape assertions before host apply
- verify `fields` masks for Google Sheets
- verify `numberFormat` 2D shape for Excel
- verify cache refresh after sheet add/delete/rename in Google Sheets

## Ordered Implementation Plan

Phase 1: Strongest near-term value
- `sheet_structure_update`
- `range_sort_plan`
- `range_filter_plan`
- richer `range_format_update`
- `data_validation_plan`
- capability-schema expansion

Phase 2: Spreadsheet editing power tools
- `range_transfer_plan`
- `data_cleanup_plan`
- `named_range_update`
- `conditional_format_plan`
- stale-plan detection
- reversible metadata

Phase 3: Analysis and reporting
- `analysis_report_plan`
- `pivot_table_plan`
- `chart_plan`
- structured formula audit

Phase 4: Workflow depth
- `composite_plan`
- undo/redo stack
- dry-run mode
- plan history

Phase 5: Attachment/import expansion
- PDF
- CSV/TSV/XLSX
- multi-region imports
- URL-based import sources

## Key Tradeoffs

- Keep workbook/sheet structure distinct from rectangular cell writes.
- Keep validation and conditional formatting separate from plain range formatting.
- Prefer capability-gated degradation over fake confirmations.
- Prefer server-side preview generation for complex plan types.
- Avoid growing one giant dispatcher switch; move to handler registry.
- Preserve a clean distinction between:
  - read-only analytic responses
  - previewable write plans
  - directly executable host operations
