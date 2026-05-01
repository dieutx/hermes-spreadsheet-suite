# Review Notes - 2026-04-22 Pass 3

## Review status
- [x] Switched to latest `origin/main` at `db683c4`.
- [x] Reran `npm test` on Windows against the latest pull.
- [x] Revalidated the earlier writeback, app, execution-control, and shared-client findings on the current `main`.
- [x] Trimmed this pass down to findings that are still reproducible on `db683c4`.

## Current test blockers
- [ ] Excel host suites still import the taskpane from an author-local path under `<repo-root>`, so the Excel host tests cannot load the add-in module on this machine.
  Evidence: `services/gateway/tests/excelWave1Plans.test.ts:17`, `services/gateway/tests/excelWave2Plans.test.ts:6`, `services/gateway/tests/excelWave3Plans.test.ts:6`, `services/gateway/tests/excelWave4Plans.test.ts:6`, `services/gateway/tests/excelWave5Plans.test.ts:6`, `services/gateway/tests/excelWave6Plans.test.ts:6`
- [ ] Invalid-response debug-artifact tests still enumerate `/tmp`, which resolves to `E:\tmp` on Windows and keeps the suite red before those assertions reach the code under test.
  Evidence: `services/gateway/tests/hermesClient.test.ts:2046-2127`

## Fixed upstream on `db683c4`
- [x] Gateway app error handling no longer maps every `SyntaxError` to a bad client JSON body; it now only treats Express `entity.parse.failed` errors as malformed request bodies.
  Evidence: `services/gateway/src/app.ts:67-109`
- [x] Shared-client conversation sanitization now reprojects each message to `{ role, content }`, so strict gateway parsing no longer rejects callers that carry extra chat-state fields.
  Evidence: `packages/shared-client/src/request.ts:23-33`
- [x] Dry-run expiry generation now comes from the injected ledger clock, so creation and freshness checks share the same time source.
  Evidence: `services/gateway/src/routes/executionControl.ts:148-174`, `services/gateway/src/lib/executionLedger.ts:35-47`, `services/gateway/src/lib/executionLedger.ts:117-129`
- [x] `/complete` now records execution history before mutating `TraceBus`, so a failed history write no longer consumes the approval token.
  Evidence: `services/gateway/src/routes/writeback.ts:1562-1590`, `services/gateway/tests/writebackFlow.test.ts:706-785`
- [x] `range_filter_plan` and `conditional_format_plan` completion now require a typed full-plan payload and reject same-family detail mismatches.
  Evidence: `services/gateway/src/routes/writeback.ts:314-337`, `services/gateway/src/routes/writeback.ts:1341-1366`, `services/gateway/tests/writebackFlow.test.ts:4422-4520`
- [x] Materialized `analysis_report` approvals now reject invalid `targetRange` anchors at request validation instead of falling through to a 500.
  Evidence: `services/gateway/tests/writebackFlow.test.ts:2393-2454`

## Confirmed findings on `db683c4`

### [x] Confirmed: [P2] `sheet_update` and `sheet_import_plan` completion still only proves the written rectangle, not the written cell content
Files:
- `services/gateway/src/routes/writeback.ts:1234-1251`
- `services/gateway/src/routes/writeback.ts:143-150`
- `packages/contracts/src/schemas.ts:1578-1746`
- `services/gateway/tests/writebackFlow.test.ts:543-645`
- `services/gateway/tests/writebackFlow.test.ts:4361-4419`

Why this matters:
- These plans approve the actual spreadsheet payload: `values`, `formulas`, `notes`, `headers`, and extracted/imported cell content.
- The completion contract still only reports `targetSheet`, `targetRange`, `writtenRows`, and `writtenColumns`.
- `/complete` therefore cannot distinguish "wrote the right rectangle with the wrong cells" from a real successful apply.
- Current tests only prove replay protection, target/shape enforcement, and retry behavior; they do not cover wrong-content writes on the approved range.

### [x] Confirmed: [P2] `range_transfer` completion ignores `pasteMode` and `transpose`, even though both hosts treat them as execution-critical semantics
Files:
- `packages/contracts/src/schemas.ts:1098-1143`
- `packages/contracts/src/schemas.ts:2171-2178`
- `services/gateway/src/routes/writeback.ts:898-922`
- `apps/excel-addin/src/taskpane/taskpane.js:1674-1775`
- `apps/google-sheets-addon/src/Code.gs:2021-2127`
- `services/gateway/tests/writebackFlow.test.ts:4523-4757`

Why this matters:
- Approved transfer plans include `pasteMode` (`values`, `formulas`, `formats`) and `transpose`, and both hosts branch heavily on those fields when computing shape and copying data.
- The completion payload only carries `sourceSheet`, `sourceRange`, `targetSheet`, `targetRange`, and `transferOperation`.
- A host can therefore copy or move the right source and target rectangles while applying the wrong transfer mode or transpose setting, and `/complete` will still record success.
- Existing tests only cover source/target/operation mismatches; they never exercise a wrong `pasteMode` or wrong `transpose` on the approved transfer.

### [x] Confirmed: [P2] Malformed-response debug artifacts are still hard-coded to `/tmp`, so the forensic dump path silently fails on Windows
Files:
- `services/gateway/src/lib/hermesClient.ts:46-97`
- `services/gateway/tests/hermesClient.test.ts:2046-2127`

Why this matters:
- The runtime debug dump prefix is still hard-coded to `/tmp/hermes-spreadsheet-invalid`.
- On Windows that location does not exist by default, so invalid-assistant payload forensics are lost even when the debug flag is enabled.
- The current red tests on Windows are reflecting the same portability gap as production code, not just a bad assertion.

## Notes for reviewer
- [x] Everything in `Fixed upstream on db683c4` was rechecked on the latest pull and should not be reopened without a fresh repro.
- [x] Everything in `Confirmed findings on db683c4` is still open on the current `main`.
- [x] The Windows test blockers remain reproducible and still prevent a clean full-suite pass locally.
