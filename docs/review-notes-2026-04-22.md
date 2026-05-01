# Review Notes - 2026-04-22 (Reconciled on `main`)

## Scope
- Original review note came from PR `#1` (`codex/review-notes-20260422`) and was written against old `main` commit `3ead367`.
- This reconciled note reflects current `main` after follow-up fixes through `20804e0`.

## Still open
- [ ] Excel dev manifest is still pinned to a local dev host.
  Evidence: [apps/excel-addin/manifest.xml](<repo-root>/apps/excel-addin/manifest.xml)
  Notes:
  - `SourceLocation`, icon URLs, and `Taskpane.Url` still point to `https://localhost:3000/...`
  - this is a portability blocker for local/shared-host scenarios, not a gateway contract bug

- [ ] Invalid-response debug dumps are still hard-coded to `/tmp`.
  Evidence: [services/gateway/src/lib/hermesClient.ts](<repo-root>/services/gateway/src/lib/hermesClient.ts)
  Notes:
  - `INVALID_HERMES_DEBUG_PREFIX` still uses `/tmp/hermes-spreadsheet-invalid`
  - this is a portability issue for Windows forensic/debug flows

- [ ] Invalid-response debug tests still enumerate `/tmp`.
  Evidence: [services/gateway/tests/hermesClient.test.ts](<repo-root>/services/gateway/tests/hermesClient.test.ts)
  Notes:
  - tests still call `fs.readdir("/tmp")` and `fs.readFile("/tmp/...")`
  - this keeps the Windows-path blocker real in the test suite design

## Fixed on current `main`
- [x] Composite raw-step normalization no longer misclassifies valid nested plans as `sheet_structure_update`.
- [x] Conditional-format normalization no longer drops valid style fields like `underline` and `strikethrough`; `numberFormat` is rejected because hosts do not support exact conditional number-format writeback.
- [x] Local loopback gateway defaults no longer expose broad wildcard CORS/private-network access by default.
- [x] `range_filter` completion now verifies approved filter semantics instead of only `targetSheet/targetRange`.
- [x] `conditional_format_plan` completion now verifies approved rule/style semantics instead of only `targetSheet/targetRange/managementMode`.
- [x] `data_cleanup_plan` completion now verifies approved cleanup parameters instead of only `targetSheet/targetRange/cleanupOperation`.

## Current assessment
- The functional findings from PR `#1` are stale on current `main`; they should not be reopened without a fresh repro.
- The remaining real items from that note are portability/documentation blockers around Excel local hosting and `/tmp`-based invalid-response debug paths.

## Verification
- `npm test`
- Result: `34/34` test files pass, `563/563` tests pass on current `main`
