# Claude Opus 4.7 First Pass Review

Date: 2026-04-19

Model used through OpenRouter:

- `anthropic/claude-opus-4.7`

Raw extracted review text was saved locally during the run to:

- `/tmp/openrouter-hermes-review.md`
- `/tmp/openrouter-hermes-review-response.json`

This file keeps only findings that were checked back against the repo source.

## Verified findings

### 1. Write-back completion state is not persisted

Files:

- [services/gateway/src/routes/writeback.ts](<repo-root>/services/gateway/src/routes/writeback.ts)

What is true:

- `approveWriteback(...)` generates an approval token but does not persist approval state.
- `completeWriteback(...)` verifies the token and returns `{ ok: true }`.
- It does not mark the token as consumed.
- It does not reject unknown completion state beyond token verification.
- It does not append run/trace state about completion.

Why it matters:

- replay resistance is weak
- auditability of confirmation/write-back is weak
- gateway state does not fully reflect the approval/completion lifecycle it exposes

### 2. Excel host is missing `referencedCells`

Files:

- [apps/excel-addin/src/taskpane/taskpane.js](<repo-root>/apps/excel-addin/src/taskpane/taskpane.js)
- [apps/google-sheets-addon/src/Code.gs](<repo-root>/apps/google-sheets-addon/src/Code.gs)

What is true:

- Google Sheets now extracts explicit A1 references from prompts and includes `context.referencedCells`.
- Excel host still sends only `selection` and `activeCell`.

Why it matters:

- formula-help / cell-specific prompts behave differently by host
- the same prompt can succeed in Google Sheets and degrade in Excel

### 3. `sheet_import_plan` write-back shape is off by one row in both hosts

Files:

- [packages/contracts/src/schemas.ts](<repo-root>/packages/contracts/src/schemas.ts)
- [apps/excel-addin/src/taskpane/taskpane.js](<repo-root>/apps/excel-addin/src/taskpane/taskpane.js)
- [apps/google-sheets-addon/src/Code.gs](<repo-root>/apps/google-sheets-addon/src/Code.gs)

What is true:

- the contract clearly defines `shape.rows = 1 + values.length` when headers are present
- both hosts validate the target range against `plan.shape.rows`
- both hosts then write `[headers] + values`, which is also `1 + values.length` rows

The risky part:

- host code returns `writtenRows: plan.shape.rows`, which matches the contract
- but this path should still be regression-tested directly because the header-row convention is easy to break and has already been a source of confusion

Status:

- this is a real high-risk area
- but the specific claim that the contract forgot `1 + values.length` was a false positive from the external review

### 4. Request normalization in the gateway is broad and brittle

Files:

- [services/gateway/src/routes/requests.ts](<repo-root>/services/gateway/src/routes/requests.ts)

What is true:

- `normalizeHermesRequestInput(...)` currently strips `null` recursively before schema validation
- this was introduced to absorb stale Google Sheets client payloads with `null` optional fields
- it is wider than ideal and could erase future legitimate nullable fields

Why it matters:

- it is a pragmatic compatibility patch, not a clean contract boundary
- future schema evolution could be masked by this normalization

## Provisional findings that still deserve verification

### 5. Demo labeling enforcement may miss some response shapes

Files:

- [services/gateway/src/lib/hermesClient.ts](<repo-root>/services/gateway/src/lib/hermesClient.ts)

Why it is provisional:

- current enforcement keys off `data.extractionMode === "demo"`
- this is fine for extraction-bearing structured data
- if product requirements ever allow demo-labeled `chat` or `formula` responses, the current guard may be too narrow

### 6. Reviewer-safe unavailable may still be narrower than ideal

Files:

- [services/gateway/src/lib/hermesClient.ts](<repo-root>/services/gateway/src/lib/hermesClient.ts)

Why it is provisional:

- the strongest protection today is around `sheet_import_plan` and `extracted_table`
- if image-derived `attachment_analysis` or `document_summary` becomes user-visible in reviewer-safe unavailable mode, enforcement may need to expand

## False positive from the external review

The external review claimed `SheetImportPlanDataSchema` did not validate header-row semantics and row counts.

That claim is false for the current repo state.

Actual source already does:

- `shape.columns === headers.length`
- `shape.rows === 1 + values.length`
- `validateTargetRangeMatchesShape(...)`

See:

- [packages/contracts/src/schemas.ts](<repo-root>/packages/contracts/src/schemas.ts)

## Recommended next review passes

1. Write-back lifecycle and replay resistance only
2. Excel/Google host parity only
3. Reviewer-safe enforcement only
