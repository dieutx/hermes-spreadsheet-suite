# Codex Implementation Prompt — Upgrade Hermes Spreadsheet Suite for Real Remote Spreadsheet Reasoning

You are working in this repository:

`/root/claude/hermes-spreadsheet-suite`

Your job is to upgrade the gateway/client flow so the spreadsheet assistant behaves like a real remote Hermes-powered spreadsheet AI, especially for formula debugging and fix/apply workflows.

This is an external integration task only.
Do NOT modify Hermes core.
Do NOT assume you can patch Hermes Agent source.
All changes must remain inside this repository and any external config/env/runtime integration points.

## Product goal

We have:
- Hermes Agent API server running on Linux server
- gateway in this repo
- Google Sheets / Excel clients that should act as thin frontends

We want this UX:

1. User types in a spreadsheet sidebar/add-in:
   - `why cell H11 error?`
2. Client sends spreadsheet context to the gateway
3. Gateway forwards the request to Hermes Agent via OpenAI-compatible Chat Completions
4. Hermes returns a structured response
5. Client renders the answer in the chat panel

Example desired answer:

> H11 is erroring because its formula is invalid: `=SUMIF(B1:F5)`. `SUMIF` needs at least 2 required arguments: `range` and `criterion`, with optional `sum_range`. But H11 only has 1 argument (`B1:F5`), so the function call is incomplete. Root cause: wrong number of arguments for `SUMIF`. Correct pattern: `=SUMIF(range, criterion)` or `=SUMIF(range, criterion, sum_range)`.

Then if the user says:
- `fix this for me. i want sumif revenue of cable product`

Hermes should return a structured write proposal (`sheet_update`) that the client can preview and confirm before applying.

## Current problem

Right now the system is still too weak for spreadsheet reasoning and often fails with:

`Hermes Agent returned a response body that does not match the structured gateway contract.`

This indicates the gateway and the Hermes prompting/response shaping are not aligned strongly enough.

## Important architecture constraints

These are mandatory:

- Hermes core is immutable.
- Do not modify Hermes core source.
- The gateway must be the place where spreadsheet product runtime rules are layered in.
- The client must not do core reasoning locally.
- The client may only:
  - collect spreadsheet context
  - collect attachment metadata
  - send requests
  - render trace/proof
  - render confirmation UI
  - execute approved write-back through host APIs
- The gateway must call the Hermes Agent API server using OpenAI-compatible `/v1/chat/completions`.
- The gateway must inject spreadsheet-specific runtime rules as a system prompt.
- No chain-of-thought or hidden reasoning may ever be exposed.

## Current environment assumptions

Assume the following are already true or should remain true:

- Hermes Agent API server base URL is configured via env:
  - `HERMES_AGENT_BASE_URL=http://127.0.0.1:8642/v1`
- Hermes API auth key is configured via env:
  - `HERMES_API_SERVER_KEY=<REPLACE_ME_OPTIONAL_API_KEY>`
- model is:
  - `HERMES_AGENT_MODEL=hermes-agent`
- spreadsheet runtime rules file exists or should exist at:
  - `services/gateway/src/hermes/runtimeRules.ts`

## What to study first

Before changing anything, inspect at least these files:

- `services/gateway/src/lib/hermesClient.ts`
- `services/gateway/src/hermes/requestTemplate.ts`
- `services/gateway/src/hermes/structuredBody.ts`
- `packages/contracts/src/schemas.ts`
- `packages/shared-client/src/request.ts`
- host adapters / Google Sheets / Excel request builders under this repo
- any routes involved in request submission, trace polling, uploads, or write-back approval

## High-level diagnosis you should confirm

The likely root issues are:

1. Hermes prompt contract is not strong enough
- Hermes is not being constrained enough to return exactly the gateway’s internal structured body format.
- It may return prose, fenced JSON, full envelope JSON, or otherwise malformed content.

2. Formula debugging lacks rich spreadsheet evidence
- For prompts like `why cell H11 error?`, the client/gateway often does not send enough cell-specific context.
- The request contract already supports:
  - `selection.formulas`
  - `activeCell`
  - `referencedCells`
- But the host/client may not actually populate them well enough.

3. The response-type routing is too simplistic
- The current request prompt appears to heuristically prefer `chat` vs `formula` based on regex.
- But we need stronger behavior:
  - explain/diagnose -> `chat` or `formula`
  - fix/apply/propose update -> `sheet_update`

4. Image flows must remain safe and reviewer-compatible
- reviewer-safe mode must never fabricate extracted tables as if they were real.
- demo and unavailable modes must be explicitly surfaced.

## Required outcome

Implement the necessary upgrades so that:

### Formula diagnosis works well
If the user asks:
- `why cell H11 error?`
and the request includes H11 formula metadata, Hermes should directly analyze the formula and explain the spreadsheet error.

### Formula fix proposal works well
If the user asks:
- `fix this for me. i want sumif revenue of cable product`
Hermes should return a valid `sheet_update` structured body when there is enough context to propose a fix.

### Gateway contract errors stop happening
The Hermes response path should be strongly aligned so valid spreadsheet outputs pass `HermesStructuredBodySchema`.

---

# Implementation requirements

## 1. Strengthen the runtime rules

Update `services/gateway/src/hermes/runtimeRules.ts` so the rules are strong and explicit about spreadsheet behavior.

The runtime rules must clearly state:

- Hermes is the real remote execution/reasoning engine
- clients are thin frontends only
- backend/gateway validates, forwards, relays proof, and controls safe write-back execution
- Hermes core is immutable and must never be modified
- responses must align to the spreadsheet response contract types:
  - `chat`
  - `formula`
  - `sheet_update`
  - `sheet_import_plan`
  - `error`
  - `attachment_analysis`
  - `extracted_table`
  - `document_summary`
- no chain-of-thought or hidden reasoning may ever be exposed
- spreadsheet write-back is always proposal-first, never automatic
- `targetRange` means the full destination rectangle, not just anchor cell
- `sheet_import_plan.headers` are separate from `values`
- `shape.rows` includes the header row for import plans
- extraction modes:
  - `real`
  - `demo`
  - `unavailable`
- reviewer-safe mode must never fabricate real extraction output
- UI is minimal and chat-first; do not emit dashboard-style noise

Make sure the runtime rules are optimized for spreadsheet diagnosis/fix behavior, not generic chatbot behavior.

## 2. Rewrite the gateway request template to be much more constraining

Update `services/gateway/src/hermes/requestTemplate.ts`.

The prompt sent to Hermes must:

- require JSON only
- require exactly one JSON object and nothing else
- forbid markdown fences
- forbid prose before/after the JSON
- explicitly say the response must match exactly one internal structured body type
- explicitly forbid external Step 1 envelope fields in the assistant output, because the gateway wraps them later
- explicitly say no hidden reasoning or chain-of-thought may be exposed

### Add stronger behavior instructions
The prompt should clearly instruct Hermes:

- If the user explicitly provides a formula or exact spreadsheet error text, treat that as authoritative evidence.
- If `context.activeCell` or `context.referencedCells` contains formula/value/note metadata, use that directly.
- Do not say “insufficient context” when formula syntax can be diagnosed from the provided formula text or cell metadata.
- Validate function signature first before using sheet preview heuristics.
- If the user asks to explain/diagnose, prefer `chat` or `formula`.
- If the user asks to fix/apply/write/change a formula and there is enough target context, prefer `sheet_update`.

### Add explicit few-shot examples in the request template
Add several realistic examples directly into the prompt template, including at minimum:

#### Example A — formula error explanation
User asks:
- `why cell H11 error?`
Context includes referenced cell formula:
- `=SUMIF(B1:F5)`
Expected structured body should be valid and should explain that `SUMIF` has too few arguments.

#### Example B — formula fix proposal
User asks:
- `fix this for me. i want sumif revenue of cable product`
Context includes:
- target cell H11
- headers including Product and Revenue
Expected result should be a valid `sheet_update` body with:
- `targetSheet`
- `targetRange`
- `operation: "set_formulas"` or another valid operation consistent with current schema
- formula matrix
- explanation
- confidence
- `requiresConfirmation: true`

#### Example C — image extraction unavailable in reviewer-safe mode
Expected result should be `error` or another contract-valid safe path, not fabricated extraction output.

## 3. Improve response-type routing logic

Review and improve any simplistic preferred response type heuristics.

The system should map requests more intelligently:

- explain / summarize / diagnose -> `chat` or `formula`
- suggest formula -> `formula`
- fix / apply / update cell / write this -> `sheet_update` if target context is sufficient
- attachment/image analysis -> `attachment_analysis`, `extracted_table`, or `sheet_import_plan`
- unavailable capability -> `error`

Do not rely only on a regex like “message contains sum/formula => formula”.
Use intent-sensitive logic.

## 4. Ensure host/client context is rich enough for spreadsheet debugging

Inspect the host/shared client request-building path.

The request contract already supports:
- `selection.formulas`
- `activeCell`
- `referencedCells`

Make sure the implementation actually populates these fields where possible.

### Required behavior for spreadsheet prompts mentioning cells
If the user message references cells like:
- `H11`
- `F12`
- `B3`

then the host/gateway path should try to include metadata for those cells in `referencedCells`, including when available:
- `a1Notation`
- `displayValue`
- `value`
- `formula`
- `note`

Also include `activeCell` where available.

For selected ranges, include:
- `selection.range`
- `selection.headers`
- `selection.values`
- `selection.formulas` when feasible

This is essential for prompts like:
- `why cell h11 error?`
- `fix this`

Without this data, Hermes will continue to underperform.

## 5. Preserve strict contract compliance with the gateway structured body parser

The gateway expects an internal structured body first, not the full public response envelope.

Make sure Hermes is constrained to output only the internal structured body shape expected by:
- `services/gateway/src/hermes/structuredBody.ts`

That means valid output should look like:
- `type`
- `data`
- optional:
  - `warnings`
  - `skillsUsed`
  - `downstreamProvider`

and nothing else.

Do not allow Hermes to emit:
- full external response envelope
- `schemaVersion`
- `requestId`
- `trace`
- `ui`
- `startedAt`
- `completedAt`
- `durationMs`
- or raw prose outside the JSON object

## 6. Keep image/extraction modes safe and explicit

Preserve and enforce:
- `real`
- `demo`
- `unavailable`

In reviewer-safe mode:
- never fabricate real extraction output
- do not pretend extraction succeeded if mode is `unavailable`
- if mode is `demo`, label it clearly via warnings and proper response typing

## 7. Preserve minimal chat-first UI compatibility

Do not redesign the product into a dashboard.
The resulting behavior must remain compatible with:
- user message
- one `Thinking...` placeholder
- final response replacement
- compact proof/status line
- optional warnings
- structured preview only when the response type requires it

## 8. Do not modify Hermes core

All fixes must remain external.
If anything currently assumes otherwise, refactor the local repo implementation so the integration remains outside Hermes core.

---

# Deliverables

Implement the changes in this repository and provide a final summary including:

1. Root cause analysis
- Why the gateway contract mismatch was happening
- Why spreadsheet reasoning was weak before

2. Exact files changed
- list all file paths

3. Diff summary
- what changed in runtime rules
- what changed in request template
- what changed in request/context building
- what changed in response routing

4. Verification results
At minimum, show or describe tests for:

### Test 1 — formula error diagnosis
A request representing:
- `why cell H11 error?`
with H11 formula `=SUMIF(B1:F5)`
should produce a valid structured body that explains the wrong number of arguments.

### Test 2 — formula fix proposal
A request representing:
- `fix this for me. i want sumif revenue of cable product`
with enough target context
should produce a valid `sheet_update` structured body.

### Test 3 — reviewer-safe unavailable extraction
A request in reviewer-safe unavailable mode must not fabricate extracted table output.

### Test 4 — structured-body validation
Show that the responses pass the internal structured body validation path and no longer trigger:
- `Hermes Agent returned a response body that does not match the structured gateway contract.`

---

# Extra notes

- Prefer small, correct changes over broad rewrites.
- Preserve the current architecture where the gateway calls Hermes Agent via `/v1/chat/completions`.
- Keep the implementation practical and production-oriented.
- If you discover a mismatch between repo docs and code, note it in your summary and fix the repo-side implementation/docs if necessary, but do not change Hermes core.

Now inspect the codebase, make the required repo-local upgrades, and return the implementation summary with verification evidence.
