# OpenRouter Claude Opus 4.7 Review Playbook

This playbook is for running an external code review on this repository using OpenRouter with Claude Opus 4.7.

Official references:

- Model page: https://openrouter.ai/anthropic/claude-opus-4.7
- Chat completions API: https://openrouter.ai/docs/api-reference/chat-completion
- API overview and optional attribution headers: https://openrouter.ai/docs/api-reference/overview

## Review objective

Ask Claude to review the repo like a strict senior engineer:

- primary focus: bugs, regressions, contract mistakes, reviewer-safe mistakes, missing tests
- secondary focus: architectural drift from the intended thin-host/thin-gateway design
- avoid generic praise, broad cleanup suggestions, or speculative feature ideas

## Required context

Always include:

- [repo-brief.md](<repo-root>/docs/review/repo-brief.md)
- [README.md](<repo-root>/README.md)
- [reviewer-checklist.md](<repo-root>/docs/reviewer-checklist.md)

Then include the most relevant code slices:

- [schemas.ts](<repo-root>/packages/contracts/src/schemas.ts)
- [structuredBody.ts](<repo-root>/services/gateway/src/hermes/structuredBody.ts)
- [hermesClient.ts](<repo-root>/services/gateway/src/lib/hermesClient.ts)
- [requests.ts](<repo-root>/services/gateway/src/routes/requests.ts)
- [uploads.ts](<repo-root>/services/gateway/src/routes/uploads.ts)
- [writeback.ts](<repo-root>/services/gateway/src/routes/writeback.ts)
- [Code.gs](<repo-root>/apps/google-sheets-addon/src/Code.gs)
- [Sidebar.js.html](<repo-root>/apps/google-sheets-addon/html/Sidebar.js.html)
- [taskpane.js](<repo-root>/apps/excel-addin/src/taskpane/taskpane.js)

## Recommended review question

Use Claude for one of these modes:

1. Full critical-path review
   Focus only on Flow 1, Flow 2 real, Flow 2 demo, Flow 2 reviewer-safe unavailable, contracts, and write-back safety.

2. Boundary review
   Focus only on request/response validation, normalization, and host-to-gateway mismatches.

3. Host parity review
   Focus only on mismatches between Google Sheets and Excel hosts.

## Required output format

Tell Claude to answer in this structure:

1. Findings
   Each finding must include severity, exact file/path, issue, impact, and smallest safe fix direction.
2. Open questions
   Only if evidence is incomplete.
3. Residual risk
   What is still risky even if no critical bug is found.

If there are no meaningful findings, Claude must say so explicitly.

## Context packing strategy

Do not dump the whole repo blindly. Pack only the files that define:

- contracts
- gateway boundary behavior
- host request generation
- tests around the critical flows

If token budget is tight, include tests after code, not before.

## Suggested local context bundle command

```bash
{
  echo '===== docs/review/repo-brief.md ====='
  cat docs/review/repo-brief.md
  echo
  echo '===== README.md ====='
  sed -n '1,260p' README.md
  echo
  echo '===== docs/reviewer-checklist.md ====='
  cat docs/reviewer-checklist.md
  echo
  echo '===== packages/contracts/src/schemas.ts ====='
  sed -n '1,320p' packages/contracts/src/schemas.ts
  echo
  echo '===== services/gateway/src/hermes/structuredBody.ts ====='
  sed -n '1,280p' services/gateway/src/hermes/structuredBody.ts
  echo
  echo '===== services/gateway/src/lib/hermesClient.ts ====='
  sed -n '1,360p' services/gateway/src/lib/hermesClient.ts
  echo
  echo '===== services/gateway/src/routes/requests.ts ====='
  sed -n '1,220p' services/gateway/src/routes/requests.ts
  echo
  echo '===== services/gateway/src/routes/uploads.ts ====='
  sed -n '1,220p' services/gateway/src/routes/uploads.ts
  echo
  echo '===== services/gateway/src/routes/writeback.ts ====='
  sed -n '1,260p' services/gateway/src/routes/writeback.ts
  echo
  echo '===== apps/google-sheets-addon/src/Code.gs ====='
  sed -n '1,320p' apps/google-sheets-addon/src/Code.gs
  echo
  echo '===== apps/google-sheets-addon/html/Sidebar.js.html ====='
  sed -n '1,320p' apps/google-sheets-addon/html/Sidebar.js.html
  echo
  echo '===== apps/excel-addin/src/taskpane/taskpane.js ====='
  sed -n '620,980p' apps/excel-addin/src/taskpane/taskpane.js
} > /tmp/hermes-review-context.txt
```

## Exact OpenRouter curl pattern

```bash
curl https://openrouter.ai/api/v1/chat/completions \
  -H "Authorization: Bearer $OPENROUTER_API_KEY" \
  -H "Content-Type: application/json" \
  -H "HTTP-Referer: https://local.repo.review" \
  -H "X-OpenRouter-Title: Hermes Spreadsheet Suite Review" \
  -d @/tmp/openrouter-hermes-review-request.json
```

## Model choice

Use:

```text
anthropic/claude-opus-4.7
```

This is the current OpenRouter model identifier for Claude Opus 4.7.
