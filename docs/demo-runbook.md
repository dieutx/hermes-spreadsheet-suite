# Demo Runbook

This runbook is for hackathon/demo presentation only. It reflects the repo as implemented after Batches 1–3.

## 1. Local setup

1. Install dependencies:

```bash
npm install
```

2. Copy `.env.example` to `.env`.

3. Set the gateway variables at minimum:

```bash
PORT=8787
GATEWAY_PUBLIC_BASE_URL=http://127.0.0.1:8787
HERMES_SERVICE_LABEL=spreadsheet-gateway
HERMES_ENVIRONMENT_LABEL=demo-review
# Required: replace with a long random value; never commit a real secret.
APPROVAL_SECRET=<REPLACE_ME_LONG_RANDOM_SECRET>
HERMES_AGENT_BASE_URL=http://127.0.0.1:9000
```

4. Start the gateway:

```bash
npm run dev:gateway
```

5. Confirm the gateway is up:

```bash
curl http://127.0.0.1:8787/health
```

Expected result:

- `ok: true`
- `service` matches `HERMES_SERVICE_LABEL`
- `environment` matches `HERMES_ENVIRONMENT_LABEL`

## 2. Connect the gateway to Hermes Agent

The remote Hermes Agent API server must be reachable at:

```text
POST {HERMES_AGENT_BASE_URL}/v1/spreadsheet-assistant
```

It must accept the Step 2 request envelope and return a Step 1 response envelope.

If your remote Hermes Agent deployment uses the example sidecars from this repo, start them separately:

Selection sidecar:

```bash
PORT=8791 npm run dev:selection-skill
```

Screenshot sidecar in real mode:

```bash
PORT=8792 IMAGE_EXTRACTION_MODE=real REVIEWER_SAFE_MODE=false VISION_EXTRACTOR_URL=https://your-vision-bridge.example/extract npm run dev:table-skill
```

## 3. Launch the Excel host

1. Serve `apps/excel-addin/src/taskpane/` and `apps/excel-addin/src/commands/` over HTTPS at the origin referenced by `apps/excel-addin/manifest.xml`.
2. Update `manifest.xml` if your origin changes.
3. Sideload the manifest into Excel.
4. Open the Hermes task pane.

Optional Excel runtime toggles can be set in the task pane origin’s `localStorage`:

```js
localStorage.setItem("hermesGatewayBaseUrl", "http://127.0.0.1:8787");
localStorage.setItem("hermesReviewerSafeMode", "false");
localStorage.setItem("hermesForceExtractionMode", "");
```

Reload the task pane after changing them.

## 4. Launch the Google Sheets host

### Google Sheets live-demo deploy

Use the local repo as the source of truth for the disposable workbook rollout.

Preflight: seed and populate every `Demo_*` tab with disposable sample data before you run any prompts. Do not skip this step; the live demo depends on those tabs existing with realistic sample content.

1. Open the provided Google Sheets demo workbook.
2. Seed/populate the `Demo_*` tabs with disposable sample data.
3. Open the bound Apps Script project from `Extensions -> Apps Script`.
4. Replace the bound project files with the local versions from:
   - `apps/google-sheets-addon/appsscript.json`
   - `apps/google-sheets-addon/src/Code.gs`
   - `apps/google-sheets-addon/src/ReferencedCells.js`
   - `apps/google-sheets-addon/src/Wave1Plans.js`
   - `apps/google-sheets-addon/html/Sidebar.html`
   - `apps/google-sheets-addon/html/Sidebar.css.html`
   - `apps/google-sheets-addon/html/Sidebar.js.html`
5. Set these Apps Script script properties:
   - `HERMES_GATEWAY_URL=http://127.0.0.1:8787`
   - `HERMES_CLIENT_VERSION=google-sheets-addon-live-demo`
   - optionally `HERMES_REVIEWER_SAFE_MODE=false`
   - optionally `HERMES_FORCE_EXTRACTION_MODE=real`
   - optionally reviewer/extraction flags required by the remote Hermes Agent deployment
6. Reload the spreadsheet and confirm the `Hermes` menu still appears.
7. Open the sidebar from the `Hermes` menu.

## 5. Flow 1 — explain/summarize current selection

### Setup

- Open Excel or Google Sheets.
- Select a small range that has visible headers and numeric data.
- Make sure the gateway and remote Hermes Agent API server are both running.

### Presenter action

Type:

```text
Explain the current selection.
```

### Expected visible UI behavior

Immediately:

- a user message bubble appears
- a single assistant placeholder appears with `Thinking...`

While processing:

- the muted status line updates from public trace events only

On completion:

- `Thinking...` is replaced by the final Hermes answer
- the proof line appears below the answer

### Expected proof/status metadata

The assistant response should visibly include:

- `requestId`
- `hermesRunId`
- `serviceLabel`
- `environmentLabel`
- `durationMs`

The response may also show:

- `skillsUsed`
- `downstreamProvider.label`
- `downstreamProvider.model`
- `confidence`

### Expected result

- No spreadsheet write occurs.
- Formula suggestions are read-only assistance and do not require confirmation or write back.
- The result is a `chat` response or equivalent read-only answer.

## 6. Flow 2 — real image extraction mode

### Setup

- Remote Hermes Agent is configured for real extraction.
- If your Hermes Agent deployment uses the example screenshot sidecar, start it with:

```bash
PORT=8792 IMAGE_EXTRACTION_MODE=real REVIEWER_SAFE_MODE=false VISION_EXTRACTOR_URL=https://your-vision-bridge.example/extract npm run dev:table-skill
```

- Excel:

```js
localStorage.setItem("hermesReviewerSafeMode", "false");
localStorage.removeItem("hermesForceExtractionMode");
```

- Google Sheets script properties:
  - `HERMES_REVIEWER_SAFE_MODE=false`
  - `HERMES_FORCE_EXTRACTION_MODE=` blank

### Presenter action

1. Attach or paste a PNG/JPG/JPEG/WEBP table screenshot.
2. Type:

```text
Extract this table and put it into Demo_ImageImport starting at B4.
```

### Expected visible UI behavior

Before send:

- image attachment chip is visible
- thumbnail preview is visible
- attachment can be removed

After send:

- `Thinking...` appears
- public trace/status updates appear

On completion:

- the response body says Hermes prepared an import preview
- the preview table renders with `headers` above `values`
- target sheet and full final `targetRange` are shown
- shape is visible in compact metadata
- a confirm button is shown

### What must happen before confirmation

- No write is performed.
- The user sees the preview only.
- The proof line shows remote Hermes Agent processing metadata.

### What must happen after confirmation

- The host requests approval from the gateway.
- The host writes only after approval.
- A success status line appears, for example:

```text
Write applied to Demo_ImageImport!B4:D8
```

## 7. Flow 2 — demo mode

### Setup

- Remote Hermes Agent is configured to emit demo-labeled extraction output.
- If your Hermes Agent deployment uses the example screenshot sidecar, start it with:

```bash
PORT=8792 IMAGE_EXTRACTION_MODE=demo REVIEWER_SAFE_MODE=false npm run dev:table-skill
```

- Excel:

```js
localStorage.setItem("hermesReviewerSafeMode", "false");
localStorage.setItem("hermesForceExtractionMode", "demo");
```

- Google Sheets script properties:
  - `HERMES_REVIEWER_SAFE_MODE=false`
  - `HERMES_FORCE_EXTRACTION_MODE=demo`

### Presenter action

1. Attach a table screenshot.
2. Type:

```text
Show me a demo extraction preview for this image in Demo_ImageImport starting at B4.
```

### Expected visible UI behavior

- `Thinking...` appears
- proof line is still shown
- the preview table may be shown
- the response warnings explicitly say the output is demo-only
- compact metadata shows `extraction demo`

### Reviewer-visible proof

The reviewer should be able to see:

- `requestId`
- `hermesRunId`
- `serviceLabel`
- `environmentLabel`
- demo warnings
- `extraction demo`

### What must not happen

- The UI must not present demo output as real extraction truth.

## 8. Flow 2 — reviewer-safe unavailable mode

### Setup

- Configure the host to force unavailable mode.

Excel:

```js
localStorage.setItem("hermesReviewerSafeMode", "true");
localStorage.setItem("hermesForceExtractionMode", "unavailable");
```

Google Sheets script properties:

- `HERMES_REVIEWER_SAFE_MODE=true`
- `HERMES_FORCE_EXTRACTION_MODE=unavailable`

- If you also use the example screenshot sidecar behind Hermes and want it to fail closed:

```bash
PORT=8792 IMAGE_EXTRACTION_MODE=disabled REVIEWER_SAFE_MODE=true npm run dev:table-skill
```

### Presenter action

1. Attach a table screenshot.
2. Type:

```text
Extract this table and put it into Demo_ImageImport starting at B4.
```

### Expected visible UI behavior

- `Thinking...` appears
- proof line is still shown
- the final assistant response is an unavailable/error path
- no extracted table preview is shown
- no sheet import preview is shown
- no confirm button is shown

### Expected proof/status metadata

The reviewer should still see:

- `requestId`
- `hermesRunId`
- `serviceLabel`
- `environmentLabel`

### What must not happen

- No fabricated extracted table
- No fabricated import plan
- No write-back control

## 9. Exact proof fields to show a reviewer live

During the demo, pause on the assistant response and point out:

- `Processed by Hermes`
- `requestId ...`
- `hermesRunId ...`
- `service ...`
- `environment ...`
- `duration ...`
- trace timeline
- warnings when demo/unavailable modes are active

## 10. Recommended presenter order

1. Start with Flow 1 on a live selection.
2. Show the `Thinking...` placeholder and compact proof line.
3. Show Flow 2 in real mode.
4. Stop before confirmation and point at the preview and full `targetRange`.
5. Confirm the write and show the final write status line.
6. Run demo mode and show explicit demo warnings.
7. Run reviewer-safe unavailable mode and show that no fake extraction preview appears.
