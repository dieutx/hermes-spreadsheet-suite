import { randomUUID } from "node:crypto";
import { Router } from "express";
import { HermesRequestSchema } from "@hermes/contracts";
import type { HermesRequest } from "@hermes/contracts";
import type { GatewayConfig } from "../lib/config.js";
import { formatClientIssuePath, sanitizeClientIssueMessage } from "../lib/publicErrors.js";
import { AttachmentStore } from "../lib/store.js";
import type { TraceBus } from "../lib/traceBus.js";

type JsonRecord = Record<string, unknown>;

const PUBLIC_RUN_ERROR_FALLBACK =
  "The gateway hit an unexpected error while processing the request.";
const UNSAFE_RUN_ERROR_PATTERN = /\b(?:APPROVAL_SECRET|HERMES_API_SERVER_KEY|HERMES_AGENT_API_KEY|HERMES_AGENT_BASE_URL|OPENAI_API_KEY|ANTHROPIC_API_KEY|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)\b|(?:^|\s)\/(?:root|srv|home|tmp)\/[^\s]+|https?:\/\/[^\s]+/i;
const NULL_OPTIONAL_PATHS = new Set([
  "context.selection.headers",
  "context.activeCell.formula",
  "context.activeCell.note",
  "context.referencedCells.*.formula",
  "context.referencedCells.*.note"
]);
const REQUEST_STATUS_ID_MAX_LENGTH = 128;
const MAX_REQUEST_ID_LENGTH = 128;
const PUBLIC_REQUEST_STATUS_ID_PATTERN = /^[A-Za-z0-9._:-]+$/;
const PUBLIC_REQUEST_ID_PATTERN = /^[A-Za-z0-9._:-]+$/;

function isObject(value: unknown): value is JsonRecord {
  return typeof value === "object" && value !== null;
}

function pathMatches(pattern: string, path: string[]): boolean {
  const expectedSegments = pattern.split(".");
  if (expectedSegments.length !== path.length) {
    return false;
  }

  return expectedSegments.every((segment, index) =>
    segment === "*" || segment === path[index]
  );
}

function shouldOmitNullAtPath(path: string[]): boolean {
  for (const pattern of NULL_OPTIONAL_PATHS) {
    if (pathMatches(pattern, path)) {
      return true;
    }
  }

  return false;
}

function normalizeHermesRequestNode(value: unknown, path: string[]): unknown {
  if (Array.isArray(value)) {
    return value.map((item, index) => normalizeHermesRequestNode(item, [...path, String(index)]));
  }

  if (!isObject(value)) {
    return value;
  }

  const normalized: JsonRecord = {};
  for (const [key, child] of Object.entries(value)) {
    const childPath = [...path, key];
    if (child === null && shouldOmitNullAtPath(childPath)) {
      continue;
    }

    normalized[key] = normalizeHermesRequestNode(child, childPath);
  }

  return normalized;
}

function normalizeHermesRequestInput(value: unknown): unknown {
  if (!isObject(value)) {
    return value;
  }

  const normalized = normalizeHermesRequestNode(value, []);
  if (!isObject(normalized)) {
    return normalized;
  }

  const reviewer = isObject(normalized.reviewer) ? { ...normalized.reviewer } : undefined;
  if (reviewer && !Object.prototype.hasOwnProperty.call(reviewer, "forceExtractionMode")) {
    reviewer.forceExtractionMode = null;
  }

  return {
    ...normalized,
    ...(reviewer ? { reviewer } : {})
  };
}

function buildCanonicalAttachments(
  attachments: HermesRequest["context"]["attachments"] | undefined,
  requestSessionId: string | undefined,
  requestWorkbookId: string | undefined,
  attachmentStore: AttachmentStore
): {
  canonicalAttachments: HermesRequest["context"]["attachments"];
  missingAttachmentId?: string;
} {
  if (!attachments || attachments.length === 0) {
    return { canonicalAttachments: undefined };
  }

  const normalizedSessionId = typeof requestSessionId === "string" && requestSessionId.trim().length > 0
    ? requestSessionId.trim()
    : undefined;
  const normalizedWorkbookId = typeof requestWorkbookId === "string" && requestWorkbookId.trim().length > 0
    ? requestWorkbookId.trim()
    : undefined;
  const canonicalAttachments: NonNullable<HermesRequest["context"]["attachments"]> = [];
  for (const attachment of attachments) {
    const stored = attachmentStore.get(attachment.id);
    const uploadToken = typeof attachment.uploadToken === "string" && attachment.uploadToken.trim().length > 0
      ? attachment.uploadToken.trim()
      : undefined;
    if (
      !stored ||
      !uploadToken ||
      uploadToken !== stored.metadata.uploadToken ||
      !normalizedSessionId ||
      !stored.sessionId ||
      stored.sessionId !== normalizedSessionId ||
      !normalizedWorkbookId ||
      !stored.workbookId ||
      stored.workbookId !== normalizedWorkbookId
    ) {
      return { canonicalAttachments: undefined, missingAttachmentId: attachment.id };
    }

    canonicalAttachments.push({
      id: stored.metadata.id,
      type: stored.metadata.type,
      mimeType: stored.metadata.mimeType,
      fileName: stored.metadata.fileName,
      size: stored.metadata.size,
      source: stored.metadata.source,
      previewUrl: stored.metadata.previewUrl,
      uploadToken: stored.metadata.uploadToken,
      storageRef: stored.metadata.storageRef
    });
  }

  return { canonicalAttachments };
}

function formatIssues(issues: Array<{ path: (string | number)[]; message: string }>) {
  return issues.map((issue) => ({
    path: formatClientIssuePath(issue.path),
    message: sanitizeClientIssueMessage(issue.message)
  }));
}

function isSingleCellRange(value: string | undefined): boolean {
  if (typeof value !== "string") {
    return false;
  }

  return /^\$?[A-Z]{1,3}\$?[1-9][0-9]*$/i.test(value.trim());
}

function parseRequiredStatusIdentifier(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const normalized = value.trim();
  if (
    normalized.length === 0 ||
    normalized.length > REQUEST_STATUS_ID_MAX_LENGTH ||
    !PUBLIC_REQUEST_STATUS_ID_PATTERN.test(normalized)
  ) {
    return undefined;
  }

  return normalized;
}

function parseOptionalStatusIdentifier(value: unknown): { ok: true; value?: string } | { ok: false } {
  if (value === undefined) {
    return { ok: true };
  }

  if (typeof value !== "string") {
    return { ok: false };
  }

  const normalized = value.trim();
  if (normalized.length === 0) {
    return { ok: true };
  }

  if (
    normalized.length > REQUEST_STATUS_ID_MAX_LENGTH ||
    !PUBLIC_REQUEST_STATUS_ID_PATTERN.test(normalized)
  ) {
    return { ok: false };
  }

  return { ok: true, value: normalized };
}

function invalidRunStatusRequest() {
  return {
    error: {
      code: "INVALID_REQUEST",
      message: "Run status identifiers are invalid.",
      userAction: "Retry status polling from the current Hermes session."
    }
  };
}

function getSafeInvalidRequestId(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 &&
    trimmed.length <= MAX_REQUEST_ID_LENGTH &&
    PUBLIC_REQUEST_ID_PATTERN.test(trimmed)
    ? trimmed
    : undefined;
}

function shouldIncludeResponseTrace(value: unknown): boolean {
  if (value === undefined) {
    return true;
  }

  if (typeof value !== "string") {
    return true;
  }

  const normalized = value.trim().toLowerCase();
  return normalized !== "0" && normalized !== "false" && normalized !== "no";
}

function getPublicRunError(error: string | undefined): string | undefined {
  if (!error) {
    return undefined;
  }

  const message = error
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!message || UNSAFE_RUN_ERROR_PATTERN.test(message)) {
    return PUBLIC_RUN_ERROR_FALLBACK;
  }

  return message.length > 4000 ? message.slice(0, 4000) : message;
}

function stripTraceFromResponse<T>(response: T, includeTrace: boolean): T {
  if (includeTrace || !response || typeof response !== "object" || Array.isArray(response)) {
    return response;
  }

  const candidate = response as { trace?: unknown; ui?: unknown };
  return {
    ...candidate,
    trace: [],
    ui:
      candidate.ui && typeof candidate.ui === "object" && !Array.isArray(candidate.ui)
        ? { ...(candidate.ui as Record<string, unknown>), showTrace: false }
        : candidate.ui
  } as T;
}

function matchesStoredRunRequestId(
  run: ReturnType<TraceBus["getRun"]>,
  requestId: string | undefined
): boolean {
  if (!run?.requestId) {
    return true;
  }

  return requestId === run.requestId;
}

function matchesStoredRunSessionId(
  run: ReturnType<TraceBus["getRun"]>,
  sessionId: string | undefined
): boolean {
  if (!run?.sessionId) {
    return true;
  }

  return sessionId === run.sessionId;
}

export function createRequestRouter(input: {
  traceBus: TraceBus;
  hermesClient: {
    processRequest: (input: {
      runId: string;
      request: HermesRequest;
      traceBus: TraceBus;
    }) => Promise<void>;
  };
  attachmentStore: AttachmentStore;
  config: GatewayConfig;
}): Router {
  const router = Router();

  router.post("/", (req, res) => {
    const normalizedBody = normalizeHermesRequestInput(req.body);
    const parsed = HermesRequestSchema.safeParse(normalizedBody);

    if (!parsed.success) {
      const requestId = isObject(normalizedBody)
        ? getSafeInvalidRequestId(normalizedBody.requestId)
        : undefined;
      console.warn("[gateway] invalid Step 2 request", {
        requestId,
        issues: formatIssues(parsed.error.issues)
      });
      res.status(400).json({
        requestId,
        error: {
          code: "INVALID_REQUEST",
          message: "Hermes couldn't prepare a valid request from the current spreadsheet state.",
          userAction: "Refresh the sheet state and try again. If it keeps failing, retry with a smaller prompt or reselect the relevant range.",
          issues: formatIssues(parsed.error.issues)
        }
      });
      return;
    }

    const requestEnvelope = parsed.data;
    const {
      canonicalAttachments,
      missingAttachmentId
    } = buildCanonicalAttachments(
      requestEnvelope.context.attachments,
      requestEnvelope.source.sessionId,
      requestEnvelope.host.workbookId,
      input.attachmentStore
    );

    if (missingAttachmentId) {
      res.status(400).json({
        requestId: requestEnvelope.requestId,
        error: {
          code: "ATTACHMENT_UNAVAILABLE",
          message: "I can't access that uploaded file anymore.",
          userAction: "Reattach the file, then tell me the target sheet or range if you want me to paste or import its contents."
        }
      });
      return;
    }

    const requestEnvelopeWithCanonicalAttachments: HermesRequest = {
      ...requestEnvelope,
      context: {
        ...requestEnvelope.context,
        attachments: canonicalAttachments
      }
    };

    const runId = `run_${randomUUID()}`;
    input.traceBus.ensureRun(
      runId,
      requestEnvelopeWithCanonicalAttachments.requestId,
      requestEnvelopeWithCanonicalAttachments.source.sessionId
    );
    input.traceBus.markStatus(runId, "accepted");

    input.traceBus.append(runId, {
      event: "request_received",
      timestamp: input.traceBus.nowIso()
    });

    const traceRange = requestEnvelopeWithCanonicalAttachments.context.currentRegion?.range &&
      isSingleCellRange(requestEnvelopeWithCanonicalAttachments.context.selection?.range)
      ? requestEnvelopeWithCanonicalAttachments.context.currentRegion.range
      : requestEnvelopeWithCanonicalAttachments.context.selection?.range ??
        requestEnvelopeWithCanonicalAttachments.host.selectedRange;

    if (traceRange || requestEnvelopeWithCanonicalAttachments.host.activeSheet) {
      input.traceBus.append(runId, {
        event: "spreadsheet_context_received",
        timestamp: input.traceBus.nowIso(),
        details: {
          range: traceRange,
          sheet: requestEnvelopeWithCanonicalAttachments.host.activeSheet
        }
      });
    }

    for (const attachment of canonicalAttachments ?? []) {
      input.traceBus.append(runId, {
        event: "attachment_received",
        timestamp: input.traceBus.nowIso(),
        details: {
          attachmentId: attachment.id
        }
      });

      input.traceBus.append(runId, {
        event: "image_received",
        timestamp: input.traceBus.nowIso(),
        details: {
          attachmentId: attachment.id
        }
      });
    }

    void input.hermesClient.processRequest({
      runId,
      request: requestEnvelopeWithCanonicalAttachments,
      traceBus: input.traceBus
    });

    res.status(202).json({
      requestId: requestEnvelopeWithCanonicalAttachments.requestId,
      runId,
      status: "accepted"
    });
  });

  router.get("/:runId", (req, res) => {
    const includeTrace = shouldIncludeResponseTrace(req.query.includeTrace);
    const runId = parseRequiredStatusIdentifier(req.params.runId);
    const requestId = parseOptionalStatusIdentifier(req.query.requestId);
    const sessionId = parseOptionalStatusIdentifier(req.query.sessionId);
    if (!runId || !requestId.ok || !sessionId.ok) {
      res.status(400).json(invalidRunStatusRequest());
      return;
    }

    const run = input.traceBus.peekRun(runId);
    if (
      !run ||
      !matchesStoredRunRequestId(run, requestId.value) ||
      !matchesStoredRunSessionId(run, sessionId.value)
    ) {
      res.status(404).json({
        error: {
          code: "RUN_NOT_FOUND",
          message: "That Hermes request is no longer available.",
          userAction: "Send the request again from the spreadsheet if you need a fresh result."
        }
      });
      return;
    }

    const activeRun = input.traceBus.getRun(runId);
    if (!activeRun) {
      res.status(404).json({
        error: {
          code: "RUN_NOT_FOUND",
          message: "That Hermes request is no longer available.",
          userAction: "Send the request again from the spreadsheet if you need a fresh result."
        }
      });
      return;
    }

    res.json({
      runId: req.params.runId,
      requestId: activeRun.requestId,
      hermesRunId: activeRun.hermesRunId,
      status: activeRun.status,
      startedAt: activeRun.startedAt,
      completedAt: activeRun.completedAt,
      response: stripTraceFromResponse(activeRun.response, includeTrace),
      error: getPublicRunError(activeRun.error)
    });
  });

  return router;
}
