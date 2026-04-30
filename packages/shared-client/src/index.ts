import {
  WritebackApprovalResponseSchema,
  WritebackCompletionResponseSchema
} from "@hermes/contracts";
import type {
  HermesRequest,
  ImageAttachment
} from "@hermes/contracts";
import { buildHermesRequest, createRequestId, sanitizeConversation } from "./request.js";
import {
  buildDataCleanupPreview,
  buildDataCleanupUpdatePreview,
  buildCompositePlanPreview,
  buildCompositeUpdatePreview,
  buildRangeTransferPreview,
  buildRangeTransferUpdatePreview,
  buildDataValidationPreview,
  buildDryRunPreview,
  buildExtractedTablePreview,
  buildPlanHistoryPreview,
  buildRangeFormatPreview,
  buildRangeFilterPreview,
  buildRangeSortPreview,
  buildWorkbookStructurePreview,
  buildNamedRangeUpdatePreview,
  buildSheetImportPreview,
  buildSheetStructurePreview,
  buildSheetUpdatePreview,
  buildStructuredPreview,
  buildWriteMatrix,
  getFollowUpSuggestions,
  getResponseBodyText,
  getResponseConfidence,
  getResponseMetaLine,
  getResponseWarnings,
  getRequiresConfirmation,
  getStructuredPreview,
  formatDryRunSummary,
  formatHistoryEntrySummary,
  isWritePlanResponse
} from "./render.js";
import {
  formatWritebackStatusLine,
  formatProofLine,
  formatTraceEvent,
  formatTraceTimeline,
  summarizeLatestTrace
} from "./trace.js";
import {
  filterSupportedImageFiles,
  isSupportedImageMimeType
} from "./attachments.js";
import type { GatewayClient } from "./types.js";

function tryParseJsonText<T>(text: string): T | undefined {
  try {
    return JSON.parse(text) as T;
  } catch {
    return undefined;
  }
}

async function parseJson<T>(response: Response): Promise<T> {
  const bodyText = await response.text();
  const trimmedBodyText = bodyText.trim();

  if (!response.ok) {
    const contentType = response.headers.get("content-type") ?? "";
    if (contentType.includes("application/json")) {
      const payload = tryParseJsonText<{
        error?: { message?: string; userAction?: string } | string;
        message?: string;
        userAction?: string;
      }>(bodyText);
      const message = typeof payload?.error === "string"
        ? payload.error
        : payload?.error?.message || payload?.message;
      const userAction = typeof payload?.error === "string"
        ? payload?.userAction
        : payload?.error?.userAction || payload?.userAction;
      if (typeof message === "string" && message.trim()) {
        const trimmedMessage = message.trim();
        const trimmedUserAction = typeof userAction === "string" ? userAction.trim() : "";
        throw new Error(
          trimmedUserAction && trimmedUserAction !== trimmedMessage
            ? `${trimmedMessage}\n\n${trimmedUserAction}`
            : trimmedMessage
        );
      }
    }

    if (/^<!doctype html/i.test(trimmedBodyText) || /^<html/i.test(trimmedBodyText)) {
      throw new Error(
        "The Hermes service returned an unexpected error page.\n\nRetry the request, then check the Hermes gateway if it keeps happening."
      );
    }

    if (trimmedBodyText && containsSensitiveGatewayErrorText(trimmedBodyText)) {
      throw new Error(formatRawGatewayTextFailure(response.status));
    }

    throw new Error(bodyText || `Request failed with ${response.status}`);
  }

  const parsed = tryParseJsonText<T>(bodyText);
  if (parsed === undefined) {
    throw new Error(
      "The Hermes service returned a response the client could not use.\n\nRetry the request, then reload the client if it keeps happening."
    );
  }

  return parsed;
}

function parseContractPayload<T>(
  payload: unknown,
  schema: { parse: (input: unknown) => T },
  invalidMessage: string
): T {
  try {
    return schema.parse(payload);
  } catch {
    throw new Error(invalidMessage);
  }
}

function containsSensitiveGatewayErrorText(text: string): boolean {
  return (
    /\b(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|approval_secret|HERMES_[A-Z0-9_]+)\s*[:=]/i.test(text) ||
    /\bBearer\s+[A-Za-z0-9._~+/-]+=*/i.test(text) ||
    /\bat\s+(?:file:\/\/\/|\/|[A-Za-z]:\\)/i.test(text) ||
    /(?:^|\s)\/(?:srv|var|tmp|root|home|Users|opt|workspace|app|mnt)\/[^\s]+(?::\d+)?/i.test(text) ||
    /(?:^|\s)[A-Za-z]:\\[^\s]+/.test(text) ||
    /https?:\/\/(?:localhost|127(?:\.\d{1,3}){3}|0\.0\.0\.0|10\.|192\.168\.|172\.(?:1[6-9]|2\d|3[01])\.|[^/\s]*internal|[^/\s]*\.local)(?:[/:]|\s|$)/i.test(text) ||
    /\b(?:stack trace|traceback)\b/i.test(text)
  );
}

function formatRawGatewayTextFailure(status: number): string {
  return `Hermes gateway request failed with HTTP ${status}.`;
}

export function createGatewayClient(baseUrl: string): GatewayClient {
  const normalizedBaseUrl = baseUrl.replace(/\/$/, "");

  return {
    async uploadImage(input) {
      const form = new FormData();
      form.set("file", input.file, input.fileName);
      form.set("source", input.source);
      form.set("sessionId", input.sessionId);
      form.set("workbookId", input.workbookId);

      const payload = await parseJson<{ attachment: ImageAttachment }>(
        await fetch(`${normalizedBaseUrl}/api/uploads/image`, {
          method: "POST",
          body: form
        })
      );

      return payload.attachment;
    },

    async startRun(request) {
      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/requests`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(request)
        })
      );
    },

    async pollRun(runId, requestId) {
      const params = new URLSearchParams();
      if (requestId) {
        params.set("requestId", requestId);
      }

      return parseJson(
        await fetch(
          `${normalizedBaseUrl}/api/requests/${runId}${params.size > 0 ? `?${params.toString()}` : ""}`
        )
      );
    },

    async pollTrace(runId, after = 0, requestId) {
      const params = new URLSearchParams({
        after: String(after)
      });
      if (requestId) {
        params.set("requestId", requestId);
      }

      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/trace/${runId}?${params.toString()}`)
      );
    },

    async dryRunPlan(input) {
      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/execution/dry-run`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(input)
        })
      );
    },

    async listPlanHistory(input) {
      const params = new URLSearchParams({
        workbookSessionKey: input.workbookSessionKey
      });

      if (input.cursor) {
        params.set("cursor", input.cursor);
      }

      if (typeof input.limit === "number") {
        params.set("limit", String(input.limit));
      }

      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/execution/history?${params.toString()}`)
      );
    },

    async undoExecution(input) {
      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/execution/undo`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(input)
        })
      );
    },

    async prepareUndoExecution(input) {
      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/execution/undo/prepare`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(input)
        })
      );
    },

    async redoExecution(input) {
      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/execution/redo`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(input)
        })
      );
    },

    async prepareRedoExecution(input) {
      return parseJson(
        await fetch(`${normalizedBaseUrl}/api/execution/redo/prepare`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify(input)
        })
      );
    },

    async approveWrite(input) {
      return parseContractPayload(
        await parseJson(
          await fetch(`${normalizedBaseUrl}/api/writeback/approve`, {
            method: "POST",
            headers: { "content-type": "application/json" },
            body: JSON.stringify(input)
          })
        ),
        WritebackApprovalResponseSchema,
        "The Hermes service returned a writeback approval response the client could not use.\n\nRetry the approval, then reload the client if it keeps happening."
      );
    },

    async completeWrite(input) {
      return parseContractPayload(
        await parseJson(
          await fetch(`${normalizedBaseUrl}/api/writeback/complete`, {
            method: "POST",
            headers: { "content-type": "application/json" },
            body: JSON.stringify(input)
          })
        ),
        WritebackCompletionResponseSchema,
        "The Hermes service returned a writeback completion response the client could not use.\n\nRetry the writeback completion, then reload the client if it keeps happening."
      );
    }
  };
}

export {
  buildDataCleanupPreview,
  buildDataCleanupUpdatePreview,
  buildCompositePlanPreview,
  buildCompositeUpdatePreview,
  buildRangeTransferPreview,
  buildRangeTransferUpdatePreview,
  buildDataValidationPreview,
  buildDryRunPreview,
  buildExtractedTablePreview,
  buildPlanHistoryPreview,
  buildHermesRequest,
  buildRangeFormatPreview,
  buildRangeFilterPreview,
  buildRangeSortPreview,
  buildWorkbookStructurePreview,
  buildNamedRangeUpdatePreview,
  buildSheetImportPreview,
  buildSheetStructurePreview,
  buildSheetUpdatePreview,
  buildStructuredPreview,
  buildWriteMatrix,
  createRequestId,
  filterSupportedImageFiles,
  formatDryRunSummary,
  formatHistoryEntrySummary,
  formatProofLine,
  formatTraceEvent,
  formatTraceTimeline,
  formatWritebackStatusLine,
  getFollowUpSuggestions,
  getResponseBodyText,
  getResponseConfidence,
  getResponseMetaLine,
  getResponseWarnings,
  getRequiresConfirmation,
  getStructuredPreview,
  isSupportedImageMimeType,
  isWritePlanResponse,
  sanitizeConversation,
  summarizeLatestTrace
};

export type {
  AnalysisReportUpdateWritebackResult,
  ChartUpdateWritebackResult,
  ConditionalFormatUpdateWritebackResult,
  DataCleanupUpdateWritebackResult,
  DataValidationWritebackResult,
  GatewayClient,
  HostBridge,
  HostRuntimeConfig,
  HostSnapshot,
  CompositeWritePlan,
  CompositeUpdateWritebackResult,
  DryRunResult,
  NamedRangeUpdateWritebackResult,
  PlanHistoryEntry,
  PlanHistoryPage,
  PivotTableUpdateWritebackResult,
  RangeFilterWritebackResult,
  RangeSortWritebackResult,
  RangeTransferUpdateWritebackResult,
  RangeWritebackResult,
  RedoRequest,
  RequestEnvelopeInput,
  RunPollResult,
  SheetStructureWritebackResult,
  StartRunAccepted,
  TracePollResult,
  UndoRequest,
  WorkbookStructureWritebackResult,
  WritePlan,
  WritebackApprovalRequest,
  WritebackCompletionRequest,
  WritebackDestructiveConfirmation,
  WritebackResult
} from "./types.js";

export type { HermesRequest };
export type { PreviewTable, StructuredPreview } from "./render.js";
