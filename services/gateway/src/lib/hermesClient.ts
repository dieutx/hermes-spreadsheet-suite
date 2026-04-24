import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import type { ZodIssue } from "zod";
import type {
  CompositePlanData,
  ExtractionMode,
  HermesRequest,
  HermesResponse,
  HermesTraceEvent,
  SheetUpdateData,
  Warning
} from "@hermes/contracts";
import {
  HermesResponseSchema,
} from "@hermes/contracts";
import { SPREADSHEET_RUNTIME_RULES } from "../hermes/runtimeRules.js";
import { buildHermesSpreadsheetRequestPrompt } from "../hermes/requestTemplate.js";
import {
  extractSingleJsonObjectText,
  normalizeHermesStructuredBodyInput,
  HermesStructuredBodySchema,
  type HermesStructuredBody
} from "../hermes/structuredBody.js";
import {
  getSpreadsheetRoutingHints,
  type SpreadsheetRoutingHints
} from "../hermes/requestTemplate.js";
import type { GatewayConfig } from "./config.js";
import type { TraceBus } from "./traceBus.js";

type ProcessRequestInput = {
  runId: string;
  request: HermesRequest;
  traceBus: TraceBus;
};

type JsonRecord = Record<string, unknown>;
type ContractSafeErrorCode =
  | "INVALID_REQUEST"
  | "UNSUPPORTED_ATTACHMENT_TYPE"
  | "ATTACHMENT_UNAVAILABLE"
  | "UNSUPPORTED_OPERATION"
  | "SPREADSHEET_CONTEXT_MISSING"
  | "EXTRACTION_UNAVAILABLE"
  | "CONFIRMATION_REQUIRED"
  | "PROVIDER_ERROR"
  | "TIMEOUT"
  | "INTERNAL_ERROR";
const INVALID_HERMES_DEBUG_PREFIX = path.join(tmpdir(), "hermes-spreadsheet-invalid");
const INTERNAL_ERROR_LANGUAGE_PATTERN = /\b(contract|schema|structured body|validation|json|payload|parse|parser|normaliz(?:e|ation))\b/i;
const MAX_GATEWAY_RESPONSE_TRACE_EVENTS = 200;
const DEFAULT_HERMES_AGENT_TIMEOUT_MS = 45_000;

class HermesProviderTimeoutError extends Error {
  constructor(timeoutMs: number) {
    super(`Hermes provider request timed out after ${timeoutMs}ms.`);
    this.name = "HermesProviderTimeoutError";
  }
}

function isObject(value: unknown): value is JsonRecord {
  return typeof value === "object" && value !== null;
}

function hasExplicitDemoWarning(warnings: Warning[] | undefined): boolean {
  return (warnings ?? []).some((warning) =>
    /demo/i.test(warning.code) || /demo/i.test(warning.message)
  );
}

function getAssistantContent(payload: unknown): string | undefined {
  if (!isObject(payload) || !Array.isArray(payload.choices) || payload.choices.length === 0) {
    return undefined;
  }

  const firstChoice = payload.choices[0];
  if (!isObject(firstChoice) || !isObject(firstChoice.message)) {
    return undefined;
  }

  const content = firstChoice.message.content;
  if (typeof content === "string") {
    return content;
  }

  if (!Array.isArray(content)) {
    return undefined;
  }

  const textParts = content.flatMap((part) => {
    if (!isObject(part) || typeof part.text !== "string") {
      return [];
    }

    return [part.text];
  });

  return textParts.length > 0 ? textParts.join("") : undefined;
}

function sanitizeForFileName(value: string): string {
  const sanitized = value.replace(/[^a-zA-Z0-9_-]+/g, "_");
  return sanitized.slice(0, 64) || "unknown";
}

function buildDebugFileBase(requestId: string): string {
  return [
    INVALID_HERMES_DEBUG_PREFIX,
    sanitizeForFileName(requestId),
    Date.now()
  ].join("-");
}

function hasCurrentRegionContext(request: HermesRequest): boolean {
  return typeof request.context.currentRegion?.range === "string" &&
    request.context.currentRegion.range.trim().length > 0;
}

type HelperSheetScaffoldLayout = {
  inputCells: string[];
  resultCell: string;
  guidanceRange: string;
};

function extractExplicitSingleCellReferences(rawMessage: string): string[] {
  const matches = [...rawMessage.matchAll(/(?:^|[^A-Za-z0-9$])(\$?[A-Za-z]{1,3}\$?\d+)(?!\s*:)(?=$|[^A-Za-z0-9$])/g)];
  return Array.from(new Set(matches.map((match) => match[1].replace(/\$/g, "").toUpperCase())));
}

function parseSingleCellReference(value: string | undefined): { row: number; column: number } | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const match = value.replace(/\$/g, "").match(/^([A-Za-z]{1,3})(\d+)$/);
  if (!match) {
    return undefined;
  }

  let column = 0;
  for (const char of match[1].toUpperCase()) {
    column = (column * 26) + (char.charCodeAt(0) - 64);
  }

  return {
    column,
    row: Number.parseInt(match[2], 10)
  };
}

function columnNumberToLetters(value: number): string {
  let column = Math.max(1, Math.floor(value));
  let letters = "";
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    column = Math.floor((column - 1) / 26);
  }
  return letters;
}

function normalizeSingleCellRange(value: string | undefined): string | undefined {
  const parsed = parseSingleCellReference(value);
  if (!parsed) {
    return undefined;
  }

  return `${columnNumberToLetters(parsed.column)}${parsed.row}`;
}

function inferHelperSheetScaffoldLayout(
  request: HermesRequest,
  formulaTargetRange: string | undefined
): HelperSheetScaffoldLayout {
  const normalizedResultCell = normalizeSingleCellRange(formulaTargetRange) ?? "C1";
  const explicitRefs = extractExplicitSingleCellReferences(request.userMessage)
    .filter((cell) => cell !== normalizedResultCell);
  const inputCells = explicitRefs.slice(0, 2);

  while (inputCells.length < 2) {
    const fallbackCell = inputCells.length === 0 ? "A1" : "B1";
    if (!inputCells.includes(fallbackCell) && fallbackCell !== normalizedResultCell) {
      inputCells.push(fallbackCell);
      continue;
    }

    const fallbackColumn = inputCells.length === 0 ? 1 : 2;
    inputCells.push(`${columnNumberToLetters(fallbackColumn)}1`);
  }

  const occupiedRows = [normalizedResultCell, ...inputCells]
    .map((cell) => parseSingleCellReference(cell)?.row ?? 1);
  const maxOccupiedRow = Math.max(...occupiedRows, 1);
  const guidanceStartRow = maxOccupiedRow + 2;
  const guidanceEndRow = guidanceStartRow + 4;

  return {
    inputCells,
    resultCell: normalizedResultCell,
    guidanceRange: `A${guidanceStartRow}:B${guidanceEndRow}`
  };
}

function isCreateSheetPlan(plan: unknown): plan is { operation: "create_sheet"; sheetName: string } {
  return isObject(plan) &&
    plan.operation === "create_sheet" &&
    typeof plan.sheetName === "string" &&
    plan.sheetName.trim().length > 0;
}

function isSheetUpdatePlan(plan: unknown): plan is SheetUpdateData {
  return isObject(plan) &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    typeof plan.operation === "string";
}

function compositeHasVisibleHelperScaffold(plan: CompositePlanData, helperSheet: string): boolean {
  return plan.steps.some((step) =>
    isSheetUpdatePlan(step.plan) &&
    step.plan.targetSheet === helperSheet &&
    step.plan.operation === "replace_range" &&
    Array.isArray(step.plan.values) &&
    step.plan.values.length > 0
  );
}

function buildHelperSheetGuidanceStep(
  request: HermesRequest,
  helperSheet: string,
  formulaTargetRange: string | undefined,
  confidence: number
): CompositePlanData["steps"][number] {
  const layout = inferHelperSheetScaffoldLayout(request, formulaTargetRange);
  const values = [
    ["Hermes helper sheet", ""],
    ["How to use", "Edit the input cells, then read the result cell."],
    [layout.inputCells[0], "Input value"],
    [layout.inputCells[1], "Parameter or field"],
    [layout.resultCell, "Result cell"]
  ];

  return {
    stepId: "seed_helper_sheet_guidance",
    dependsOn: [],
    continueOnError: false,
    plan: {
      targetSheet: helperSheet,
      targetRange: layout.guidanceRange,
      operation: "replace_range",
      values,
      explanation: "Seed helper-sheet guidance for the input and output cells.",
      confidence,
      requiresConfirmation: true,
      overwriteRisk: "low",
      shape: {
        rows: values.length,
        columns: values[0]?.length ?? 0
      }
    }
  };
}

function augmentToolScaffoldingCompositePlan(
  request: HermesRequest,
  plan: CompositePlanData
): CompositePlanData {
  const createSheetStepIndex = plan.steps.findIndex((step) => isCreateSheetPlan(step.plan));
  if (createSheetStepIndex < 0) {
    return plan;
  }

  const createSheetStep = plan.steps[createSheetStepIndex];
  const createSheetPlan = createSheetStep?.plan;
  if (!isCreateSheetPlan(createSheetPlan)) {
    return plan;
  }

  const helperSheet = createSheetPlan.sheetName;
  if (compositeHasVisibleHelperScaffold(plan, helperSheet)) {
    return plan;
  }

  const formulaStepIndex = plan.steps.findIndex((step, index) =>
    index > createSheetStepIndex &&
    isSheetUpdatePlan(step.plan) &&
    step.plan.targetSheet === helperSheet &&
    (step.plan.operation === "set_formulas" || step.plan.operation === "mixed_update") &&
    Array.isArray(step.plan.formulas) &&
    step.plan.formulas.length > 0
  );

  if (formulaStepIndex < 0) {
    return plan;
  }

  const formulaStep = plan.steps[formulaStepIndex];
  const scaffoldStep = buildHelperSheetGuidanceStep(
    request,
    helperSheet,
    isSheetUpdatePlan(formulaStep.plan) ? formulaStep.plan.targetRange : undefined,
    plan.confidence
  );
  scaffoldStep.dependsOn = [createSheetStep.stepId];

  const stepIdSet = new Set(plan.steps.map((step) => step.stepId));
  let scaffoldStepId = scaffoldStep.stepId;
  let suffix = 1;
  while (stepIdSet.has(scaffoldStepId)) {
    suffix += 1;
    scaffoldStepId = `seed_helper_sheet_guidance_${suffix}`;
  }
  scaffoldStep.stepId = scaffoldStepId;

  const nextSteps = [...plan.steps];
  nextSteps.splice(createSheetStepIndex + 1, 0, scaffoldStep);

  for (let index = createSheetStepIndex + 2; index < nextSteps.length; index += 1) {
    const step = nextSteps[index];
    if (
      isSheetUpdatePlan(step.plan) &&
      step.plan.targetSheet === helperSheet &&
      !step.dependsOn.includes(scaffoldStepId)
    ) {
      step.dependsOn = [...step.dependsOn, scaffoldStepId];
    }
  }

  return {
    ...plan,
    steps: nextSteps
  };
}

export class HermesAgentClient {
  constructor(private readonly config: GatewayConfig) {}

  private nowIso(traceBus: TraceBus): string {
    return traceBus.nowIso();
  }

  async processRequest(input: ProcessRequestInput): Promise<void> {
    input.traceBus.markStatus(input.runId, "processing");

    try {
      const response = await this.fetchHermesAgentResponse(input);
      input.traceBus.setResponse(input.runId, response);
    } catch (error) {
      console.error("[gateway] Hermes request failed unexpectedly", {
        requestId: input.request.requestId,
        runId: input.runId,
        error
      });
      const fallback = this.makeUnexpectedAgentFailureResponse(input);
      input.traceBus.setResponse(input.runId, fallback);
    }
  }

  private async fetchHermesAgentResponse(input: ProcessRequestInput): Promise<HermesResponse> {
    const startedAt = this.getRunStartedAt(input);
    const endpoint = this.resolveEndpoint();
    const timeoutMs = this.config.hermesAgentTimeoutMs ?? DEFAULT_HERMES_AGENT_TIMEOUT_MS;
    const controller = typeof AbortController === "function"
      ? new AbortController()
      : undefined;
    let timeoutHandle: ReturnType<typeof globalThis.setTimeout> | undefined;
    let httpResponse: Response;

    try {
      httpResponse = await Promise.race([
        fetch(endpoint, {
          method: "POST",
          headers: this.buildHeaders(),
          body: JSON.stringify(this.buildChatCompletionsBody(input.request)),
          ...(controller ? { signal: controller.signal } : {})
        }),
        new Promise<Response>((_resolve, reject) => {
          timeoutHandle = globalThis.setTimeout(() => {
            controller?.abort();
            reject(new HermesProviderTimeoutError(timeoutMs));
          }, timeoutMs);
        })
      ]);
    } catch (error) {
      if (error instanceof HermesProviderTimeoutError) {
        return this.makeGatewayErrorResponse({
          request: input.request,
          runId: input.runId,
          traceBus: input.traceBus,
          startedAt,
          code: "TIMEOUT",
          message: `The Hermes service did not respond before the ${Math.ceil(timeoutMs / 1000)} second timeout.`,
          retryable: true,
          userAction: "Retry the request. If it keeps timing out, reduce the scope or check that the Hermes service is healthy.",
          trace: [
            ...input.traceBus.list(input.runId, 0),
            {
              event: "failed",
              timestamp: this.nowIso(input.traceBus)
            }
          ]
        });
      }

      throw error;
    } finally {
      if (timeoutHandle) {
        globalThis.clearTimeout(timeoutHandle);
      }
    }

    let payload: unknown;
    const contentType = httpResponse.headers.get("content-type") ?? "";

    if (contentType.includes("application/json")) {
      payload = await httpResponse.json();
    } else {
      const text = await httpResponse.text();
      payload = { message: text };
    }

    if (!httpResponse.ok) {
      return this.makeGatewayErrorResponse({
        request: input.request,
        runId: input.runId,
        traceBus: input.traceBus,
        startedAt,
        code: "PROVIDER_ERROR",
        message: this.errorMessageFromPayload(payload, httpResponse.status),
        retryable: httpResponse.status >= 500,
        userAction: httpResponse.status >= 500
          ? "Retry the request after the remote Hermes Agent service recovers."
          : undefined,
        trace: [
          ...input.traceBus.list(input.runId, 0),
          {
            event: "failed",
            timestamp: this.nowIso(input.traceBus)
          }
        ]
      });
    }

    const assistantContent = getAssistantContent(payload);
    if (!assistantContent) {
      return this.makeMalformedAgentPayloadResponse(input, startedAt);
    }

    const normalizedAssistantJson = extractSingleJsonObjectText(assistantContent);
    if (!normalizedAssistantJson) {
      await this.writeInvalidAssistantContent({
        requestId: input.request.requestId,
        runId: input.runId,
        reason: "assistant_content_not_single_json_object",
        rawContent: assistantContent
      });

      return this.makeMalformedAgentPayloadResponse(input, startedAt);
    }

    let responsePayload: unknown;
    try {
      responsePayload = JSON.parse(normalizedAssistantJson);
    } catch {
      await this.writeInvalidAssistantContent({
        requestId: input.request.requestId,
        runId: input.runId,
        reason: "assistant_content_not_valid_json",
        rawContent: assistantContent
      });

      return this.makeMalformedAgentPayloadResponse(input, startedAt);
    }

    const normalizedResponsePayload = normalizeHermesStructuredBodyInput(responsePayload);
    const parsed = HermesStructuredBodySchema.safeParse(normalizedResponsePayload);
    if (!parsed.success) {
      const debugArtifacts = await this.writeStructuredBodyValidationDebug({
        requestId: input.request.requestId,
        runId: input.runId,
        rawContent: assistantContent,
        issues: parsed.error.issues
      });
      console.warn("[gateway] Hermes structured body normalization failed", {
        requestId: input.request.requestId,
        runId: input.runId,
        rawAssistantFilePath: debugArtifacts?.rawFilePath,
        validationIssuePaths: parsed.error.issues.map((issue) => issue.path)
      });

      return this.makeMalformedAgentPayloadResponse(input, startedAt);
    }
    console.info(`[gateway] normalized Hermes structured body type=${parsed.data.type}`);

    const routingHints = getSpreadsheetRoutingHints(input.request);
    const reviewedBody = (
      parsed.data.type === "composite_plan" &&
      (routingHints.toolScaffoldingOpportunity || routingHints.inputLayoutConflictRisk)
    )
      ? {
        ...parsed.data,
        data: augmentToolScaffoldingCompositePlan(input.request, parsed.data.data)
      }
      : parsed.data;

    const response = this.buildGatewayResponseEnvelope({
      request: input.request,
      runId: input.runId,
      traceBus: input.traceBus,
      startedAt,
      trace: input.traceBus.list(input.runId, 0),
      body: reviewedBody
    });

    const reviewerError = this.ensureContractSafety(input.request, response, input.traceBus);
    if (reviewerError) {
      return reviewerError;
    }

    return response;
  }

  private ensureContractSafety(
    request: HermesRequest,
    response: HermesResponse,
    traceBus: TraceBus
  ): HermesResponse | undefined {
    const data = response.data as JsonRecord;
    const routingHints = getSpreadsheetRoutingHints(request);
    const hasImageAttachment = Array.isArray(request.context.attachments) &&
      request.context.attachments.some((attachment) => attachment.type === "image");
    const dataExtractionMode = typeof data.extractionMode === "string"
      ? data.extractionMode as ExtractionMode
      : undefined;
    const responseErrorCode = response.type === "error" && typeof data.code === "string"
      ? data.code
      : undefined;
    const warnings = response.warnings ?? [];
    const dataWarnings = Array.isArray(data.warnings)
      ? data.warnings as Warning[]
      : [];

    if (request.reviewer.forceExtractionMode === "unavailable" &&
      hasImageAttachment &&
      responseErrorCode !== "EXTRACTION_UNAVAILABLE") {
      return this.makeGatewayErrorResponse({
        request,
        runId: response.hermesRunId,
        traceBus,
        startedAt: response.startedAt,
        code: "EXTRACTION_UNAVAILABLE",
        message: "Real image extraction is unavailable in the current reviewer-safe runtime.",
        retryable: false,
        userAction: "Switch to a runtime with real extraction or disable reviewer-safe forced unavailable mode.",
        baseEnvelope: response,
        trace: this.withFailedTrace(response.trace, traceBus)
      });
    }

    if (request.reviewer.forceExtractionMode === "unavailable" &&
      (response.type === "extracted_table" || response.type === "sheet_import_plan")) {
      return this.makeGatewayErrorResponse({
        request,
        runId: response.hermesRunId,
        traceBus,
        startedAt: response.startedAt,
        code: "EXTRACTION_UNAVAILABLE",
        message: "Real image extraction is unavailable in the current reviewer-safe runtime.",
        retryable: false,
        userAction: "Switch to a runtime with real extraction or disable reviewer-safe forced unavailable mode.",
        baseEnvelope: response,
        trace: this.withFailedTrace(response.trace, traceBus)
      });
    }

    if (request.reviewer.reviewerSafeMode &&
      dataExtractionMode === "unavailable" &&
      (response.type === "extracted_table" || response.type === "sheet_import_plan")) {
      return this.makeGatewayErrorResponse({
        request,
        runId: response.hermesRunId,
        traceBus,
        startedAt: response.startedAt,
        code: "EXTRACTION_UNAVAILABLE",
        message: "Real image extraction is unavailable in the current reviewer-safe runtime.",
        retryable: false,
        userAction: "Switch to a runtime with real extraction or disable reviewer-safe forced unavailable mode.",
        baseEnvelope: response,
        trace: this.withFailedTrace(response.trace, traceBus)
      });
    }

    if (dataExtractionMode === "demo") {
      const topLevelDemo = hasExplicitDemoWarning(warnings);
      const dataDemo = hasExplicitDemoWarning(dataWarnings);

      if (!topLevelDemo && !dataDemo) {
        return this.makeGatewayErrorResponse({
          request,
          runId: response.hermesRunId,
          traceBus,
          startedAt: response.startedAt,
          code: "INTERNAL_ERROR",
          message: "Demo extraction responses must be explicitly labeled as demo in warnings or data warnings.",
          retryable: true,
          baseEnvelope: response,
          trace: this.withFailedTrace(response.trace, traceBus)
        });
      }
    }

    if (response.type === "error") {
      const sanitized = this.sanitizeGatewayErrorFields({
        request,
        code: response.data.code,
        message: response.data.message,
        userAction: response.data.userAction
      });

      if (
        sanitized.message !== response.data.message ||
        sanitized.userAction !== response.data.userAction
      ) {
        return this.makeGatewayErrorResponse({
          request,
          runId: response.hermesRunId,
          traceBus,
          startedAt: response.startedAt,
          code: response.data.code,
          message: sanitized.message,
          retryable: response.data.retryable,
          userAction: sanitized.userAction,
          baseEnvelope: response,
          trace: response.trace
        });
      }
    }

    const droppedExplicitWriteAction = routingHints.explicitWriteIntent && (
      response.type === "chat" ||
      response.type === "formula" ||
      (
        response.type === "analysis_report_plan" &&
        response.data.outputMode === "chat_only"
      )
    );

    if (droppedExplicitWriteAction) {
      return this.makeGatewayErrorResponse({
        request,
        runId: response.hermesRunId,
        traceBus,
        startedAt: response.startedAt,
        code: "INTERNAL_ERROR",
        message: "I couldn't prepare an actionable spreadsheet plan for that request.",
        retryable: true,
        userAction: this.getWriteRecoveryUserAction(request, routingHints),
        baseEnvelope: response,
        trace: this.withFailedTrace(response.trace, traceBus)
      });
    }

    return undefined;
  }

  private withFailedTrace(trace: HermesTraceEvent[], traceBus?: TraceBus): HermesTraceEvent[] {
    const lastEvent = trace.at(-1);
    if (lastEvent?.event === "failed") {
      return this.capGatewayTrace(trace);
    }

    return this.capGatewayTrace([
      ...trace,
      {
        event: "failed",
        timestamp: traceBus ? this.nowIso(traceBus) : new Date().toISOString()
      }
    ]);
  }

  private makeGatewayErrorResponse(input: {
    request: HermesRequest;
    runId: string;
    traceBus: TraceBus;
    startedAt?: string;
        code:
      | "INVALID_REQUEST"
      | "UNSUPPORTED_ATTACHMENT_TYPE"
      | "ATTACHMENT_UNAVAILABLE"
      | "UNSUPPORTED_OPERATION"
      | "SPREADSHEET_CONTEXT_MISSING"
      | "EXTRACTION_UNAVAILABLE"
      | "CONFIRMATION_REQUIRED"
      | "PROVIDER_ERROR"
      | "TIMEOUT"
      | "INTERNAL_ERROR";
    message: string;
    retryable: boolean;
    userAction?: string;
    baseEnvelope?: Partial<HermesResponse>;
    trace: HermesTraceEvent[];
  }): HermesResponse {
    const startedAt = input.baseEnvelope?.startedAt ?? input.startedAt ?? this.nowIso(input.traceBus);
    const completedAt = this.nowIso(input.traceBus);
    const hermesRunId = input.baseEnvelope?.hermesRunId ?? input.runId;
    const sanitized = this.sanitizeGatewayErrorFields({
      request: input.request,
      code: input.code,
      message: input.message,
      userAction: input.userAction
    });

    return HermesResponseSchema.parse({
      schemaVersion: "1.0.0",
      type: "error",
      requestId: input.request.requestId,
      hermesRunId,
      processedBy: "hermes",
      serviceLabel: input.baseEnvelope?.serviceLabel ?? this.config.serviceLabel,
      environmentLabel: input.baseEnvelope?.environmentLabel ?? this.config.environmentLabel,
      startedAt,
      completedAt,
      durationMs: Math.max(0, Date.parse(completedAt) - Date.parse(startedAt)),
      skillsUsed: input.baseEnvelope?.skillsUsed ?? [],
      downstreamProvider: input.baseEnvelope?.downstreamProvider ?? null,
      warnings: input.baseEnvelope?.warnings ?? [],
      trace: this.capGatewayTrace(input.trace),
      ui: {
        displayMode: "error",
        showTrace: true,
        showWarnings: true,
        showConfidence: false,
        showRequiresConfirmation: false
      },
      data: {
        code: input.code,
        message: sanitized.message,
        retryable: input.retryable,
        userAction: sanitized.userAction
      }
    });
  }

  private sanitizeGatewayErrorFields(input: {
    request?: HermesRequest;
    code: ContractSafeErrorCode;
    message: string;
    userAction?: string;
  }): { message: string; userAction?: string } {
    const resolvedMessage = String(input.message ?? "").trim();
    const resolvedUserAction = typeof input.userAction === "string"
      ? input.userAction.trim()
      : "";
    const messageLooksInternal = !resolvedMessage || INTERNAL_ERROR_LANGUAGE_PATTERN.test(resolvedMessage);
    const userActionLooksInternal = resolvedUserAction.length > 0 &&
      INTERNAL_ERROR_LANGUAGE_PATTERN.test(resolvedUserAction);

    if (!messageLooksInternal && !userActionLooksInternal) {
      return {
        message: resolvedMessage,
        userAction: resolvedUserAction || undefined
      };
    }

    const fallback = this.getFallbackErrorCopy(input.code, input.request);
    return {
      message: messageLooksInternal ? fallback.message : resolvedMessage,
      userAction: userActionLooksInternal || !resolvedUserAction
        ? fallback.userAction
        : resolvedUserAction
    };
  }

  private getFallbackErrorCopy(
    code: ContractSafeErrorCode,
    request?: HermesRequest
  ): {
    message: string;
    userAction?: string;
  } {
    const routingHints = request ? getSpreadsheetRoutingHints(request) : undefined;
    const currentRegionAvailable = request ? hasCurrentRegionContext(request) : false;

    switch (code) {
      case "UNSUPPORTED_OPERATION":
        return {
          message: "I can't do that exact spreadsheet action here.",
          userAction: routingHints?.explicitWriteIntent
            ? this.getWriteRecoveryUserAction(request, routingHints)
            : currentRegionAvailable
              ? "Tell me how you want me to use the current table or range, or ask for the closest supported alternative."
              : "Tell me the target sheet, range, cell, or output location you want me to use, or ask for the closest supported alternative."
        };
      case "SPREADSHEET_CONTEXT_MISSING":
        return {
          message: "I need a bit more spreadsheet context before I can do that.",
          userAction: currentRegionAvailable
            ? "Tell me whether to use the current table or a different target, and include any output sheet or anchor if needed."
            : "Tell me the target sheet, range, cell, or attachment you want me to use."
        };
      case "ATTACHMENT_UNAVAILABLE":
        return {
          message: "I can't access that uploaded file anymore.",
          userAction: currentRegionAvailable
            ? "Reattach the file, then tell me whether to use the current table or a different target for the import."
            : "Reattach the file, then tell me the target sheet or range if you want me to paste or import its contents."
        };
      case "PROVIDER_ERROR":
      case "TIMEOUT":
        return {
          message: "The Hermes service couldn't complete that request right now.",
          userAction: "Retry the request in a moment. If it keeps failing, check that the Hermes service is online."
        };
      case "INVALID_REQUEST":
      case "CONFIRMATION_REQUIRED":
        return {
          message: "I need a little more information before I can do that.",
          userAction: routingHints?.explicitWriteIntent
            ? this.getWriteRecoveryUserAction(request, routingHints)
            : currentRegionAvailable
              ? "Tell me whether to use the current table or a different target, or restate the request more explicitly."
              : "Tell me the target sheet, range, cell, or output location, or restate the request more explicitly."
        };
      default:
        return {
          message: "I couldn't prepare a valid spreadsheet response for that request.",
          userAction: routingHints?.explicitWriteIntent
            ? this.getWriteRecoveryUserAction(request, routingHints)
            : currentRegionAvailable
              ? "Try again for the current table or range, or tell me a different target if you want to change something else."
              : "Try again with the target sheet, range, cell, or attachment, or split the request into smaller steps."
        };
    }
  }

  private getWriteRecoveryUserAction(
    request: HermesRequest | undefined,
    routingHints: SpreadsheetRoutingHints | undefined
  ): string {
    if (!request || !routingHints) {
      return "Tell me the target sheet, range, cell, or output location you want me to use, or retry the request more explicitly.";
    }

    const currentRegionAvailable = hasCurrentRegionContext(request);

    if (routingHints.inputLayoutConflictRisk) {
      return currentRegionAvailable
        ? "Your requested input or result cells overlap the current source table. I can create a separate helper sheet for the inputs and output, or you can tell me a different target sheet or cells."
        : "The cells you named appear to overlap the source table. I can create a separate helper sheet for the inputs and output, or you can tell me a different target sheet or cells.";
    }

    if (routingHints.mixedAdvisoryAndWriteRequest) {
      switch (routingHints.preferredResponseType) {
        case "sheet_update":
          return "Split the analysis and writeback into separate steps, or tell me the exact cell or range you want me to change now.";
        case "workbook_structure_update":
          return "Split the analysis and writeback into separate steps, or tell me the sheet action you want me to apply now.";
        default:
          return currentRegionAvailable
            ? "Split the analysis and writeback into separate steps, or tell me to apply the write on the current table now."
            : "Split the analysis and writeback into separate steps, or tell me the target sheet, range, or output location you want me to change now.";
      }
    }

    switch (routingHints.preferredResponseType) {
      case "workbook_structure_update":
        return "Retry the sheet action explicitly, or tell me the sheet/tab name you want me to create, rename, move, hide, or delete.";
      case "sheet_structure_update":
        return currentRegionAvailable
          ? "Retry the structure change for the current range, or tell me the exact rows, columns, or range you want me to change."
          : "Tell me the exact rows, columns, or range you want me to change, or retry the request more explicitly.";
      case "sheet_update":
        return routingHints.generatedDataRequest
          ? currentRegionAvailable
            ? "Retry the data update for the current table, or tell me a different target sheet or range."
            : "Tell me the target sheet or range where you want the generated data to go, or retry the request more explicitly."
          : "Tell me the exact cell or range you want me to change, or retry the write request more explicitly.";
      case "pivot_table_plan":
      case "chart_plan":
      case "external_data_plan":
      case "analysis_report_plan":
        return currentRegionAvailable
          ? "Retry the request for the current table, or tell me a different output sheet, range, or anchor if you want the result somewhere else."
          : "Tell me the source table or range and the output sheet or anchor you want me to use, or retry the request more explicitly.";
      case "range_sort_plan":
      case "range_filter_plan":
      case "data_cleanup_plan":
      case "conditional_format_plan":
      case "data_validation_plan":
      case "range_transfer_plan":
      case "composite_plan":
        return currentRegionAvailable
          ? "Retry the request for the current table or range, or tell me a different target if you want the write somewhere else."
          : "Tell me the target sheet or range you want me to change, or retry the request more explicitly.";
      case "named_range_update":
        return "Tell me the named range and target sheet or range you want me to update, or retry the request more explicitly.";
      default:
        return currentRegionAvailable
          ? "Retry the request for the current table or range, or tell me a different target if you want the write somewhere else."
          : "Tell me the target sheet, range, cell, or output location you want me to use, or retry the request more explicitly.";
    }
  }

  private makeMalformedAgentPayloadResponse(
    input: ProcessRequestInput,
    startedAt: string
  ): HermesResponse {
    const fallback = this.getFallbackErrorCopy("INTERNAL_ERROR", input.request);
    return this.makeGatewayErrorResponse({
      request: input.request,
      runId: input.runId,
      traceBus: input.traceBus,
      startedAt,
      code: "INTERNAL_ERROR",
      message: fallback.message,
      retryable: true,
      userAction: fallback.userAction,
      trace: [
        ...input.traceBus.list(input.runId, 0),
        {
          event: "failed",
          timestamp: this.nowIso(input.traceBus)
        }
      ]
    });
  }

  private makeUnexpectedAgentFailureResponse(input: ProcessRequestInput): HermesResponse {
    return this.makeGatewayErrorResponse({
      request: input.request,
      runId: input.runId,
      traceBus: input.traceBus,
      startedAt: this.getRunStartedAt(input),
      code: "INTERNAL_ERROR",
      message: "I couldn't complete that spreadsheet request right now.",
      retryable: true,
      userAction: "Retry the request in a moment. If it keeps failing, check that the Hermes service is online.",
      trace: [
        ...input.traceBus.list(input.runId, 0),
        {
          event: "failed",
          timestamp: this.nowIso(input.traceBus)
        }
      ]
    });
  }

  private getRunStartedAt(input: ProcessRequestInput): string {
    return input.traceBus.getRun(input.runId)?.startedAt ?? this.nowIso(input.traceBus);
  }

  private buildGatewayResponseEnvelope(input: {
    request: HermesRequest;
    runId: string;
    traceBus: TraceBus;
    startedAt: string;
    trace: HermesTraceEvent[];
    body: HermesStructuredBody;
  }): HermesResponse {
    const completedAt = this.nowIso(input.traceBus);
    const envelope = {
      schemaVersion: "1.0.0",
      type: input.body.type,
      requestId: input.request.requestId,
      hermesRunId: input.runId,
      processedBy: "hermes",
      serviceLabel: this.config.serviceLabel,
      environmentLabel: this.config.environmentLabel,
      startedAt: input.startedAt,
      completedAt,
      durationMs: Math.max(0, Date.parse(completedAt) - Date.parse(input.startedAt)),
      skillsUsed: input.body.skillsUsed ?? [],
      downstreamProvider: input.body.downstreamProvider ?? null,
      warnings: input.body.warnings ?? [],
      trace: this.buildResponseTrace(input.body.type, input.trace, completedAt),
      ui: this.buildUiContract(input.body),
      data: input.body.data
    };

    return HermesResponseSchema.parse(envelope);
  }

  private buildUiContract(body: HermesStructuredBody): HermesResponse["ui"] {
    switch (body.type) {
      case "chat":
        return {
          displayMode: "chat-first",
          showTrace: true,
          showWarnings: true,
          showConfidence: true,
          showRequiresConfirmation: false
        };
      case "formula":
        return {
          displayMode: "structured-preview",
          showTrace: true,
          showWarnings: true,
          showConfidence: true,
          showRequiresConfirmation: Boolean(body.data.requiresConfirmation)
        };
      case "composite_plan":
      case "analysis_report_plan":
        return {
          displayMode: "structured-preview",
          showTrace: true,
          showWarnings: true,
          showConfidence: true,
          showRequiresConfirmation: Boolean(body.data.requiresConfirmation)
        };
      case "workbook_structure_update":
      case "range_format_update":
      case "conditional_format_plan":
      case "sheet_structure_update":
      case "range_sort_plan":
      case "range_filter_plan":
      case "data_validation_plan":
      case "pivot_table_plan":
      case "chart_plan":
      case "named_range_update":
      case "range_transfer_plan":
      case "data_cleanup_plan":
      case "sheet_update":
      case "sheet_import_plan":
      case "external_data_plan":
        return {
          displayMode: "structured-preview",
          showTrace: true,
          showWarnings: true,
          showConfidence: true,
          showRequiresConfirmation: true
        };
      case "attachment_analysis":
      case "extracted_table":
        return {
          displayMode: "structured-preview",
          showTrace: true,
          showWarnings: true,
          showConfidence: true,
          showRequiresConfirmation: false
        };
      case "analysis_report_update":
      case "pivot_table_update":
      case "chart_update":
        return {
          displayMode: "structured-preview",
          showTrace: true,
          showWarnings: true,
          showConfidence: false,
          showRequiresConfirmation: false
        };
      case "document_summary":
        return {
          displayMode: "chat-first",
          showTrace: true,
          showWarnings: true,
          showConfidence: true,
          showRequiresConfirmation: false
        };
      case "error":
        return {
          displayMode: "error",
          showTrace: true,
          showWarnings: true,
          showConfidence: false,
          showRequiresConfirmation: false
        };
    }
  }

  private buildResponseTrace(
    type: HermesStructuredBody["type"],
    existingTrace: HermesTraceEvent[],
    completedAt: string
  ): HermesTraceEvent[] {
    const trace = [...existingTrace];

    const pushIfMissing = (event: HermesTraceEvent) => {
      if (!trace.some((existing) => JSON.stringify(existing) === JSON.stringify(event))) {
        trace.push(event);
      }
    };

    switch (type) {
      case "chat":
      case "formula":
      case "attachment_analysis":
      case "extracted_table":
      case "document_summary":
        pushIfMissing({ event: "result_generated", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "sheet_update":
        pushIfMissing({ event: "sheet_update_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "composite_plan":
        pushIfMissing({ event: "composite_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "sheet_import_plan":
        pushIfMissing({ event: "sheet_import_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "external_data_plan":
        pushIfMissing({ event: "external_data_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "workbook_structure_update":
        pushIfMissing({ event: "workbook_structure_update_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "range_format_update":
        pushIfMissing({ event: "range_format_update_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "conditional_format_plan":
        pushIfMissing({ event: "conditional_format_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "sheet_structure_update":
        pushIfMissing({ event: "sheet_structure_update_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "range_sort_plan":
        pushIfMissing({ event: "range_sort_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "range_filter_plan":
        pushIfMissing({ event: "range_filter_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "data_validation_plan":
        pushIfMissing({ event: "data_validation_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "analysis_report_plan":
        pushIfMissing({ event: "analysis_report_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "pivot_table_plan":
        pushIfMissing({ event: "pivot_table_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "chart_plan":
        pushIfMissing({ event: "chart_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "named_range_update":
        pushIfMissing({ event: "named_range_update_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "range_transfer_plan":
        pushIfMissing({ event: "range_transfer_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "data_cleanup_plan":
        pushIfMissing({ event: "data_cleanup_plan_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "analysis_report_update":
        pushIfMissing({ event: "analysis_report_update_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "pivot_table_update":
        pushIfMissing({ event: "pivot_table_update_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "chart_update":
        pushIfMissing({ event: "chart_update_ready", timestamp: completedAt });
        pushIfMissing({ event: "completed", timestamp: completedAt });
        break;
      case "error":
        pushIfMissing({ event: "failed", timestamp: completedAt });
        break;
    }

    return this.capGatewayTrace(trace);
  }

  private capGatewayTrace(trace: HermesTraceEvent[]): HermesTraceEvent[] {
    if (!Array.isArray(trace) || trace.length <= MAX_GATEWAY_RESPONSE_TRACE_EVENTS) {
      return trace;
    }

    return trace.slice(-MAX_GATEWAY_RESPONSE_TRACE_EVENTS);
  }

  private errorMessageFromPayload(payload: unknown, status: number): string {
    if (isObject(payload) && typeof payload.message === "string") {
      return payload.message;
    }

    if (isObject(payload) && typeof payload.error === "string") {
      return payload.error;
    }

    if (isObject(payload) && isObject(payload.error) && typeof payload.error.message === "string") {
      return payload.error.message;
    }

    return `Hermes Agent request failed with status ${status}.`;
  }

  private async writeInvalidAssistantContent(input: {
    requestId: string;
    runId: string;
    reason: string;
    rawContent: string;
  }): Promise<void> {
    if (!this.config.saveInvalidHermesDebugArtifacts) {
      return;
    }

    const filePath = `${buildDebugFileBase(input.requestId)}.txt`;

    const debugContents = [
      `reason: ${input.reason}`,
      `requestId: ${input.requestId}`,
      `runId: ${input.runId}`,
      "--- raw assistant content ---",
      input.rawContent
    ].join("\n");

    try {
      await fs.writeFile(filePath, debugContents, "utf8");
    } catch {
      // Debug logging must never change the user-visible error path.
    }
  }

  private async writeStructuredBodyValidationDebug(input: {
    requestId: string;
    runId: string;
    rawContent: string;
    issues: ZodIssue[];
  }): Promise<{ rawFilePath: string; validationFilePath: string } | undefined> {
    if (!this.config.saveInvalidHermesDebugArtifacts) {
      return undefined;
    }

    const fileBase = buildDebugFileBase(input.requestId);
    const rawFilePath = `${fileBase}.txt`;
    const validationFilePath = `${fileBase}.validation.txt`;

    const rawDebugContents = [
      "reason: assistant_json_failed_structured_body_validation",
      `requestId: ${input.requestId}`,
      `runId: ${input.runId}`,
      "--- raw assistant content ---",
      input.rawContent
    ].join("\n");

    const validationDebugContents = [
      `requestId: ${input.requestId}`,
      `runId: ${input.runId}`,
      `rawAssistantFilePath: ${rawFilePath}`,
      "validator: HermesStructuredBodySchema",
      "--- issues ---",
      JSON.stringify(
        input.issues.map((issue) => ({
          path: issue.path,
          message: issue.message,
          code: issue.code,
          expected: "expected" in issue ? issue.expected : undefined,
          received: "received" in issue ? issue.received : undefined,
          keys: "keys" in issue ? issue.keys : undefined
        })),
        null,
        2
      )
    ].join("\n");

    try {
      await Promise.all([
        fs.writeFile(rawFilePath, rawDebugContents, "utf8"),
        fs.writeFile(validationFilePath, validationDebugContents, "utf8")
      ]);
      return { rawFilePath, validationFilePath };
    } catch {
      // Debug logging must never change the user-visible error path.
      return undefined;
    }
  }

  private buildHeaders(): Record<string, string> {
    const headers: Record<string, string> = {
      "content-type": "application/json"
    };

    if (this.config.hermesAgentApiKey?.trim()) {
      headers.authorization = `Bearer ${this.config.hermesAgentApiKey.trim()}`;
    }

    return headers;
  }

  private buildChatCompletionsBody(request: HermesRequest): {
    model: string;
    messages: Array<{ role: "system" | "user"; content: string }>;
  } {
    return {
      model: this.config.hermesAgentModel?.trim() || "hermes-agent",
      messages: [
        {
          role: "system",
          content: SPREADSHEET_RUNTIME_RULES
        },
        {
          role: "user",
          content: buildHermesSpreadsheetRequestPrompt(request)
        }
      ]
    };
  }

  private resolveEndpoint(): string {
    const baseUrl = this.config.hermesAgentBaseUrl?.trim();
    if (!baseUrl) {
      throw new Error(
        "HERMES_AGENT_BASE_URL (or legacy HERMES_BASE_URL) must be configured for backend-to-Hermes-Agent forwarding."
      );
    }

    return `${baseUrl.replace(/\/+$/, "")}/chat/completions`;
  }
}

export { HermesAgentClient as HermesAdapterClient };
