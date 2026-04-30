import cors from "cors";
import express from "express";
import { getConfig } from "./lib/config.js";
import { ExecutionLedger } from "./lib/executionLedger.js";
import { HermesAgentClient } from "./lib/hermesClient.js";
import { AttachmentStore } from "./lib/store.js";
import { TraceBus } from "./lib/traceBus.js";
import { createExecutionControlRouter } from "./routes/executionControl.js";
import { createRequestRouter } from "./routes/requests.js";
import { createTraceRouter } from "./routes/trace.js";
import { createUploadRouter } from "./routes/uploads.js";
import { createWritebackRouter } from "./routes/writeback.js";

function tryNormalizeOrigin(value: string | undefined): string | undefined {
  const text = String(value || "").trim();
  if (!text) {
    return undefined;
  }

  try {
    return new URL(text).origin;
  } catch {
    return undefined;
  }
}

export function isCorsOriginAllowed(
  origin: string | undefined,
  allowedOrigins: string[] | undefined
): boolean {
  if (!origin) {
    return true;
  }

  const normalizedOrigin = tryNormalizeOrigin(origin);
  if (!normalizedOrigin) {
    return false;
  }

  const normalizedAllowedOrigins = (allowedOrigins || [])
    .map((value) => tryNormalizeOrigin(value) ?? (value === "*" ? "*" : undefined))
    .filter((value): value is string => Boolean(value));

  if (normalizedAllowedOrigins.includes("*")) {
    return true;
  }

  return normalizedAllowedOrigins.includes(normalizedOrigin);
}

export function allowPrivateNetworkPreflight(allowedOrigins: string[] | undefined) {
  return function privateNetworkPreflightMiddleware(
    req: express.Request,
    res: express.Response,
    next: express.NextFunction
  ) {
    if (
      req.headers["access-control-request-private-network"] === "true" &&
      isCorsOriginAllowed(req.headers.origin, allowedOrigins)
    ) {
      res.header("Access-Control-Allow-Private-Network", "true");
    }
    next();
  };
}

export function enforceAllowedOrigin(allowedOrigins: string[] | undefined) {
  return function allowedOriginMiddleware(
    req: express.Request,
    res: express.Response,
    next: express.NextFunction
  ) {
    if (isCorsOriginAllowed(req.headers.origin, allowedOrigins)) {
      next();
      return;
    }

    res.status(403).json({
      error: {
        code: "ORIGIN_NOT_ALLOWED",
        message: "This Hermes gateway origin is not allowed.",
        userAction: "Open Hermes from an approved Excel or Google Sheets host, then retry."
      }
    });
  };
}

export function handleGatewayAppError(
  error: unknown,
  _req: express.Request,
  res: express.Response,
  next: express.NextFunction
) {
  if (res.headersSent) {
    next(error);
    return;
  }

  const typedError = error as { type?: string } | undefined;
  if (typedError?.type === "entity.too.large") {
    res.status(413).json({
      error: {
        code: "INVALID_REQUEST",
        message: "That request is too large for the gateway to process safely.",
        userAction: "Retry with a smaller prompt, fewer attachments, or a smaller spreadsheet selection."
      }
    });
    return;
  }

  if (typedError?.type === "entity.parse.failed") {
    res.status(400).json({
      error: {
        code: "INVALID_REQUEST",
        message: "Hermes couldn't read that request body.",
        userAction: "Retry the request. If it keeps failing, refresh the host and try again."
      }
    });
    return;
  }

  console.error("[gateway] unhandled app error", error);
  res.status(500).json({
    error: {
      code: "INTERNAL_ERROR",
      message: "The gateway hit an unexpected error while processing the request.",
      userAction: "Retry the request. If it keeps failing, restart the gateway or check the server logs."
    }
  });
}

export function createApp() {
  const config = getConfig();
  const traceBus = new TraceBus();
  const executionLedger = new ExecutionLedger();
  const attachmentStore = new AttachmentStore();
  const hermesClient = new HermesAgentClient(config);
  const app = express();

  app.use(allowPrivateNetworkPreflight(config.allowedCorsOrigins));
  app.use(cors({
    origin(origin, callback) {
      callback(null, isCorsOriginAllowed(origin, config.allowedCorsOrigins));
    }
  }));
  app.use(express.json({ limit: "2mb" }));

  app.get("/health", (_req, res) => {
    res.json({
      ok: true,
      service: config.serviceLabel,
      environment: config.environmentLabel
    });
  });

  app.use(enforceAllowedOrigin(config.allowedCorsOrigins));
  app.use("/api/uploads", createUploadRouter({ attachmentStore, config }));
  app.use("/api/requests", createRequestRouter({
    traceBus,
    hermesClient,
    attachmentStore,
    config
  }));
  app.use("/api/trace", createTraceRouter({ traceBus }));
  app.use("/api/execution", createExecutionControlRouter({ executionLedger, config }));
  app.use("/api/writeback", createWritebackRouter({ traceBus, executionLedger, config }));
  app.use(handleGatewayAppError);

  return { app, config };
}
