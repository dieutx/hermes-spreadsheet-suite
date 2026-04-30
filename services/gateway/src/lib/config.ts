import { config as loadEnv } from "dotenv";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

loadEnv();
loadEnv({
  path: resolve(dirname(fileURLToPath(import.meta.url)), "../../../../.env")
});

export type GatewayConfig = {
  port: number;
  environmentLabel: string;
  serviceLabel: string;
  gatewayPublicBaseUrl: string;
  allowedCorsOrigins?: string[];
  maxUploadBytes: number;
  approvalSecret: string;
  saveInvalidHermesDebugArtifacts: boolean;
  hermesAgentBaseUrl?: string;
  hermesAgentApiKey?: string;
  hermesAgentModel?: string;
  hermesAgentTimeoutMs?: number;
  skillRegistryPath: string;
};

const LOCAL_GATEWAY_DEFAULT_ALLOWED_ORIGINS = [
  "https://docs.google.com",
  "https://excel.officeapps.live.com",
  "https://localhost:3000",
  "https://127.0.0.1:3000"
] as const;
const DEFAULT_HERMES_AGENT_TIMEOUT_MS = 45_000;
const SAFE_PUBLIC_LABEL_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._ -]{0,79}$/;
const UNSAFE_PUBLIC_LABEL_PATTERN = /(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|approval_secret|APPROVAL_SECRET|HERMES_[A-Z0-9_]+|OPENAI_API_KEY|ANTHROPIC_API_KEY|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)|\/(?:root|srv|home|tmp|var|opt|workspace|app|mnt)\/[^\s]+|https?:\/\/[^\s]+/i;

function isLoopbackBaseUrl(value: string): boolean {
  try {
    const url = new URL(value);
    return url.hostname === "127.0.0.1" || url.hostname === "localhost" || url.hostname === "::1";
  } catch {
    return false;
  }
}

function parseBooleanEnv(value: string | undefined): boolean {
  return value === "true" || value === "1";
}

function parseRequiredPositiveIntegerEnv(name: string, value: string | undefined, fallback: number): number {
  const raw = String(value ?? fallback).trim();
  if (!/^[1-9]\d*$/.test(raw)) {
    throw new Error(`${name} must be a positive integer.`);
  }

  const parsed = Number.parseInt(raw, 10);
  if (!Number.isSafeInteger(parsed)) {
    throw new Error(`${name} must be a positive integer.`);
  }

  return parsed;
}

function sanitizePublicLabel(value: string | undefined, fallback: string): string {
  const label = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();

  if (!label) {
    return fallback;
  }

  if (UNSAFE_PUBLIC_LABEL_PATTERN.test(label) || !SAFE_PUBLIC_LABEL_PATTERN.test(label)) {
    return fallback;
  }

  return label;
}

function parseHttpUrlEnv(name: string, value: string): string {
  const trimmed = value.trim();

  try {
    const url = new URL(trimmed);
    if (url.protocol !== "http:" && url.protocol !== "https:") {
      throw new Error("unsupported protocol");
    }
    return trimmed;
  } catch {
    throw new Error(`${name} must be a valid http(s) URL.`);
  }
}

function tryNormalizeOrigin(value: string): string | undefined {
  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  if (trimmed === "*") {
    return "*";
  }

  try {
    return new URL(trimmed).origin;
  } catch {
    return undefined;
  }
}

function getAllowedCorsOrigins(
  gatewayPublicBaseUrl: string,
  configuredOrigins: string | undefined
): string[] {
  const configuredValues = String(configuredOrigins || "")
    .split(",")
    .map((value) => value.trim())
    .filter((value) => value.length > 0);

  if (configuredValues.length > 0) {
    const configured = configuredValues.map((value) => {
      const normalized = tryNormalizeOrigin(value);
      if (!normalized) {
        throw new Error(`GATEWAY_ALLOWED_ORIGINS contains an invalid origin: ${value}.`);
      }
      return normalized;
    });
    return [...new Set(configured)];
  }

  if (isLoopbackBaseUrl(gatewayPublicBaseUrl)) {
    return [...LOCAL_GATEWAY_DEFAULT_ALLOWED_ORIGINS];
  }

  const publicOrigin = tryNormalizeOrigin(gatewayPublicBaseUrl);
  return publicOrigin ? [publicOrigin] : [];
}

export function getConfig(): GatewayConfig {
  const gatewayPublicBaseUrl = parseHttpUrlEnv(
    "GATEWAY_PUBLIC_BASE_URL",
    process.env.GATEWAY_PUBLIC_BASE_URL ?? "http://127.0.0.1:8787"
  );
  const approvalSecret = process.env.APPROVAL_SECRET?.trim() ?? "";
  const allowedCorsOrigins = getAllowedCorsOrigins(
    gatewayPublicBaseUrl,
    process.env.GATEWAY_ALLOWED_ORIGINS
  );

  if (!approvalSecret) {
    throw new Error(
      "APPROVAL_SECRET must be configured before the gateway can approve writeback plans."
    );
  }

  if (allowedCorsOrigins.includes("*") && !isLoopbackBaseUrl(gatewayPublicBaseUrl)) {
    throw new Error(
      "GATEWAY_ALLOWED_ORIGINS must not contain * when the gateway public base URL is not local."
    );
  }

  return {
    port: parseRequiredPositiveIntegerEnv("PORT", process.env.PORT, 8787),
    environmentLabel: sanitizePublicLabel(process.env.HERMES_ENVIRONMENT_LABEL, "local-dev"),
    serviceLabel: sanitizePublicLabel(process.env.HERMES_SERVICE_LABEL, "hermes-gateway-local"),
    gatewayPublicBaseUrl,
    allowedCorsOrigins,
    maxUploadBytes: parseRequiredPositiveIntegerEnv(
      "MAX_UPLOAD_BYTES",
      process.env.MAX_UPLOAD_BYTES,
      8_000_000
    ),
    approvalSecret,
    saveInvalidHermesDebugArtifacts: parseBooleanEnv(process.env.HERMES_DEBUG_INVALID_RESPONSES),
    hermesAgentBaseUrl:
      process.env.HERMES_AGENT_BASE_URL ??
      process.env.HERMES_BASE_URL ??
      "http://127.0.0.1:8642/v1",
    hermesAgentApiKey: process.env.HERMES_API_SERVER_KEY ?? process.env.HERMES_AGENT_API_KEY,
    hermesAgentModel: process.env.HERMES_AGENT_MODEL ?? process.env.HERMES_AGENT_ID,
    hermesAgentTimeoutMs: parseRequiredPositiveIntegerEnv(
      "HERMES_AGENT_TIMEOUT_MS",
      process.env.HERMES_AGENT_TIMEOUT_MS,
      DEFAULT_HERMES_AGENT_TIMEOUT_MS
    ),
    skillRegistryPath: process.env.SKILL_REGISTRY_PATH ??
      "../../extensions/registry/skill-registry.json"
  };
}
