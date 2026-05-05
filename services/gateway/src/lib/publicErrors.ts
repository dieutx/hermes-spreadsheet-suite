const CLIENT_UNSAFE_VALIDATION_PATTERN =
  /(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|approval_secret|HERMES_[A-Z0-9_]+|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)|\/(?:root|srv|home|tmp|var|opt|workspace|app|mnt|Users)\/[^\s]+|[A-Za-z]:\\[^\s]+|(?:^|[\s.(["'=:])\\\\[^\s]+|https?:\/\/(?:internal(?:[.\w-]*)?|localhost|127(?:\.\d{1,3}){3}|0\.0\.0\.0|10\.\d{1,3}\.\d{1,3}\.\d{1,3}|169\.254\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[0-1])\.\d{1,3}\.\d{1,3}|\[(?:::ffff:|(?:0:){5}ffff:)(?:(?:127|10)(?:\.\d{1,3}){3}|169\.254\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[01])\.\d{1,3}\.\d{1,3}|(?:0{1,4}|7f[0-9a-f]{2}|0?a[0-9a-f]{2}|a9fe|c0a8|ac1[0-9a-f]):[0-9a-f]{1,4})\]|\[(?:::|::1|f[cd][0-9a-f:]*|fe[89ab][0-9a-f:]*)\]|[^/\s]*\.local)[^\s]*/i;
const CLIENT_UNSAFE_ENV_NAME_PATTERN =
  /(?:[A-Z][A-Z0-9_]*(?:SECRET|TOKEN|PASSWORD|PRIVATE|CREDENTIAL|API_KEY|SERVER_KEY|BASE_URL)[A-Z0-9_]*|\b(?:SECRET|TOKEN|PASSWORD|PRIVATE|CREDENTIAL|API_KEY|SERVER_KEY|BASE_URL)\b)/;
const CLIENT_UNSAFE_NUMERIC_IPV4_URL_PATTERN =
  /https?:\/\/(?:0x[0-9a-f]+|0[0-7]+|\d+)(?:\.(?:0x[0-9a-f]+|0[0-7]+|\d+)){0,3}(?::\d+)?(?:[/?#]|\s|$)/i;

function normalizePublicErrorText(value: unknown): string {
  return String(value ?? "")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function clampPublicErrorText(value: string, limit: number): string {
  return value.length > limit ? value.slice(0, limit) : value;
}

export function sanitizeClientIssueMessage(
  message: unknown,
  fallback = "Invalid request field."
): string {
  const normalized = normalizePublicErrorText(message);
  if (
    !normalized ||
    CLIENT_UNSAFE_VALIDATION_PATTERN.test(normalized) ||
    CLIENT_UNSAFE_ENV_NAME_PATTERN.test(normalized) ||
    CLIENT_UNSAFE_NUMERIC_IPV4_URL_PATTERN.test(normalized)
  ) {
    return fallback;
  }

  return clampPublicErrorText(normalized, 1000);
}

export function formatClientIssuePath(path: ReadonlyArray<string | number>): string {
  const normalized = normalizePublicErrorText(path.map((segment) => String(segment)).join("."));
  if (!normalized) {
    return "";
  }

  if (
    CLIENT_UNSAFE_VALIDATION_PATTERN.test(normalized) ||
    CLIENT_UNSAFE_ENV_NAME_PATTERN.test(normalized) ||
    CLIENT_UNSAFE_NUMERIC_IPV4_URL_PATTERN.test(normalized)
  ) {
    return "(redacted)";
  }

  return clampPublicErrorText(normalized, 512);
}
