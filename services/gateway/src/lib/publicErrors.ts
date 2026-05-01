const CLIENT_UNSAFE_VALIDATION_PATTERN =
  /(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|approval_secret|APPROVAL_SECRET|HERMES_[A-Z0-9_]+|OPENAI_API_KEY|ANTHROPIC_API_KEY|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)|\/(?:root|srv|home|tmp|var|opt|workspace|app|mnt)\/[^\s]+|[A-Za-z]:\\[^\s]+|(?:^|[\s.])\\\\[^\s]+|https?:\/\/(?:internal(?:[.\w-]*)?|localhost|127(?:\.\d{1,3}){3}|0\.0\.0\.0|10\.\d{1,3}\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[0-1])\.\d{1,3}\.\d{1,3}|[^/\s]*\.local)[^\s]*/i;

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
  if (!normalized || CLIENT_UNSAFE_VALIDATION_PATTERN.test(normalized)) {
    return fallback;
  }

  return clampPublicErrorText(normalized, 1000);
}

export function formatClientIssuePath(path: ReadonlyArray<string | number>): string {
  const normalized = normalizePublicErrorText(path.map((segment) => String(segment)).join("."));
  if (!normalized) {
    return "";
  }

  if (CLIENT_UNSAFE_VALIDATION_PATTERN.test(normalized)) {
    return "(redacted)";
  }

  return clampPublicErrorText(normalized, 512);
}
