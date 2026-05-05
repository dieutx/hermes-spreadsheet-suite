import { createHash, createHmac, randomBytes, timingSafeEqual } from "node:crypto";

type ApprovalTokenInput = {
  requestId: string;
  runId: string;
  planDigest: string;
  issuedAt: string;
  secret: string;
};

type VerifyApprovalTokenInput = {
  token: string;
  requestId: string;
  runId: string;
  planDigest: string;
  secret: string;
  maxAgeMs?: number;
  nowMs?: number;
};

function signPayload(payload: string, secret: string): string {
  return createHmac("sha256", secret).update(payload).digest("hex");
}

const TOKEN_DELIMITER = ".";
const APPROVAL_TOKEN_NONCE_BYTES = 16;
const APPROVAL_TOKEN_NONCE_BASE64URL_LENGTH = 22;

export function canonicalizePlan(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalizePlan);
  }

  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, nestedValue]) => [key, canonicalizePlan(nestedValue)])
    );
  }

  return value;
}

export function digestPlan(plan: unknown): string {
  const serialized = JSON.stringify(plan);
  return createHash("sha256").update(serialized).digest("hex");
}

export function digestCanonicalPlan(plan: unknown): string {
  return digestPlan(canonicalizePlan(plan));
}

export function createApprovalToken(input: ApprovalTokenInput): string {
  const payload = JSON.stringify({
    requestId: input.requestId,
    runId: input.runId,
    planDigest: input.planDigest,
    issuedAt: input.issuedAt,
    nonce: randomBytes(APPROVAL_TOKEN_NONCE_BYTES).toString("base64url")
  });
  const encodedPayload = Buffer.from(payload).toString("base64url");
  const signature = signPayload(encodedPayload, input.secret);
  return `${encodedPayload}${TOKEN_DELIMITER}${signature}`;
}

export function verifyApprovalToken(input: VerifyApprovalTokenInput): {
  valid: boolean;
  issuedAt?: string;
  expired?: boolean;
} {
  const delimiterIndex = input.token.indexOf(TOKEN_DELIMITER);
  if (delimiterIndex <= 0 || delimiterIndex === input.token.length - 1) {
    return { valid: false };
  }

  const encodedPayload = input.token.slice(0, delimiterIndex);
  const signature = input.token.slice(delimiterIndex + 1);
  const expectedSignature = signPayload(encodedPayload, input.secret);

  if (signature.length !== expectedSignature.length) {
    return { valid: false };
  }

  const matchesSignature = timingSafeEqual(
    Buffer.from(signature),
    Buffer.from(expectedSignature)
  );

  if (!matchesSignature) {
    return { valid: false };
  }

  let parsedPayload:
    | {
        requestId: string;
        runId: string;
        planDigest: string;
        issuedAt: string;
        nonce?: string;
      }
    | undefined;
  try {
    parsedPayload = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8"));
  } catch {
    return { valid: false };
  }

  if (!parsedPayload) {
    return { valid: false };
  }

  const { requestId, runId, planDigest, issuedAt, nonce } = parsedPayload;

  if (typeof nonce !== "string" || nonce.length < APPROVAL_TOKEN_NONCE_BASE64URL_LENGTH) {
    return { valid: false };
  }

  if (typeof input.maxAgeMs === "number") {
    const issuedAtMs = Date.parse(issuedAt);
    if (!Number.isFinite(issuedAtMs)) {
      return { valid: false };
    }

    const nowMs = typeof input.nowMs === "number" ? input.nowMs : Date.now();
    if (issuedAtMs > nowMs) {
      return { valid: false, expired: false };
    }

    if ((nowMs - issuedAtMs) > input.maxAgeMs) {
      return { valid: false, expired: true };
    }
  }

  const valid =
    requestId === input.requestId &&
    runId === input.runId &&
    planDigest === input.planDigest;

  return { valid, issuedAt: valid ? issuedAt : undefined, expired: false };
}

export function signReviewerProof(input: {
  requestId: string;
  runId: string;
  serviceLabel: string;
  environment: string;
  startedAt: string;
  completedAt: string;
  responseHash: string;
  secret: string;
}): string {
  const payload = [
    input.requestId,
    input.runId,
    input.serviceLabel,
    input.environment,
    input.startedAt,
    input.completedAt,
    input.responseHash
  ].join(".");

  return signPayload(payload, input.secret);
}
