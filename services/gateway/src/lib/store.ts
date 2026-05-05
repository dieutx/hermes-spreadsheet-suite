import { randomBytes, randomUUID } from "node:crypto";
import type { ImageAttachment } from "@hermes/contracts";

type StoredAttachment = {
  metadata: ImageAttachment;
  buffer: Buffer;
  createdAt: string;
  createdAtMs: number;
  sessionId?: string;
  workbookId?: string;
};

type AttachmentStoreOptions = {
  ttlMs?: number;
  maxEntries?: number;
  now?: () => number;
};

const UPLOAD_TOKEN_BYTES = 32;
const MAX_ATTACHMENT_FILE_NAME_LENGTH = 128;
const UNSAFE_ATTACHMENT_FILE_NAME_PATTERN = /(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|APPROVAL_SECRET|HERMES_API_SERVER_KEY|HERMES_AGENT_API_KEY|HERMES_AGENT_BASE_URL|OPENAI_API_KEY|ANTHROPIC_API_KEY|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)|\/(?:root|srv|home|tmp|var|opt|workspace|app|mnt|Users)\/[^\s]+|[A-Za-z]:\\[^\s]+|(?:^|[\s=:])\\\\[^\s]+|https?:\/\/[^\s]+/i;
const ATTACHMENT_FILE_EXTENSION_BY_MIME: Record<ImageAttachment["mimeType"], string> = {
  "image/png": ".png",
  "image/jpeg": ".jpg",
  "image/jpg": ".jpg",
  "image/webp": ".webp"
};

function createUploadToken(): string {
  return `upl_${randomBytes(UPLOAD_TOKEN_BYTES).toString("base64url")}`;
}

function getDefaultAttachmentFileName(mimeType: ImageAttachment["mimeType"]): string {
  return `uploaded-image${ATTACHMENT_FILE_EXTENSION_BY_MIME[mimeType] ?? ""}`;
}

function sanitizeAttachmentFileName(
  fileName: string,
  mimeType: ImageAttachment["mimeType"]
): string {
  const fallback = getDefaultAttachmentFileName(mimeType);
  const original = String(fileName || "")
    .replace(/[\u0000-\u001f\u007f]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!original || UNSAFE_ATTACHMENT_FILE_NAME_PATTERN.test(original)) {
    return fallback;
  }

  const pathParts = original.split(/[\\/]+/).filter(Boolean);
  const baseName = pathParts.length > 0 ? pathParts[pathParts.length - 1] : original;
  const sanitized = baseName
    .replace(/[^A-Za-z0-9._ ()+-]/g, "_")
    .replace(/_+/g, "_")
    .trim()
    .slice(0, MAX_ATTACHMENT_FILE_NAME_LENGTH);

  if (!sanitized || sanitized === "." || sanitized === ".." || UNSAFE_ATTACHMENT_FILE_NAME_PATTERN.test(sanitized)) {
    return fallback;
  }

  return sanitized;
}

export class AttachmentStore {
  private readonly attachments = new Map<string, StoredAttachment>();
  private readonly ttlMs: number;
  private readonly maxEntries: number;
  private readonly now: () => number;

  constructor(options: AttachmentStoreOptions = {}) {
    this.ttlMs = Math.max(1, options.ttlMs ?? 15 * 60 * 1000);
    this.maxEntries = Math.max(1, options.maxEntries ?? 200);
    this.now = options.now ?? (() => Date.now());
  }

  private isExpired(entry: StoredAttachment): boolean {
    return (this.now() - entry.createdAtMs) > this.ttlMs;
  }

  private pruneExpired(): void {
    for (const [attachmentId, entry] of this.attachments.entries()) {
      if (this.isExpired(entry)) {
        this.attachments.delete(attachmentId);
      }
    }
  }

  private evictOverflow(): void {
    if (this.attachments.size <= this.maxEntries) {
      return;
    }

    const overflow = this.attachments.size - this.maxEntries;
    const oldestIds = [...this.attachments.entries()]
      .sort((left, right) => left[1].createdAtMs - right[1].createdAtMs)
      .slice(0, overflow)
      .map(([attachmentId]) => attachmentId);

    for (const attachmentId of oldestIds) {
      this.attachments.delete(attachmentId);
    }
  }

  save(input: {
    buffer: Buffer;
    mimeType: ImageAttachment["mimeType"];
    fileName: string;
    size: number;
    source: ImageAttachment["source"];
    previewUrl: string;
    sessionId?: string;
    workbookId?: string;
  }): ImageAttachment {
    const id = `att_${randomUUID()}`;
    const uploadToken = createUploadToken();
    const storageRef = `blob://${id}`;
    const createdAtMs = this.now();
    const metadata: ImageAttachment = {
      id,
      type: "image",
      mimeType: input.mimeType,
      fileName: sanitizeAttachmentFileName(input.fileName, input.mimeType),
      size: input.size,
      source: input.source,
      previewUrl: input.previewUrl,
      uploadToken,
      storageRef
    };

    this.pruneExpired();
    this.attachments.set(id, {
      metadata,
      buffer: input.buffer,
      createdAt: new Date(createdAtMs).toISOString(),
      createdAtMs,
      sessionId: typeof input.sessionId === "string" && input.sessionId.trim().length > 0
        ? input.sessionId.trim()
        : undefined,
      workbookId: typeof input.workbookId === "string" && input.workbookId.trim().length > 0
        ? input.workbookId.trim()
        : undefined
    });
    this.evictOverflow();

    return metadata;
  }

  get(attachmentId: string): StoredAttachment | undefined {
    const entry = this.attachments.get(attachmentId);
    if (!entry) {
      return undefined;
    }

    if (this.isExpired(entry)) {
      this.attachments.delete(attachmentId);
      return undefined;
    }

    return entry;
  }
}
