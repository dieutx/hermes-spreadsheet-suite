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

function createUploadToken(): string {
  return `upl_${randomBytes(UPLOAD_TOKEN_BYTES).toString("base64url")}`;
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
      fileName: input.fileName,
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
