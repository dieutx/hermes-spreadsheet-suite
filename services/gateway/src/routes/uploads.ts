import { Router } from "express";
import type { NextFunction, Request, Response } from "express";
import multer from "multer";
import { MvpImageMimeTypes } from "@hermes/contracts";
import type { GatewayConfig } from "../lib/config.js";
import { AttachmentStore } from "../lib/store.js";

type UploadRouteErrorPayload = {
  error: {
    code: string;
    message: string;
    userAction?: string;
  };
};

const UPLOAD_CONTENT_ATTACHMENT_ID_MAX_LENGTH = 128;
const UPLOAD_CONTENT_TOKEN_MAX_LENGTH = 256;

function matchesImageMimeType(
  buffer: Buffer,
  mimeType: (typeof MvpImageMimeTypes)[number]
): boolean {
  if (mimeType === "image/png") {
    return buffer.length >= 8 &&
      buffer[0] === 0x89 &&
      buffer[1] === 0x50 &&
      buffer[2] === 0x4e &&
      buffer[3] === 0x47 &&
      buffer[4] === 0x0d &&
      buffer[5] === 0x0a &&
      buffer[6] === 0x1a &&
      buffer[7] === 0x0a;
  }

  if (mimeType === "image/jpeg" || mimeType === "image/jpg") {
    return buffer.length >= 2 &&
      buffer[0] === 0xff &&
      buffer[1] === 0xd8;
  }

  if (mimeType === "image/webp") {
    return buffer.length >= 12 &&
      buffer.subarray(0, 4).toString("ascii") === "RIFF" &&
      buffer.subarray(8, 12).toString("ascii") === "WEBP";
  }

  return false;
}

function formatUploadRouterError(error: unknown, maxUploadBytes: number): {
  status: number;
  body: UploadRouteErrorPayload;
} {
  if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
    return {
      status: 413,
      body: {
        error: {
          code: "INVALID_REQUEST",
          message: `Images must be ${maxUploadBytes} bytes or smaller.`,
          userAction: "Compress the image or upload a smaller file, then try again."
        }
      }
    };
  }

  const message = error instanceof Error ? error.message : String(error || "");
  if (/unexpected end of form|multipart/i.test(message)) {
    return {
      status: 400,
      body: {
        error: {
          code: "INVALID_REQUEST",
          message: "The upload was incomplete or malformed.",
          userAction: "Retry the upload. If it keeps failing, choose the file again before sending it."
        }
      }
    };
  }

  return {
    status: 400,
    body: {
      error: {
        code: "INVALID_REQUEST",
        message: "The image upload could not be processed.",
        userAction: "Retry the upload with a supported PNG, JPG/JPEG, or WEBP image."
      }
    }
  };
}

function getQueryString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function getBoundUploadContentString(value: unknown, maxLength: number): { ok: true; value?: string } | { ok: false } {
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

  if (normalized.length > maxLength) {
    return { ok: false };
  }

  return { ok: true, value: normalized };
}

function invalidUploadContentCredentials() {
  return {
    error: {
      code: "INVALID_REQUEST",
      message: "Upload content credentials are invalid.",
      userAction: "Upload the file again if you still need it."
    }
  };
}

export function createUploadRouter(input: {
  attachmentStore: AttachmentStore;
  config: GatewayConfig;
}): Router {
  const router = Router();
  const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: input.config.maxUploadBytes }
  });

  router.post("/image", upload.single("file"), (req, res) => {
    if (!req.file) {
      res.status(400).json({
        error: {
          code: "INVALID_REQUEST",
          message: "Expected a file upload in the 'file' field.",
          userAction: "Choose an image file, then try the upload again."
        }
      });
      return;
    }

    if (!MvpImageMimeTypes.includes(req.file.mimetype as (typeof MvpImageMimeTypes)[number])) {
      res.status(400).json({
        error: {
          code: "UNSUPPORTED_ATTACHMENT_TYPE",
          message: "Only PNG, JPG/JPEG, and WEBP are supported in the MVP.",
          userAction: "Retry the upload with a supported PNG, JPG/JPEG, or WEBP image."
        }
      });
      return;
    }

    if (!matchesImageMimeType(
      req.file.buffer,
      req.file.mimetype as (typeof MvpImageMimeTypes)[number]
    )) {
      res.status(400).json({
        error: {
          code: "INVALID_REQUEST",
          message: "Uploaded file bytes do not match the declared image mime type.",
          userAction: "Choose the image file again, then retry the upload."
        }
      });
      return;
    }

    const sessionId = typeof req.body.sessionId === "string" && req.body.sessionId.trim().length > 0
      ? req.body.sessionId.trim()
      : undefined;
    if (!sessionId) {
      res.status(400).json({
        error: {
          code: "INVALID_REQUEST",
          message: "Image uploads require a Hermes session id.",
          userAction: "Reload the spreadsheet sidebar or add-in, then try the upload again."
        }
      });
      return;
    }

    const workbookId = typeof req.body.workbookId === "string" && req.body.workbookId.trim().length > 0
      ? req.body.workbookId.trim()
      : undefined;
    if (!workbookId) {
      res.status(400).json({
        error: {
          code: "INVALID_REQUEST",
          message: "Image uploads require a workbook id.",
          userAction: "Reload the spreadsheet, then try the upload again from the workbook where you want to use it."
        }
      });
      return;
    }

    const source = typeof req.body.source === "string" ? req.body.source : "upload";
    const attachment = input.attachmentStore.save({
      buffer: req.file.buffer,
      mimeType: req.file.mimetype as (typeof MvpImageMimeTypes)[number],
      fileName: req.file.originalname,
      size: req.file.size,
      source: source === "clipboard" || source === "drag_drop" ? source : "upload",
      previewUrl: `${input.config.gatewayPublicBaseUrl}/api/uploads/pending-preview`,
      sessionId,
      workbookId
    });

    const uploadToken = attachment.uploadToken;
    if (!uploadToken) {
      throw new Error("Upload token missing for saved attachment.");
    }

    const previewUrl =
      `${input.config.gatewayPublicBaseUrl}/api/uploads/${attachment.id}/content` +
      `?uploadToken=${encodeURIComponent(uploadToken)}` +
      `&sessionId=${encodeURIComponent(sessionId)}` +
      `&workbookId=${encodeURIComponent(workbookId)}`;
    attachment.previewUrl = previewUrl;
    res.status(201).json({
      attachment
    });
  });

  router.get("/:attachmentId/content", (req, res) => {
    const attachmentId = getBoundUploadContentString(
      req.params.attachmentId,
      UPLOAD_CONTENT_ATTACHMENT_ID_MAX_LENGTH
    );
    const uploadToken = getBoundUploadContentString(
      req.query.uploadToken,
      UPLOAD_CONTENT_TOKEN_MAX_LENGTH
    );
    if (!attachmentId.ok || !attachmentId.value || !uploadToken.ok) {
      res.status(400).json(invalidUploadContentCredentials());
      return;
    }

    const attachment = input.attachmentStore.get(attachmentId.value);
    const sessionId = getQueryString(req.query.sessionId);
    const workbookId = getQueryString(req.query.workbookId);
    if (
      !attachment ||
      uploadToken.value !== attachment.metadata.uploadToken ||
      (attachment.sessionId && sessionId !== attachment.sessionId) ||
      (attachment.workbookId && workbookId !== attachment.workbookId)
    ) {
      res.status(404).json({
        error: {
          code: "ATTACHMENT_UNAVAILABLE",
          message: "Attachment not found.",
          userAction: "Upload the file again if you still need it."
        }
      });
      return;
    }

    res.setHeader("content-type", attachment.metadata.mimeType);
    res.setHeader("cache-control", "private, max-age=60");
    res.setHeader("x-content-type-options", "nosniff");
    res.send(attachment.buffer);
  });

  router.use((error: unknown, _req: Request, res: Response, next: NextFunction) => {
    if (res.headersSent) {
      next(error);
      return;
    }

    const formatted = formatUploadRouterError(error, input.config.maxUploadBytes);
    res.status(formatted.status).json(formatted.body);
  });

  return router;
}
