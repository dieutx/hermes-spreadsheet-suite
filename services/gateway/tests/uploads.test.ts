import express from "express";
import { describe, expect, it, vi } from "vitest";
import multer from "multer";
import { AttachmentStore } from "../src/lib/store.ts";
import { createRequestRouter } from "../src/routes/requests.ts";
import { createUploadRouter } from "../src/routes/uploads.ts";
import { TraceBus } from "../src/lib/traceBus.ts";

const PNG_SIGNATURE_BYTES = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00
]);

function createTestApp() {
  const traceBus = new TraceBus();
  const attachmentStore = new AttachmentStore();
  const hermesClient = {
    processRequest: vi.fn(async () => undefined)
  };
  const config = {
    port: 8787,
    environmentLabel: "review",
    serviceLabel: "spreadsheet-gateway",
    gatewayPublicBaseUrl: "http://127.0.0.1:8787",
    maxUploadBytes: 8_000_000,
    approvalSecret: "secret",
    hermesAgentBaseUrl: "http://agent.test",
    hermesAgentApiKey: undefined,
    hermesAgentModel: undefined,
    skillRegistryPath: ""
  };

  return {
    uploadRouter: createUploadRouter({ attachmentStore, config }),
    requestRouter: createRequestRouter({
      traceBus,
      hermesClient,
      attachmentStore,
      config
    }),
    attachmentStore,
    hermesClient
  };
}

async function invokeUploadImage(
  router: express.Router,
  input: {
    body?: Record<string, unknown>;
    file?: Partial<Express.Multer.File>;
  }
) {
  const layer = router.stack.find((entry) =>
    entry.route?.path === "/image" && entry.route.methods.post
  );

  if (!layer) {
    throw new Error("POST /image route not found.");
  }

  let statusCode = 200;
  let jsonBody: unknown;

  const req = {
    body: input.body ?? {},
    file: input.file
  } as express.Request;
  const res = {
    status(code: number) {
      statusCode = code;
      return this;
    },
    json(payload: unknown) {
      jsonBody = payload;
      return this;
    }
  } as unknown as express.Response;

  await Promise.resolve(layer.route.stack[1]?.handle(req, res));

  return { statusCode, body: jsonBody };
}

async function invokeUploadRouterError(
  router: express.Router,
  error: unknown
) {
  const errorLayer = [...router.stack].reverse().find((entry) => entry.handle?.length === 4);

  if (!errorLayer) {
    throw new Error("Upload router error middleware not found.");
  }

  let statusCode = 200;
  let jsonBody: unknown;

  const req = {} as express.Request;
  const res = {
    headersSent: false,
    status(code: number) {
      statusCode = code;
      return this;
    },
    json(payload: unknown) {
      jsonBody = payload;
      return this;
    }
  } as unknown as express.Response;
  const next = vi.fn();

  await Promise.resolve(errorLayer.handle(error, req, res, next));

  return { statusCode, body: jsonBody, next };
}

async function invokeRequestPost(router: express.Router, body: unknown) {
  const layer = router.stack.find((entry) =>
    entry.route?.path === "/" && entry.route.methods.post
  );

  if (!layer) {
    throw new Error("POST / route not found.");
  }

  let statusCode = 200;
  let jsonBody: unknown;

  const req = { body } as express.Request;
  const res = {
    status(code: number) {
      statusCode = code;
      return this;
    },
    json(payload: unknown) {
      jsonBody = payload;
      return this;
    }
  } as unknown as express.Response;

  await Promise.resolve(layer.route.stack[0]?.handle(req, res));

  return { statusCode, body: jsonBody };
}

async function invokeUploadContent(
  router: express.Router,
  attachmentId: string,
  query: Record<string, unknown> = {}
) {
  const layer = router.stack.find((entry) =>
    entry.route?.path === "/:attachmentId/content" && entry.route.methods.get
  );

  if (!layer) {
    throw new Error("GET /:attachmentId/content route not found.");
  }

  let statusCode = 200;
  let jsonBody: unknown;
  let sentBody: unknown;
  const headers = new Map<string, string>();

  const req = {
    params: { attachmentId },
    query
  } as unknown as express.Request;
  const res = {
    status(code: number) {
      statusCode = code;
      return this;
    },
    json(payload: unknown) {
      jsonBody = payload;
      return this;
    },
    setHeader(name: string, value: string) {
      headers.set(name.toLowerCase(), value);
      return this;
    },
    send(payload: unknown) {
      sentBody = payload;
      return this;
    }
  } as unknown as express.Response;

  await Promise.resolve(layer.route.stack[0]?.handle(req, res));

  return { statusCode, body: jsonBody, sentBody, headers };
}

describe("upload router", () => {
  it("uploads one MVP image and the returned attachment reference is accepted by Flow 2", async () => {
    const { uploadRouter, requestRouter, hermesClient, attachmentStore } = createTestApp();

    const uploadResponse = await invokeUploadImage(uploadRouter, {
      body: { source: "upload", sessionId: "sess_img_real_001", workbookId: "sheet_import_demo" },
      file: {
        buffer: PNG_SIGNATURE_BYTES,
        mimetype: "image/png",
        originalname: "table.png",
        size: PNG_SIGNATURE_BYTES.length
      }
    });

    expect(uploadResponse.statusCode).toBe(201);
    expect((uploadResponse.body as any).attachment).toMatchObject({
      id: expect.stringMatching(/^att_/),
      type: "image",
      mimeType: "image/png",
      fileName: "table.png",
      size: PNG_SIGNATURE_BYTES.length,
      source: "upload",
      storageRef: expect.stringMatching(/^blob:\/\/att_/)
    });
    expect(
      attachmentStore.get((uploadResponse.body as any).attachment.id)?.metadata.previewUrl
    ).toBe((uploadResponse.body as any).attachment.previewUrl);
    expect((uploadResponse.body as any).attachment.previewUrl).toContain("uploadToken=");
    expect((uploadResponse.body as any).attachment.previewUrl).toContain("sessionId=sess_img_real_001");
    expect((uploadResponse.body as any).attachment.previewUrl).toContain("workbookId=sheet_import_demo");

    const flow2Response = await invokeRequestPost(requestRouter, {
        schemaVersion: "1.0.0",
        requestId: "req_img_real_001",
        source: {
          channel: "google_sheets",
          clientVersion: "0.1.0",
          sessionId: "sess_img_real_001"
        },
        host: {
          platform: "google_sheets",
          workbookTitle: "Import Demo",
          workbookId: "sheet_import_demo",
          activeSheet: "Sheet1",
          locale: "en-US",
          timeZone: "UTC"
        },
        userMessage: "Extract this table and prepare it for import into Sheet3 at B4:D8",
        conversation: [],
        context: {
          attachments: [(uploadResponse.body as any).attachment]
        },
        capabilities: {
          canRenderTrace: true,
          canRenderStructuredPreview: true,
          canConfirmWriteBack: true,
          supportsImageInputs: true,
          supportsWriteBackExecution: true
        },
        reviewer: {
          reviewerSafeMode: false,
          forceExtractionMode: "real"
        },
        confirmation: {
          state: "none"
        }
      });

    expect(flow2Response.statusCode).toBe(202);
    expect((flow2Response.body as any).requestId).toBe("req_img_real_001");
    expect((flow2Response.body as any).runId).toMatch(/^run_/);
    expect((flow2Response.body as any).error).toBeUndefined();
    expect(hermesClient.processRequest).toHaveBeenCalledTimes(1);
  });

  it("rejects files whose bytes do not match the declared image mime type", async () => {
    const { uploadRouter } = createTestApp();

    const uploadResponse = await invokeUploadImage(uploadRouter, {
      body: { source: "upload", sessionId: "sess_img_real_001", workbookId: "sheet_import_demo" },
      file: {
        buffer: Buffer.from("not-a-real-png"),
        mimetype: "image/png",
        originalname: "table.png",
        size: Buffer.byteLength("not-a-real-png")
      }
    });

    expect(uploadResponse.statusCode).toBe(400);
    expect(uploadResponse.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Uploaded file bytes do not match the declared image mime type.",
        userAction: "Choose the image file again, then retry the upload."
      }
    });
  });

  it("requires a workbook id for image uploads", async () => {
    const { uploadRouter } = createTestApp();

    const uploadResponse = await invokeUploadImage(uploadRouter, {
      body: { source: "upload", sessionId: "sess_img_real_001" },
      file: {
        buffer: PNG_SIGNATURE_BYTES,
        mimetype: "image/png",
        originalname: "table.png",
        size: PNG_SIGNATURE_BYTES.length
      }
    });

    expect(uploadResponse.statusCode).toBe(400);
    expect(uploadResponse.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Image uploads require a workbook id.",
        userAction: "Reload the spreadsheet, then try the upload again from the workbook where you want to use it."
      }
    });
  });

  it("returns a JSON upload error when multer rejects an oversized image", async () => {
    const { uploadRouter } = createTestApp();
    const error = new multer.MulterError("LIMIT_FILE_SIZE");

    const response = await invokeUploadRouterError(uploadRouter, error);

    expect(response.statusCode).toBe(413);
    expect(response.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Images must be 8000000 bytes or smaller.",
        userAction: "Compress the image or upload a smaller file, then try again."
      }
    });
    expect(response.next).not.toHaveBeenCalled();
  });

  it("returns a JSON upload error when multipart parsing fails mid-stream", async () => {
    const { uploadRouter } = createTestApp();

    const response = await invokeUploadRouterError(
      uploadRouter,
      new Error("Unexpected end of form")
    );

    expect(response.statusCode).toBe(400);
    expect(response.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "The upload was incomplete or malformed.",
        userAction: "Retry the upload. If it keeps failing, choose the file again before sending it."
      }
    });
    expect(response.next).not.toHaveBeenCalled();
  });

  it("requires matching upload token, session id, and workbook id to fetch uploaded preview content", async () => {
    const { uploadRouter } = createTestApp();

    const uploadResponse = await invokeUploadImage(uploadRouter, {
      body: { source: "upload", sessionId: "sess_img_real_001", workbookId: "sheet_import_demo" },
      file: {
        buffer: PNG_SIGNATURE_BYTES,
        mimetype: "image/png",
        originalname: "table.png",
        size: PNG_SIGNATURE_BYTES.length
      }
    });

    const attachment = (uploadResponse.body as any).attachment;

    const missingToken = await invokeUploadContent(uploadRouter, attachment.id);
    expect(missingToken.statusCode).toBe(404);
    expect(missingToken.body).toEqual({
      error: {
        code: "ATTACHMENT_UNAVAILABLE",
        message: "Attachment not found.",
        userAction: "Upload the file again if you still need it."
      }
    });

    const validToken = await invokeUploadContent(uploadRouter, attachment.id, {
      uploadToken: attachment.uploadToken
    });
    expect(validToken.statusCode).toBe(404);
    expect(validToken.body).toEqual({
      error: {
        code: "ATTACHMENT_UNAVAILABLE",
        message: "Attachment not found.",
        userAction: "Upload the file again if you still need it."
      }
    });

    const wrongSession = await invokeUploadContent(uploadRouter, attachment.id, {
      uploadToken: attachment.uploadToken,
      sessionId: "sess_other",
      workbookId: "sheet_import_demo"
    });
    expect(wrongSession.statusCode).toBe(404);

    const wrongWorkbook = await invokeUploadContent(uploadRouter, attachment.id, {
      uploadToken: attachment.uploadToken,
      sessionId: "sess_img_real_001",
      workbookId: "sheet_other"
    });
    expect(wrongWorkbook.statusCode).toBe(404);

    const validOwnership = await invokeUploadContent(uploadRouter, attachment.id, {
      uploadToken: attachment.uploadToken,
      sessionId: "sess_img_real_001",
      workbookId: "sheet_import_demo"
    });
    expect(validOwnership.statusCode).toBe(200);
    expect(validOwnership.body).toBeUndefined();
    expect(validOwnership.sentBody).toEqual(PNG_SIGNATURE_BYTES);
    expect(validOwnership.headers.get("content-type")).toBe("image/png");
  });

  it("sanitizes unsafe uploaded file names before returning attachment metadata", async () => {
    const { uploadRouter, attachmentStore } = createTestApp();

    const uploadResponse = await invokeUploadImage(uploadRouter, {
      body: { source: "upload", sessionId: "sess_img_real_001", workbookId: "sheet_import_demo" },
      file: {
        buffer: PNG_SIGNATURE_BYTES,
        mimetype: "image/png",
        originalname: "/home/alice/HERMES_API_SERVER_KEY=secret.png",
        size: PNG_SIGNATURE_BYTES.length
      }
    });

    const attachment = (uploadResponse.body as any).attachment;

    expect(uploadResponse.statusCode).toBe(201);
    expect(attachment.fileName).toBe("uploaded-image.png");
    expect(attachment.fileName).not.toContain("HERMES_API_SERVER_KEY");
    expect(attachment.fileName).not.toContain("/home/alice");
    expect(attachmentStore.get(attachment.id)?.metadata.fileName).toBe("uploaded-image.png");
  });

  it("rejects image uploads without a Hermes session id", async () => {
    const { uploadRouter } = createTestApp();

    const uploadResponse = await invokeUploadImage(uploadRouter, {
      body: { source: "upload" },
      file: {
        buffer: PNG_SIGNATURE_BYTES,
        mimetype: "image/png",
        originalname: "table.png",
        size: PNG_SIGNATURE_BYTES.length
      }
    });

    expect(uploadResponse.statusCode).toBe(400);
    expect(uploadResponse.body).toEqual({
      error: {
        code: "INVALID_REQUEST",
        message: "Image uploads require a Hermes session id.",
        userAction: "Reload the spreadsheet sidebar or add-in, then try the upload again."
      }
    });
  });
});

describe("attachment store", () => {
  it("generates high-entropy upload tokens for attachment content access", () => {
    const store = new AttachmentStore();

    const first = store.save({
      buffer: PNG_SIGNATURE_BYTES,
      mimeType: "image/png",
      fileName: "one.png",
      size: PNG_SIGNATURE_BYTES.length,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/one"
    });
    const second = store.save({
      buffer: PNG_SIGNATURE_BYTES,
      mimeType: "image/png",
      fileName: "two.png",
      size: PNG_SIGNATURE_BYTES.length,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/two"
    });

    expect(first.uploadToken).toMatch(/^upl_[A-Za-z0-9_-]{43}$/);
    expect(second.uploadToken).toMatch(/^upl_[A-Za-z0-9_-]{43}$/);
    expect(first.uploadToken).not.toBe(second.uploadToken);
  });

  it("expires stale attachments on access", () => {
    let nowMs = Date.UTC(2026, 3, 22, 6, 0, 0);
    const store = new AttachmentStore({
      ttlMs: 1_000,
      now: () => nowMs
    });

    const attachment = store.save({
      buffer: PNG_SIGNATURE_BYTES,
      mimeType: "image/png",
      fileName: "table.png",
      size: PNG_SIGNATURE_BYTES.length,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/demo",
      sessionId: "sess_img_real_001",
      workbookId: "sheet_import_demo"
    });

    expect(store.get(attachment.id)?.metadata.id).toBe(attachment.id);

    nowMs += 1_500;
    expect(store.get(attachment.id)).toBeUndefined();
  });

  it("evicts the oldest attachments when the in-memory store exceeds its cap", () => {
    let nowMs = Date.UTC(2026, 3, 22, 6, 0, 0);
    const store = new AttachmentStore({
      maxEntries: 2,
      ttlMs: 60_000,
      now: () => nowMs
    });

    const first = store.save({
      buffer: PNG_SIGNATURE_BYTES,
      mimeType: "image/png",
      fileName: "one.png",
      size: PNG_SIGNATURE_BYTES.length,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/one"
    });
    nowMs += 1_000;
    const second = store.save({
      buffer: PNG_SIGNATURE_BYTES,
      mimeType: "image/png",
      fileName: "two.png",
      size: PNG_SIGNATURE_BYTES.length,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/two"
    });
    nowMs += 1_000;
    const third = store.save({
      buffer: PNG_SIGNATURE_BYTES,
      mimeType: "image/png",
      fileName: "three.png",
      size: PNG_SIGNATURE_BYTES.length,
      source: "upload",
      previewUrl: "http://127.0.0.1:8787/api/uploads/three"
    });

    expect(store.get(first.id)).toBeUndefined();
    expect(store.get(second.id)?.metadata.id).toBe(second.id);
    expect(store.get(third.id)?.metadata.id).toBe(third.id);
  });
});
