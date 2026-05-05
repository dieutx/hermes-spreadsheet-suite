import { describe, expect, it } from "vitest";
import { AttachmentStore } from "../src/lib/store.ts";

describe("AttachmentStore", () => {
  it("falls back when uploaded file names contain embedded secret markers", () => {
    const store = new AttachmentStore();

    const attachment = store.save({
      buffer: Buffer.from("png"),
      mimeType: "image/png",
      fileName: "table_HERMES_API_SERVER_KEY=secret.png",
      size: 3,
      source: "upload",
      previewUrl: "/api/uploads/att_1/content"
    });

    expect(attachment.fileName).toBe("uploaded-image.png");
    expect(attachment.fileName).not.toContain("HERMES_API_SERVER_KEY");
    expect(attachment.fileName).not.toContain("secret");
    expect(store.get(attachment.id)?.metadata.fileName).toBe("uploaded-image.png");
  });

  it("falls back when uploaded file names contain generic secret assignments", () => {
    const store = new AttachmentStore();

    const attachment = store.save({
      buffer: Buffer.from("png"),
      mimeType: "image/png",
      fileName: "DATABASE_PASSWORD=secret_123.png",
      size: 3,
      source: "upload",
      previewUrl: "/api/uploads/att_1/content"
    });

    expect(attachment.fileName).toBe("uploaded-image.png");
    expect(attachment.fileName).not.toContain("DATABASE_PASSWORD");
    expect(attachment.fileName).not.toContain("secret_123");
    expect(store.get(attachment.id)?.metadata.fileName).toBe("uploaded-image.png");
  });

  it("falls back when uploaded file names contain local filesystem paths", () => {
    const store = new AttachmentStore();

    const drivePathAttachment = store.save({
      buffer: Buffer.from("png"),
      mimeType: "image/png",
      fileName: String.raw`C:\Users\runner\work\private-table.png`,
      size: 3,
      source: "upload",
      previewUrl: "/api/uploads/att_2/content"
    });
    const uncPathAttachment = store.save({
      buffer: Buffer.from("png"),
      mimeType: "image/png",
      fileName: String.raw`\\runner\share\private-table.png`,
      size: 3,
      source: "upload",
      previewUrl: "/api/uploads/att_3/content"
    });
    const macPathAttachment = store.save({
      buffer: Buffer.from("png"),
      mimeType: "image/png",
      fileName: "/Users/runner/work/private-table.png",
      size: 3,
      source: "upload",
      previewUrl: "/api/uploads/att_4/content"
    });

    expect(drivePathAttachment.fileName).toBe("uploaded-image.png");
    expect(uncPathAttachment.fileName).toBe("uploaded-image.png");
    expect(macPathAttachment.fileName).toBe("uploaded-image.png");
    expect(drivePathAttachment.fileName).not.toContain("private-table");
    expect(uncPathAttachment.fileName).not.toContain("private-table");
    expect(macPathAttachment.fileName).not.toContain("private-table");
  });

  it("falls back when uploaded file names contain wrapped UNC paths", () => {
    const store = new AttachmentStore();

    const attachment = store.save({
      buffer: Buffer.from("png"),
      mimeType: "image/png",
      fileName: String.raw`Screenshot (\\runner\share\private-table.png)`,
      size: 3,
      source: "upload",
      previewUrl: "/api/uploads/att_5/content"
    });

    expect(attachment.fileName).toBe("uploaded-image.png");
    expect(attachment.fileName).not.toContain("private-table");
    expect(store.get(attachment.id)?.metadata.fileName).toBe("uploaded-image.png");
  });
});
