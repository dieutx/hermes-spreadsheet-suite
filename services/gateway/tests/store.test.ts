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
});
