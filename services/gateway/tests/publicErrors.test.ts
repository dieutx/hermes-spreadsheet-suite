import { describe, expect, it } from "vitest";
import {
  formatClientIssuePath,
  sanitizeClientIssueMessage
} from "../src/lib/publicErrors.ts";

describe("public error formatting", () => {
  it("sanitizes embedded secret markers in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage("Invalid value qa_HERMES_API_SERVER_KEY")).toBe(
      "Invalid request field."
    );
    expect(formatClientIssuePath(["context", "selection_HERMES_AGENT_BASE_URL", "values"])).toBe(
      "(redacted)"
    );
  });
});
