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

  it("sanitizes UNC paths in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage(
      String.raw`Invalid value produced near \\runner\share\schema.ts:42`
    )).toBe("Invalid request field.");
    expect(sanitizeClientIssueMessage(
      String.raw`Invalid value produced at path=\\runner\share\schema.ts:42`
    )).toBe("Invalid request field.");
    expect(formatClientIssuePath([
      "context",
      String.raw`\\runner\share\debug`,
      "values"
    ])).toBe("(redacted)");
  });

  it("sanitizes wrapped UNC paths in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage(
      String.raw`Invalid value produced near ("\\runner\share\schema.ts:42")`
    )).toBe("Invalid request field.");
    expect(formatClientIssuePath([
      "context",
      String.raw`("\\runner\share\debug")`,
      "values"
    ])).toBe("(redacted)");
  });

  it("sanitizes macOS user paths in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage(
      "Invalid value produced near /Users/runner/work/schema.ts:42"
    )).toBe("Invalid request field.");
    expect(formatClientIssuePath([
      "context",
      "/Users/runner/work/debug",
      "values"
    ])).toBe("(redacted)");
  });

  it("sanitizes link-local metadata URLs in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage(
      "Invalid value produced near http://169.254.169.254/latest/meta-data"
    )).toBe("Invalid request field.");
    expect(formatClientIssuePath([
      "context",
      "http://169.254.169.254/latest/meta-data",
      "values"
    ])).toBe("(redacted)");
  });

  it("sanitizes private IPv6 metadata URLs in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage(
      "Invalid value produced near http://[fd00::1]/latest/meta-data"
    )).toBe("Invalid request field.");
    expect(formatClientIssuePath([
      "context",
      "http://[fe80::1]/latest/meta-data",
      "values"
    ])).toBe("(redacted)");
  });

  it("sanitizes IPv4-mapped private metadata URLs in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage(
      "Invalid value produced near http://[::ffff:7f00:1]/latest/meta-data"
    )).toBe("Invalid request field.");
    expect(formatClientIssuePath([
      "context",
      "http://[::ffff:192.168.1.10]/latest/meta-data",
      "values"
    ])).toBe("(redacted)");
  });

  it("sanitizes numeric private IPv4 metadata URLs in validation issue messages and paths", () => {
    expect(sanitizeClientIssueMessage(
      "Invalid value produced near http://2130706433/latest/meta-data"
    )).toBe("Invalid request field.");
    expect(formatClientIssuePath([
      "context",
      "http://2852039166/latest/meta-data",
      "values"
    ])).toBe("(redacted)");
  });
});
