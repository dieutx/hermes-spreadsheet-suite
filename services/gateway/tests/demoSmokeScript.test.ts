import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

describe("demo smoke harness script", () => {
  const script = readFileSync(resolve(process.cwd(), "scripts/run_demo_smoke.py"), "utf8");

  it("derives the repository root from the checkout instead of an author-local path", () => {
    expect(script).not.toContain("/root/claude/hermes-spreadsheet-suite");
    expect(script).toContain("HERMES_DEMO_REPO_ROOT");
    expect(script).toContain("Path(__file__).resolve().parents[1]");
  });
});
