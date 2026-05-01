import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";
import { describe, expect, it } from "vitest";

const AUTHOR_LOCAL_REPO_PATH = "/root/claude/hermes-spreadsheet-suite";
const DOCUMENTATION_ROOTS = ["docs", ".github"];
const ROOT_DOCUMENTS = ["README.md", "CODEX_IMPLEMENTATION_PROMPT.md", ".env.example"];

function collectFiles(path: string): string[] {
  const absolutePath = resolve(process.cwd(), path);
  const stats = statSync(absolutePath);
  if (stats.isFile()) {
    return [absolutePath];
  }

  return readdirSync(absolutePath)
    .flatMap((entry) => collectFiles(join(path, entry)));
}

describe("documentation sensitive path scan", () => {
  it("does not expose author-local checkout paths", () => {
    const files = [
      ...DOCUMENTATION_ROOTS.flatMap((root) => collectFiles(root)),
      ...ROOT_DOCUMENTS.map((path) => resolve(process.cwd(), path))
    ];

    const offenders = files
      .filter((path) => readFileSync(path, "utf8").includes(AUTHOR_LOCAL_REPO_PATH))
      .map((path) => path.replace(`${process.cwd()}/`, ""));

    expect(offenders).toEqual([]);
  });
});
