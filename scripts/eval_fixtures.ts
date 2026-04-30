import fs from "node:fs";
import path from "node:path";
import { isDeepStrictEqual } from "node:util";
import {
  HermesStructuredBodySchema,
  normalizeHermesStructuredBodyInput,
  type HermesStructuredBody
} from "../services/gateway/src/hermes/structuredBody.ts";
import { HermesRequestSchema } from "../packages/contracts/src/index.ts";

export type CapabilityFixtureHost = "excel" | "google_sheets" | "gateway";

export interface CapabilityFixtureExpectation {
  valid?: boolean;
  paths?: Record<string, unknown>;
  errorIncludes?: string;
}

export interface CapabilityFixture {
  id: string;
  host: CapabilityFixtureHost;
  family: string;
  description?: string;
  structuredBody?: unknown;
  request?: unknown;
  expect?: CapabilityFixtureExpectation;
}

export interface CapabilityFixturePack {
  version: 1;
  fixtures: CapabilityFixture[];
}

export interface CapabilityFixtureResult {
  id: string;
  host: CapabilityFixtureHost | "unknown";
  family: string;
  status: "passed" | "failed";
  expectedValid: boolean;
  actualValid: boolean;
  checks: string[];
  errors: string[];
}

export interface CapabilityFixtureSummary {
  source: string;
  total: number;
  passed: number;
  failed: number;
  results: CapabilityFixtureResult[];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function stringifyValue(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

function getPathValue(value: unknown, pathExpression: string): unknown {
  return pathExpression.split(".").reduce<unknown>((current, segment) => {
    if (Array.isArray(current) && /^\d+$/.test(segment)) {
      return current[Number(segment)];
    }

    if (isRecord(current)) {
      return current[segment];
    }

    return undefined;
  }, value);
}

function formatSchemaIssues(error: { issues: Array<{ path: Array<string | number>; message: string }> }): string[] {
  return error.issues.map((issue) => {
    const issuePath = issue.path.length > 0 ? issue.path.join(".") : "(root)";
    return `${issuePath}: ${issue.message}`;
  });
}

function countExtractedRows(body: HermesStructuredBody): number {
  if (body.type === "extracted_table") {
    return Array.isArray(body.data.rows) ? body.data.rows.length : 0;
  }

  if (body.type === "sheet_import_plan") {
    return Array.isArray(body.data.values) ? body.data.values.length : 0;
  }

  return 0;
}

function validateReviewerSafeUnavailableInvariant(body: HermesStructuredBody): string[] {
  if (
    (body.type === "extracted_table" || body.type === "sheet_import_plan") &&
    body.data.extractionMode === "unavailable" &&
    countExtractedRows(body) > 0
  ) {
    return [
      "reviewer-safe unavailable mode must not include extracted table rows or import values"
    ];
  }

  return [];
}

function evaluateStructuredBody(fixture: CapabilityFixture): {
  normalized?: unknown;
  actualValid: boolean;
  validationErrors: string[];
} {
  const normalized = normalizeHermesStructuredBodyInput(fixture.structuredBody);
  const parsed = HermesStructuredBodySchema.safeParse(normalized);

  if (!parsed.success) {
    return {
      normalized,
      actualValid: false,
      validationErrors: formatSchemaIssues(parsed.error)
    };
  }

  const invariantErrors = validateReviewerSafeUnavailableInvariant(parsed.data);
  return {
    normalized: parsed.data,
    actualValid: invariantErrors.length === 0,
    validationErrors: invariantErrors
  };
}

function evaluateRequest(fixture: CapabilityFixture): {
  normalized?: unknown;
  actualValid: boolean;
  validationErrors: string[];
} {
  const parsed = HermesRequestSchema.safeParse(fixture.request);

  if (!parsed.success) {
    return {
      normalized: fixture.request,
      actualValid: false,
      validationErrors: formatSchemaIssues(parsed.error)
    };
  }

  return {
    normalized: parsed.data,
    actualValid: true,
    validationErrors: []
  };
}

function evaluateFixturePayload(fixture: CapabilityFixture): {
  normalized?: unknown;
  actualValid: boolean;
  validationErrors: string[];
  checks: string[];
} {
  if (fixture.request !== undefined) {
    return {
      ...evaluateRequest(fixture),
      checks: ["request_contract"]
    };
  }

  if (fixture.structuredBody !== undefined) {
    return {
      ...evaluateStructuredBody(fixture),
      checks: ["structured_body"]
    };
  }

  return {
    actualValid: false,
    validationErrors: ["Fixture must include request or structuredBody."],
    checks: ["fixture_shape"]
  };
}

export function evaluateCapabilityFixture(fixture: CapabilityFixture): CapabilityFixtureResult {
  const expectedValid = fixture.expect?.valid ?? true;
  const { normalized, actualValid, validationErrors, checks } = evaluateFixturePayload(fixture);
  const assertionErrors: string[] = [];

  if (actualValid !== expectedValid) {
    assertionErrors.push(`Expected valid=${expectedValid}, got valid=${actualValid}.`);
  }

  if (fixture.expect?.errorIncludes) {
    const expectedText = fixture.expect.errorIncludes;
    if (!validationErrors.some((error) => error.includes(expectedText))) {
      assertionErrors.push(`Expected an error containing "${expectedText}".`);
    }
  }

  if (fixture.expect?.paths) {
    if (!actualValid) {
      assertionErrors.push("Path expectations require a valid structured body.");
    } else {
      for (const [pathExpression, expectedValue] of Object.entries(fixture.expect.paths)) {
        const actualValue = getPathValue(normalized, pathExpression);
        if (!isDeepStrictEqual(actualValue, expectedValue)) {
          assertionErrors.push(
            `Expected ${pathExpression} to equal ${stringifyValue(expectedValue)}, got ${stringifyValue(actualValue)}.`
          );
        }
      }
    }
  }

  return {
    id: fixture.id || "(missing id)",
    host: fixture.host || "unknown",
    family: fixture.family || "(missing family)",
    status: assertionErrors.length === 0 ? "passed" : "failed",
    expectedValid,
    actualValid,
    checks,
    errors: [...validationErrors, ...assertionErrors]
  };
}

export function evaluateCapabilityFixturePack(
  pack: CapabilityFixturePack,
  options: { source?: string } = {}
): CapabilityFixtureSummary {
  const source = options.source ?? "inline";
  const results = (pack.fixtures || []).map((fixture) => evaluateCapabilityFixture(fixture));
  const passed = results.filter((result) => result.status === "passed").length;

  return {
    source,
    total: results.length,
    passed,
    failed: results.length - passed,
    results
  };
}

function mergeSummaries(source: string, summaries: CapabilityFixtureSummary[]): CapabilityFixtureSummary {
  const results = summaries.flatMap((summary) => summary.results);
  const passed = results.filter((result) => result.status === "passed").length;

  return {
    source,
    total: results.length,
    passed,
    failed: results.length - passed,
    results
  };
}

function loadJsonFile(filePath: string): unknown {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function normalizePack(value: unknown, filePath: string): CapabilityFixturePack {
  if (isRecord(value) && value.version === 1 && Array.isArray(value.fixtures)) {
    return value as unknown as CapabilityFixturePack;
  }

  if (Array.isArray(value)) {
    return {
      version: 1,
      fixtures: value as CapabilityFixture[]
    };
  }

  throw new Error(`Fixture file ${filePath} must contain a versioned pack or fixture array.`);
}

export function evaluateCapabilityFixtureDirectory(
  directory: string,
  options: { source?: string } = {}
): CapabilityFixtureSummary {
  const source = options.source ?? directory;
  const files = fs.readdirSync(directory)
    .filter((fileName) => fileName.endsWith(".json"))
    .sort();

  const summaries = files.map((fileName) => {
    const filePath = path.join(directory, fileName);
    return evaluateCapabilityFixturePack(normalizePack(loadJsonFile(filePath), filePath), {
      source: filePath
    });
  });

  return mergeSummaries(source, summaries);
}

function parseArgs(argv: string[]): { directory: string; json: boolean } {
  let directory = path.join(process.cwd(), "fixtures/capability-eval");
  let json = false;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--json") {
      json = true;
    } else if (arg === "--dir") {
      directory = path.resolve(argv[index + 1] || "");
      index += 1;
    } else if (arg.startsWith("--dir=")) {
      directory = path.resolve(arg.slice("--dir=".length));
    }
  }

  return { directory, json };
}

function renderTextSummary(summary: CapabilityFixtureSummary): string {
  const header = `Capability fixture eval: ${summary.passed}/${summary.total} passed`;
  const lines = summary.results.map((result) => {
    const suffix = result.errors.length > 0 ? ` - ${result.errors.join("; ")}` : "";
    return `${result.status.toUpperCase()} ${result.id} [${result.host}/${result.family}]${suffix}`;
  });

  return [header, ...lines].join("\n");
}

export function runCapabilityFixtureCli(argv: string[] = process.argv.slice(2)): number {
  const { directory, json } = parseArgs(argv);
  const summary = evaluateCapabilityFixtureDirectory(directory);

  if (json) {
    console.log(JSON.stringify(summary, null, 2));
  } else {
    console.log(renderTextSummary(summary));
  }

  return summary.failed === 0 && summary.total > 0 ? 0 : 1;
}

if (process.argv[1] && path.basename(process.argv[1]) === "eval_fixtures.ts") {
  process.exitCode = runCapabilityFixtureCli();
}
