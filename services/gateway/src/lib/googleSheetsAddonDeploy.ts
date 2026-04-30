import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

export type GoogleSheetsForceExtractionMode = "real" | "demo" | "unavailable" | null;

export type GoogleSheetsAddonDeploymentOptions = {
  gatewayBaseUrl: string;
  clientVersion?: string;
  reviewerSafeMode?: boolean;
  forceExtractionMode?: GoogleSheetsForceExtractionMode;
};

export type ClaspAuthorizedUser = {
  client_id: string;
  client_secret: string;
  refresh_token: string;
  access_token?: string;
  type?: string;
};

export const GOOGLE_SHEETS_ADDON_SOURCE_FILES = [
  "apps/google-sheets-addon/appsscript.json",
  "apps/google-sheets-addon/src/Code.gs",
  "apps/google-sheets-addon/src/ReferencedCells.js",
  "apps/google-sheets-addon/src/Wave1Plans.js",
  "apps/google-sheets-addon/html/Sidebar.html",
  "apps/google-sheets-addon/html/Sidebar.css.html",
  "apps/google-sheets-addon/html/Sidebar.js.html"
] as const;

export const GOOGLE_SHEETS_ADDON_STAGE_FILES = [
  {
    sourcePath: "apps/google-sheets-addon/appsscript.json",
    targetPath: "appsscript.json"
  },
  {
    sourcePath: "apps/google-sheets-addon/src/Code.gs",
    targetPath: "src/Code.gs"
  },
  {
    sourcePath: "apps/google-sheets-addon/src/ReferencedCells.js",
    targetPath: "src/ReferencedCells.js"
  },
  {
    sourcePath: "apps/google-sheets-addon/src/Wave1Plans.js",
    targetPath: "src/Wave1Plans.js"
  },
  {
    sourcePath: "apps/google-sheets-addon/html/Sidebar.html",
    targetPath: "html/Sidebar.html"
  },
  {
    sourcePath: "apps/google-sheets-addon/html/Sidebar.css.html",
    targetPath: "html/Sidebar.css.html"
  },
  {
    sourcePath: "apps/google-sheets-addon/html/Sidebar.js.html",
    targetPath: "html/Sidebar.js.html"
  }
] as const;

export const GOOGLE_SHEETS_DEPLOYMENT_CONFIG_PATH = "src/HermesDeploymentConfig.gs";

function toPortableStagePath(stageDir: string, targetPath: string): string {
  return path.relative(stageDir, targetPath).split(path.sep).join("/");
}

export function getDefaultClaspCredentialsPath(): string {
  return path.join(os.homedir(), ".clasprc.json");
}

function formatMissingClaspCredentialsMessage(user: string): string {
  return `No usable clasp OAuth credentials were found for user "${user}".`;
}

export function extractGoogleSpreadsheetId(input: string): string {
  const value = String(input || "").trim();
  if (!value) {
    throw new Error("A Google Sheets spreadsheet id or URL is required.");
  }

  const urlMatch = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch?.[1]) {
    return urlMatch[1];
  }

  if (/^[a-zA-Z0-9-_]+$/.test(value)) {
    return value;
  }

  throw new Error("Could not extract a Google Sheets spreadsheet id from the provided input.");
}

export function isAppsScriptReachableGatewayBaseUrl(baseUrl: string): boolean {
  const value = String(baseUrl || "").trim();
  let parsedUrl: URL;
  try {
    parsedUrl = new URL(value);
  } catch {
    return false;
  }

  const protocol = String(parsedUrl.protocol || "").toLowerCase();
  const host = String(parsedUrl.hostname || "").toLowerCase();
  if (protocol !== "https:") {
    return false;
  }
  if (!host) {
    return false;
  }

  if (host === "localhost" || host === "127.0.0.1" || host === "0.0.0.0") {
    return false;
  }

  if (/^10\./.test(host) || /^192\.168\./.test(host) || /^172\.(1[6-9]|2\d|3[0-1])\./.test(host)) {
    return false;
  }

  const decodeMappedIpv4 = (suffix: string): string | null => {
    if (/^\d{1,3}(?:\.\d{1,3}){3}$/.test(suffix)) {
      return suffix;
    }

    const hexMatch = suffix.match(/^([0-9a-f]{1,4}):([0-9a-f]{1,4})$/i);
    if (!hexMatch) {
      return null;
    }

    const high = Number.parseInt(hexMatch[1], 16);
    const low = Number.parseInt(hexMatch[2], 16);
    if (!Number.isFinite(high) || !Number.isFinite(low)) {
      return null;
    }

    return [
      (high >> 8) & 0xff,
      high & 0xff,
      (low >> 8) & 0xff,
      low & 0xff
    ].join(".");
  };

  const normalizedIpv6Host = host.replace(/^\[|\]$/g, "");
  if (normalizedIpv6Host.includes(":")) {
    if (normalizedIpv6Host === "::" || normalizedIpv6Host === "::1") {
      return false;
    }

    const mappedIpv4 = /^::ffff:/i.test(normalizedIpv6Host)
      ? decodeMappedIpv4(normalizedIpv6Host.replace(/^::ffff:/i, ""))
      : null;
    if (mappedIpv4) {
      if (
        mappedIpv4 === "127.0.0.1" ||
        mappedIpv4 === "0.0.0.0" ||
        /^10\./.test(mappedIpv4) ||
        /^192\.168\./.test(mappedIpv4) ||
        /^172\.(1[6-9]|2\d|3[0-1])\./.test(mappedIpv4)
      ) {
        return false;
      }
    }

    const firstHextet = normalizedIpv6Host
      .split(":")
      .find((segment) => segment.length > 0);
    if (firstHextet) {
      const firstValue = Number.parseInt(firstHextet, 16);
      if (Number.isFinite(firstValue)) {
        if ((firstValue & 0xfe00) === 0xfc00) {
          return false;
        }
        if ((firstValue & 0xffc0) === 0xfe80) {
          return false;
        }
      }
    }
  }

  return true;
}

export function buildGoogleSheetsDeploymentConfigSource(
  input: GoogleSheetsAddonDeploymentOptions
): string {
  const gatewayBaseUrl = String(input.gatewayBaseUrl || "").trim();
  if (!gatewayBaseUrl) {
    throw new Error("gatewayBaseUrl is required.");
  }

  const forceExtractionMode = input.forceExtractionMode ?? null;
  if (
    forceExtractionMode !== null &&
    forceExtractionMode !== "real" &&
    forceExtractionMode !== "demo" &&
    forceExtractionMode !== "unavailable"
  ) {
    throw new Error("forceExtractionMode must be null, real, demo, or unavailable.");
  }

  const overrides = {
    gatewayBaseUrl,
    clientVersion: String(input.clientVersion || "google-sheets-addon-live-demo").trim(),
    reviewerSafeMode: Boolean(input.reviewerSafeMode),
    forceExtractionMode
  };

  return [
    "function getHermesDeploymentOverrides() {",
    `  return ${JSON.stringify(overrides, null, 2).replace(/\n/g, "\n  ")};`,
    "}",
    ""
  ].join("\n");
}

export function buildClaspConfig(scriptId: string): string {
  const resolvedScriptId = String(scriptId || "").trim();
  if (!resolvedScriptId) {
    throw new Error("scriptId is required.");
  }

  return `${JSON.stringify({ scriptId: resolvedScriptId, rootDir: "." }, null, 2)}\n`;
}

export async function readClaspAuthorizedUser(options?: {
  credentialsPath?: string;
  user?: string;
}): Promise<ClaspAuthorizedUser> {
  const credentialsPath = path.resolve(
    options?.credentialsPath || getDefaultClaspCredentialsPath()
  );
  const user = String(options?.user || "default").trim() || "default";
  let raw: string;
  try {
    raw = await fs.readFile(credentialsPath, "utf8");
  } catch {
    throw new Error(formatMissingClaspCredentialsMessage(user));
  }

  let parsed: {
    tokens?: Record<string, Partial<ClaspAuthorizedUser>>;
  };
  try {
    parsed = JSON.parse(raw) as {
      tokens?: Record<string, Partial<ClaspAuthorizedUser>>;
    };
  } catch {
    throw new Error(formatMissingClaspCredentialsMessage(user));
  }

  const tokenRecord = parsed.tokens?.[user];

  if (
    !tokenRecord ||
    !tokenRecord.client_id ||
    !tokenRecord.client_secret ||
    !tokenRecord.refresh_token
  ) {
    throw new Error(formatMissingClaspCredentialsMessage(user));
  }

  return {
    client_id: String(tokenRecord.client_id),
    client_secret: String(tokenRecord.client_secret),
    refresh_token: String(tokenRecord.refresh_token),
    access_token: tokenRecord.access_token ? String(tokenRecord.access_token) : undefined,
    type: tokenRecord.type ? String(tokenRecord.type) : undefined
  };
}

export async function refreshClaspAccessToken(
  auth: ClaspAuthorizedUser,
  fetchImpl: typeof fetch = fetch
): Promise<string> {
  const response = await fetchImpl("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "content-type": "application/x-www-form-urlencoded"
    },
    body: new URLSearchParams({
      client_id: auth.client_id,
      client_secret: auth.client_secret,
      refresh_token: auth.refresh_token,
      grant_type: "refresh_token"
    })
  });

  if (!response.ok) {
    throw new Error(`Failed to refresh clasp OAuth access token (${response.status}).`);
  }

  const data = (await response.json()) as { access_token?: string };
  if (!data.access_token) {
    throw new Error("Google OAuth token refresh did not return an access_token.");
  }

  return String(data.access_token);
}

export async function createBoundGoogleSheetsScriptProject(options: {
  accessToken: string;
  spreadsheetId: string;
  title: string;
  fetchImpl?: typeof fetch;
}): Promise<{ scriptId: string }> {
  const spreadsheetId = extractGoogleSpreadsheetId(options.spreadsheetId);
  const title = String(options.title || "").trim() || "Hermes Agent";
  const fetchImpl = options.fetchImpl || fetch;

  const response = await fetchImpl("https://script.googleapis.com/v1/projects", {
    method: "POST",
    headers: {
      authorization: `Bearer ${options.accessToken}`,
      "content-type": "application/json"
    },
    body: JSON.stringify({
      title,
      parentId: spreadsheetId
    })
  });

  if (!response.ok) {
    throw new Error(`Failed to create a bound Apps Script project (${response.status}).`);
  }

  const data = (await response.json()) as { scriptId?: string };
  if (!data.scriptId) {
    throw new Error("Apps Script project creation did not return a scriptId.");
  }

  return {
    scriptId: String(data.scriptId)
  };
}

export async function stageGoogleSheetsAddonProject(options: {
  repoRoot: string;
  stageDir: string;
  deployment: GoogleSheetsAddonDeploymentOptions;
  scriptId?: string;
}): Promise<string[]> {
  const repoRoot = path.resolve(options.repoRoot);
  const stageDir = path.resolve(options.stageDir);

  await fs.mkdir(stageDir, { recursive: true });

  const writtenFiles: string[] = [];

  for (const file of GOOGLE_SHEETS_ADDON_STAGE_FILES) {
    const sourcePath = path.join(repoRoot, file.sourcePath);
    const targetPath = path.join(stageDir, file.targetPath);
    await fs.mkdir(path.dirname(targetPath), { recursive: true });
    await fs.copyFile(sourcePath, targetPath);
    writtenFiles.push(toPortableStagePath(stageDir, targetPath));
  }

  const configTargetPath = path.join(stageDir, GOOGLE_SHEETS_DEPLOYMENT_CONFIG_PATH);
  await fs.mkdir(path.dirname(configTargetPath), { recursive: true });
  await fs.writeFile(
    configTargetPath,
    buildGoogleSheetsDeploymentConfigSource(options.deployment),
    "utf8"
  );
  writtenFiles.push(toPortableStagePath(stageDir, configTargetPath));

  if (options.scriptId) {
    const claspConfigPath = path.join(stageDir, ".clasp.json");
    await fs.writeFile(claspConfigPath, buildClaspConfig(options.scriptId), "utf8");
    writtenFiles.push(toPortableStagePath(stageDir, claspConfigPath));
  }

  return writtenFiles.sort();
}
