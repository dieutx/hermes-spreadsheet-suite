import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import {
  createBoundGoogleSheetsScriptProject,
  GOOGLE_SHEETS_ADDON_SOURCE_FILES,
  GOOGLE_SHEETS_ADDON_STAGE_FILES,
  GOOGLE_SHEETS_DEPLOYMENT_CONFIG_PATH,
  buildGoogleSheetsDeploymentConfigSource,
  buildClaspConfig,
  extractGoogleSpreadsheetId,
  isAppsScriptReachableGatewayBaseUrl,
  readClaspAuthorizedUser,
  refreshClaspAccessToken,
  stageGoogleSheetsAddonProject
} from "../src/lib/googleSheetsAddonDeploy";

const REPO_ROOT = path.resolve(__dirname, "../../..");
const tempDirs: string[] = [];

afterEach(async () => {
  await Promise.all(
    tempDirs.splice(0).map((dir) => fs.rm(dir, { recursive: true, force: true }))
  );
});

describe("Google Sheets add-on deploy helpers", () => {
  it("extracts the spreadsheet id from both raw ids and full Google Sheets URLs", () => {
    expect(extractGoogleSpreadsheetId("1Smz8ctI5gpahOuAP-p0l8VTE7oGeppnMsSVnzVU1YmA")).toBe(
      "1Smz8ctI5gpahOuAP-p0l8VTE7oGeppnMsSVnzVU1YmA"
    );
    expect(
      extractGoogleSpreadsheetId(
        "https://docs.google.com/spreadsheets/d/1Smz8ctI5gpahOuAP-p0l8VTE7oGeppnMsSVnzVU1YmA/edit#gid=0"
      )
    ).toBe("1Smz8ctI5gpahOuAP-p0l8VTE7oGeppnMsSVnzVU1YmA");
  });

  it("rejects malformed spreadsheet inputs", () => {
    expect(() => extractGoogleSpreadsheetId("")).toThrow("required");
    expect(() => extractGoogleSpreadsheetId("not a spreadsheet url")).toThrow(
      "Could not extract"
    );
  });

  it("accepts public gateway hosts and rejects localhost or private-network ones", () => {
    expect(isAppsScriptReachableGatewayBaseUrl("https://gateway.example.com")).toBe(true);
    expect(isAppsScriptReachableGatewayBaseUrl("https://[2606:4700::1111]")).toBe(true);
    expect(isAppsScriptReachableGatewayBaseUrl("http://gateway.example.com")).toBe(false);
    expect(isAppsScriptReachableGatewayBaseUrl("http://localhost:8787")).toBe(false);
    expect(isAppsScriptReachableGatewayBaseUrl("http://127.0.0.1:8787")).toBe(false);
    expect(isAppsScriptReachableGatewayBaseUrl("http://192.168.1.10:8787")).toBe(false);
    expect(isAppsScriptReachableGatewayBaseUrl("https://[::1]:8787")).toBe(false);
    expect(isAppsScriptReachableGatewayBaseUrl("https://[fc00::1]:8787")).toBe(false);
    expect(isAppsScriptReachableGatewayBaseUrl("https://[fe80::1]:8787")).toBe(false);
    expect(isAppsScriptReachableGatewayBaseUrl("https://[::ffff:127.0.0.1]:8787")).toBe(false);
  });

  it("builds a deployment config source file with the requested overrides", () => {
    const source = buildGoogleSheetsDeploymentConfigSource({
      gatewayBaseUrl: "https://gateway.example.com",
      clientVersion: "google-sheets-addon-live-demo",
      reviewerSafeMode: true,
      forceExtractionMode: "demo"
    });

    expect(source).toContain("function getHermesDeploymentOverrides()");
    expect(source).toContain('"gatewayBaseUrl": "https://gateway.example.com"');
    expect(source).toContain('"clientVersion": "google-sheets-addon-live-demo"');
    expect(source).toContain('"reviewerSafeMode": true');
    expect(source).toContain('"forceExtractionMode": "demo"');
  });

  it("rejects non-public gateway urls when building deployment config", () => {
    for (const gatewayBaseUrl of [
      "http://gateway.example.com",
      "http://127.0.0.1:8787",
      "https://[::1]:8787"
    ]) {
      expect(() => buildGoogleSheetsDeploymentConfigSource({
        gatewayBaseUrl
      })).toThrow(
        "gatewayBaseUrl must be a public HTTPS URL reachable from Google Apps Script."
      );
    }
  });

  it("sanitizes clasp credential loading failures", async () => {
    const stageDir = await fs.mkdtemp(path.join(os.tmpdir(), "hermes-gs-addon-creds-test-"));
    tempDirs.push(stageDir);

    const missingPath = path.join(stageDir, "private", ".clasprc.json");
    let missingMessage = "";
    try {
      await readClaspAuthorizedUser({ credentialsPath: missingPath, user: "default" });
    } catch (error) {
      missingMessage = (error as Error).message;
    }

    expect(missingMessage).toBe('No usable clasp OAuth credentials were found for user "default".');
    expect(missingMessage).not.toContain(stageDir);
    expect(missingMessage).not.toContain(".clasprc.json");

    const incompletePath = path.join(stageDir, ".clasprc.json");
    await fs.writeFile(
      incompletePath,
      JSON.stringify({ tokens: { default: { client_id: "client_123" } } }),
      "utf8"
    );

    let incompleteMessage = "";
    try {
      await readClaspAuthorizedUser({ credentialsPath: incompletePath, user: "default" });
    } catch (error) {
      incompleteMessage = (error as Error).message;
    }

    expect(incompleteMessage).toBe('No usable clasp OAuth credentials were found for user "default".');
    expect(incompleteMessage).not.toContain(stageDir);
    expect(incompleteMessage).not.toContain(".clasprc.json");
  });

  it("refreshes the clasp access token from the stored OAuth credentials", async () => {
    const fetchMock = async () =>
      ({
        ok: true,
        json: async () => ({ access_token: "access_123" })
      }) as Response;

    await expect(
      refreshClaspAccessToken(
        {
          client_id: "client_123",
          client_secret: "secret_123",
          refresh_token: "refresh_123"
        },
        fetchMock
      )
    ).resolves.toBe("access_123");
  });

  it("sanitizes OAuth refresh failure messages", async () => {
    const fetchMock = async () =>
      ({
        ok: false,
        status: 401,
        text: async () =>
          "invalid_grant refresh_token=refresh_123 client_secret=secret_123 stack=/srv/internal"
      }) as Response;

    let message = "";
    try {
      await refreshClaspAccessToken(
        {
          client_id: "client_123",
          client_secret: "secret_123",
          refresh_token: "refresh_123"
        },
        fetchMock
      );
    } catch (error) {
      message = (error as Error).message;
    }

    expect(message).toBe("Failed to refresh clasp OAuth access token (401).");
    expect(message).not.toContain("refresh_123");
    expect(message).not.toContain("secret_123");
    expect(message).not.toContain("/srv/internal");
  });

  it("creates a bound Apps Script project for the target spreadsheet through the official API", async () => {
    const fetchMock = async (url: string, init?: RequestInit) => {
      expect(url).toBe("https://script.googleapis.com/v1/projects");
      expect(init?.method).toBe("POST");
      expect(init?.headers).toMatchObject({
        authorization: "Bearer access_123"
      });
      expect(JSON.parse(String(init?.body))).toEqual({
        title: "Hermes Agent",
        parentId: "1Smz8ctI5gpahOuAP-p0l8VTE7oGeppnMsSVnzVU1YmA"
      });
      return {
        ok: true,
        json: async () => ({ scriptId: "script_123" })
      } as Response;
    };

    await expect(
      createBoundGoogleSheetsScriptProject({
        accessToken: "access_123",
        spreadsheetId:
          "https://docs.google.com/spreadsheets/d/1Smz8ctI5gpahOuAP-p0l8VTE7oGeppnMsSVnzVU1YmA/edit#gid=0",
        title: "Hermes Agent",
        fetchImpl: fetchMock
      })
    ).resolves.toEqual({ scriptId: "script_123" });
  });

  it("sanitizes Apps Script project creation failure messages", async () => {
    const spreadsheetId = "1Smz8ctI5gpahOuAP-p0l8VTE7oGeppnMsSVnzVU1YmA";
    const fetchMock = async () =>
      ({
        ok: false,
        status: 403,
        text: async () =>
          `denied parentId=${spreadsheetId} access_token=access_123 stack=/srv/internal`
      }) as Response;

    let message = "";
    try {
      await createBoundGoogleSheetsScriptProject({
        accessToken: "access_123",
        spreadsheetId,
        title: "Hermes Agent",
        fetchImpl: fetchMock
      });
    } catch (error) {
      message = (error as Error).message;
    }

    expect(message).toBe("Failed to create a bound Apps Script project (403).");
    expect(message).not.toContain(spreadsheetId);
    expect(message).not.toContain("access_123");
    expect(message).not.toContain("/srv/internal");
  });

  it("stages the add-on project with repo files, generated deployment config, and clasp config", async () => {
    const stageDir = await fs.mkdtemp(path.join(os.tmpdir(), "hermes-gs-addon-test-"));
    tempDirs.push(stageDir);

    const writtenFiles = await stageGoogleSheetsAddonProject({
      repoRoot: REPO_ROOT,
      stageDir,
      scriptId: "script_123",
      deployment: {
        gatewayBaseUrl: "https://gateway.example.com"
      }
    });

    for (const file of GOOGLE_SHEETS_ADDON_STAGE_FILES) {
      await expect(fs.readFile(path.join(stageDir, file.targetPath), "utf8")).resolves.toBeTruthy();
    }

    await expect(
      fs.readFile(path.join(stageDir, GOOGLE_SHEETS_DEPLOYMENT_CONFIG_PATH), "utf8")
    ).resolves.toContain("https://gateway.example.com");

    expect(await fs.readFile(path.join(stageDir, ".clasp.json"), "utf8")).toBe(
      buildClaspConfig("script_123")
    );
    expect(writtenFiles).toContain(".clasp.json");
    expect(writtenFiles).toContain(GOOGLE_SHEETS_DEPLOYMENT_CONFIG_PATH);
    expect(writtenFiles).toContain("appsscript.json");
    expect(writtenFiles).toContain("src/Code.gs");
    expect(writtenFiles).toContain("html/Sidebar.html");
    expect(writtenFiles).not.toContain(GOOGLE_SHEETS_ADDON_SOURCE_FILES[0]);
    expect(writtenFiles.every((file) => !file.includes("\\"))).toBe(true);
  });
});
