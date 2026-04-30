/* global Office, Excel */

import "../commands/commands.js?v=20260423w";

import {
  normalizeExcelFormulaText,
  normalizeExcelHeaderText,
  normalizeExcelCellValue,
  normalizeExcelMatrixValues,
  truncateContextText
} from "./cellValues.js?v=20260423w";
import { getPromptReferencedA1Notations } from "./referencedCells.js?v=20260423w";
import { rangeHasExistingContent } from "./rangeSafety.js?v=20260423w";
import {
  buildAnalysisReportMatrix,
  getAnalysisReportPreviewSummary,
  getAnalysisReportStatusSummary,
  getChartPreviewSummary,
  getChartStatusSummary,
  getPivotTablePreviewSummary,
  getPivotTableStatusSummary,
  isChartPlan,
  isMaterializedAnalysisReportPlan,
  isPivotTablePlan,
  resolveMaterializedAnalysisReportPlan
} from "./analysisArtifactsPlan.js?v=20260423w";
import {
  buildCompositeStepPreview,
  getCompositeStatusSummary,
  getCompositePreviewSummary,
  isCompositePlan
} from "./compositePlan.js?v=20260423w";
import {
  expandRangeBorderLines,
  getConditionalFormatStatusSummary,
  getDataCleanupStatusSummary,
  getRangeFormatStatusSummary,
  getRangeTransferStatusSummary,
  getTablePreviewSummary,
  getTableStatusSummary,
  getWorkbookStructureStatusSummary,
  getDataValidationStatusSummary,
  getNamedRangeStatusSummary,
  getCompositeStepWritebackStatusLine,
  getWritebackStatusLine,
  isRangeFormatPlan,
  isTablePlan,
  isWorkbookStructurePlan,
  mapHorizontalAlignmentToExcel,
  mapVerticalAlignmentToExcel,
  mapWrapStrategyToExcel
} from "./writePlan.js?v=20260423w";
import {
  getSheetStructureStatusSummary,
  isSheetStructurePlan
} from "./structurePlan.js?v=20260423w";
import {
  buildExcelSortFields,
  getRangeFilterStatusSummary,
  getRangeSortStatusSummary,
  isRangeFilterPlan,
  isRangeSortPlan
} from "./sortFilterPlan.js?v=20260423w";

const SUPPORTED_IMAGE_TYPES = new Set([
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/webp"
]);
const WORKBOOK_SESSION_ID_SETTING = "Hermes.WorkbookSessionId";
const WORKBOOK_EPHEMERAL_ID_STORAGE_KEY = "Hermes.WorkbookEphemeralId";
let inMemoryWorkbookEphemeralId = null;

function getDefaultGatewayBaseUrl() {
  if (window.location.protocol === "https:") {
    return `${window.location.origin}/hermes-gateway`;
  }

  return "http://127.0.0.1:8787";
}

function normalizeGatewayBaseUrl(value) {
  const trimmed = String(value || "").trim().replace(/\/+$/, "");
  if (!trimmed) {
    return "";
  }

  try {
    const url = new URL(trimmed);
    return url.protocol === "http:" || url.protocol === "https:" ? trimmed : "";
  } catch {
    return "";
  }
}

function generateClientUuid() {
  const cryptoObject = globalThis.crypto;
  if (cryptoObject && typeof cryptoObject.randomUUID === "function") {
    return cryptoObject.randomUUID();
  }

  if (cryptoObject && typeof cryptoObject.getRandomValues === "function") {
    const bytes = new Uint8Array(16);
    cryptoObject.getRandomValues(bytes);
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = Array.from(bytes, (value) => value.toString(16).padStart(2, "0"));
    return `${hex.slice(0, 4).join("")}-${hex.slice(4, 6).join("")}-${hex.slice(6, 8).join("")}-${hex.slice(8, 10).join("")}-${hex.slice(10, 16).join("")}`;
  }

  return `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;
}

function resolveGatewayBaseUrl() {
  const configuredGateway = new URLSearchParams(window.location.search).get("gateway");
  if (configuredGateway && configuredGateway.trim()) {
    return normalizeGatewayBaseUrl(configuredGateway);
  }

  const storedGateway = safeStorageGetItem(window.localStorage, "hermesGatewayBaseUrl");
  if (storedGateway && storedGateway.trim()) {
    return normalizeGatewayBaseUrl(storedGateway);
  }

  return normalizeGatewayBaseUrl(getDefaultGatewayBaseUrl());
}

const gatewayBaseUrl = resolveGatewayBaseUrl();
const sessionId = safeStorageGetItem(window.localStorage, "hermesSessionId") ||
  `sess_${generateClientUuid()}`;
safeStorageSetItem(window.localStorage, "hermesSessionId", sessionId);
const MESSAGE_SCROLL_BOTTOM_THRESHOLD_PX = 40;
const MESSAGE_SCROLL_FOLLOWUP_DELAYS_MS = [0, 32, 120, 320, 640];
const MESSAGE_POLL_INTERVAL_MS = 900;
const MESSAGE_POLL_MAX_INTERVAL_MS = 5000;
const AUTO_OPEN_SETTING = "Hermes.EnableAutoOpen";
const MAX_REQUEST_MESSAGE_LENGTH = 16000;
const MAX_CONVERSATION_MESSAGES = 50;
const MAX_STORED_MESSAGES = 100;
const MAX_MESSAGE_TRACE_EVENTS = 200;
const LOCAL_EXECUTION_SNAPSHOT_STORAGE_PREFIX = "Hermes.ReversibleExecutions.v1::";
const MAX_LOCAL_EXECUTION_SNAPSHOTS = 100;
const EXECUTION_HISTORY_SHORTCUT_LIMIT = 20;
const REQUEST_TRUNCATION_SUFFIX = "...";
const TRACE_POLL_EVERY_N_ATTEMPTS = 3;
const UNDO_PROMPT_PATTERN = /^\s*(?:please\s+)?undo(?:\s+(?:that|it|this|last|latest|previous)(?:\s+(?:write|change|update))?)?\s*[.!?]*\s*$/i;
const REDO_PROMPT_PATTERN = /^\s*(?:please\s+)?redo(?:\s+(?:that|it|this|last|latest|previous)(?:\s+(?:write|change|update))?)?\s*[.!?]*\s*$/i;
const UNDER_SPECIFIED_AFFIRMATION_PATTERN = /^\s*(?:yes|yep|yeah|ok|okay|sure|please do|do it|go ahead|continue)\s*[.!?]*\s*$/i;

const state = {
  messages: [],
  pendingAttachments: [],
  messageScrollPinned: true,
  messageScrollTimeoutIds: [],
  messageScrollListenersBound: false,
  messageLayoutObserver: null,
  messageMutationObserver: null
};

const elements = {
  app: document.getElementById("app"),
  messages: document.getElementById("messages"),
  prompt: document.getElementById("prompt"),
  sendButton: document.getElementById("send-button"),
  fileInput: document.getElementById("file-input"),
  attachmentStrip: document.getElementById("attachment-strip")
};

function saveDocumentSettingsAsync() {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(result.error);
    });
  });
}

async function ensureDemoStartupDefaults() {
  const operations = [];
  let shouldPersistSettings = false;
  const settings = Office.context?.document?.settings;
  const runtimeConfig = getRuntimeConfig();
  const autoOpenEnabled = runtimeConfig.enableDocumentAutoOpen ||
    settings?.get?.(AUTO_OPEN_SETTING) === true;

  if (ensureWorkbookSessionIdentity()) {
    shouldPersistSettings = true;
  }

  if (autoOpenEnabled) {
    if (Office.addin?.setStartupBehavior && Office.StartupBehavior?.load) {
      operations.push(Office.addin.setStartupBehavior(Office.StartupBehavior.load));
    }

    if (settings?.get?.("Office.AutoShowTaskpaneWithDocument") !== true) {
      settings?.set?.("Office.AutoShowTaskpaneWithDocument", true);
      shouldPersistSettings = true;
    }

    if (settings?.get?.(AUTO_OPEN_SETTING) !== true) {
      settings?.set?.(AUTO_OPEN_SETTING, true);
      shouldPersistSettings = true;
    }
  }

  if (shouldPersistSettings) {
    operations.push(saveDocumentSettingsAsync());
  }

  if (operations.length === 0) {
    return;
  }

  const results = await Promise.allSettled(operations);
  for (const result of results) {
    if (result.status === "rejected") {
      console.warn("Hermes startup default could not be persisted.", result.reason);
    }
  }
}

function parseBooleanSetting(value) {
  return value === "true" || value === "1";
}

function getRuntimeConfig() {
  const params = new URLSearchParams(window.location.search);
  const forceExtractionMode = params.get("forceExtractionMode") ||
    safeStorageGetItem(window.localStorage, "hermesForceExtractionMode");

  return {
    gatewayBaseUrl,
    clientVersion: Office.context?.diagnostics?.version || "excel-addin-dev",
    enableDocumentAutoOpen: parseBooleanSetting(
      params.get("enableDocumentAutoOpen") ||
      safeStorageGetItem(window.localStorage, "hermesEnableDocumentAutoOpen")
    ),
    reviewerSafeMode: parseBooleanSetting(
      params.get("reviewerSafeMode") ||
      safeStorageGetItem(window.localStorage, "hermesReviewerSafeMode")
    ),
    forceExtractionMode: forceExtractionMode === "real" ||
      forceExtractionMode === "demo" ||
      forceExtractionMode === "unavailable"
      ? forceExtractionMode
      : null
  };
}

function detectExcelPlatform() {
  if (Office.context?.platform === Office.PlatformType.Mac) {
    return "excel_macos";
  }

  return "excel_windows";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeHtmlAttribute(value) {
  return escapeHtml(value)
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeA1(address) {
  const text = String(address || "");
  const withoutSheet = text.includes("!") ? text.split("!").pop() : text;
  return withoutSheet.replaceAll("$", "");
}

function normalizeFormulaMatrix(formulas) {
  return (formulas || []).map((row) =>
    row.map((cell) => normalizeExcelFormulaText(cell))
  );
}

function shouldIncludeRegionMatrix(bounds) {
  return Boolean(bounds) && (bounds.rowCount * bounds.columnCount) <= 400;
}

function buildImplicitRegionTargets(rangeA1) {
  if (!rangeA1) {
    return {};
  }

  const bounds = parseA1RangeReference(rangeA1);
  return {
    currentRegionArtifactTarget: buildA1RangeFromBounds({
      startRow: bounds.endRow + 2,
      endRow: bounds.endRow + 2,
      startColumn: bounds.startColumn,
      endColumn: bounds.startColumn
    }),
    currentRegionAppendTarget: buildA1RangeFromBounds({
      startRow: bounds.endRow + 1,
      endRow: bounds.endRow + 1,
      startColumn: bounds.startColumn,
      endColumn: bounds.endColumn
    })
  };
}

function getSelectionHeaders(values) {
  const firstRow = values?.[0] || [];
  if (firstRow.length === 0) {
    return undefined;
  }

  return firstRow.every((cell) => typeof cell === "string" && cell.trim().length > 0)
    ? firstRow
      .map((cell) => normalizeExcelHeaderText(cell))
      .filter((cell) => typeof cell === "string")
    : undefined;
}

function createRequestId() {
  return `req_${generateClientUuid()}`;
}

function hashWorkbookIdentity(value) {
  let hash = 2166136261;
  const text = String(value || "").trim().toLowerCase();

  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }

  return (hash >>> 0).toString(36);
}

function getEphemeralWorkbookIdentity() {
  const sessionStorage = window.sessionStorage;
  const existing = sessionStorage?.getItem?.(WORKBOOK_EPHEMERAL_ID_STORAGE_KEY);
  if (typeof existing === "string" && existing.trim().length > 0) {
    return existing.trim();
  }

  if (typeof inMemoryWorkbookEphemeralId === "string" && inMemoryWorkbookEphemeralId.trim().length > 0) {
    return inMemoryWorkbookEphemeralId.trim();
  }

  const generated = `local_${generateClientUuid()}`;
  inMemoryWorkbookEphemeralId = generated;
  sessionStorage?.setItem?.(WORKBOOK_EPHEMERAL_ID_STORAGE_KEY, generated);
  return generated;
}

function ensureWorkbookSessionIdentity() {
  const settings = Office.context?.document?.settings;
  if (!settings?.get || !settings?.set) {
    return false;
  }

  const existing = settings.get(WORKBOOK_SESSION_ID_SETTING);
  if (typeof existing === "string" && existing.trim().length > 0) {
    return false;
  }

  const rawUrl = Office.context?.document?.url;
  if (typeof rawUrl === "string" && rawUrl.trim().length > 0) {
    return false;
  }

  settings.set(WORKBOOK_SESSION_ID_SETTING, `local_${generateClientUuid()}`);
  return true;
}

function getWorkbookIdentity() {
  const settings = Office.context?.document?.settings;
  const persisted = settings?.get?.(WORKBOOK_SESSION_ID_SETTING);
  if (typeof persisted === "string" && persisted.trim().length > 0) {
    return persisted.trim();
  }

  const rawUrl = Office.context?.document?.url;
  if (typeof rawUrl === "string" && rawUrl.trim().length > 0) {
    return `url_${hashWorkbookIdentity(rawUrl)}`;
  }

  ensureWorkbookSessionIdentity();
  const generated = settings?.get?.(WORKBOOK_SESSION_ID_SETTING);
  if (typeof generated === "string" && generated.trim().length > 0) {
    return generated.trim();
  }

  return getEphemeralWorkbookIdentity();
}

function getWorkbookSessionKey() {
  return `${detectExcelPlatform()}::${getWorkbookIdentity()}`;
}

function safeStorageGetItem(storage, key) {
  try {
    return storage?.getItem?.(key) ?? null;
  } catch {
    return null;
  }
}

function safeStorageSetItem(storage, key, value) {
  try {
    storage?.setItem?.(key, value);
    return true;
  } catch {
    return false;
  }
}

function getLocalExecutionSnapshotStoreKey(workbookSessionKey) {
  return `${LOCAL_EXECUTION_SNAPSHOT_STORAGE_PREFIX}${workbookSessionKey}`;
}

function createEmptyLocalExecutionSnapshotStore() {
  return {
    version: 1,
    order: [],
    executions: {},
    bases: {}
  };
}

function normalizeLocalExecutionSnapshotStore(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return createEmptyLocalExecutionSnapshotStore();
  }

  return {
    version: 1,
    order: Array.isArray(value.order)
      ? value.order.filter((entry) => typeof entry === "string" && entry.trim().length > 0)
      : [],
    executions: value.executions && typeof value.executions === "object" && !Array.isArray(value.executions)
      ? value.executions
      : {},
    bases: value.bases && typeof value.bases === "object" && !Array.isArray(value.bases)
      ? value.bases
      : {}
  };
}

function readLocalExecutionSnapshotStore(workbookSessionKey) {
  const raw = safeStorageGetItem(
    window.localStorage,
    getLocalExecutionSnapshotStoreKey(workbookSessionKey)
  );
  if (!raw) {
    return createEmptyLocalExecutionSnapshotStore();
  }

  try {
    return normalizeLocalExecutionSnapshotStore(JSON.parse(raw));
  } catch {
    return createEmptyLocalExecutionSnapshotStore();
  }
}

function pruneLocalExecutionSnapshotStore(store) {
  const normalized = normalizeLocalExecutionSnapshotStore(store);
  while (normalized.order.length > MAX_LOCAL_EXECUTION_SNAPSHOTS) {
    const removedExecutionId = normalized.order.shift();
    if (!removedExecutionId) {
      break;
    }
    delete normalized.executions[removedExecutionId];
  }

  const referencedBaseIds = new Set(
    Object.values(normalized.executions)
      .map((entry) => entry && typeof entry === "object" ? entry.baseExecutionId : null)
      .filter((entry) => typeof entry === "string" && entry.trim().length > 0)
  );

  for (const baseExecutionId of Object.keys(normalized.bases)) {
    if (!referencedBaseIds.has(baseExecutionId)) {
      delete normalized.bases[baseExecutionId];
    }
  }

  return normalized;
}

function writeLocalExecutionSnapshotStore(workbookSessionKey, store) {
  return safeStorageSetItem(
    window.localStorage,
    getLocalExecutionSnapshotStoreKey(workbookSessionKey),
    JSON.stringify(pruneLocalExecutionSnapshotStore(store))
  );
}

function serializeExecutionSnapshotScalar(value) {
  if (isDateObject(value)) {
    return {
      type: "date",
      value: value.toISOString()
    };
  }

  if (value === null) {
    return {
      type: "null"
    };
  }

  if (value === undefined) {
    return {
      type: "blank"
    };
  }

  if (typeof value === "number" || typeof value === "string" || typeof value === "boolean") {
    return {
      type: typeof value,
      value
    };
  }

  return {
    type: "string",
    value: String(value)
  };
}

function deserializeExecutionSnapshotScalar(serialized) {
  if (!serialized || typeof serialized !== "object") {
    return "";
  }

  switch (serialized.type) {
    case "date":
      return typeof serialized.value === "string" ? new Date(serialized.value) : "";
    case "null":
      return null;
    case "number":
    case "string":
    case "boolean":
      return serialized.value;
    case "blank":
    default:
      return "";
  }
}

function normalizeExcelNoteContent(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value);
}

function buildExecutionSnapshotCellMatrix(values, formulas, notes) {
  const includeNotes = Array.isArray(notes);
  return (values || []).map((row, rowIndex) =>
    (row || []).map((value, columnIndex) => {
      const formulaValue = normalizeExcelFormulaText(formulas?.[rowIndex]?.[columnIndex]);
      const snapshotCell = typeof formulaValue === "string" && formulaValue.trim().startsWith("=")
        ? {
            kind: "formula",
            formula: formulaValue
          }
        : {
            kind: "value",
            value: serializeExecutionSnapshotScalar(value)
          };

      if (includeNotes) {
        snapshotCell.note = normalizeExcelNoteContent(notes?.[rowIndex]?.[columnIndex]);
      }

      return snapshotCell;
    })
  );
}

function hasExecutionSnapshotNote(snapshotCell) {
  return Object.prototype.hasOwnProperty.call(snapshotCell || {}, "note");
}

function executionSnapshotCellsHaveNotes(cells) {
  return (cells || []).some((row) =>
    (row || []).some((cell) => hasExecutionSnapshotNote(cell))
  );
}

function getExcelNoteCollection(workbook, sheet) {
  const collection = sheet?.notes || workbook?.notes;
  if (
    !collection ||
    typeof collection.add !== "function" ||
    typeof collection.getItemOrNullObject !== "function"
  ) {
    throw new Error("This Excel host does not support exact note write-back.");
  }

  return {
    collection,
    useWorksheetRelativeAddress: Boolean(sheet?.notes)
  };
}

function getExcelContextNoteCollection(workbook, sheet) {
  const collection = sheet?.notes || workbook?.notes;
  if (!collection || typeof collection.getItemOrNullObject !== "function") {
    return null;
  }

  return {
    collection,
    useWorksheetRelativeAddress: Boolean(sheet?.notes)
  };
}

function getExcelNoteAddress(cell, noteCollection) {
  const rawAddress = typeof cell?.address === "string" ? cell.address.trim() : "";
  if (!rawAddress) {
    throw new Error("Excel note write-back could not resolve the target cell address.");
  }

  if (!noteCollection.useWorksheetRelativeAddress) {
    return rawAddress;
  }

  return rawAddress
    .slice(rawAddress.lastIndexOf("!") + 1)
    .replace(/\$/g, "");
}

function queueExcelContextNoteRead(noteCollection, cell) {
  if (!noteCollection || !cell) {
    return null;
  }

  try {
    const address = getExcelNoteAddress(cell, noteCollection);
    const note = noteCollection.collection.getItemOrNullObject(address);
    if (typeof note?.load === "function") {
      note.load("content");
    }
    return note || null;
  } catch {
    return null;
  }
}

function readExcelContextNote(note) {
  if (!note || note.isNullObject) {
    return undefined;
  }

  const content = truncateContextText(normalizeExcelNoteContent(note.content));
  return content.length > 0 ? content : undefined;
}

async function prepareExcelNoteTargets(context, noteCollection, target, rowCount, columnCount) {
  const noteTargets = Array.from({ length: rowCount }, () => Array.from({ length: columnCount }, () => null));

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const cell = target.getCell(rowIndex, columnIndex);
      if (typeof cell.load === "function") {
        cell.load("address");
      }

      noteTargets[rowIndex][columnIndex] = {
        cell,
        note: null
      };
    }
  }

  await context.sync();

  for (const row of noteTargets) {
    for (const noteTarget of row) {
      const address = getExcelNoteAddress(noteTarget.cell, noteCollection);
      const note = noteCollection.collection.getItemOrNullObject(address);
      if (typeof note.load === "function") {
        note.load("content");
      }

      noteTarget.note = note;
    }
  }

  await context.sync();
  return noteTargets;
}

function readPreparedExcelNoteMatrix(noteTargets) {
  return (noteTargets || []).map((row) =>
    (row || []).map((noteTarget) => {
      const note = noteTarget?.note;
      if (!note || note.isNullObject) {
        return "";
      }

      return normalizeExcelNoteContent(note.content);
    })
  );
}

function getExcelNoteMatrixAfterWrite(beforeNotes, notes, options = {}) {
  const afterNotes = cloneMatrix(beforeNotes);
  for (let rowIndex = 0; rowIndex < afterNotes.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < (afterNotes[rowIndex] || []).length; columnIndex += 1) {
      const nextNote = normalizeExcelNoteContent(notes?.[rowIndex]?.[columnIndex]);
      if (options.clearBlankNotes || nextNote.length > 0) {
        afterNotes[rowIndex][columnIndex] = nextNote;
      }
    }
  }

  return afterNotes;
}

function queueExcelPreparedNoteWrite(noteCollection, noteTarget, noteValue, options = {}) {
  const content = normalizeExcelNoteContent(noteValue);
  if (!options.clearBlankNote && content.length === 0) {
    return;
  }

  const note = noteTarget?.note;
  if (note && !note.isNullObject) {
    if (content.length > 0) {
      note.content = content;
      return;
    }

    if (typeof note.delete === "function") {
      note.delete();
      return;
    }

    throw new Error("This Excel host does not support clearing notes exactly.");
  }

  if (content.length > 0) {
    noteCollection.collection.add(noteTarget.cell, content);
  }
}

function queueExcelNoteMatrixWrites(noteCollection, noteTargets, notes, options = {}) {
  for (let rowIndex = 0; rowIndex < (noteTargets || []).length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < (noteTargets[rowIndex] || []).length; columnIndex += 1) {
      queueExcelPreparedNoteWrite(
        noteCollection,
        noteTargets[rowIndex][columnIndex],
        notes?.[rowIndex]?.[columnIndex],
        { clearBlankNote: Boolean(options.clearBlankNotes) }
      );
    }
  }
}

function restoreExecutionSnapshotNotes(noteCollection, noteTargets, cells) {
  for (let rowIndex = 0; rowIndex < (cells || []).length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < (cells[rowIndex] || []).length; columnIndex += 1) {
      const snapshotCell = cells[rowIndex][columnIndex];
      if (!hasExecutionSnapshotNote(snapshotCell)) {
        continue;
      }

      queueExcelPreparedNoteWrite(
        noteCollection,
        noteTargets[rowIndex]?.[columnIndex],
        snapshotCell.note,
        { clearBlankNote: true }
      );
    }
  }
}

function isExcelNoteWriteSupported() {
  const requirements = Office?.context?.requirements;
  if (!requirements || typeof requirements.isSetSupported !== "function") {
    return false;
  }

  return requirements.isSetSupported("ExcelApi", "1.18");
}

function createLocalExecutionSnapshot({
  executionId,
  targetSheet,
  targetRange,
  beforeValues,
  beforeFormulas,
  beforeNotes,
  afterValues,
  afterFormulas,
  afterNotes
}) {
  if (!executionId || !targetSheet || !targetRange) {
    return null;
  }

  return {
    baseExecutionId: executionId,
    targetSheet,
    targetRange,
    beforeCells: buildExecutionSnapshotCellMatrix(beforeValues, beforeFormulas, beforeNotes),
    afterCells: buildExecutionSnapshotCellMatrix(afterValues, afterFormulas, afterNotes)
  };
}

function getExcelRangeFormatSnapshotFields(format) {
  const fields = [];
  const addField = (planField, snapshotField) => {
    if (format?.[planField] !== undefined) {
      fields.push(snapshotField || planField);
    }
  };

  addField("backgroundColor");
  addField("textColor");
  addField("fontFamily");
  addField("fontSize");
  addField("bold");
  addField("italic");
  addField("underline");
  addField("strikethrough");
  addField("horizontalAlignment");
  addField("verticalAlignment");
  addField("wrapStrategy", "wrapText");
  addField("numberFormat");
  addField("columnWidth");
  addField("rowHeight");

  return fields;
}

function getExcelRangeFormatSnapshotBorderSides(format) {
  if (!format?.border) {
    return [];
  }

  return [...new Set(
    expandRangeBorderLines(format.border)
      .map(({ side }) => side)
      .filter(Boolean)
  )];
}

function getExcelRangeFormatSnapshotLoadPaths(fields) {
  const paths = [];

  if (fields.includes("backgroundColor")) {
    paths.push("format/fill/color");
  }

  if (fields.includes("textColor")) {
    paths.push("format/font/color");
  }

  if (fields.includes("fontFamily")) {
    paths.push("format/font/name");
  }

  if (fields.includes("fontSize")) {
    paths.push("format/font/size");
  }

  if (fields.includes("bold")) {
    paths.push("format/font/bold");
  }

  if (fields.includes("italic")) {
    paths.push("format/font/italic");
  }

  if (fields.includes("underline")) {
    paths.push("format/font/underline");
  }

  if (fields.includes("strikethrough")) {
    paths.push("format/font/strikethrough");
  }

  if (fields.includes("horizontalAlignment")) {
    paths.push("format/horizontalAlignment");
  }

  if (fields.includes("verticalAlignment")) {
    paths.push("format/verticalAlignment");
  }

  if (fields.includes("wrapText")) {
    paths.push("format/wrapText");
  }

  if (fields.includes("numberFormat")) {
    paths.push("numberFormat");
  }

  if (fields.includes("columnWidth")) {
    paths.push("format/columnWidth");
  }

  if (fields.includes("rowHeight")) {
    paths.push("format/rowHeight");
  }

  return paths;
}

function prepareExcelRangeFormatSnapshotTargets(target, format) {
  const fields = getExcelRangeFormatSnapshotFields(format);
  const borderSides = getExcelRangeFormatSnapshotBorderSides(format);
  if (fields.length === 0 && borderSides.length === 0) {
    return null;
  }

  const rowCount = Math.max(0, Number(target?.rowCount) || 0);
  const columnCount = Math.max(0, Number(target?.columnCount) || 0);
  const cells = Array.from({ length: rowCount }, (_, rowIndex) =>
    Array.from({ length: columnCount }, (_, columnIndex) => target.getCell(rowIndex, columnIndex))
  );
  const borders = {};
  for (const side of borderSides) {
    const excelSide = mapRangeBorderSideToExcel(side);
    if (excelSide) {
      borders[side] = target.format.borders.getItem(excelSide);
    }
  }

  return {
    fields,
    borderSides,
    cells,
    borders
  };
}

function loadExcelRangeFormatSnapshotTargets(snapshotTargets) {
  if (!snapshotTargets) {
    return;
  }

  const loadPaths = getExcelRangeFormatSnapshotLoadPaths(snapshotTargets.fields);
  for (const row of snapshotTargets.cells) {
    for (const cell of row) {
      if (loadPaths.length > 0 && typeof cell?.load === "function") {
        cell.load(loadPaths);
      }
    }
  }

  for (const border of Object.values(snapshotTargets.borders || {})) {
    if (typeof border?.load === "function") {
      border.load(["lineStyle", "color", "weight"]);
    }
  }
}

function cloneExcelNumberFormatSnapshot(numberFormat) {
  return Array.isArray(numberFormat) ? cloneMatrix(numberFormat) : numberFormat;
}

function readExcelRangeFormatCellSnapshot(cell, fields) {
  const format = cell?.format || {};
  const font = format.font || {};
  const fill = format.fill || {};
  const snapshot = {};

  if (fields.includes("backgroundColor")) {
    snapshot.backgroundColor = fill.color;
  }

  if (fields.includes("textColor")) {
    snapshot.textColor = font.color;
  }

  if (fields.includes("fontFamily")) {
    snapshot.fontFamily = font.name;
  }

  if (fields.includes("fontSize")) {
    snapshot.fontSize = font.size;
  }

  if (fields.includes("bold")) {
    snapshot.bold = font.bold;
  }

  if (fields.includes("italic")) {
    snapshot.italic = font.italic;
  }

  if (fields.includes("underline")) {
    snapshot.underline = font.underline;
  }

  if (fields.includes("strikethrough")) {
    snapshot.strikethrough = font.strikethrough;
  }

  if (fields.includes("horizontalAlignment")) {
    snapshot.horizontalAlignment = format.horizontalAlignment;
  }

  if (fields.includes("verticalAlignment")) {
    snapshot.verticalAlignment = format.verticalAlignment;
  }

  if (fields.includes("wrapText")) {
    snapshot.wrapText = format.wrapText;
  }

  if (fields.includes("numberFormat")) {
    snapshot.numberFormat = cloneExcelNumberFormatSnapshot(cell.numberFormat);
  }

  if (fields.includes("columnWidth")) {
    snapshot.columnWidth = format.columnWidth;
  }

  if (fields.includes("rowHeight")) {
    snapshot.rowHeight = format.rowHeight;
  }

  return snapshot;
}

function readExcelRangeFormatBorderSnapshot(snapshotTargets) {
  const snapshot = {};
  for (const [side, border] of Object.entries(snapshotTargets?.borders || {})) {
    snapshot[side] = {
      lineStyle: border?.lineStyle,
      color: border?.color,
      weight: border?.weight
    };
  }

  return snapshot;
}

function readExcelRangeFormatCellsSnapshot(snapshotTargets) {
  return (snapshotTargets?.cells || []).map((row) =>
    (row || []).map((cell) => readExcelRangeFormatCellSnapshot(cell, snapshotTargets.fields))
  );
}

function createRangeFormatLocalExecutionSnapshot({
  executionId,
  targetSheet,
  targetRange,
  target,
  beforeFormatCells,
  afterFormatCells,
  beforeBorders,
  afterBorders
}) {
  if (
    !executionId ||
    !targetSheet ||
    !targetRange ||
    !Array.isArray(beforeFormatCells) ||
    beforeFormatCells.length === 0 ||
    !Array.isArray(afterFormatCells) ||
    afterFormatCells.length === 0
  ) {
    return null;
  }

  return {
    baseExecutionId: executionId,
    kind: "range_format",
    targetSheet,
    targetRange,
    shape: {
      rows: Number(target?.rowCount) || beforeFormatCells.length,
      columns: Number(target?.columnCount) || (beforeFormatCells[0]?.length || 0)
    },
    beforeFormatCells,
    afterFormatCells,
    ...(beforeBorders && Object.keys(beforeBorders).length > 0 ? { beforeBorders } : {}),
    ...(afterBorders && Object.keys(afterBorders).length > 0 ? { afterBorders } : {})
  };
}

function cloneLocalExecutionSnapshotValue(value) {
  if (value === undefined) {
    return undefined;
  }

  return JSON.parse(JSON.stringify(value));
}

function loadExcelDataValidationSnapshotTarget(target) {
  if (typeof target?.dataValidation?.load === "function") {
    target.dataValidation.load(["rule", "ignoreBlanks", "prompt", "errorAlert"]);
  }
}

function readExcelDataValidationSnapshot(target) {
  const dataValidation = target?.dataValidation;
  if (!dataValidation) {
    return null;
  }

  return {
    rule: cloneLocalExecutionSnapshotValue(dataValidation.rule),
    ignoreBlanks: dataValidation.ignoreBlanks,
    prompt: cloneLocalExecutionSnapshotValue(dataValidation.prompt),
    errorAlert: cloneLocalExecutionSnapshotValue(dataValidation.errorAlert)
  };
}

function createDataValidationLocalExecutionSnapshot({
  executionId,
  targetSheet,
  targetRange,
  target,
  beforeValidation,
  afterValidation
}) {
  if (!executionId || !targetSheet || !targetRange || !beforeValidation || !afterValidation) {
    return null;
  }

  return {
    baseExecutionId: executionId,
    kind: "data_validation",
    targetSheet,
    targetRange,
    shape: {
      rows: Number(target?.rowCount) || 0,
      columns: Number(target?.columnCount) || 0
    },
    beforeValidation,
    afterValidation
  };
}

function normalizeExcelNamedRangeReference(reference) {
  const resolvedReference = typeof reference === "string" ? reference.trim() : "";
  return resolvedReference.length > 0 ? resolvedReference : undefined;
}

function readExcelNamedRangeSnapshotState(namedRange, fallbackName) {
  if (!namedRange || namedRange.isNullObject === true) {
    return {
      exists: false,
      name: fallbackName
    };
  }

  const reference = normalizeExcelNamedRangeReference(namedRange.reference);
  return {
    exists: true,
    name: typeof namedRange.name === "string" && namedRange.name.length > 0 ? namedRange.name : fallbackName,
    ...(reference ? { reference } : {})
  };
}

async function readExcelNamedRangeSnapshotStateForName(context, collection, name) {
  if (!collection || !name) {
    return null;
  }

  let namedRange;
  if (typeof collection.getItemOrNullObject === "function") {
    namedRange = collection.getItemOrNullObject(name);
  } else if (typeof collection.getItem === "function") {
    namedRange = collection.getItem(name);
  }

  if (!namedRange) {
    return {
      exists: false,
      name
    };
  }

  if (typeof namedRange.load === "function") {
    namedRange.load(["name", "reference"]);
    await context.sync();
  }

  return readExcelNamedRangeSnapshotState(namedRange, name);
}

async function createNamedRangeLocalExecutionSnapshot({
  context,
  collection,
  executionId,
  plan,
  resolvedReference
}) {
  if (!executionId || !plan?.name || !plan?.scope) {
    return null;
  }

  const normalizedReference = normalizeExcelNamedRangeReference(resolvedReference);
  let before;
  let after;

  switch (plan.operation) {
    case "create":
      if (!normalizedReference) {
        return null;
      }
      before = {
        exists: false,
        name: plan.name
      };
      after = {
        exists: true,
        name: plan.name,
        reference: normalizedReference
      };
      break;
    case "retarget":
      if (!normalizedReference) {
        return null;
      }
      before = await readExcelNamedRangeSnapshotStateForName(context, collection, plan.name);
      if (!before?.exists || !before.reference) {
        return null;
      }
      after = {
        exists: true,
        name: before.name || plan.name,
        reference: normalizedReference
      };
      break;
    case "rename":
      before = await readExcelNamedRangeSnapshotStateForName(context, collection, plan.name);
      if (!before?.exists || !before.reference || !plan.newName) {
        return null;
      }
      after = {
        exists: true,
        name: plan.newName,
        reference: before.reference
      };
      break;
    case "delete":
      before = await readExcelNamedRangeSnapshotStateForName(context, collection, plan.name);
      if (!before?.exists || !before.reference) {
        return null;
      }
      after = {
        exists: false,
        name: before.name || plan.name
      };
      break;
    default:
      return null;
  }

  return {
    baseExecutionId: executionId,
    kind: "named_range",
    scope: plan.scope,
    ...(plan.scope === "sheet" && (plan.sheetName || plan.targetSheet)
      ? { sheetName: plan.sheetName || plan.targetSheet }
      : {}),
    before,
    after
  };
}

function createCompositeLocalExecutionSnapshot({ executionId, entries }) {
  if (!executionId || !Array.isArray(entries) || entries.length === 0) {
    return null;
  }

  return {
    baseExecutionId: executionId,
    entries
  };
}

function isCompositeLocalExecutionSnapshot(snapshot) {
  return Array.isArray(snapshot?.entries) && snapshot.entries.length > 0;
}

function isRangeFormatLocalExecutionSnapshot(snapshot) {
  return snapshot?.kind === "range_format";
}

function isDataValidationLocalExecutionSnapshot(snapshot) {
  return snapshot?.kind === "data_validation";
}

function isNamedRangeLocalExecutionSnapshot(snapshot) {
  return snapshot?.kind === "named_range";
}

function getLocalExecutionSnapshotEntriesForMode(snapshot, mode) {
  if (!isCompositeLocalExecutionSnapshot(snapshot)) {
    return [snapshot];
  }

  const entries = snapshot.entries.filter((entry) => entry && typeof entry === "object");
  return mode === "undo" ? [...entries].reverse() : entries;
}

function attachLocalExecutionSnapshot(result, snapshot) {
  if (!snapshot) {
    return result;
  }

  return {
    ...result,
    __hermesLocalExecutionSnapshot: snapshot
  };
}

function stripLocalExecutionSnapshot(result) {
  if (!result || typeof result !== "object" || Array.isArray(result)) {
    return result;
  }

  const {
    __hermesLocalExecutionSnapshot: _localExecutionSnapshot,
    ...rest
  } = result;
  return rest;
}

function persistLocalExecutionSnapshot(workbookSessionKey, executionId, snapshot) {
  if (!workbookSessionKey || !executionId || !snapshot?.baseExecutionId) {
    return false;
  }

  const store = readLocalExecutionSnapshotStore(workbookSessionKey);
  store.bases[snapshot.baseExecutionId] = snapshot;
  store.executions[executionId] = {
    baseExecutionId: snapshot.baseExecutionId
  };
  store.order = store.order.filter((entry) => entry !== executionId);
  store.order.push(executionId);
  return writeLocalExecutionSnapshotStore(workbookSessionKey, store);
}

function getLocalExecutionSnapshot(workbookSessionKey, executionId) {
  if (!workbookSessionKey || !executionId) {
    return null;
  }

  const store = readLocalExecutionSnapshotStore(workbookSessionKey);
  const executionEntry = store.executions[executionId];
  if (!executionEntry?.baseExecutionId) {
    return null;
  }

  return store.bases[executionEntry.baseExecutionId] || null;
}

function linkLocalExecutionSnapshot(workbookSessionKey, executionId, previousExecutionId) {
  const snapshot = getLocalExecutionSnapshot(workbookSessionKey, previousExecutionId);
  if (!snapshot) {
    return false;
  }

  return persistLocalExecutionSnapshot(workbookSessionKey, executionId, snapshot);
}

function assertLocalExecutionSnapshotStoreWritable(workbookSessionKey, snapshot, mode) {
  if (!workbookSessionKey || !snapshot?.baseExecutionId) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  const probeExecutionId = `probe_${mode}_${generateClientUuid()}`;
  const store = readLocalExecutionSnapshotStore(workbookSessionKey);
  const nextStore = {
    ...store,
    order: [...(store.order || []), probeExecutionId],
    executions: {
      ...(store.executions || {}),
      [probeExecutionId]: {
        baseExecutionId: snapshot.baseExecutionId
      }
    },
    bases: {
      ...(store.bases || {}),
      [snapshot.baseExecutionId]: snapshot
    }
  };

  if (!writeLocalExecutionSnapshotStore(workbookSessionKey, nextStore)) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  const cleanedStore = readLocalExecutionSnapshotStore(workbookSessionKey);
  delete cleanedStore.executions?.[probeExecutionId];
  cleanedStore.order = (cleanedStore.order || []).filter((entry) => entry !== probeExecutionId);
  writeLocalExecutionSnapshotStore(workbookSessionKey, cleanedStore);
}

function prepareGatewayWritebackResult(result, executionId, workbookSessionKey) {
  const snapshot = result?.__hermesLocalExecutionSnapshot;
  const strippedResult = stripLocalExecutionSnapshot(result);

  if (!snapshot || !executionId || !workbookSessionKey) {
    return strippedResult;
  }

  if (!persistLocalExecutionSnapshot(workbookSessionKey, executionId, snapshot)) {
    return strippedResult;
  }

  return {
    ...strippedResult,
    undoReady: true
  };
}

function hasSnapshotField(snapshot, field) {
  return Object.prototype.hasOwnProperty.call(snapshot || {}, field) &&
    snapshot[field] !== undefined;
}

function getRangeFormatSnapshotCellsForMode(snapshot, mode) {
  return mode === "undo" ? snapshot?.beforeFormatCells : snapshot?.afterFormatCells;
}

function getRangeFormatSnapshotBordersForMode(snapshot, mode) {
  return mode === "undo" ? snapshot?.beforeBorders : snapshot?.afterBorders;
}

function assertRangeFormatSnapshotShape(snapshot, target, cells) {
  const expectedRows = Number(snapshot?.shape?.rows) || cells?.length || 0;
  const expectedColumns = Number(snapshot?.shape?.columns) || cells?.[0]?.length || 0;
  if (!snapshot?.targetSheet || !snapshot?.targetRange || !Array.isArray(cells) || expectedRows <= 0 || expectedColumns <= 0) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  if (
    target.rowCount !== expectedRows ||
    target.columnCount !== expectedColumns ||
    cells.length !== expectedRows ||
    (cells[0]?.length || 0) !== expectedColumns
  ) {
    throw new Error("The saved undo snapshot no longer matches the current range shape.");
  }
}

function applyExcelRangeFormatCellSnapshot(cell, snapshot) {
  const format = cell?.format || {};

  if (hasSnapshotField(snapshot, "backgroundColor") && format.fill) {
    format.fill.color = snapshot.backgroundColor;
  }

  if (hasSnapshotField(snapshot, "textColor") && format.font) {
    format.font.color = snapshot.textColor;
  }

  if (hasSnapshotField(snapshot, "fontFamily") && format.font) {
    format.font.name = snapshot.fontFamily;
  }

  if (hasSnapshotField(snapshot, "fontSize") && format.font) {
    format.font.size = snapshot.fontSize;
  }

  if (hasSnapshotField(snapshot, "bold") && format.font) {
    format.font.bold = snapshot.bold;
  }

  if (hasSnapshotField(snapshot, "italic") && format.font) {
    format.font.italic = snapshot.italic;
  }

  if (hasSnapshotField(snapshot, "underline") && format.font) {
    format.font.underline = snapshot.underline;
  }

  if (hasSnapshotField(snapshot, "strikethrough") && format.font) {
    format.font.strikethrough = snapshot.strikethrough;
  }

  if (hasSnapshotField(snapshot, "horizontalAlignment")) {
    format.horizontalAlignment = snapshot.horizontalAlignment;
  }

  if (hasSnapshotField(snapshot, "verticalAlignment")) {
    format.verticalAlignment = snapshot.verticalAlignment;
  }

  if (hasSnapshotField(snapshot, "wrapText")) {
    format.wrapText = snapshot.wrapText;
  }

  if (hasSnapshotField(snapshot, "numberFormat")) {
    cell.numberFormat = cloneExcelNumberFormatSnapshot(snapshot.numberFormat);
  }

  if (hasSnapshotField(snapshot, "columnWidth")) {
    format.columnWidth = snapshot.columnWidth;
  }

  if (hasSnapshotField(snapshot, "rowHeight")) {
    format.rowHeight = snapshot.rowHeight;
  }
}

function applyExcelRangeFormatBorderSnapshot(target, borders) {
  if (!borders || typeof borders !== "object") {
    return;
  }

  for (const [side, snapshot] of Object.entries(borders)) {
    const excelSide = mapRangeBorderSideToExcel(side);
    if (!excelSide) {
      continue;
    }

    const border = target.format.borders.getItem(excelSide);
    if (hasSnapshotField(snapshot, "lineStyle")) {
      border.lineStyle = snapshot.lineStyle;
    }

    if (hasSnapshotField(snapshot, "color")) {
      border.color = snapshot.color;
    }

    if (hasSnapshotField(snapshot, "weight")) {
      border.weight = snapshot.weight;
    }
  }
}

async function restoreRangeFormatLocalExecutionSnapshotForMode(snapshot, mode) {
  const cells = getRangeFormatSnapshotCellsForMode(snapshot, mode);
  const borders = getRangeFormatSnapshotBordersForMode(snapshot, mode);

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const sheet = worksheets.getItem(snapshot.targetSheet);
    const target = sheet.getRange(snapshot.targetRange);
    target.load(["rowCount", "columnCount"]);
    await context.sync();

    assertRangeFormatSnapshotShape(snapshot, target, cells);

    for (let rowIndex = 0; rowIndex < cells.length; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < (cells[rowIndex] || []).length; columnIndex += 1) {
        applyExcelRangeFormatCellSnapshot(target.getCell(rowIndex, columnIndex), cells[rowIndex][columnIndex]);
      }
    }

    applyExcelRangeFormatBorderSnapshot(target, borders);
    await context.sync();
  });
}

async function validateRangeFormatLocalExecutionSnapshotForMode(snapshot, mode) {
  const cells = getRangeFormatSnapshotCellsForMode(snapshot, mode);

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const sheet = worksheets.getItem(snapshot.targetSheet);
    const target = sheet.getRange(snapshot.targetRange);
    target.load(["rowCount", "columnCount"]);
    await context.sync();

    assertRangeFormatSnapshotShape(snapshot, target, cells);
  });
}

function getDataValidationSnapshotForMode(snapshot, mode) {
  return mode === "undo" ? snapshot?.beforeValidation : snapshot?.afterValidation;
}

function assertDataValidationSnapshotShape(snapshot, target, validation) {
  const expectedRows = Number(snapshot?.shape?.rows) || 0;
  const expectedColumns = Number(snapshot?.shape?.columns) || 0;
  if (!snapshot?.targetSheet || !snapshot?.targetRange || !validation || expectedRows <= 0 || expectedColumns <= 0) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  if (target.rowCount !== expectedRows || target.columnCount !== expectedColumns) {
    throw new Error("The saved undo snapshot no longer matches the current range shape.");
  }
}

function applyExcelDataValidationSnapshot(target, validation) {
  const dataValidation = target?.dataValidation;
  if (!dataValidation) {
    throw new Error("Excel host does not support data validation on this range.");
  }

  if (validation.rule === undefined) {
    if (typeof dataValidation.clear === "function") {
      dataValidation.clear();
      return;
    }

    throw new Error("Excel host cannot clear data validation exactly for this undo snapshot.");
  }

  dataValidation.rule = cloneLocalExecutionSnapshotValue(validation.rule);
  if (validation.ignoreBlanks !== undefined) {
    dataValidation.ignoreBlanks = validation.ignoreBlanks;
  }
  if (validation.prompt !== undefined) {
    dataValidation.prompt = cloneLocalExecutionSnapshotValue(validation.prompt);
  }
  if (validation.errorAlert !== undefined) {
    dataValidation.errorAlert = cloneLocalExecutionSnapshotValue(validation.errorAlert);
  }
}

async function restoreDataValidationLocalExecutionSnapshotForMode(snapshot, mode) {
  const validation = getDataValidationSnapshotForMode(snapshot, mode);

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const sheet = worksheets.getItem(snapshot.targetSheet);
    const target = sheet.getRange(snapshot.targetRange);
    target.load(["rowCount", "columnCount"]);
    await context.sync();

    assertDataValidationSnapshotShape(snapshot, target, validation);
    applyExcelDataValidationSnapshot(target, validation);
    await context.sync();
  });
}

async function validateDataValidationLocalExecutionSnapshotForMode(snapshot, mode) {
  const validation = getDataValidationSnapshotForMode(snapshot, mode);

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const sheet = worksheets.getItem(snapshot.targetSheet);
    const target = sheet.getRange(snapshot.targetRange);
    target.load(["rowCount", "columnCount"]);
    await context.sync();

    assertDataValidationSnapshotShape(snapshot, target, validation);
  });
}

function getNamedRangeSnapshotTransitionForMode(snapshot, mode) {
  return mode === "undo"
    ? {
      from: snapshot?.after,
      to: snapshot?.before
    }
    : {
      from: snapshot?.before,
      to: snapshot?.after
    };
}

function assertNamedRangeSnapshotState(state) {
  if (!state || typeof state.name !== "string" || state.name.length === 0) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  if (state.exists === true && !normalizeExcelNamedRangeReference(state.reference)) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }
}

function getNamedRangeCollectionForSnapshot(context, snapshot) {
  if (!snapshot || (snapshot.scope !== "workbook" && snapshot.scope !== "sheet")) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  if (snapshot.scope === "sheet") {
    if (!snapshot.sheetName) {
      throw new Error("That history entry is no longer available in this spreadsheet session.");
    }

    const worksheet = context.workbook.worksheets.getItem(snapshot.sheetName);
    if (!worksheet?.names) {
      throw new Error("Excel host does not support named ranges on this scope.");
    }
    return worksheet.names;
  }

  if (!context.workbook.names) {
    throw new Error("Excel host does not support named ranges on this scope.");
  }
  return context.workbook.names;
}

async function resolveExcelNamedRangeForSnapshotRestore(context, collection, name) {
  if (!collection || !name) {
    return undefined;
  }

  let namedRange;
  if (typeof collection.getItemOrNullObject === "function") {
    namedRange = collection.getItemOrNullObject(name);
  } else if (typeof collection.getItem === "function") {
    namedRange = collection.getItem(name);
  }

  if (!namedRange) {
    return undefined;
  }

  if (typeof namedRange.load === "function") {
    namedRange.load(["name", "reference"]);
    await context.sync();
  }

  return namedRange.isNullObject === true ? undefined : namedRange;
}

async function applyNamedRangeSnapshotTransition(context, collection, from, to) {
  assertNamedRangeSnapshotState(from);
  assertNamedRangeSnapshotState(to);

  if (from.exists === false && to.exists === false) {
    return;
  }

  if (from.exists === false && to.exists === true) {
    if (typeof collection.add !== "function") {
      throw new Error("Excel host does not support creating named ranges on this scope.");
    }

    collection.add(to.name, to.reference);
    return;
  }

  const namedRange = await resolveExcelNamedRangeForSnapshotRestore(context, collection, from.name);
  if (!namedRange) {
    throw new Error(`Named range not found: ${from.name}`);
  }

  if (from.exists === true && to.exists === false) {
    if (typeof namedRange.delete !== "function") {
      throw new Error("Excel host does not support deleting named ranges on this scope.");
    }

    namedRange.delete();
    return;
  }

  if (from.name !== to.name) {
    namedRange.name = to.name;
  }

  if (normalizeExcelNamedRangeReference(namedRange.reference) !== normalizeExcelNamedRangeReference(to.reference)) {
    namedRange.reference = to.reference;
  }
}

async function restoreNamedRangeLocalExecutionSnapshotForMode(snapshot, mode) {
  const transition = getNamedRangeSnapshotTransitionForMode(snapshot, mode);

  return Excel.run(async (context) => {
    const collection = getNamedRangeCollectionForSnapshot(context, snapshot);
    await applyNamedRangeSnapshotTransition(context, collection, transition.from, transition.to);
    await context.sync();
  });
}

async function validateNamedRangeLocalExecutionSnapshotForMode(snapshot, mode) {
  const transition = getNamedRangeSnapshotTransitionForMode(snapshot, mode);

  return Excel.run(async (context) => {
    const collection = getNamedRangeCollectionForSnapshot(context, snapshot);
    assertNamedRangeSnapshotState(transition.from);
    assertNamedRangeSnapshotState(transition.to);

    if (transition.from.exists === true) {
      const namedRange = await resolveExcelNamedRangeForSnapshotRestore(context, collection, transition.from.name);
      if (!namedRange) {
        throw new Error(`Named range not found: ${transition.from.name}`);
      }
    }
  });
}

async function restoreLocalExecutionSnapshotForMode(snapshot, mode) {
  if (isCompositeLocalExecutionSnapshot(snapshot)) {
    for (const entry of getLocalExecutionSnapshotEntriesForMode(snapshot, mode)) {
      await restoreLocalExecutionSnapshotForMode(entry, mode);
    }
    return;
  }

  if (isRangeFormatLocalExecutionSnapshot(snapshot)) {
    return restoreRangeFormatLocalExecutionSnapshotForMode(snapshot, mode);
  }

  if (isDataValidationLocalExecutionSnapshot(snapshot)) {
    return restoreDataValidationLocalExecutionSnapshotForMode(snapshot, mode);
  }

  if (isNamedRangeLocalExecutionSnapshot(snapshot)) {
    return restoreNamedRangeLocalExecutionSnapshotForMode(snapshot, mode);
  }

  const cells = mode === "undo" ? snapshot?.beforeCells : snapshot?.afterCells;
  if (!snapshot?.targetSheet || !snapshot?.targetRange || !Array.isArray(cells) || cells.length === 0) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const sheet = worksheets.getItem(snapshot.targetSheet);
    const target = sheet.getRange(snapshot.targetRange);
    target.load(["rowCount", "columnCount"]);
    await context.sync();

    if (target.rowCount !== cells.length || target.columnCount !== (cells[0]?.length || 0)) {
      throw new Error("The saved undo snapshot no longer matches the current range shape.");
    }

    const shouldRestoreNotes = executionSnapshotCellsHaveNotes(cells);
    const noteCollection = shouldRestoreNotes ? getExcelNoteCollection(context.workbook, sheet) : null;
    const noteTargets = shouldRestoreNotes
      ? await prepareExcelNoteTargets(context, noteCollection, target, cells.length, cells[0]?.length || 0)
      : null;

    for (let rowIndex = 0; rowIndex < cells.length; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < (cells[rowIndex] || []).length; columnIndex += 1) {
        const cell = target.getCell(rowIndex, columnIndex);
        const snapshotCell = cells[rowIndex][columnIndex];
        if (snapshotCell?.kind === "formula" && typeof snapshotCell.formula === "string") {
          cell.formulas = [[snapshotCell.formula]];
        } else {
          cell.values = [[deserializeExecutionSnapshotScalar(snapshotCell?.value)]];
        }
      }
    }

    if (shouldRestoreNotes) {
      restoreExecutionSnapshotNotes(noteCollection, noteTargets, cells);
    }

    await context.sync();
  });
}

async function validateLocalExecutionSnapshotForMode(snapshot, mode) {
  if (isCompositeLocalExecutionSnapshot(snapshot)) {
    for (const entry of getLocalExecutionSnapshotEntriesForMode(snapshot, mode)) {
      await validateLocalExecutionSnapshotForMode(entry, mode);
    }
    return;
  }

  if (isRangeFormatLocalExecutionSnapshot(snapshot)) {
    return validateRangeFormatLocalExecutionSnapshotForMode(snapshot, mode);
  }

  if (isDataValidationLocalExecutionSnapshot(snapshot)) {
    return validateDataValidationLocalExecutionSnapshotForMode(snapshot, mode);
  }

  if (isNamedRangeLocalExecutionSnapshot(snapshot)) {
    return validateNamedRangeLocalExecutionSnapshotForMode(snapshot, mode);
  }

  const cells = mode === "undo" ? snapshot?.beforeCells : snapshot?.afterCells;
  if (!snapshot?.targetSheet || !snapshot?.targetRange || !Array.isArray(cells) || cells.length === 0) {
    throw new Error("That history entry is no longer available in this spreadsheet session.");
  }

  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const sheet = worksheets.getItem(snapshot.targetSheet);
    const target = sheet.getRange(snapshot.targetRange);
    target.load(["rowCount", "columnCount"]);
    await context.sync();

    if (target.rowCount !== cells.length || target.columnCount !== (cells[0]?.length || 0)) {
      throw new Error("The saved undo snapshot no longer matches the current range shape.");
    }
  });
}

function sanitizeConversation(messages) {
  return messages
    .filter((message) =>
      (message.role === "user" || message.role === "assistant" || message.role === "system") &&
      message.content &&
      message.content !== "Thinking..."
    )
    .map((message) => ({
      role: message.role,
      content: truncateRequestText(message.content)
    }))
    .slice(-MAX_CONVERSATION_MESSAGES);
}

function pruneStoredMessages(messages) {
  return Array.isArray(messages)
    ? messages.slice(-MAX_STORED_MESSAGES)
    : [];
}

function appendStoredMessage(message) {
  state.messages = pruneStoredMessages([...(state.messages || []), message]);
  return message;
}

function trimMessageTraceEvents(trace) {
  return Array.isArray(trace)
    ? trace.slice(-MAX_MESSAGE_TRACE_EVENTS)
    : [];
}

function setMessageResponse(message, response) {
  if (!message) {
    return;
  }

  if (!response || !Array.isArray(response.trace)) {
    message.response = response;
    return;
  }

  const trimmedTrace = trimMessageTraceEvents(response.trace);
  message.response = {
    ...response,
    trace: trimmedTrace
  };
  message.trace = trimmedTrace;
}

function truncateRequestText(value) {
  const text = String(value || "");
  if (text.length <= MAX_REQUEST_MESSAGE_LENGTH) {
    return text;
  }

  return `${text.slice(0, MAX_REQUEST_MESSAGE_LENGTH - REQUEST_TRUNCATION_SUFFIX.length)}${REQUEST_TRUNCATION_SUFFIX}`;
}

function isSupportedImageFile(file) {
  return SUPPORTED_IMAGE_TYPES.has(String(file?.type || "").toLowerCase());
}

function filterSupportedImageFiles(files) {
  return Array.from(files || []).filter((file) => isSupportedImageFile(file));
}

function getTraceLabel(event) {
  const labels = {
    request_received: "Request received by Hermes",
    spreadsheet_context_received: "Spreadsheet context received",
    attachment_received: "Attachment received",
    image_received: "Image received",
    skill_selected: "Skill selected",
    tool_selected: "Tool selected",
    downstream_provider_called: "Downstream provider called",
    ocr_started: "OCR started",
    table_extraction_started: "Table extraction started",
    result_generated: "Result generated",
    workbook_structure_update_ready: "Workbook update plan ready",
    sheet_structure_update_ready: "Sheet structure plan ready",
    range_sort_plan_ready: "Range sort plan ready",
    range_filter_plan_ready: "Range filter plan ready",
    range_format_update_ready: "Formatting plan ready",
    conditional_format_plan_ready: "Conditional formatting plan ready",
    data_validation_plan_ready: "Validation plan ready",
    named_range_update_ready: "Named range update plan ready",
    range_transfer_plan_ready: "Range transfer plan ready",
    data_cleanup_plan_ready: "Data cleanup plan ready",
    analysis_report_plan_ready: "Analysis report plan ready",
    pivot_table_plan_ready: "Pivot table plan ready",
    chart_plan_ready: "Chart plan ready",
    table_plan_ready: "Table plan ready",
    external_data_plan_ready: "External data plan ready",
    analysis_report_update_ready: "Analysis report update ready",
    pivot_table_update_ready: "Pivot table update ready",
    chart_update_ready: "Chart update ready",
    table_update_ready: "Table update ready",
    sheet_update_plan_ready: "Sheet update plan ready",
    sheet_import_plan_ready: "Sheet import plan ready",
    completed: "Completed",
    failed: "Failed"
  };
  return event.label || labels[event.event] || "Waiting for Hermes";
}

function summarizeLatestTrace(trace) {
  if (!trace || trace.length === 0) {
    return "Waiting for Hermes";
  }

  return getTraceLabel(trace[trace.length - 1]);
}

function formatTraceTimeline(trace) {
  if (!trace || trace.length === 0) {
    return "";
  }

  return trace
    .map(getTraceLabel)
    .filter((label, index, labels) => index === 0 || label !== labels[index - 1])
    .join(" -> ");
}

function formatUserFacingErrorText(message, userAction) {
  const resolvedMessage = String(message || "").trim();
  const resolvedUserAction = typeof userAction === "string" ? userAction.trim() : "";

  if (!resolvedUserAction || resolvedUserAction === resolvedMessage) {
    return resolvedMessage;
  }

  return `${resolvedMessage}\n\n${resolvedUserAction}`;
}

function formatGatewayIssueSummary(issues) {
  if (!Array.isArray(issues) || issues.length === 0) {
    return "";
  }

  return issues
    .slice(0, 3)
    .map((issue) => {
      const path = typeof issue?.path === "string" && issue.path.trim().length > 0
        ? issue.path.trim()
        : "request";
      const detail = typeof issue?.message === "string" && issue.message.trim().length > 0
        ? issue.message.trim()
        : "invalid value";
      return `${path}: ${detail}`;
    })
    .join("\n");
}

function appendGatewayIssueSummary(message, issues) {
  const issueSummary = formatGatewayIssueSummary(issues);
  return issueSummary ? `${message}\n\nRequest details:\n${issueSummary}` : message;
}

function isExcelPreviewSupportCheckedWritePlanType(responseType) {
  return responseType === "pivot_table_plan" ||
    responseType === "chart_plan" ||
    responseType === "table_plan" ||
    responseType === "external_data_plan" ||
    responseType === "range_format_update" ||
    responseType === "conditional_format_plan" ||
    responseType === "range_filter_plan" ||
    responseType === "data_validation_plan" ||
    responseType === "range_transfer_plan" ||
      responseType === "data_cleanup_plan";
}

function inferExcelPreviewSupportKind(preview) {
  if (!preview || typeof preview !== "object") {
    return "";
  }

  if (typeof preview.kind === "string" && preview.kind.trim().length > 0) {
    return preview.kind;
  }

  if (typeof preview.chartType === "string" && Array.isArray(preview.series)) {
    return "chart_plan";
  }

  if (typeof preview.hasHeaders === "boolean" &&
    typeof preview.targetSheet === "string" &&
    typeof preview.targetRange === "string") {
    return "table_plan";
  }

  if (Array.isArray(preview.rowGroups) && Array.isArray(preview.valueAggregations)) {
    return "pivot_table_plan";
  }

  if (typeof preview.managementMode === "string") {
    return "conditional_format_plan";
  }

  if (typeof preview.ruleType === "string" && typeof preview.invalidDataBehavior === "string") {
    return "data_validation_plan";
  }

  if (typeof preview.targetSheet === "string" &&
    typeof preview.targetRange === "string" &&
    Array.isArray(preview.conditions)) {
    return "range_filter_plan";
  }

  if (typeof preview.sourceSheet === "string" &&
    typeof preview.sourceRange === "string" &&
    typeof preview.pasteMode === "string") {
    return "range_transfer_plan";
  }

  if (typeof preview.sourceType === "string" &&
    typeof preview.provider === "string" &&
    typeof preview.formula === "string" &&
    typeof preview.targetSheet === "string" &&
    typeof preview.targetRange === "string") {
    return "external_data_plan";
  }

  if (typeof preview.operation === "string" &&
    [
      "trim_whitespace",
      "remove_blank_rows",
      "remove_duplicate_rows",
      "normalize_case",
      "split_column",
      "join_columns",
      "fill_down",
      "standardize_format"
    ].includes(preview.operation)) {
    return "data_cleanup_plan";
  }

  return "";
}

function isSingleCellA1Anchor(range) {
  try {
    const bounds = parseA1RangeReference(range);
    return bounds.rowCount === 1 && bounds.columnCount === 1;
  } catch {
    return false;
  }
}

function mapPivotAggregationToExcel(aggregation) {
  switch (aggregation) {
    case "sum":
      return "Sum";
    case "count":
      return "Count";
    case "average":
      return "Average";
    case "min":
      return "Min";
    case "max":
      return "Max";
    default:
      throw new Error(`Unsupported pivot aggregation: ${aggregation}`);
  }
}

function mapPivotSortDirectionToExcel(direction) {
  if (direction === "desc") {
    return Excel?.SortBy?.descending || Excel?.SortBy?.Descending || "Descending";
  }

  return Excel?.SortBy?.ascending || Excel?.SortBy?.Ascending || "Ascending";
}

function mapPivotLabelFilterConditionToExcel(operator) {
  switch (operator) {
    case "equal_to":
    case "not_equal_to":
      return Excel?.LabelFilterCondition?.equals || Excel?.LabelFilterCondition?.Equals || "Equals";
    case "greater_than":
      return Excel?.LabelFilterCondition?.greaterThan || Excel?.LabelFilterCondition?.GreaterThan || "GreaterThan";
    case "greater_than_or_equal_to":
      return Excel?.LabelFilterCondition?.greaterThanOrEqualTo || Excel?.LabelFilterCondition?.GreaterThanOrEqualTo || "GreaterThanOrEqualTo";
    case "less_than":
      return Excel?.LabelFilterCondition?.lessThan || Excel?.LabelFilterCondition?.LessThan || "LessThan";
    case "less_than_or_equal_to":
      return Excel?.LabelFilterCondition?.lessThanOrEqualTo || Excel?.LabelFilterCondition?.LessThanOrEqualTo || "LessThanOrEqualTo";
    case "between":
    case "not_between":
      return Excel?.LabelFilterCondition?.between || Excel?.LabelFilterCondition?.Between || "Between";
    default:
      throw new Error(`Unsupported pivot filter operator: ${operator}`);
  }
}

function getExcelPivotStructureSupportError(plan) {
  const rowGroups = Array.isArray(plan?.rowGroups) ? plan.rowGroups : [];
  const columnGroups = Array.isArray(plan?.columnGroups) ? plan.columnGroups : [];
  const valueAggregations = Array.isArray(plan?.valueAggregations) ? plan.valueAggregations : [];
  const filters = Array.isArray(plan?.filters) ? plan.filters : [];
  const groupedFields = new Set([...rowGroups, ...columnGroups].map((field) => String(field || "").trim()).filter(Boolean));
  const aggregatedFields = new Set(valueAggregations.map((aggregation) => String(aggregation?.field || "").trim()).filter(Boolean));

  if (plan?.sort) {
    if (plan.sort.sortOn === "group_field") {
      if (!groupedFields.has(String(plan.sort.field || "").trim())) {
        return "This Excel runtime can only sort an existing row or column group field.";
      }
    } else if (plan.sort.sortOn === "aggregated_value") {
      if (!aggregatedFields.has(String(plan.sort.field || "").trim())) {
        return "This Excel runtime can only sort by an existing pivot value field.";
      }

      if (rowGroups.length > 0 && columnGroups.length > 0) {
        return "This Excel runtime can't sort pivot values when both row and column groups are present yet.";
      }
    }
  }

  for (const filter of filters) {
    if (!filter || typeof filter !== "object") {
      return "This Excel runtime can't apply that pivot filter.";
    }

    if (![
      "equal_to",
      "not_equal_to",
      "greater_than",
      "greater_than_or_equal_to",
      "less_than",
      "less_than_or_equal_to",
      "between",
      "not_between"
    ].includes(filter.operator)) {
      return "This Excel runtime can't apply that pivot filter.";
    }

    if ((filter.operator === "between" || filter.operator === "not_between") && filter.value2 === undefined) {
      return "This Excel runtime requires two bounds for between pivot filters.";
    }
  }

  return "";
}

function getSupportedDateTextPatternSpec(formatPattern) {
  if (typeof formatPattern !== "string") {
    return null;
  }

  const trimmed = formatPattern.trim();
  const match = trimmed.match(/^[Yy]{4}([\-/.])[Mm]{2}\1[Dd]{2}$/);
  if (!match) {
    return null;
  }

  return {
    formatType: "date_text",
    separator: match[1],
    formatPattern: trimmed
  };
}

function getSupportedNumberTextPatternSpec(formatPattern) {
  if (typeof formatPattern !== "string") {
    return null;
  }

  const trimmed = formatPattern.trim();
  const match = trimmed.match(/^(#,##0|0)(?:\.(0+))?$/);
  if (!match) {
    return null;
  }

  return {
    formatType: "number_text",
    useGrouping: match[1] === "#,##0",
    decimals: match[2] ? match[2].length : 0,
    formatPattern: trimmed
  };
}

function getSupportedStandardizeFormatSpec(formatType, formatPattern) {
  if (formatType === "date_text") {
    return getSupportedDateTextPatternSpec(formatPattern);
  }

  if (formatType === "number_text") {
    return getSupportedNumberTextPatternSpec(formatPattern);
  }

  return null;
}

function getStandardizeFormatSupportError(formatType, formatPattern, hostLabel) {
  const resolvedHostLabel = hostLabel || "This host";

  if (
    typeof formatType !== "string" ||
    !formatType.trim() ||
    typeof formatPattern !== "string" ||
    !formatPattern.trim()
  ) {
    return `${resolvedHostLabel} requires an exact formatType and formatPattern for standardize_format.`;
  }

  if (getSupportedStandardizeFormatSpec(formatType, formatPattern)) {
    return "";
  }

  if (formatType === "date_text") {
    return `${resolvedHostLabel} only supports exact year-first date text patterns like YYYY-MM-DD, YYYY/MM/DD, or YYYY.MM.DD.`;
  }

  if (formatType === "number_text") {
    return `${resolvedHostLabel} only supports simple fixed-decimal number text patterns like #,##0.00 or 0.00.`;
  }

  return `${resolvedHostLabel} can't standardize ${formatType} with pattern ${formatPattern} exactly.`;
}

function isDateObject(value) {
  return Object.prototype.toString.call(value) === "[object Date]" &&
    typeof value.getTime === "function";
}

function isValidDateParts(year, month, day) {
  const candidate = new Date(year, month - 1, day);
  return candidate.getFullYear() === year &&
    candidate.getMonth() === month - 1 &&
    candidate.getDate() === day;
}

function normalizeIntegerDigits(integerDigits) {
  const normalized = String(integerDigits || "").replace(/^0+(?=\d)/, "");
  return normalized.length > 0 ? normalized : "0";
}

function formatGroupedIntegerDigits(integerDigits) {
  return normalizeIntegerDigits(integerDigits).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function parseExactNumericParts(value, hostLabel) {
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw new Error(`${hostLabel} cannot standardize non-finite numbers exactly.`);
    }

    const serialized = String(value);
    if (/e/i.test(serialized)) {
      throw new Error(`${hostLabel} cannot standardize scientific-notation numbers exactly.`);
    }

    return parseExactNumericParts(serialized, hostLabel);
  }

  if (typeof value !== "string" || value !== value.trim()) {
    throw new Error(`${hostLabel} cannot standardize non-text numeric values exactly.`);
  }

  const match = value.match(/^([+-]?)(?:(\d{1,3}(?:,\d{3})+)|(\d+))(?:\.(\d+))?$/);
  if (!match) {
    throw new Error(`${hostLabel} cannot standardize numeric text exactly for value ${JSON.stringify(value)}.`);
  }

  return {
    sign: match[1] === "-" ? "-" : "",
    integerDigits: normalizeIntegerDigits((match[2] || match[3] || "").replace(/,/g, "")),
    fractionDigits: match[4] || ""
  };
}

function standardizeDateTextValueExact(value, spec, hostLabel) {
  if (isBlankCellValue(value)) {
    return value;
  }

  let year;
  let month;
  let day;

  if (isDateObject(value)) {
    if (Number.isNaN(value.getTime())) {
      throw new Error(`${hostLabel} cannot standardize invalid dates exactly.`);
    }

    if (
      value.getHours() !== 0 ||
      value.getMinutes() !== 0 ||
      value.getSeconds() !== 0 ||
      value.getMilliseconds() !== 0
    ) {
      throw new Error(`${hostLabel} cannot rewrite date-time values as date_text without losing precision.`);
    }

    year = value.getFullYear();
    month = value.getMonth() + 1;
    day = value.getDate();
  } else if (typeof value === "string" && value === value.trim()) {
    const match = value.match(/^(\d{4})([-/.])(\d{1,2})\2(\d{1,2})$/);
    if (!match) {
      throw new Error(`${hostLabel} cannot standardize date text exactly for value ${JSON.stringify(value)}.`);
    }

    year = Number(match[1]);
    month = Number(match[3]);
    day = Number(match[4]);
  } else {
    throw new Error(`${hostLabel} cannot standardize non-date values as date_text exactly.`);
  }

  if (!isValidDateParts(year, month, day)) {
    throw new Error(`${hostLabel} cannot standardize invalid calendar dates exactly.`);
  }

  return `${String(year).padStart(4, "0")}${spec.separator}${String(month).padStart(2, "0")}${spec.separator}${String(day).padStart(2, "0")}`;
}

function standardizeNumberTextValueExact(value, spec, hostLabel) {
  if (isBlankCellValue(value)) {
    return value;
  }

  const parsed = parseExactNumericParts(value, hostLabel);
  const discardedFraction = parsed.fractionDigits.slice(spec.decimals);
  if (discardedFraction.replace(/0/g, "").length > 0) {
    throw new Error(`${hostLabel} cannot standardize numeric text exactly without rounding.`);
  }

  const integerDigits = spec.useGrouping
    ? formatGroupedIntegerDigits(parsed.integerDigits)
    : normalizeIntegerDigits(parsed.integerDigits);
  const fractionDigits = parsed.fractionDigits.slice(0, spec.decimals).padEnd(spec.decimals, "0");
  return parsed.sign + integerDigits + (spec.decimals > 0 ? `.${fractionDigits}` : "");
}

function standardizeFormatMatrixExact(plan, values, hostLabel) {
  const spec = getSupportedStandardizeFormatSpec(plan?.formatType, plan?.formatPattern);
  if (!spec) {
    throw new Error(
      getStandardizeFormatSupportError(plan?.formatType, plan?.formatPattern, hostLabel)
    );
  }

  return values.map((row) => row.map((value) =>
    spec.formatType === "date_text"
      ? standardizeDateTextValueExact(value, spec, hostLabel)
      : standardizeNumberTextValueExact(value, spec, hostLabel)
  ));
}

function normalizeFilterPreviewColumnRef(columnRef) {
  if (Number.isInteger(columnRef)) {
    return `#${columnRef}`;
  }

  if (typeof columnRef !== "string") {
    return "";
  }

  const trimmed = columnRef.trim();
  if (!trimmed) {
    return "";
  }

  if (/^\d+$/.test(trimmed)) {
    return `#${Number(trimmed)}`;
  }

  return `s:${trimmed.toLocaleLowerCase()}`;
}

function isExcelCustomFilterOperator(operator) {
  return [
    "equals",
    "notEquals",
    "contains",
    "startsWith",
    "endsWith",
    "greaterThan",
    "greaterThanOrEqual",
    "lessThan",
    "lessThanOrEqual",
    "isEmpty",
    "isNotEmpty"
  ].includes(operator);
}

function hasUnsupportedRepeatedExcelFilterPreviewColumns(conditions) {
  const groupedConditions = new Map();

  for (const condition of conditions || []) {
    const normalizedColumnRef = normalizeFilterPreviewColumnRef(condition?.columnRef);
    if (!normalizedColumnRef) {
      continue;
    }

    if (!groupedConditions.has(normalizedColumnRef)) {
      groupedConditions.set(normalizedColumnRef, []);
    }

    groupedConditions.get(normalizedColumnRef).push(condition);
  }

  for (const group of groupedConditions.values()) {
    if (group.length <= 1) {
      continue;
    }

    if (group.length !== 2) {
      return true;
    }

    if (!group.every((condition) => isExcelCustomFilterOperator(condition?.operator))) {
      return true;
    }
  }

  return false;
}

function getExcelChartTypeConfig(chartType) {
  switch (chartType) {
    case "bar":
      return { chartType: Excel?.ChartType?.barClustered || "BarClustered" };
    case "column":
      return { chartType: Excel?.ChartType?.columnClustered || "ColumnClustered" };
    case "stacked_bar":
      return { chartType: Excel?.ChartType?.barStacked || "BarStacked" };
    case "stacked_column":
      return { chartType: Excel?.ChartType?.columnStacked || "ColumnStacked" };
    case "line":
      return { chartType: Excel?.ChartType?.line || "Line" };
    case "area":
      return { chartType: Excel?.ChartType?.area || "Area" };
    case "pie":
      return { chartType: Excel?.ChartType?.pie || "Pie" };
    case "scatter":
      return { chartType: Excel?.ChartType?.xyScatter || "XYScatter" };
    default:
      throw new Error(`Excel host does not support chart type: ${chartType}.`);
  }
}

function getExcelChartLegendConfig(legendPosition) {
  if (legendPosition === undefined || legendPosition === null || legendPosition === "") {
    return null;
  }

  switch (String(legendPosition).trim()) {
    case "hidden":
      return { visible: false };
    case "top":
      return {
        visible: true,
        position: Excel?.ChartLegendPosition?.top || "Top"
      };
    case "bottom":
      return {
        visible: true,
        position: Excel?.ChartLegendPosition?.bottom || "Bottom"
      };
    case "left":
      return {
        visible: true,
        position: Excel?.ChartLegendPosition?.left || "Left"
      };
    case "right":
      return {
        visible: true,
        position: Excel?.ChartLegendPosition?.right || "Right"
      };
    default:
      throw new Error(`Excel host does not support exact-safe chart legend positioning for ${legendPosition}.`);
  }
}

function normalizeExcelChartField(value) {
  return normalizeExcelHeaderText(String(value || "")) || "";
}

function getExcelChartFieldSequence(plan) {
  const categoryField = normalizeExcelChartField(plan?.categoryField);
  if (!categoryField) {
    throw new Error("Excel host requires categoryField for exact-safe chart creation.");
  }

  if (!Array.isArray(plan?.series) || plan.series.length === 0) {
    throw new Error("Excel host requires at least one series for exact-safe chart creation.");
  }

  const fields = [categoryField];
  const seenFields = new Set(fields);

  for (const series of plan.series) {
    const field = normalizeExcelChartField(series?.field);
    if (!field) {
      throw new Error("Excel host requires exact-safe chart series fields.");
    }

    if (seenFields.has(field)) {
      throw new Error("Excel host requires chart fields to be unique.");
    }

    seenFields.add(field);
    fields.push(field);
  }

  return fields;
}

function validateExcelChartSeriesLabels(plan) {
  if (!Array.isArray(plan?.series)) {
    return;
  }

  for (const series of plan.series) {
    const label = typeof series?.label === "string" ? series.label.trim() : "";
    if (label.length > 255) {
      throw new Error("Excel host requires chart series labels to be 255 characters or fewer.");
    }
  }
}

function hasExcelChartAxisTitles(plan) {
  return Boolean(
    (typeof plan?.horizontalAxisTitle === "string" && plan.horizontalAxisTitle.trim().length > 0) ||
    (typeof plan?.verticalAxisTitle === "string" && plan.verticalAxisTitle.trim().length > 0)
  );
}

function assertExcelChartPlanSupport(plan) {
  getExcelChartTypeConfig(plan?.chartType);
  getExcelChartLegendConfig(plan?.legendPosition);
  const fieldSequence = getExcelChartFieldSequence(plan);
  validateExcelChartSeriesLabels(plan);

  if (plan?.chartType === "pie" && fieldSequence.length !== 2) {
    throw new Error("Excel host only supports a single series when creating pie charts.");
  }

  if (plan?.chartType === "pie" && hasExcelChartAxisTitles(plan)) {
    throw new Error("Excel host does not support axis titles for pie charts.");
  }
}

function getExcelChartSupportError(preview) {
  if (!isSingleCellA1Anchor(preview?.targetRange)) {
    return "This Excel runtime requires a single-cell target anchor for charts.";
  }

  try {
    assertExcelChartPlanSupport(preview);
    return "";
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error || "");

    if (/chart fields to be unique/i.test(message) ||
      /requires categoryField/i.test(message) ||
      /requires at least one series/i.test(message) ||
      /requires exact-safe chart series fields/i.test(message)) {
      return "This Excel runtime requires one category field and at least one unique series field for exact-safe chart creation.";
    }

    if (/legend positioning/i.test(message)) {
      return "This Excel runtime can't place the chart legend exactly there. Use top, bottom, left, right, or hidden.";
    }

    if (/series labels to be 255 characters or fewer/i.test(message)) {
      return "This Excel runtime requires chart series labels to be 255 characters or fewer.";
    }

    if (/single series when creating pie charts/i.test(message)) {
      return "This Excel runtime only supports a single series when creating pie charts.";
    }

    if (/axis titles for pie charts/i.test(message)) {
      return "Pie charts do not support axis titles. Remove the horizontal and vertical axis titles or choose a chart type with axes.";
    }

    if (/chart type/i.test(message)) {
      return "This Excel runtime can't create that chart type exactly yet. Try line, column, bar, area, pie, or scatter.";
    }

    return "This Excel runtime can't create that chart safely yet. Keep it as a preview or ask for a simpler supported output.";
  }
}

function getExcelPlanSupportError(preview) {
  if (!preview || typeof preview !== "object") {
    return "";
  }

  const kind = inferExcelPreviewSupportKind(preview);

  if (kind === "pivot_table_plan") {
    if (!isSingleCellA1Anchor(preview.targetRange)) {
      return "This Excel runtime requires a single-cell target anchor for pivot tables.";
    }

    const pivotSupportError = getExcelPivotStructureSupportError(preview);
    if (pivotSupportError) {
      return pivotSupportError;
    }

    try {
      (preview.valueAggregations || []).forEach((aggregation) => {
        mapPivotAggregationToExcel(aggregation?.aggregation);
      });
    } catch (error) {
      return error instanceof Error
        ? error.message
        : "This Excel runtime can't apply that pivot aggregation.";
    }

    return "";
  }

  if (kind === "chart_plan") {
    return getExcelChartSupportError(preview);
  }

  if (kind === "table_plan") {
    if (typeof preview.name === "string" &&
      preview.name.trim().length > 0 &&
      !/^[A-Za-z_][A-Za-z0-9_]*$/.test(preview.name.trim())) {
      return "Excel table names must start with a letter or underscore and contain only letters, numbers, and underscores.";
    }

    return "";
  }

  if (kind === "external_data_plan") {
    return "This Excel runtime can't create first-class external data imports yet. Use Google Sheets for GOOGLEFINANCE or web-table imports.";
  }

  if (kind === "range_format_update") {
    return getExcelRangeFormatSupportError(preview.format, "This Excel runtime");
  }

  if (kind === "conditional_format_plan") {
    const unsupportedStyleFields = getUnsupportedConditionalFormatStyleFields(preview.style);
    if (unsupportedStyleFields.length > 0) {
      return `This Excel runtime can't apply that conditional formatting style exactly. Remove ${unsupportedStyleFields.join(", ")} and try again.`;
    }

    const supportedRuleTypes = new Set([
      "single_color",
      "number_compare",
      "date_compare",
      "text_contains",
      "duplicate_values",
      "custom_formula",
      "top_n",
      "average_compare",
      "color_scale"
    ]);
    if (!supportedRuleTypes.has(preview.ruleType)) {
      return "This Excel runtime can't apply that conditional formatting rule exactly. Try text contains, value comparison, duplicate values, custom formula, top/bottom, above/below average, or color scale.";
    }

    if (preview.ruleType === "color_scale") {
      try {
        buildExcelConditionalColorScaleCriteria(preview);
      } catch (error) {
        return error instanceof Error
          ? error.message
          : "This Excel runtime can't apply that color scale exactly.";
      }
    }

    if (
      (preview.ruleType === "single_color" ||
        preview.ruleType === "number_compare" ||
        preview.ruleType === "date_compare") &&
      ![
        "between",
        "not_between",
        "equal_to",
        "not_equal_to",
        "greater_than",
        "greater_than_or_equal_to",
        "less_than",
        "less_than_or_equal_to"
      ].includes(preview.comparator)
    ) {
      return "This Excel runtime can't apply that conditional formatting comparison exactly. Use a standard comparison such as between, equal to, or greater than.";
    }

    return "";
  }

  if (kind === "range_filter_plan") {
    if (preview.combiner !== "and") {
      return "This Excel runtime can't combine those filter conditions exactly. Use a single AND filter step instead.";
    }

    if (hasUnsupportedRepeatedExcelFilterPreviewColumns(preview.conditions)) {
      return "This Excel runtime can't apply multiple conditions to the same filter column in one exact step.";
    }

    return "";
  }

  if (kind === "data_validation_plan") {
    const supportedRuleTypes = new Set([
      "whole_number",
      "decimal",
      "date",
      "text_length",
      "custom_formula",
      "list",
      "checkbox"
    ]);
    if (!supportedRuleTypes.has(preview.ruleType)) {
      return "This Excel runtime can't apply that validation rule. Try list, whole number, decimal, date, text length, checkbox, or custom formula instead.";
    }

    if (preview.ruleType === "checkbox" && hasUnsupportedExcelCheckboxValues(preview)) {
      return "This Excel runtime only supports checkbox values as true and false.";
    }

    if (preview.ruleType === "checkbox" && preview.allowBlank === true) {
      return "This Excel runtime can't preserve allowBlank=true checkbox semantics.";
    }

    if (preview.ruleType === "checkbox" && hasCustomValidationPromptOrAlert(preview)) {
      return "This Excel runtime can't apply custom validation prompt or error text to checkbox controls.";
    }

    return "";
  }

  if (kind === "range_transfer_plan") {
    if (!["values", "formulas", "formats"].includes(preview.pasteMode)) {
      return "This Excel runtime can't apply that transfer mode. Use values, formulas, or formats instead.";
    }

    return "";
  }

  if (kind === "data_cleanup_plan") {
    if (preview.operation === "normalize_case" &&
      preview.mode !== "upper" &&
      preview.mode !== "lower" &&
      preview.mode !== "title" &&
      preview.mode !== "sentence") {
      return `This Excel runtime only supports upper, lower, title, and sentence case normalization in cleanup plans, not ${preview.mode}.`;
    }

    if (preview.operation === "standardize_format") {
      return getStandardizeFormatSupportError(
        preview.formatType,
        preview.formatPattern,
        "This Excel runtime"
      );
    }

    return "";
  }

  return "";
}

function getCompositePreviewSupportErrors(steps) {
  return (steps || [])
    .map((step) => ({
      stepId: step?.stepId,
      supportError: step?.supportError || (step?.plan ? getExcelPlanSupportError(step.plan) : "")
    }))
    .filter((step) => typeof step.supportError === "string" && step.supportError.trim().length > 0);
}

function containsSensitiveHostDiagnostics(message) {
  const text = String(message || "");
  return (
    /(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|approval_secret|HERMES_[A-Z0-9_]+)/i.test(text) ||
    /\bBearer\s+[A-Za-z0-9._~+/-]+=*/i.test(text) ||
    /\bat\s+(?:file:\/\/\/|\/|[A-Za-z]:\\)/i.test(text) ||
    /(?:^|\s)\/(?:srv|var|tmp|root|home|Users|opt|workspace|app|mnt)\/[^\s]+(?::\d+)?/i.test(text) ||
    /(?:^|\s)[A-Za-z]:\\[^\s]+/.test(text) ||
    /https?:\/\/(?:localhost|127(?:\.\d{1,3}){3}|0\.0\.0\.0|10\.|192\.168\.|172\.(?:1[6-9]|2\d|3[01])\.|[^/\s]*internal|[^/\s]*\.local)(?:[/:]|\s|$)/i.test(text) ||
    /\b(?:stack trace|traceback)\b/i.test(text)
  );
}

function sanitizeHostExecutionError(error, fallbackMessage = "Write-back failed.") {
  const rawMessage = error instanceof Error ? error.message : String(error || "");
  const message = String(rawMessage || "").trim().replace(/^Error:\s*/i, "");

  if (!message) {
    return fallbackMessage;
  }

  if (containsSensitiveHostDiagnostics(message)) {
    return fallbackMessage;
  }

  if (/Hermes gateway URL is not configured/i.test(message)) {
    return formatUserFacingErrorText(
      "The Hermes connection is not configured for this workbook.",
      "Set the Hermes gateway URL, reload the add-in, and retry."
    );
  }

  if (
    /Hermes gateway returned invalid JSON/i.test(message) ||
    /structured gateway contract/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "The Hermes service returned a response the add-in could not use.",
      "Retry the request. If it keeps happening, reload the add-in or check the Hermes gateway."
    );
  }

  if (/Hermes gateway proxy requires a request path/i.test(message)) {
    return formatUserFacingErrorText(
      "The Hermes request could not be sent correctly.",
      "Retry the action. If it keeps happening, reload the add-in and try again."
    );
  }

  if (/^The requested resource doesn't exist\.?$/i.test(message)) {
    return formatUserFacingErrorText(
      "Hermes could not read the current workbook or selection.",
      "Select a normal worksheet cell, reload the add-in, and retry. If it keeps happening, reopen the workbook in Excel and try again."
    );
  }

  if (/Failed to fetch/i.test(message) || /NetworkError/i.test(message)) {
    return formatUserFacingErrorText(
      "The Hermes service could not be reached.",
      "Check the network connection or Hermes gateway, then retry."
    );
  }

  if (/Destructive confirmation is unavailable in this host/i.test(message)) {
    return formatUserFacingErrorText(
      "This spreadsheet app cannot approve destructive changes inline.",
      "Ask Hermes for a safer alternative, or use a non-destructive step first."
    );
  }

  if (/exact (?:undo|redo) snapshot/i.test(message) || /saved undo snapshot/i.test(message)) {
    return formatUserFacingErrorText(
      "That history entry is no longer available for exact undo or redo in this workbook session.",
      "Re-run the change, then use undo or redo again from the current session."
    );
  }

  const targetSheetMatch = message.match(/^Target sheet not found:\s*(.+)$/i);
  if (targetSheetMatch) {
    return formatUserFacingErrorText(
      `Sheet "${targetSheetMatch[1].trim()}" was not found.`,
      "Create or select that sheet, then retry."
    );
  }

  const sourceSheetMatch = message.match(/^(?:Validation )?Source sheet not found:\s*(.+)$/i);
  if (sourceSheetMatch) {
    return formatUserFacingErrorText(
      `Sheet "${sourceSheetMatch[1].trim()}" was not found.`,
      "Select a valid source sheet, then retry."
    );
  }

  const namedRangeMatch = message.match(/^Named range not found:\s*(.+)$/i);
  if (namedRangeMatch) {
    return formatUserFacingErrorText(
      `Named range "${namedRangeMatch[1].trim()}" was not found.`,
      "Check the range name or create it first, then retry."
    );
  }

  const invalidRangeMatch = message.match(/^Unsupported A1 reference:\s*(.+)$/i);
  if (invalidRangeMatch) {
    return formatUserFacingErrorText(
      `Range "${invalidRangeMatch[1].trim()}" is not a valid A1 reference.`,
      "Use a valid cell or range address, then retry."
    );
  }

  const duplicateHeaderMatch = message.match(/duplicate header:\s*(.+?)\.?$/i);
  if (duplicateHeaderMatch) {
    return formatUserFacingErrorText(
      `Column "${duplicateHeaderMatch[1].trim()}" appears more than once in the header row.`,
      "Rename duplicate columns or select a table with unique headers, then retry."
    );
  }

  if (/requires a header row/i.test(message)) {
    return formatUserFacingErrorText(
      "This action needs a table with a header row.",
      "Select or create a table with column headers, then retry."
    );
  }

  const missingHeaderFieldMatch = message.match(/cannot find (?:pivot|chart) field in header row:\s*(.+?)\.?$/i);
  if (missingHeaderFieldMatch) {
    return formatUserFacingErrorText(
      `Column "${missingHeaderFieldMatch[1].trim()}" was not found in the header row.`,
      "Select the full table with headers, or use the exact column name in the request and retry."
    );
  }

  if (
    /could not resolve any valid sort keys/i.test(message) ||
    /could not resolve a filter column inside the target range/i.test(message) ||
    /Column .* is outside /i.test(message)
  ) {
    return formatUserFacingErrorText(
      "The selected range does not include the columns this step needs.",
      "Select the full table, or update the request to use columns inside the chosen range."
    );
  }

  const invalidDateMatch = message.match(/^Invalid date literal:\s*(.+)$/i);
  if (invalidDateMatch) {
    return formatUserFacingErrorText(
      `The date "${invalidDateMatch[1].trim()}" is not valid.`,
      "Use a real calendar date such as 2026-04-22, then retry."
    );
  }

  if (
    /Unsupported filter operator/i.test(message) ||
    /grid filters cannot represent operator "topN" exactly/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This filter condition is not supported here.",
      "Try a simpler operator such as equals or contains, or ask Hermes for a different filter."
    );
  }

  if (
    /Unsupported filter combiner/i.test(message) ||
    /filter combiners other than and/i.test(message) ||
    /multiple conditions for the same column/i.test(message) ||
    /cannot represent combiner "or" exactly/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This spreadsheet app cannot combine those filter conditions in one exact step.",
      "Use a single filter rule per column, or split the filter into smaller steps."
    );
  }

  if (
    /named ranges? on this scope/i.test(message) ||
    /sheet-scoped named ranges/i.test(message) ||
    /does not support creating named ranges/i.test(message) ||
    /does not support renaming named ranges/i.test(message) ||
    /does not support deleting named ranges/i.test(message) ||
    /does not support retargeting named ranges/i.test(message) ||
    /Unsupported named range update/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This named range action is not supported in this spreadsheet app.",
      "Use a workbook-level named range or ask Hermes for a simpler named range update."
    );
  }

  if (
    /Named range create and retarget require targetSheet and targetRange/i.test(message) ||
    /Named range rename requires newName/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This named range request is missing required details.",
      "Include the destination sheet and range, or provide the new name, then retry."
    );
  }

  if (/approved targetRange does not match/i.test(message)) {
    return formatUserFacingErrorText(
      "The spreadsheet changed, so the approved destination no longer matches the intended shape.",
      "Refresh the spreadsheet state and run the request again."
    );
  }

  if (
    /cannot append exactly when the approved target range contains internal gaps/i.test(message) ||
    /cannot append exactly within the approved target range/i.test(message) ||
    /requires append targetRange to match the full destination rectangle/i.test(message) ||
    /cannot split this column exactly within the approved target range/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "The chosen destination range cannot accept this write safely.",
      "Choose a clean target range or ask Hermes to write into a blank area."
    );
  }

  if (
    /cannot apply an overlapping .* transfer exactly/i.test(message) ||
    /cannot clear the source range for this move/i.test(message) ||
    /Unsupported transfer pasteMode/i.test(message) ||
    /does not support exact-safe transfer pasteMode/i.test(message) ||
    /cannot append when the approved target range width does not match/i.test(message) ||
    /cannot expand the approved append anchor exactly/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This transfer cannot be applied safely on the current source and destination ranges.",
      "Choose a simpler target range or ask Hermes for a different transfer plan."
    );
  }

  if (/does not support exact-safe pivot table creation yet/i.test(message)) {
    return formatUserFacingErrorText(
      "This spreadsheet app cannot create that pivot table safely yet.",
      "Ask for a preview only, or target a simpler transformation first."
    );
  }

  if (/does not support exact-safe chart/i.test(message)) {
    return formatUserFacingErrorText(
      "This spreadsheet app cannot create that chart safely yet.",
      "Ask for a preview only, or request a simpler supported chart."
    );
  }

  if (
    /does not support exact-safe formula transfers on this range/i.test(message) ||
    /does not support exact format transfers on this range/i.test(message) ||
    /does not support exact format append transfers on this range/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This spreadsheet app cannot apply that transfer safely on the current range.",
      "Try a simpler target range or ask Hermes for a direct cell update instead."
    );
  }

  if (
    /does not support exact-safe cleanup semantics/i.test(message) ||
    /cannot apply cleanup plans exactly when the target range contains formulas/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This cleanup action cannot be applied safely on the current range.",
      "Try a narrower range or ask Hermes for a simpler cleanup step."
    );
  }

  if (/clip overflowing text exactly/i.test(message)) {
    return formatUserFacingErrorText(
      "This formatting setup cannot be represented safely in Excel.",
      "Use wrap or overflow text behavior instead, then retry."
    );
  }

  if (
    /Unsupported Excel data validation rule type/i.test(message) ||
    /Unsupported .*validation comparator/i.test(message) ||
    /Unsupported invalidDataBehavior/i.test(message) ||
    /List validation requires values, sourceRange, or namedRangeName/i.test(message) ||
    /Custom formula validation requires/i.test(message) ||
    /checkbox .* only support boolean true\/false/i.test(message) ||
    /cannot represent allowBlank/i.test(message) ||
    /uncheckedValue without checkedValue/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This validation setup cannot be represented safely here.",
      "Try a simpler dropdown, checkbox, or date rule, then retry."
    );
  }

  if (
    /requires a valid target range for /i.test(message) ||
    /requires a single-cell target anchor for /i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This action needs a valid destination cell or anchor.",
      "Choose a single target cell or a valid destination range, then retry."
    );
  }

  if (
    /conditional-format/i.test(message) ||
    /conditional formatting/i.test(message) ||
    /text_contains conditional formatting/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This conditional formatting step is not supported here.",
      "Try a simpler highlight rule, or ask Hermes for a preview-only result first."
    );
  }

  if (
    /does not support data validation on this range/i.test(message) ||
    /does not support checkbox cell controls on this range/i.test(message) ||
    /does not expose checkbox cell control support/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This validation action cannot run on the current range.",
      "Choose a standard editable cell range, then retry."
    );
  }

  if (
    /does not support range sort on this selection/i.test(message) ||
    /does not support range filters on this selection/i.test(message) ||
    /does not support conditional formatting on this range/i.test(message) ||
    /does not support conditional formatting on this sheet/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This action cannot run on the current selection.",
      "Choose a standard table or cell range, then retry."
    );
  }

  if (/Cannot hide the only visible worksheet/i.test(message)) {
    return formatUserFacingErrorText(
      "At least one worksheet must stay visible.",
      "Keep another sheet visible or unhide one first, then retry."
    );
  }

  if (
    /Unsupported workbook structure update/i.test(message) ||
    /Unsupported sheet structure update/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This sheet change is not supported in this spreadsheet app.",
      "Ask Hermes for a simpler sheet change, or try a different supported operation."
    );
  }

  if (
    /pivot/i.test(message) &&
    (
      /can't apply pivot/i.test(message) ||
      /does not support/i.test(message) ||
      /requires /i.test(message) ||
      /Unsupported pivot aggregation/i.test(message) ||
      /only supports equal_to pivot filters/i.test(message) ||
      /pivot filter criteria builders/i.test(message) ||
      /does not expose pivot creation/i.test(message)
    )
  ) {
    return formatUserFacingErrorText(
      "This pivot configuration is not supported here yet.",
      "Try a simpler pivot, or ask Hermes for a preview-only result first."
    );
  }

  if (
    /chart/i.test(message) &&
    (
      /does not support/i.test(message) ||
      /requires /i.test(message) ||
      /chart type/i.test(message) ||
      /legend positioning/i.test(message) ||
      /series fields/i.test(message) ||
      /series labels/i.test(message)
    )
  ) {
    return formatUserFacingErrorText(
      "This chart configuration is not supported here yet.",
      "Try a simpler chart, or ask Hermes for a preview-only result first."
    );
  }

  if (/Target range already contains content/i.test(message)) {
    return formatUserFacingErrorText(
      "The destination already contains data.",
      "Clear that range or choose a blank destination, then retry."
    );
  }

  if (/chat-only analysis reports are not writeback eligible/i.test(message)) {
    return formatUserFacingErrorText(
      "This result is analysis only and cannot be applied directly.",
      "Ask Hermes to turn it into a specific writeback on a sheet or range."
    );
  }

  if (
    /Composite workflow execution requires executionId/i.test(message) ||
    /Dependency .* has not completed before this step/i.test(message)
  ) {
    return formatUserFacingErrorText(
      "This workflow is no longer valid for the current spreadsheet state.",
      "Run the request again so Hermes can rebuild the workflow before applying it."
    );
  }

  return message;
}

function isTraceUnavailablePollError(error) {
  const rawMessage = error instanceof Error ? error.message : String(error || "");
  return /Hermes trace is no longer available/i.test(rawMessage) ||
    (/trace/i.test(rawMessage) && /fresh trace/i.test(rawMessage));
}

function isTraceBandwidthPollError(error) {
  const rawMessage = error instanceof Error ? error.message : String(error || "");
  return /bandwidth quota exceeded/i.test(rawMessage) ||
    /reducing the rate of data transfer/i.test(rawMessage) ||
    /too much data/i.test(rawMessage);
}

function shouldPollTraceForMessage(message) {
  if (!message || message.tracePollingDisabled) {
    return false;
  }

  const attempt = Number(message.pollAttemptCount || 0);
  return attempt <= 1 || attempt % TRACE_POLL_EVERY_N_ATTEMPTS === 0;
}

function getNextMessagePollDelay(message) {
  const currentDelay = Number(message && message.pollDelayMs) > 0
    ? Number(message.pollDelayMs)
    : MESSAGE_POLL_INTERVAL_MS;
  return Math.min(Math.round(currentDelay * 1.5), MESSAGE_POLL_MAX_INTERVAL_MS);
}

function getResponseBodyText(response) {
  const resolvedAnalysisPlan = response.type === "analysis_report_plan" &&
    response.data.outputMode === "materialize_report"
    ? resolveMaterializedAnalysisReportPlan(response.data)
    : null;

  switch (response.type) {
    case "chat":
      return response.data.message;
    case "formula":
      return response.data.explanation;
    case "workbook_structure_update":
      switch (response.data.operation) {
        case "create_sheet":
          return `Prepared a workbook update to create sheet ${response.data.sheetName}.`;
        case "delete_sheet":
          return `Prepared a workbook update to delete sheet ${response.data.sheetName}.`;
        case "rename_sheet":
          return `Prepared a workbook update to rename ${response.data.sheetName} to ${response.data.newSheetName}.`;
        case "duplicate_sheet":
          return `Prepared a workbook update to duplicate sheet ${response.data.sheetName}.`;
        case "move_sheet":
          return `Prepared a workbook update to move sheet ${response.data.sheetName}.`;
        case "hide_sheet":
          return `Prepared a workbook update to hide sheet ${response.data.sheetName}.`;
        case "unhide_sheet":
          return `Prepared a workbook update to unhide sheet ${response.data.sheetName}.`;
        default:
          return "Prepared a workbook update.";
      }
    case "sheet_structure_update":
      return `Prepared a sheet structure update for ${response.data.targetSheet}.`;
    case "range_sort_plan":
      return `Prepared a sort plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "range_filter_plan":
      return `Prepared a filter plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "range_format_update":
      return `Prepared a formatting update for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "conditional_format_plan":
      return `Prepared a conditional formatting plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "data_validation_plan":
      return `Prepared a validation plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "named_range_update":
      return `Prepared a named range update for ${response.data.name}.`;
    case "range_transfer_plan":
      return `Prepared a transfer plan from ${response.data.sourceSheet}!${response.data.sourceRange} to ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "data_cleanup_plan":
      return `Prepared a cleanup plan for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "composite_plan":
      return `Prepared a workflow preview with ${response.data.steps.length} step${response.data.steps.length === 1 ? "" : "s"}.`;
    case "analysis_report_plan":
      return response.data.outputMode === "materialize_report"
        ? `Prepared an analysis report preview for ${resolvedAnalysisPlan.targetSheet}!${resolvedAnalysisPlan.targetRange}.`
        : `Prepared a chat-only analysis report for ${response.data.sourceSheet}!${response.data.sourceRange}.`;
    case "pivot_table_plan":
      return `Prepared a pivot table preview for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "chart_plan":
      return `Prepared a chart preview for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "table_plan":
      return `Prepared a table preview for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "external_data_plan":
      return `Prepared an external data preview for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "sheet_update":
      return response.data.explanation;
    case "sheet_import_plan":
      return `Prepared an import preview for ${response.data.targetSheet}!${response.data.targetRange}.`;
    case "attachment_analysis":
      return response.data.summary;
    case "extracted_table":
      return "Prepared an extracted table preview from the uploaded image.";
    case "document_summary":
      return response.data.summary;
    case "error":
      return formatUserFacingErrorText(response.data.message, response.data.userAction);
    default:
      return "Processed by Hermes.";
  }
}

function applyWritebackResultToMessage(message, result) {
  message.content = getWritebackStatusLine(result);
  message.response = null;
  message.statusLine = "";
  delete message.pendingCompletion;
}

function buildPendingWritebackCompletionRequest(message, approval, result, workbookSessionKey) {
  return {
    requestId: message.requestId,
    runId: message.runId,
    ...(workbookSessionKey ? { workbookSessionKey } : {}),
    approvalToken: approval.approvalToken,
    planDigest: approval.planDigest,
    result
  };
}

function getPendingCompletionRetryStatus() {
  return "Applied locally. Retry confirm to finish syncing Hermes history.";
}

function getResponseWarnings(response) {
  const warnings = [...(response.warnings || [])];
  if (response.data && Array.isArray(response.data.warnings)) {
    warnings.push(...response.data.warnings);
  }
  return warnings;
}

function getResponseConfidence(response) {
  if (typeof response.data?.confidence === "number") {
    return response.data.confidence;
  }
  return undefined;
}

function getRequiresConfirmation(response) {
  if (response.type === "formula") {
    return Boolean(response.data.requiresConfirmation);
  }

  if (
    response.type === "workbook_structure_update" ||
    response.type === "sheet_structure_update" ||
    response.type === "range_sort_plan" ||
    response.type === "range_filter_plan" ||
    response.type === "range_format_update" ||
    response.type === "conditional_format_plan" ||
    response.type === "data_validation_plan" ||
    response.type === "named_range_update" ||
    response.type === "range_transfer_plan" ||
    response.type === "data_cleanup_plan" ||
    response.type === "external_data_plan" ||
    response.type === "table_plan" ||
    response.type === "composite_plan" ||
    response.type === "sheet_update" ||
    response.type === "sheet_import_plan"
  ) {
    return Boolean(response.data.requiresConfirmation);
  }

  if (response.type === "analysis_report_plan") {
    return response.data.outputMode === "materialize_report";
  }

  if (response.type === "pivot_table_plan") {
    return Boolean(response.data.requiresConfirmation);
  }

  if (response.type === "chart_plan") {
    return Boolean(response.data.requiresConfirmation);
  }

  return false;
}

const UNSAFE_PROOF_METADATA_PATTERN = /(?:APPROVAL_SECRET|HERMES_API_SERVER_KEY|HERMES_AGENT_API_KEY|HERMES_AGENT_BASE_URL|OPENAI_API_KEY|ANTHROPIC_API_KEY|stack trace|traceback|ReferenceError|TypeError|SyntaxError|RangeError)|(?:^|\s)\/(?:root|srv|home|tmp)\/[^\s]+|https?:\/\/(?:internal(?:[.\w-]*)?|localhost|127\.0\.0\.1|10\.\d{1,3}\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[0-1])\.\d{1,3}\.\d{1,3})[^\s]*/i;

function hasUnsafeProofMetadata(value) {
  return typeof value === "string" && UNSAFE_PROOF_METADATA_PATTERN.test(value);
}

function getSafeSkillsUsed(skillsUsed) {
  return Array.isArray(skillsUsed)
    ? skillsUsed.filter((skill) => !hasUnsafeProofMetadata(skill))
    : [];
}

function getSafeDownstreamProvider(provider) {
  if (!provider || hasUnsafeProofMetadata(provider.label)) {
    return null;
  }

  if (provider.model && hasUnsafeProofMetadata(provider.model)) {
    return { label: provider.label };
  }

  return provider;
}

function getResponseMetaLine(response) {
  const parts = [];
  const skillsUsed = getSafeSkillsUsed(response.skillsUsed);
  const downstreamProvider = getSafeDownstreamProvider(response.downstreamProvider);

  if (skillsUsed.length > 0) {
    parts.push(`skills ${skillsUsed.join(", ")}`);
  }

  if (downstreamProvider?.label) {
    parts.push(
      downstreamProvider.model
        ? `provider ${downstreamProvider.label}/${downstreamProvider.model}`
        : `provider ${downstreamProvider.label}`
    );
  }

  const confidence = getResponseConfidence(response);
  if (typeof confidence === "number" && response.ui?.showConfidence) {
    parts.push(`confidence ${Math.round(confidence * 100)}%`);
  }

  if (response.ui?.showRequiresConfirmation && getRequiresConfirmation(response)) {
    parts.push("confirmation required");
  }

  if (response.data?.extractionMode) {
    parts.push(`extraction ${response.data.extractionMode}`);
  }

  return parts.join(" • ");
}

function requiresDestructiveWriteApproval(plan) {
  if (!plan || typeof plan !== "object") {
    return false;
  }

  if (plan.confirmationLevel === "destructive") {
    return true;
  }

  if (isCompositePlan(plan)) {
    return plan.steps.some((step) => step?.plan?.confirmationLevel === "destructive");
  }

  return false;
}

function getWriteApprovalConfirmFunction() {
  if (typeof window?.confirm === "function") {
    return window.confirm.bind(window);
  }

  if (typeof globalThis.confirm === "function") {
    return globalThis.confirm.bind(globalThis);
  }

  throw new Error("Destructive confirmation is unavailable in this host.");
}

function getDestructiveWriteApprovalMessage(plan) {
  if (plan && typeof plan === "object" && "sourceSheet" in plan && "sourceRange" in plan) {
    return [
      "This transfer is destructive and requires a second confirmation.",
      "",
      `Move ${plan.sourceSheet}!${plan.sourceRange} to ${plan.targetSheet}!${plan.targetRange}?`,
      "",
      "Select OK to approve and execute this destructive plan."
    ].join("\n");
  }

  if (plan && typeof plan === "object" && "targetSheet" in plan && "targetRange" in plan) {
    return [
      "This cleanup is destructive and requires a second confirmation.",
      "",
      `Apply ${plan.operation} to ${plan.targetSheet}!${plan.targetRange}?`,
      "",
      "Select OK to approve and execute this destructive plan."
    ].join("\n");
  }

  return [
    "This write-back is destructive and requires a second confirmation.",
    "",
    "Select OK to approve and execute this destructive plan."
  ].join("\n");
}

function buildWriteApprovalRequest(input) {
  const resolvedPlan = isMaterializedAnalysisReportPlan(input.plan)
    ? resolveMaterializedAnalysisReportPlan(input.plan)
    : input.plan;
  const workbookSessionKey = typeof input.workbookSessionKey === "string" &&
    input.workbookSessionKey.trim().length > 0
    ? input.workbookSessionKey.trim()
    : undefined;

  const destructiveConfirmation = requiresDestructiveWriteApproval(resolvedPlan)
    ? (getWriteApprovalConfirmFunction()(getDestructiveWriteApprovalMessage(resolvedPlan))
      ? { confirmed: true }
      : null)
    : undefined;

  if (requiresDestructiveWriteApproval(resolvedPlan) && !destructiveConfirmation) {
    return null;
  }

  return destructiveConfirmation
    ? {
        requestId: input.requestId,
        runId: input.runId,
        ...(workbookSessionKey ? { workbookSessionKey } : {}),
        plan: resolvedPlan,
        destructiveConfirmation
      }
    : {
        requestId: input.requestId,
        runId: input.runId,
        ...(workbookSessionKey ? { workbookSessionKey } : {}),
        plan: resolvedPlan
      };
}

function formatProofLine(response) {
  return [
    "Processed by Hermes",
    `requestId ${response.requestId}`,
    `hermesRunId ${response.hermesRunId}`,
    `service ${response.serviceLabel}`,
    `environment ${response.environmentLabel}`,
    `${response.durationMs}ms`
  ].join(" • ");
}

function getFollowUpSuggestions(response) {
  return response.type === "chat"
    ? response.data.followUpSuggestions || []
    : [];
}

function isWritePlanResponse(response) {
  if (!response) {
    return false;
  }

  if (response.type === "analysis_report_plan" &&
    response.data?.outputMode === "materialize_report") {
    return true;
  }

  if (response.type === "composite_plan") {
    return Array.isArray(response.data?.steps) &&
      getCompositePreviewSupportErrors(response.data.steps).length === 0;
  }

  if (isExcelPreviewSupportCheckedWritePlanType(response.type)) {
    if (!response.data || typeof response.data !== "object") {
      return true;
    }
    return !getExcelPlanSupportError(getStructuredPreview(response));
  }

  return response.type === "workbook_structure_update" ||
    response.type === "sheet_structure_update" ||
    response.type === "range_sort_plan" ||
    response.type === "range_filter_plan" ||
    response.type === "range_format_update" ||
    response.type === "conditional_format_plan" ||
    response.type === "data_validation_plan" ||
    response.type === "named_range_update" ||
    response.type === "range_transfer_plan" ||
    response.type === "data_cleanup_plan" ||
    response.type === "composite_plan" ||
    response.type === "sheet_update" ||
    response.type === "sheet_import_plan";
}

function buildWriteMatrix(plan) {
  if (Array.isArray(plan.headers)) {
    return [plan.headers, ...plan.values];
  }

  if (Array.isArray(plan.values)) {
    return plan.values;
  }

  if (Array.isArray(plan.formulas)) {
    return plan.formulas;
  }

  if (Array.isArray(plan.notes)) {
    return plan.notes;
  }

  return [];
}

function formatWorkbookPositionLabel(position) {
  if (position === undefined) {
    return "";
  }

  if (position === "start" || position === "end") {
    return ` • ${position}`;
  }

  return ` • index ${position}`;
}

function formatRangeFormatFields(format) {
  return Object.entries(format || {})
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => `${key}=${formatRangeFormatValue(value)}`)
    .join(" • ");
}

function formatRangeFormatValue(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return String(value);
  }

  return Object.entries(value)
    .filter(([, nestedValue]) => nestedValue !== undefined)
    .map(([nestedKey, nestedValue]) => {
      if (!nestedValue || typeof nestedValue !== "object") {
        return `${nestedKey}:${nestedValue}`;
      }

      return `${nestedKey}:${Object.entries(nestedValue)
        .filter(([, lineValue]) => lineValue !== undefined)
        .map(([lineKey, lineValue]) => `${lineKey}=${lineValue}`)
        .join(",")}`;
    })
    .join(";");
}

function mapRangeBorderSideToExcel(side) {
  const borderIndex = Excel?.BorderIndex || {};
  switch (side) {
    case "top":
      return borderIndex.edgeTop || "EdgeTop";
    case "bottom":
      return borderIndex.edgeBottom || "EdgeBottom";
    case "left":
      return borderIndex.edgeLeft || "EdgeLeft";
    case "right":
      return borderIndex.edgeRight || "EdgeRight";
    case "innerHorizontal":
      return borderIndex.insideHorizontal || "InsideHorizontal";
    case "innerVertical":
      return borderIndex.insideVertical || "InsideVertical";
    default:
      return null;
  }
}

function mapRangeBorderLineStyleToExcel(style) {
  const borderLineStyle = Excel?.BorderLineStyle || Excel?.LineStyle || {};
  switch (style) {
    case "none":
      return borderLineStyle.none || "None";
    case "dashed":
      return borderLineStyle.dash || "Dash";
    case "dotted":
      return borderLineStyle.dot || "Dot";
    case "double":
      return borderLineStyle.double || "Double";
    case "solid":
    case "medium":
    case "thick":
      return borderLineStyle.continuous || "Continuous";
    default:
      return null;
  }
}

function mapRangeBorderWeightToExcel(style) {
  const borderWeight = Excel?.BorderWeight || {};
  switch (style) {
    case "medium":
      return borderWeight.medium || "Medium";
    case "thick":
      return borderWeight.thick || "Thick";
    case "solid":
    case "dashed":
    case "dotted":
    case "double":
      return borderWeight.thin || "Thin";
    default:
      return null;
  }
}

function applyExcelRangeBorder(target, borderPlan) {
  for (const { side, line } of expandRangeBorderLines(borderPlan)) {
    const excelSide = mapRangeBorderSideToExcel(side);
    const lineStyle = mapRangeBorderLineStyleToExcel(line.style);
    if (!excelSide || !lineStyle) {
      continue;
    }

    const border = target.format.borders.getItem(excelSide);
    border.lineStyle = lineStyle;

    if (line.style !== "none") {
      if (line.color) {
        border.color = line.color;
      }

      const weight = mapRangeBorderWeightToExcel(line.style);
      if (weight) {
        border.weight = weight;
      }
    }
  }
}

function isDataValidationPlan(plan) {
  return Boolean(
    plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    typeof plan.ruleType === "string"
  );
}

function isConditionalFormatPlan(plan) {
  if (!plan ||
    typeof plan.targetSheet !== "string" ||
    typeof plan.targetRange !== "string" ||
    typeof plan.managementMode !== "string") {
    return false;
  }

  if (plan.managementMode === "clear_on_target") {
    return true;
  }

  return typeof plan.ruleType === "string";
}

function getConditionalFormatPreviewSummary(plan) {
  switch (plan?.managementMode) {
    case "add":
      return `Will add conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
    case "replace_all_on_target":
      return `Will replace conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
    case "clear_on_target":
      return `Will clear conditional formatting on ${plan.targetSheet}!${plan.targetRange}.`;
    default:
      return "Will update conditional formatting.";
  }
}

function formatConditionalFormatFields(preview) {
  const fields = [`management ${preview.managementMode}`];

  if (preview.ruleType) {
    fields.push(`rule ${preview.ruleType}`);
  }

  if (preview.comparator) {
    fields.push(`comparator ${preview.comparator}`);
  }

  return fields.join(" • ");
}

function formatConditionalFormatDetails(preview) {
  const details = [];

  if (preview.text) {
    details.push(`text ${preview.text}`);
  }

  if (preview.formula) {
    details.push(`formula ${preview.formula}`);
  }

  if (preview.value !== undefined) {
    details.push(`value ${preview.value}`);
  }

  if (preview.value2 !== undefined) {
    details.push(`value2 ${preview.value2}`);
  }

  if (preview.rank !== undefined) {
    details.push(`rank ${preview.rank}`);
  }

  if (preview.direction) {
    details.push(`direction ${preview.direction}`);
  }

  if (Array.isArray(preview.points) && preview.points.length > 0) {
    details.push(`points ${preview.points.map((point) => {
      const label = point?.value === undefined ? point?.type : `${point?.type}:${point.value}`;
      return `${label} ${point?.color || ""}`.trim();
    }).join(", ")}`);
  }

  return details.join(" • ");
}

function isNamedRangeUpdatePlan(plan) {
  return Boolean(
    plan &&
    typeof plan.name === "string" &&
    typeof plan.operation === "string"
  );
}

function isRangeTransferPlan(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === "string" &&
    typeof plan.sourceRange === "string" &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    (plan.operation === "copy" || plan.operation === "move" || plan.operation === "append")
  );
}

function isDataCleanupPlan(plan) {
  return Boolean(
    plan &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    typeof plan.operation === "string" &&
    [
      "trim_whitespace",
      "remove_blank_rows",
      "remove_duplicate_rows",
      "normalize_case",
      "split_column",
      "join_columns",
      "fill_down",
      "standardize_format"
    ].includes(plan.operation)
  );
}

function mapDataValidationComparatorToExcel(comparator) {
  const operator = Excel?.DataValidationOperator;
  const map = {
    between: operator?.between || "Between",
    not_between: operator?.notBetween || "NotBetween",
    equal_to: operator?.equalTo || "EqualTo",
    not_equal_to: operator?.notEqualTo || "NotEqualTo",
    greater_than: operator?.greaterThan || "GreaterThan",
    greater_than_or_equal_to: operator?.greaterThanOrEqualTo || "GreaterThanOrEqualTo",
    less_than: operator?.lessThan || "LessThan",
    less_than_or_equal_to: operator?.lessThanOrEqualTo || "LessThanOrEqualTo"
  };

  const mapped = map[comparator];
  if (!mapped) {
    throw new Error(`Unsupported Excel validation comparator: ${comparator}`);
  }

  return mapped;
}

function getExcelListValidationSource(plan) {
  if (plan.namedRangeName) {
    return plan.namedRangeName;
  }

  if (plan.sourceRange) {
    return plan.sourceRange;
  }

  if (Array.isArray(plan.values)) {
    const hasCommaValue = plan.values.some((value) => String(value).includes(","));
    if (hasCommaValue) {
      throw new Error("Excel inline list validation values cannot contain commas exactly.");
    }

    return plan.values.join(",");
  }

  throw new Error("List validation requires values, sourceRange, or namedRangeName.");
}

function buildExcelValidationRule(plan) {
  const operator = "comparator" in plan
    ? mapDataValidationComparatorToExcel(plan.comparator)
    : undefined;

  switch (plan.ruleType) {
    case "whole_number":
      return {
        wholeNumber: {
          operator,
          formula1: plan.value,
          formula2: plan.value2
        }
      };
    case "decimal":
      return {
        decimal: {
          operator,
          formula1: plan.value,
          formula2: plan.value2
        }
      };
    case "date":
      return {
        date: {
          operator,
          formula1: plan.value,
          formula2: plan.value2
        }
      };
    case "text_length":
      return {
        textLength: {
          operator,
          formula1: plan.value,
          formula2: plan.value2
        }
      };
    case "custom_formula":
      return {
        custom: {
          formula: plan.formula
        }
      };
    case "list":
      return {
        list: {
          source: getExcelListValidationSource(plan),
          inCellDropDown: typeof plan.showDropdown === "boolean" ? plan.showDropdown : undefined
        }
      };
    default:
      throw new Error("Unsupported Excel data validation rule type.");
  }
}

function getValidationInputMessage(plan) {
  if (typeof plan.inputMessage === "string" && plan.inputMessage.trim().length > 0) {
    return plan.inputMessage.trim();
  }

  if (typeof plan.helpText === "string" && plan.helpText.trim().length > 0) {
    return plan.helpText.trim();
  }

  return plan.explanation;
}

function hasCustomValidationPromptOrAlert(plan) {
  return Boolean(
    plan.helpText ||
    plan.inputTitle ||
    plan.inputMessage ||
    plan.errorTitle ||
    plan.errorMessage
  );
}

function hasUnsupportedExcelCheckboxValues(plan) {
  return (
    (plan.checkedValue !== undefined && plan.checkedValue !== true) ||
    (plan.uncheckedValue !== undefined && plan.uncheckedValue !== false)
  );
}

function getUnsupportedExcelCheckboxSemanticsMessage(plan) {
  if (plan.allowBlank === true) {
    return "Excel checkbox controls cannot represent allowBlank=true checkbox semantics.";
  }

  return "";
}

function getExcelRangeFormatSupportError(format, hostLabel = "Excel host") {
  if (format?.wrapStrategy === "clip") {
    if (hostLabel === "This Excel runtime") {
      return "This Excel runtime can't clip overflowing text exactly.";
    }

    return `${hostLabel} cannot clip overflowing text exactly.`;
  }

  return "";
}

function applyExcelCheckboxValidation(target, plan) {
  if (hasUnsupportedExcelCheckboxValues(plan)) {
    throw new Error("Excel checkbox controls only support boolean true/false values.");
  }

  const semanticsError = getUnsupportedExcelCheckboxSemanticsMessage(plan);
  if (semanticsError) {
    throw new Error(semanticsError);
  }

  if (!Object.prototype.hasOwnProperty.call(target, "control")) {
    throw new Error("Excel host does not support checkbox cell controls on this range.");
  }

  if (!Excel?.CellControlType?.checkbox) {
    throw new Error("Excel host does not expose checkbox cell control support.");
  }

  target.control = {
    type: Excel.CellControlType.checkbox
  };
}

async function assertExcelNamedRangeCreateDoesNotCollide(context, collection, name) {
  if (!collection || typeof collection.getItemOrNullObject !== "function") {
    return;
  }

  const existingNamedRange = collection.getItemOrNullObject(name);
  if (!existingNamedRange) {
    return;
  }

  if (typeof existingNamedRange.load === "function") {
    existingNamedRange.load("name");
  }

  await context.sync();

  if (existingNamedRange.isNullObject !== true) {
    throw new Error(`Named range already exists: ${name}`);
  }
}

async function resolveExcelNamedRangeForUpdate(context, collection, name) {
  if (typeof collection.getItem === "function") {
    return collection.getItem(name);
  }

  if (typeof collection.getItemOrNullObject !== "function") {
    return undefined;
  }

  const namedRange = collection.getItemOrNullObject(name);
  if (!namedRange) {
    return undefined;
  }

  if (typeof namedRange.load === "function") {
    namedRange.load("name");
  }

  await context.sync();

  return namedRange.isNullObject === true ? undefined : namedRange;
}

async function applyExcelNamedRangeUpdate(context, workbook, worksheet, plan, targetRange) {
  const collection = plan.scope === "sheet"
    ? worksheet.names
    : workbook.names;

  if (!collection) {
    throw new Error("Excel host does not support named ranges on this scope.");
  }

  if ((plan.operation === "create" || plan.operation === "retarget") &&
    (typeof plan.targetSheet !== "string" || plan.targetSheet.length === 0 ||
      typeof plan.targetRange !== "string" || plan.targetRange.length === 0)) {
    throw new Error("Named range create and retarget require targetSheet and targetRange.");
  }

  const resolvedReference = targetRange.address || `${plan.targetSheet}!${plan.targetRange}`;

  switch (plan.operation) {
    case "create":
      if (!collection.add) {
        throw new Error("Excel host does not support creating named ranges on this scope.");
      }
      collection.add(plan.name, targetRange);
      return;
    case "retarget": {
      const namedRange = await resolveExcelNamedRangeForUpdate(context, collection, plan.name);
      if (!namedRange) {
        throw new Error(`Named range not found: ${plan.name}`);
      }
      namedRange.reference = resolvedReference;
      return;
    }
    case "rename": {
      const namedRange = await resolveExcelNamedRangeForUpdate(context, collection, plan.name);
      if (!namedRange) {
        throw new Error(`Named range not found: ${plan.name}`);
      }
      if (!plan.newName) {
        throw new Error("Named range rename requires newName.");
      }
      namedRange.name = plan.newName;
      return;
    }
    case "delete": {
      const namedRange = await resolveExcelNamedRangeForUpdate(context, collection, plan.name);
      if (!namedRange) {
        throw new Error(`Named range not found: ${plan.name}`);
      }
      if (!namedRange.delete) {
        throw new Error("Excel host does not support deleting named ranges on this scope.");
      }
      namedRange.delete();
      return;
    }
    default:
      throw new Error(`Unsupported named range update: ${plan.operation}`);
  }
}

function cloneMatrix(matrix) {
  return (matrix || []).map((row) => Array.isArray(row) ? [...row] : []);
}

function transposeMatrix(matrix) {
  const rowCount = matrix?.length || 0;
  const columnCount = Math.max(0, ...(matrix || []).map((row) => row?.length || 0));
  return Array.from({ length: columnCount }, (_, columnIndex) =>
    Array.from({ length: rowCount }, (_, rowIndex) => matrix?.[rowIndex]?.[columnIndex] ?? null)
  );
}

function isBlankCellValue(value) {
  return value === null || value === undefined || value === "";
}

function parseA1CellReference(reference) {
  const match = String(reference || "").trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Unsupported A1 reference: ${reference}`);
  }

  const [, columnLetters, rowText] = match;
  let column = 0;
  for (const character of columnLetters) {
    column = (column * 26) + (character.charCodeAt(0) - 64);
  }

  return {
    row: Number(rowText),
    column
  };
}

function parseA1RangeReference(reference) {
  const normalized = normalizeA1(reference).trim().toUpperCase();
  const [startRef, endRef = startRef] = normalized.split(":");
  const start = parseA1CellReference(startRef);
  const end = parseA1CellReference(endRef);

  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
    rowCount: Math.abs(end.row - start.row) + 1,
    columnCount: Math.abs(end.column - start.column) + 1
  };
}

function columnNumberToLetters(columnNumber) {
  let value = Number(columnNumber);
  let letters = "";

  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }

  return letters;
}

function buildA1RangeFromBounds(bounds) {
  const startCell = `${columnNumberToLetters(bounds.startColumn)}${bounds.startRow}`;
  const endCell = `${columnNumberToLetters(bounds.endColumn)}${bounds.endRow}`;
  return startCell === endCell ? startCell : `${startCell}:${endCell}`;
}

function buildExcelFullRowOrColumnAddress(startIndex, count, isRowOperation) {
  const start = Number(startIndex) + 1;
  const span = Number(count);

  if (!Number.isInteger(start) || !Number.isInteger(span) || start < 1 || span < 1) {
    throw new Error("Sheet structure row or column operation requires a valid startIndex and count.");
  }

  const end = start + span - 1;
  if (isRowOperation) {
    return `${start}:${end}`;
  }

  return `${columnNumberToLetters(start)}:${columnNumberToLetters(end)}`;
}

function getExcelFullRowOrColumnRange(sheet, plan, isRowOperation) {
  if (typeof sheet.getRange !== "function") {
    throw new Error("Excel host does not support full row or column sheet structure ranges.");
  }

  return sheet.getRange(buildExcelFullRowOrColumnAddress(plan.startIndex, plan.count, isRowOperation));
}

function buildSizedA1RangeFromAnchor(anchorRange, rowCount, columnCount) {
  const bounds = parseA1RangeReference(anchorRange);
  return buildA1RangeFromBounds({
    startRow: bounds.startRow,
    endRow: bounds.startRow + rowCount - 1,
    startColumn: bounds.startColumn,
    endColumn: bounds.startColumn + columnCount - 1
  });
}

function rangesOverlap(left, right) {
  return !(left.endRow < right.startRow ||
    right.endRow < left.startRow ||
    left.endColumn < right.startColumn ||
    right.endColumn < left.startColumn);
}

function resolveColumnOffsetWithinRange(columnRef, targetRange) {
  const bounds = parseA1RangeReference(targetRange);
  const trimmed = String(columnRef || "").trim().toUpperCase();
  const column = /^\d+$/.test(trimmed)
    ? Number(trimmed)
    : parseA1CellReference(`${trimmed}1`).column;
  const offset = column - bounds.startColumn;

  if (offset < 0 || offset >= bounds.columnCount) {
    throw new Error(`Column ${columnRef} is outside ${targetRange}.`);
  }

  return offset;
}

function getResolvedTransferShape(sourceRange, plan) {
  return {
    rows: plan.transpose ? sourceRange.columnCount : sourceRange.rowCount,
    columns: plan.transpose ? sourceRange.rowCount : sourceRange.columnCount
  };
}

function resolveExactMatrixTargetRange(targetRange, expectedRows, expectedColumns, shapeLabel = "write shape") {
  if (targetRange.rowCount === expectedRows && targetRange.columnCount === expectedColumns) {
    return targetRange;
  }

  if (targetRange.rowCount === 1 &&
    targetRange.columnCount === 1 &&
    typeof targetRange.getResizedRange === "function") {
    return targetRange.getResizedRange(expectedRows - 1, expectedColumns - 1);
  }

  throw new Error(`The approved targetRange does not match the proposed ${shapeLabel}.`);
}

function resolveFullMatrixTargetRange(targetRange, expectedRows, expectedColumns, shapeLabel) {
  if (targetRange.rowCount === expectedRows && targetRange.columnCount === expectedColumns) {
    return targetRange;
  }

  throw new Error(`The approved targetRange does not match the ${shapeLabel}.`);
}

function resolveAppendMatrixTargetRange(targetRange, expectedRows, expectedColumns) {
  if (targetRange.rowCount === expectedRows && targetRange.columnCount === expectedColumns) {
    return targetRange;
  }

  throw new Error("Excel host requires append targetRange to match the full destination rectangle.");
}

function normalizeTransferMatrix(matrix, plan) {
  const base = cloneMatrix(matrix);
  return plan.transpose ? transposeMatrix(base) : base;
}

function normalizeFormulaTransferMatrix(formulas, plan) {
  const base = (formulas || []).map((row) =>
    (row || []).map((value) => (value == null ? "" : value))
  );
  return plan.transpose ? transposeMatrix(base) : base;
}

function deriveTransferTargetBounds(plan, resolvedTargetRange) {
  const planBounds = parseA1RangeReference(plan.targetRange);
  if (planBounds.rowCount === resolvedTargetRange.rowCount &&
    planBounds.columnCount === resolvedTargetRange.columnCount) {
    return planBounds;
  }

  if (planBounds.rowCount === 1 && planBounds.columnCount === 1) {
    return {
      startRow: planBounds.startRow,
      endRow: planBounds.startRow + resolvedTargetRange.rowCount - 1,
      startColumn: planBounds.startColumn,
      endColumn: planBounds.startColumn + resolvedTargetRange.columnCount - 1,
      rowCount: resolvedTargetRange.rowCount,
      columnCount: resolvedTargetRange.columnCount
    };
  }

  return parseA1RangeReference(normalizeA1(resolvedTargetRange.address || plan.targetRange));
}

function assertNonOverlappingTransfer(plan, resolvedTargetRange) {
  if (plan.sourceSheet !== plan.targetSheet) {
    return;
  }

  if (plan.operation === "copy") {
    return;
  }

  const sourceBounds = parseA1RangeReference(plan.sourceRange);
  const targetBounds = deriveTransferTargetBounds(plan, resolvedTargetRange);

  if (rangesOverlap(sourceBounds, targetBounds)) {
    throw new Error(`Excel host cannot apply an overlapping ${plan.operation} transfer exactly.`);
  }
}

function writeTransferValues(targetRange, sourceRange, plan) {
  if (plan.pasteMode === "values") {
    targetRange.values = normalizeTransferMatrix(sourceRange.values, plan);
    return;
  }

  if (plan.pasteMode === "formulas") {
    if (typeof targetRange.copyFrom !== "function") {
      throw new Error("Excel host does not support exact-safe formula transfers on this range.");
    }

    targetRange.copyFrom(
      sourceRange,
      Excel?.RangeCopyType?.formulas || "Formulas",
      false,
      Boolean(plan.transpose)
    );
    return;
  }

  if (plan.pasteMode === "formats") {
    if (typeof targetRange.copyFrom !== "function") {
      throw new Error("Excel host does not support exact format transfers on this range.");
    }

    targetRange.copyFrom(
      sourceRange,
      Excel?.RangeCopyType?.formats || "Formats",
      false,
      plan.transpose
    );
    return;
  }

  throw new Error(`Unsupported transfer pasteMode: ${plan.pasteMode}`);
}

function clearTransferredSource(sourceRange, plan) {
  if (typeof sourceRange.clear === "function") {
    sourceRange.clear(plan.pasteMode === "formats"
      ? (Excel?.ClearApplyTo?.formats || "Formats")
      : (Excel?.ClearApplyTo?.contents || "Contents"));
    return;
  }

  if (plan.pasteMode === "formulas" && "formulas" in sourceRange) {
    sourceRange.formulas = Array.from({ length: sourceRange.rowCount }, () =>
      Array.from({ length: sourceRange.columnCount }, () => "")
    );
    return;
  }

  if ("values" in sourceRange) {
    sourceRange.values = Array.from({ length: sourceRange.rowCount }, () =>
      Array.from({ length: sourceRange.columnCount }, () => "")
    );
    return;
  }

  throw new Error("Excel host cannot clear the source range for this move.");
}

function getCleanupColumnOffsets(plan) {
  switch (plan.operation) {
    case "remove_blank_rows":
    case "remove_duplicate_rows":
      return (plan.keyColumns || []).map((columnRef) =>
        resolveColumnOffsetWithinRange(columnRef, plan.targetRange)
      );
    case "split_column":
      return {
        source: resolveColumnOffsetWithinRange(plan.sourceColumn, plan.targetRange),
        targetStart: resolveColumnOffsetWithinRange(plan.targetStartColumn, plan.targetRange)
      };
    case "join_columns":
      return {
        source: plan.sourceColumns.map((columnRef) =>
          resolveColumnOffsetWithinRange(columnRef, plan.targetRange)
        ),
        target: resolveColumnOffsetWithinRange(plan.targetColumn, plan.targetRange)
      };
    case "fill_down":
      return (plan.columns || []).map((columnRef) =>
        resolveColumnOffsetWithinRange(columnRef, plan.targetRange)
      );
    default:
      return [];
  }
}

function fillTrailingBlankRows(rows, targetColumnCount, targetRowCount) {
  const paddedRows = rows.map((row) => {
    const nextRow = Array.from({ length: targetColumnCount }, (_, index) => row[index] ?? "");
    return nextRow;
  });

  while (paddedRows.length < targetRowCount) {
    paddedRows.push(Array.from({ length: targetColumnCount }, () => ""));
  }

  return paddedRows.slice(0, targetRowCount);
}

function hasAnyRealFormula(formulas) {
  return (formulas || []).some((row) =>
    (row || []).some((value) => typeof value === "string" && value.trim().startsWith("="))
  );
}

function getRangeOccupancyMatrix(values, formulas) {
  return (values || []).map((row, rowIndex) =>
    (row || []).map((value, columnIndex) => {
      const formulaValue = formulas?.[rowIndex]?.[columnIndex];
      return (typeof formulaValue === "string" && formulaValue.trim().startsWith("=")) ||
        !isBlankCellValue(value);
    })
  );
}

function toTitleCaseText(value) {
  const lowerCased = String(value ?? "").toLocaleLowerCase();
  return lowerCased.replace(/(^|[^A-Za-z0-9])([A-Za-z0-9])/g, (match, prefix, character) =>
    `${prefix}${character.toLocaleUpperCase()}`
  );
}

function toSentenceCaseText(value) {
  const lowerCased = String(value ?? "").toLocaleLowerCase();
  if (!lowerCased) {
    return lowerCased;
  }

  return `${lowerCased.charAt(0).toLocaleUpperCase()}${lowerCased.slice(1)}`;
}

function quoteCleanupFormulaString(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

function getStandardizeFormatFormulaPattern(spec) {
  if (spec.formatType === "date_text") {
    const separator = spec.separator || "-";
    return `yyyy${separator}mm${separator}dd`;
  }

  return spec.formatPattern;
}

function buildStandardizeFormatFormulaExpression(spec) {
  const quotedPattern = quoteCleanupFormulaString(getStandardizeFormatFormulaPattern(spec));

  if (spec.formatType === "date_text") {
    return `IF(_hermes_value="", "", IF(ISNUMBER(_hermes_value), TEXT(_hermes_value, ${quotedPattern}), TEXT(DATEVALUE(_hermes_value), ${quotedPattern})))`;
  }

  return `IF(_hermes_value="", "", TEXT(VALUE(SUBSTITUTE(_hermes_value, ",", "")), ${quotedPattern}))`;
}

function getFormulaAwareCleanupTransform(plan, hostLabel) {
  switch (plan.operation) {
    case "trim_whitespace":
      return {
        applyToValue(value) {
          return typeof value === "string" ? value.trim() : value;
        },
        formulaFunction: "TRIM"
      };
    case "normalize_case":
      switch (plan.mode) {
        case "upper":
          return {
            applyToValue(value) {
              return typeof value === "string" ? value.toLocaleUpperCase() : value;
            },
            formulaFunction: "UPPER"
          };
        case "lower":
          return {
            applyToValue(value) {
              return typeof value === "string" ? value.toLocaleLowerCase() : value;
            },
            formulaFunction: "LOWER"
          };
        case "title":
          return {
            applyToValue(value) {
              return typeof value === "string" ? toTitleCaseText(value) : value;
            },
            formulaFunction: "PROPER"
          };
        case "sentence":
          return {
            applyToValue(value) {
              return typeof value === "string" ? toSentenceCaseText(value) : value;
            },
            formulaExpression: "LET(_hermes_text, LOWER(_hermes_value), UPPER(LEFT(_hermes_text,1)) & MID(_hermes_text,2,LEN(_hermes_text)))"
          };
        default:
          throw new Error(`${hostLabel} does not support exact-safe cleanup semantics for normalize_case mode ${plan.mode}.`);
      }
    case "standardize_format": {
      const spec = getSupportedStandardizeFormatSpec(plan?.formatType, plan?.formatPattern);
      if (!spec) {
        throw new Error(getStandardizeFormatSupportError(plan?.formatType, plan?.formatPattern, hostLabel));
      }

      return {
        applyToValue(value) {
          return spec.formatType === "date_text"
            ? standardizeDateTextValueExact(value, spec, hostLabel)
            : standardizeNumberTextValueExact(value, spec, hostLabel);
        },
        formulaCondition: "OR(ISTEXT(_hermes_value), ISNUMBER(_hermes_value))",
        formulaExpression: buildStandardizeFormatFormulaExpression(spec)
      };
    }
    default:
      return null;
  }
}

function wrapFormulaWithCleanupTransform(formula, formulaAwareTransform) {
  const normalizedFormula = normalizeExcelFormulaText(formula)?.trim();
  if (!normalizedFormula?.startsWith("=")) {
    return formula;
  }

  const expression = normalizedFormula.slice(1);
  const transformedExpression = formulaAwareTransform.formulaExpression ||
    `${formulaAwareTransform.formulaFunction}(_hermes_value)`;
  const condition = formulaAwareTransform.formulaCondition || "ISTEXT(_hermes_value)";
  return `=LET(_hermes_value, ${expression}, IF(${condition}, ${transformedExpression}, _hermes_value))`;
}

function getFormulaCellText(formulas, rowIndex, columnIndex) {
  const formulaValue = normalizeExcelFormulaText(formulas?.[rowIndex]?.[columnIndex]);
  return typeof formulaValue === "string" && formulaValue.trim().startsWith("=")
    ? formulaValue.trim()
    : "";
}

function applyFormulaAwareFillDownCleanup(target, plan, inputValues, inputFormulas, hostLabel) {
  if (plan.operation !== "fill_down" || !hasAnyRealFormula(inputFormulas)) {
    return false;
  }

  if (typeof target?.getCell !== "function") {
    throw new Error(`${hostLabel} cannot apply formula-aware fill_down without cell-level range access.`);
  }

  const values = cloneMatrix(inputValues);
  const formulas = cloneMatrix(inputFormulas);
  const explicitOffsets = getCleanupColumnOffsets(plan);
  const targetOffsets = explicitOffsets.length > 0
    ? explicitOffsets
    : Array.from({ length: values[0]?.length || 0 }, (_, index) => index);

  for (const columnIndex of targetOffsets) {
    let lastSeen = null;

    for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
      const formula = getFormulaCellText(formulas, rowIndex, columnIndex);
      const value = values[rowIndex]?.[columnIndex];

      if (!formula && isBlankCellValue(value)) {
        if (!lastSeen) {
          continue;
        }

        const destination = target.getCell(rowIndex, columnIndex);
        if (lastSeen.kind === "formula") {
          const source = target.getCell(lastSeen.rowIndex, columnIndex);
          if (typeof destination?.copyFrom !== "function") {
            throw new Error(`${hostLabel} cannot apply formula-aware fill_down without formula copy support.`);
          }

          destination.copyFrom(source, Excel?.RangeCopyType?.formulas || "Formulas", false, false);
        } else {
          destination.values = [[lastSeen.value]];
        }
        continue;
      }

      lastSeen = formula
        ? { kind: "formula", rowIndex }
        : { kind: "value", value };
    }
  }

  return true;
}

function buildFormulaPreservingRow(row, rowIndex, formulas, targetColumnCount) {
  return Array.from({ length: targetColumnCount }, (_, columnIndex) => {
    const formula = getFormulaCellText(formulas, rowIndex, columnIndex);
    return formula || (row?.[columnIndex] ?? "");
  });
}

function isFormulaAwareRowBlank(row, rowIndex, formulas, columnOffsets) {
  return columnOffsets.every((columnIndex) =>
    !getFormulaCellText(formulas, rowIndex, columnIndex) &&
    isBlankCellValue(row?.[columnIndex])
  );
}

function buildFormulaAwareRemoveBlankRowsMatrix(plan, inputValues, inputFormulas) {
  const values = cloneMatrix(inputValues);
  const formulas = cloneMatrix(inputFormulas);
  const targetColumnCount = values[0]?.length || formulas[0]?.length || 0;
  const explicitOffsets = getCleanupColumnOffsets(plan);
  const columnOffsets = explicitOffsets.length > 0
    ? explicitOffsets
    : Array.from({ length: targetColumnCount }, (_, index) => index);
  const retainedRows = [];

  values.forEach((row, rowIndex) => {
    if (isFormulaAwareRowBlank(row, rowIndex, formulas, columnOffsets)) {
      return;
    }

    retainedRows.push(buildFormulaPreservingRow(row, rowIndex, formulas, targetColumnCount));
  });

  return fillTrailingBlankRows(retainedRows, targetColumnCount, values.length);
}

function getFormulaAwareRowDigest(row, rowIndex, formulas, columnOffsets) {
  return JSON.stringify(columnOffsets.map((columnIndex) => {
    const formula = getFormulaCellText(formulas, rowIndex, columnIndex);
    return formula || (row?.[columnIndex] ?? "");
  }));
}

function buildFormulaAwareRemoveDuplicateRowsMatrix(plan, inputValues, inputFormulas) {
  const values = cloneMatrix(inputValues);
  const formulas = cloneMatrix(inputFormulas);
  const targetColumnCount = values[0]?.length || formulas[0]?.length || 0;
  const explicitOffsets = getCleanupColumnOffsets(plan);
  const columnOffsets = explicitOffsets.length > 0
    ? explicitOffsets
    : Array.from({ length: targetColumnCount }, (_, index) => index);
  const retainedRows = [];
  const seen = new Set();

  values.forEach((row, rowIndex) => {
    const digest = getFormulaAwareRowDigest(row, rowIndex, formulas, columnOffsets);
    if (seen.has(digest)) {
      return;
    }

    seen.add(digest);
    retainedRows.push(buildFormulaPreservingRow(row, rowIndex, formulas, targetColumnCount));
  });

  return fillTrailingBlankRows(retainedRows, targetColumnCount, values.length);
}

function quoteSpreadsheetFormulaString(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

function buildFormulaAwareSplitPartFormula(sourceFormula, delimiter, partIndex) {
  const expression = sourceFormula.slice(1);
  const quotedDelimiter = quoteSpreadsheetFormulaString(delimiter);
  return `=LET(_hermes_value, ${expression}, IFERROR(INDEX(SPLIT(TO_TEXT(_hermes_value), ${quotedDelimiter}), 1, ${partIndex}), ""))`;
}

function buildFormulaAwareSplitColumnMatrix(plan, inputValues, inputFormulas) {
  const values = cloneMatrix(inputValues);
  const formulas = cloneMatrix(inputFormulas);
  const targetColumnCount = values[0]?.length || formulas[0]?.length || 0;
  const { source, targetStart } = getCleanupColumnOffsets(plan);
  const targetCapacity = targetColumnCount - targetStart;

  return values.map((row, rowIndex) => {
    const nextRow = buildFormulaPreservingRow(row, rowIndex, formulas, targetColumnCount);
    const sourceFormula = getFormulaCellText(formulas, rowIndex, source);
    const currentParts = String(row?.[source] ?? "").split(plan.delimiter);
    if (currentParts.length > targetCapacity) {
      throw new Error("Excel host cannot split this column exactly within the approved target range.");
    }

    for (let offset = 0; offset < targetCapacity; offset += 1) {
      nextRow[targetStart + offset] = sourceFormula
        ? buildFormulaAwareSplitPartFormula(sourceFormula, plan.delimiter, offset + 1)
        : currentParts[offset] ?? "";
    }

    return nextRow;
  });
}

function buildFormulaAwareJoinColumnMatrix(plan, inputValues, inputFormulas) {
  const values = cloneMatrix(inputValues);
  const formulas = cloneMatrix(inputFormulas);
  const targetColumnCount = values[0]?.length || formulas[0]?.length || 0;
  const { source, target } = getCleanupColumnOffsets(plan);
  const quotedDelimiter = quoteSpreadsheetFormulaString(plan.delimiter);

  return values.map((row, rowIndex) => {
    const nextRow = buildFormulaPreservingRow(row, rowIndex, formulas, targetColumnCount);
    const sourceParts = source.map((columnIndex) => {
      const formula = getFormulaCellText(formulas, rowIndex, columnIndex);
      return formula ? formula.slice(1) : quoteSpreadsheetFormulaString(row?.[columnIndex] ?? "");
    });
    const hasSourceFormula = source.some((columnIndex) => getFormulaCellText(formulas, rowIndex, columnIndex));

    nextRow[target] = hasSourceFormula
      ? `=TEXTJOIN(${quotedDelimiter}, FALSE, ${sourceParts.join(", ")})`
      : source.map((columnIndex) => String(row?.[columnIndex] ?? "")).join(plan.delimiter);

    return nextRow;
  });
}

function buildCleanupWriteMatrix(plan, inputValues, inputFormulas, hostLabel) {
  const values = cloneMatrix(inputValues);
  const formulas = cloneMatrix(inputFormulas);
  const formulaAwareTransform = getFormulaAwareCleanupTransform(plan, hostLabel);

  if (!hasAnyRealFormula(formulas)) {
    return {
      kind: "values",
      matrix: applyCleanupTransform(plan, values, hostLabel)
    };
  }

  if (plan.operation === "remove_blank_rows") {
    return {
      kind: "formulas",
      matrix: buildFormulaAwareRemoveBlankRowsMatrix(plan, values, formulas)
    };
  }

  if (plan.operation === "remove_duplicate_rows") {
    return {
      kind: "formulas",
      matrix: buildFormulaAwareRemoveDuplicateRowsMatrix(plan, values, formulas)
    };
  }

  if (plan.operation === "split_column") {
    return {
      kind: "formulas",
      matrix: buildFormulaAwareSplitColumnMatrix(plan, values, formulas)
    };
  }

  if (plan.operation === "join_columns") {
    return {
      kind: "formulas",
      matrix: buildFormulaAwareJoinColumnMatrix(plan, values, formulas)
    };
  }

  if (!formulaAwareTransform) {
    throw new Error(`${hostLabel} cannot apply cleanup plans exactly when the target range contains formulas.`);
  }

  return {
    kind: "formulas",
    matrix: values.map((row, rowIndex) =>
      (row || []).map((value, columnIndex) => {
        const formulaValue = formulas?.[rowIndex]?.[columnIndex];
        if (typeof formulaValue === "string" && formulaValue.trim().startsWith("=")) {
          return wrapFormulaWithCleanupTransform(formulaValue, formulaAwareTransform);
        }

        return formulaAwareTransform.applyToValue(value);
      })
    )
  };
}

function buildExcelHeaderLookup(values, kindLabel) {
  const headerRow = Array.isArray(values) && Array.isArray(values[0]) ? values[0] : null;
  if (!headerRow) {
    throw new Error(`Excel host requires a header row for ${kindLabel}.`);
  }

  const exact = new Map();
  const lower = new Map();

  for (const [index, rawHeader] of headerRow.entries()) {
    const header = normalizeExcelHeaderText(rawHeader);
    if (!header) {
      continue;
    }

    if (exact.has(header)) {
      throw new Error(`Excel host found duplicate header: ${header}.`);
    }

    const lowerKey = header.toLocaleLowerCase();
    const existingHeader = lower.get(lowerKey);
    if (existingHeader && existingHeader !== header) {
      throw new Error(`Excel host found duplicate header: ${header}.`);
    }

    const entry = { header, columnIndex: index + 1 };
    exact.set(header, entry);
    lower.set(lowerKey, header);
  }

  if (exact.size === 0) {
    throw new Error(`Excel host requires a header row for ${kindLabel}.`);
  }

  return { exact, lower };
}

function resolveExcelHeaderEntry(headerLookup, requestedField, kindLabel) {
  const normalizedField = normalizeExcelHeaderText(requestedField);
  if (!normalizedField) {
    throw new Error(`Excel host cannot find ${kindLabel} field in header row: ${requestedField}.`);
  }

  const exactMatch = headerLookup.exact.get(normalizedField);
  if (exactMatch) {
    return exactMatch;
  }

  const lowerMatch = headerLookup.lower.get(normalizedField.toLocaleLowerCase());
  if (lowerMatch) {
    return headerLookup.exact.get(lowerMatch);
  }

  throw new Error(`Excel host cannot find ${kindLabel} field in header row: ${requestedField}.`);
}

function preflightExcelPivotTableStructure(plan) {
  const rowGroups = Array.isArray(plan.rowGroups) ? plan.rowGroups : [];
  const columnGroups = Array.isArray(plan.columnGroups) ? plan.columnGroups : [];
  const valueAggregations = Array.isArray(plan.valueAggregations) ? plan.valueAggregations : [];
  const filters = Array.isArray(plan.filters) ? plan.filters : [];

  if (!isSingleCellA1Anchor(plan.targetRange)) {
    throw new Error("Excel host requires a single-cell target anchor for pivot tables.");
  }

  const supportError = getExcelPivotStructureSupportError({
    ...plan,
    rowGroups,
    columnGroups,
    valueAggregations,
    filters
  });
  if (supportError) {
    throw new Error(supportError);
  }

  valueAggregations.forEach((aggregation) => {
    mapPivotAggregationToExcel(aggregation?.aggregation);
  });

  return {
    rowGroups,
    columnGroups,
    valueAggregations,
    filters,
    sort: plan.sort
  };
}

function resolveExcelPivotPlanFields(headerLookup, planState) {
  return {
    rowGroups: planState.rowGroups.map((field) =>
      resolveExcelHeaderEntry(headerLookup, field, "pivot").header
    ),
    columnGroups: planState.columnGroups.map((field) =>
      resolveExcelHeaderEntry(headerLookup, field, "pivot").header
    ),
    valueAggregations: planState.valueAggregations.map((aggregation) => ({
      ...aggregation,
      field: resolveExcelHeaderEntry(headerLookup, aggregation?.field, "pivot").header
    })),
    filters: planState.filters.map((filter) => ({
      ...filter,
      field: resolveExcelHeaderEntry(headerLookup, filter?.field, "pivot").header
    })),
    sort: planState.sort
      ? {
        ...planState.sort,
        field: resolveExcelHeaderEntry(headerLookup, planState.sort?.field, "pivot").header
      }
      : undefined
  };
}

function getExcelPivotField(pivotTable, fieldName) {
  const hierarchy = pivotTable?.hierarchies?.getItem?.(fieldName);
  const fieldCollection = hierarchy?.fields;
  if (fieldCollection?.getItem) {
    return fieldCollection.getItem(fieldName);
  }

  if (typeof hierarchy?.getPivotField === "function") {
    return hierarchy.getPivotField(fieldName);
  }

  if (typeof hierarchy?.applyFilter === "function" || typeof hierarchy?.sortByLabels === "function" || typeof hierarchy?.sortByValues === "function") {
    return hierarchy;
  }

  throw new Error(`Excel host cannot access pivot field: ${fieldName}.`);
}

function buildExcelPivotLabelFilter(filter) {
  const labelFilter = {
    condition: mapPivotLabelFilterConditionToExcel(filter.operator)
  };
  const value = filter?.value == null ? "" : String(filter.value);
  const value2 = filter?.value2 == null ? "" : String(filter.value2);

  if (filter.operator === "between" || filter.operator === "not_between") {
    labelFilter.lowerBound = value;
    labelFilter.upperBound = value2;
    if (filter.operator === "not_between") {
      labelFilter.exclusive = true;
    }
    return labelFilter;
  }

  if (filter.operator === "not_equal_to") {
    labelFilter.comparator = value;
    labelFilter.exclusive = true;
    return labelFilter;
  }

  labelFilter.comparator = value;
  return labelFilter;
}

function applyExcelPivotFilters(pivotTable, resolvedFields) {
  const groupedFields = new Set([
    ...(resolvedFields.rowGroups || []),
    ...(resolvedFields.columnGroups || [])
  ]);

  (resolvedFields.filters || []).forEach((filter) => {
    const hierarchy = pivotTable?.hierarchies?.getItem?.(filter.field);
    if (!groupedFields.has(filter.field) && pivotTable?.filterHierarchies?.add && hierarchy) {
      pivotTable.filterHierarchies.add(hierarchy);
    }

    const pivotField = getExcelPivotField(pivotTable, filter.field);
    if (typeof pivotField?.applyFilter !== "function") {
      throw new Error(`Excel host cannot apply a pivot filter on ${filter.field}.`);
    }

    pivotField.applyFilter({
      labelFilter: buildExcelPivotLabelFilter(filter)
    });
  });
}

function resolveExcelPivotSortState(resolvedFields) {
  if (!resolvedFields.sort) {
    return null;
  }

  const rowGroups = resolvedFields.rowGroups || [];
  const columnGroups = resolvedFields.columnGroups || [];
  const valueAggregations = resolvedFields.valueAggregations || [];
  const groupedFields = new Set([...rowGroups, ...columnGroups]);
  const aggregatedFields = new Set(valueAggregations.map((aggregation) => aggregation.field));
  const direction = mapPivotSortDirectionToExcel(resolvedFields.sort.direction);

  if (resolvedFields.sort.sortOn === "group_field") {
    if (!groupedFields.has(resolvedFields.sort.field)) {
      throw new Error("Excel host can only sort an existing pivot group field.");
    }

    return {
      mode: "group_field",
      field: resolvedFields.sort.field,
      direction
    };
  }

  if (!aggregatedFields.has(resolvedFields.sort.field)) {
    throw new Error("Excel host can only sort by an existing pivot value field.");
  }

  if (rowGroups.length > 0 && columnGroups.length > 0) {
    throw new Error("Excel host can't sort pivot values when both row and column groups are present yet.");
  }

  return {
    mode: "aggregated_value",
    targetField: rowGroups.length > 0 ? rowGroups[rowGroups.length - 1] : columnGroups[columnGroups.length - 1],
    valueField: resolvedFields.sort.field,
    direction
  };
}

function applyExcelPivotSort(pivotTable, resolvedFields, createdDataHierarchies) {
  const sortState = resolveExcelPivotSortState(resolvedFields);
  if (!sortState) {
    return;
  }

  if (sortState.mode === "group_field") {
    const pivotField = getExcelPivotField(pivotTable, sortState.field);
    if (typeof pivotField?.sortByLabels !== "function") {
      throw new Error(`Excel host cannot sort pivot field ${sortState.field}.`);
    }

    pivotField.sortByLabels(sortState.direction);
    return;
  }

  const pivotField = getExcelPivotField(pivotTable, sortState.targetField);
  const dataHierarchy = createdDataHierarchies[sortState.valueField] || pivotTable?.dataHierarchies?.getItem?.(sortState.valueField);

  if (!dataHierarchy) {
    throw new Error(`Excel host cannot sort by pivot value ${sortState.valueField}.`);
  }

  if (typeof pivotField?.sortByValues !== "function") {
    throw new Error(`Excel host cannot sort pivot values on ${sortState.targetField}.`);
  }

  pivotField.sortByValues(sortState.direction, dataHierarchy, []);
}

function createExcelPivotTableName(plan) {
  const normalizedSheet = String(plan?.targetSheet || "Pivot")
    .replace(/[^A-Za-z0-9_]+/g, "_")
    .replace(/^_+|_+$/g, "");
  const normalizedTarget = normalizeA1(plan?.targetRange || "A1")
    .replace(/[^A-Za-z0-9_]+/g, "_")
    .replace(/^_+|_+$/g, "");
  const uniqueSuffix = generateClientUuid().replace(/[^A-Za-z0-9]/g, "").slice(0, 10);
  return `HermesPivot_${normalizedSheet || "Pivot"}_${normalizedTarget || "A1"}_${uniqueSuffix}`.slice(0, 128);
}

async function applyExcelPivotTablePlan({ context, worksheets, platform, plan }) {
  const planState = preflightExcelPivotTableStructure(plan);
  const sourceWorksheet = worksheets.getItem(plan.sourceSheet);
  const targetWorksheet = worksheets.getItem(plan.targetSheet);
  const sourceRange = sourceWorksheet.getRange(plan.sourceRange);
  const anchorRange = targetWorksheet.getRange(plan.targetRange);

  sourceRange.load(["values"]);
  anchorRange.load(["rowCount", "columnCount", "values", "formulas"]);
  await context.sync();

  if (anchorRange.rowCount !== 1 || anchorRange.columnCount !== 1) {
    throw new Error("Excel host requires a single-cell target anchor for pivot tables.");
  }

  if (rangeHasExistingContent(anchorRange.values) || hasAnyRealFormula(anchorRange.formulas)) {
    throw new Error("Target range already contains content.");
  }

  const headerLookup = buildExcelHeaderLookup(sourceRange.values, "pivot tables");
  const resolvedFields = resolveExcelPivotPlanFields(headerLookup, planState);
  const pivotTableCollection = targetWorksheet?.pivotTables || context.workbook?.pivotTables;

  if (!pivotTableCollection || typeof pivotTableCollection.add !== "function") {
    throw new Error("Excel host does not expose pivot creation on this range.");
  }

  const pivotTable = pivotTableCollection.add(
    createExcelPivotTableName(plan),
    sourceRange,
    anchorRange
  );
  await context.sync();

  const createdDataHierarchies = {};

  resolvedFields.rowGroups.forEach((field) => {
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
  });

  resolvedFields.columnGroups.forEach((field) => {
    pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field));
  });

  resolvedFields.valueAggregations.forEach((aggregation) => {
    const dataHierarchy = pivotTable.dataHierarchies.add(
      pivotTable.hierarchies.getItem(aggregation.field)
    );
    dataHierarchy.summarizeBy = mapPivotAggregationToExcel(aggregation.aggregation);
    createdDataHierarchies[aggregation.field] = dataHierarchy;
  });

  applyExcelPivotFilters(pivotTable, resolvedFields);
  applyExcelPivotSort(pivotTable, resolvedFields, createdDataHierarchies);

  await context.sync();

  return {
    kind: "pivot_table_update",
    operation: "pivot_table_update",
    hostPlatform: platform,
    ...plan,
    summary: getPivotTableStatusSummary(plan)
  };
}

function getActualAppendTargetRange(targetRangeAddress, startRowOffset, rowCount, columnCount) {
  const bounds = parseA1RangeReference(normalizeA1(targetRangeAddress));
  return buildA1RangeFromBounds({
    startRow: bounds.startRow + startRowOffset,
    endRow: bounds.startRow + startRowOffset + rowCount - 1,
    startColumn: bounds.startColumn,
    endColumn: bounds.startColumn + columnCount - 1
  });
}

function applyCleanupTransform(plan, inputValues, hostLabel = "Excel host") {
  const values = cloneMatrix(inputValues);
  const formulaAwareTransform = getFormulaAwareCleanupTransform(plan, hostLabel);

  if (formulaAwareTransform) {
    return values.map((row) =>
      row.map((value) => formulaAwareTransform.applyToValue(value))
    );
  }

  switch (plan.operation) {
    case "remove_blank_rows": {
      const keyOffsets = getCleanupColumnOffsets(plan);
      const retainedRows = values.filter((row) => {
        const candidateValues = keyOffsets.length > 0
          ? keyOffsets.map((index) => row[index])
          : row;
        return candidateValues.some((value) => !isBlankCellValue(value));
      });
      return fillTrailingBlankRows(retainedRows, values[0]?.length || 0, values.length);
    }
    case "remove_duplicate_rows": {
      const keyOffsets = getCleanupColumnOffsets(plan);
      const seen = new Set();
      const retainedRows = [];

      for (const row of values) {
        const keyValues = keyOffsets.length > 0 ? keyOffsets.map((index) => row[index]) : row;
        const digest = JSON.stringify(keyValues);
        if (seen.has(digest)) {
          continue;
        }
        seen.add(digest);
        retainedRows.push(row);
      }

      return fillTrailingBlankRows(retainedRows, values[0]?.length || 0, values.length);
    }
    case "split_column": {
      const { source, targetStart } = getCleanupColumnOffsets(plan);
      const targetCapacity = values[0]?.length - targetStart;
      return values.map((row) => {
        const parts = String(row[source] ?? "").split(plan.delimiter);
        if (parts.length > targetCapacity) {
          throw new Error("Excel host cannot split this column exactly within the approved target range.");
        }

        const nextRow = [...row];
        for (let offset = 0; offset < targetCapacity; offset += 1) {
          nextRow[targetStart + offset] = parts[offset] ?? "";
        }
        return nextRow;
      });
    }
    case "join_columns": {
      const { source, target } = getCleanupColumnOffsets(plan);
      return values.map((row) => {
        const nextRow = [...row];
        nextRow[target] = source.map((index) => String(row[index] ?? "")).join(plan.delimiter);
        return nextRow;
      });
    }
    case "fill_down": {
      const explicitOffsets = getCleanupColumnOffsets(plan);
      const targetOffsets = explicitOffsets.length > 0
        ? explicitOffsets
        : Array.from({ length: values[0]?.length || 0 }, (_, index) => index);
      const nextValues = cloneMatrix(values);

      for (const columnIndex of targetOffsets) {
        let lastSeen = null;
        for (let rowIndex = 0; rowIndex < nextValues.length; rowIndex += 1) {
          const currentValue = nextValues[rowIndex][columnIndex];
          if (isBlankCellValue(currentValue)) {
            if (lastSeen !== null) {
              nextValues[rowIndex][columnIndex] = lastSeen;
            }
          } else {
            lastSeen = currentValue;
          }
        }
      }

      return nextValues;
    }
    case "standardize_format":
      return standardizeFormatMatrixExact(plan, values, "Excel host");
    default:
      throw new Error(`Excel host does not support exact-safe cleanup semantics for ${plan.operation}.`);
  }
}

function mapConditionalFormatComparatorToExcel(comparator) {
  const operator = Excel?.ConditionalCellValueOperator;
  const map = {
    between: operator?.between || "between",
    not_between: operator?.notBetween || "notBetween",
    equal_to: operator?.equalTo || "equalTo",
    not_equal_to: operator?.notEqualTo || "notEqualTo",
    greater_than: operator?.greaterThan || "greaterThan",
    greater_than_or_equal_to: operator?.greaterThanOrEqual || "greaterThanOrEqual",
    less_than: operator?.lessThan || "lessThan",
    less_than_or_equal_to: operator?.lessThanOrEqual || "lessThanOrEqual"
  };

  return map[comparator];
}

function mapConditionalColorScalePointTypeToExcel(pointType) {
  const pointTypes = Excel?.ConditionalFormatColorCriterionType;
  switch (pointType) {
    case "min":
      return pointTypes?.lowestValue || "lowestValue";
    case "max":
      return pointTypes?.highestValue || "highestValue";
    case "number":
      return pointTypes?.number || "number";
    case "percent":
      return pointTypes?.percent || "percent";
    case "percentile":
      return pointTypes?.percentile || "percentile";
    default:
      throw new Error(`Excel host does not support exact color_scale point type ${pointType}.`);
  }
}

function buildExcelConditionalColorScaleCriterion(point) {
  if (!point || typeof point !== "object") {
    throw new Error("Excel host requires valid color_scale points.");
  }

  const type = mapConditionalColorScalePointTypeToExcel(point.type);
  const requiresFormula = point.type === "number" || point.type === "percent" || point.type === "percentile";

  if (requiresFormula && typeof point.value !== "number") {
    throw new Error(`Excel host requires numeric value for color_scale point type ${point.type}.`);
  }

  return {
    formula: requiresFormula ? String(point.value) : null,
    type,
    color: point.color
  };
}

function buildExcelConditionalColorScaleCriteria(plan) {
  if (!Array.isArray(plan.points) || plan.points.length < 2 || plan.points.length > 3) {
    throw new Error("Excel host requires 2 or 3 color_scale points.");
  }

  const criteria = {
    minimum: buildExcelConditionalColorScaleCriterion(plan.points[0]),
    maximum: buildExcelConditionalColorScaleCriterion(plan.points[plan.points.length - 1])
  };

  if (plan.points.length === 3) {
    criteria.midpoint = buildExcelConditionalColorScaleCriterion(plan.points[1]);
  }

  return criteria;
}

function assignConditionalRule(target, values) {
  const rule = target && typeof target === "object" ? target : {};
  Object.assign(rule, values);
  return rule;
}

function getUnsupportedConditionalFormatStyleFields(style) {
  if (!style) {
    return [];
  }

  return [
    "numberFormat"
  ].filter((field) => style[field] !== undefined);
}

function validateConditionalFormatStyle(style) {
  const unsupportedFields = [
    ...getUnsupportedConditionalFormatStyleFields(style)
  ];

  if (unsupportedFields.length > 0) {
    throw new Error(
      `Excel host does not support exact conditional-format style mapping for fields: ${unsupportedFields.join(", ")}.`
    );
  }
}

function applyConditionalFormatStyle(format, style) {
  if (!format || !style) {
    return;
  }

  if (style.backgroundColor && format.fill) {
    format.fill.color = style.backgroundColor;
  }

  if (style.textColor && format.font) {
    format.font.color = style.textColor;
  }

  if (typeof style.bold === "boolean" && format.font) {
    format.font.bold = style.bold;
  }

  if (typeof style.italic === "boolean" && format.font) {
    format.font.italic = style.italic;
  }

  if (typeof style.underline === "boolean" && format.font) {
    format.font.underline = style.underline
      ? (Excel?.RangeUnderlineStyle?.single || "Single")
      : (Excel?.RangeUnderlineStyle?.none || "None");
  }

  if (typeof style.strikethrough === "boolean" && format.font) {
    format.font.strikethrough = style.strikethrough;
  }
}

function resolveExcelConditionalFormatBinding(plan) {
  const conditionalType = Excel?.ConditionalFormatType;

  switch (plan.ruleType) {
    case "single_color":
    case "number_compare":
    case "date_compare":
      return {
        type: conditionalType?.cellValue || "cellValue",
        property: "cellValue"
      };
    case "text_contains":
      return {
        type: conditionalType?.containsText || "containsText",
        property: "containsText"
      };
    case "duplicate_values":
      return {
        type: conditionalType?.duplicateValues || "duplicateValues",
        property: "duplicateValues"
      };
    case "custom_formula":
      return {
        type: conditionalType?.custom || "custom",
        property: "custom"
      };
    case "top_n":
      return {
        type: conditionalType?.topBottom || "topBottom",
        property: "topBottom"
      };
    case "average_compare":
      return {
        type: conditionalType?.aboveAverage || "aboveAverage",
        property: "aboveAverage"
      };
    case "color_scale":
      return {
        type: conditionalType?.colorScale || "colorScale",
        property: "colorScale"
      };
    default:
      throw new Error(`Excel host does not support exact conditional-format mapping for ruleType ${plan.ruleType}.`);
  }
}

function applyExcelConditionalFormatRule(configuration, plan) {
  switch (plan.ruleType) {
    case "single_color":
    case "number_compare":
    case "date_compare": {
      const operator = mapConditionalFormatComparatorToExcel(plan.comparator);
      if (!operator) {
        throw new Error(`Excel host does not support exact conditional-format mapping for comparator ${plan.comparator}.`);
      }

      configuration.rule = assignConditionalRule(configuration.rule, {
        operator,
        formula1: plan.value,
        formula2: plan.value2
      });
      return;
    }
    case "text_contains":
      configuration.rule = assignConditionalRule(configuration.rule, {
        text: plan.text
      });
      return;
    case "duplicate_values":
      return;
    case "custom_formula":
      configuration.rule = assignConditionalRule(configuration.rule, {
        formula: plan.formula
      });
      return;
    case "top_n":
      configuration.rule = assignConditionalRule(configuration.rule, {
        rank: plan.rank,
        type: plan.direction === "bottom"
          ? (Excel?.ConditionalTopBottomCriterionType?.bottomItems || "bottomItems")
          : (Excel?.ConditionalTopBottomCriterionType?.topItems || "topItems")
      });
      return;
    case "average_compare":
      configuration.rule = assignConditionalRule(configuration.rule, {
        criterion: plan.direction === "below"
          ? (Excel?.ConditionalAverageCriterion?.belowAverage || "belowAverage")
          : (Excel?.ConditionalAverageCriterion?.aboveAverage || "aboveAverage")
      });
      return;
    case "color_scale":
      if (!("criteria" in configuration)) {
        throw new Error("Excel host does not expose colorScale conditional-format criteria.");
      }
      configuration.criteria = buildExcelConditionalColorScaleCriteria(plan);
      return;
    default:
      throw new Error(`Excel host does not support exact conditional-format mapping for ruleType ${plan.ruleType}.`);
  }
}

function applyExcelConditionalFormat(target, plan) {
  const conditionalFormats = target?.conditionalFormats;
  if (!conditionalFormats) {
    throw new Error("Excel host does not support conditional formatting on this range.");
  }

  if ((plan.managementMode === "replace_all_on_target" || plan.managementMode === "clear_on_target") &&
    typeof conditionalFormats.clearAll !== "function") {
    throw new Error("Excel host does not support clearing conditional formatting on this range.");
  }

  if (plan.managementMode === "replace_all_on_target" || plan.managementMode === "clear_on_target") {
    conditionalFormats.clearAll();
  }

  if (plan.managementMode === "clear_on_target") {
    return;
  }

  if (typeof conditionalFormats.add !== "function") {
    throw new Error("Excel host does not support adding conditional formatting on this range.");
  }

  validateConditionalFormatStyle(plan.style);

  const binding = resolveExcelConditionalFormatBinding(plan);
  const conditionalFormat = conditionalFormats.add(binding.type);
  const configuration = conditionalFormat?.[binding.property];
  if (!configuration) {
    throw new Error(`Excel host does not expose ${binding.property} conditional-format configuration.`);
  }

  applyExcelConditionalFormatRule(configuration, plan);
  applyConditionalFormatStyle(configuration.format, plan.style);
}

function getExcelChartHeaderSequence(sourceRange) {
  const headerRow = Array.isArray(sourceRange?.values?.[0]) ? sourceRange.values[0] : [];
  return headerRow.map((value) => normalizeExcelChartField(value));
}

function validateExcelChartSourceLayout(sourceRange, plan) {
  if (!sourceRange || typeof sourceRange.rowCount !== "number" || typeof sourceRange.columnCount !== "number") {
    throw new Error("Excel host requires a valid source range for charts.");
  }

  if (sourceRange.rowCount < 2) {
    throw new Error("Excel host requires at least one data row for exact-safe chart creation.");
  }

  const expectedFields = getExcelChartFieldSequence(plan);
  const headerFields = getExcelChartHeaderSequence(sourceRange);
  const layoutMatchesExactly =
    headerFields.length === expectedFields.length &&
    expectedFields.every((field, index) => headerFields[index] === field);

  if (!layoutMatchesExactly) {
    throw new Error("Excel host requires sourceRange header row to match categoryField and series fields exactly for chart creation.");
  }
}

function validateExcelChartTargetAnchor(targetRange) {
  if (!targetRange || typeof targetRange.rowCount !== "number" || typeof targetRange.columnCount !== "number") {
    throw new Error("Excel host requires a valid target range for charts.");
  }

  if (targetRange.rowCount !== 1 || targetRange.columnCount !== 1) {
    throw new Error("Excel host requires a single-cell target anchor for charts.");
  }
}

function applyExcelChartTitle(chart, title) {
  if (!title) {
    return;
  }

  if (!chart?.title || typeof chart.title !== "object" || !("text" in chart.title)) {
    throw new Error("Excel host does not support exact-safe chart title options.");
  }

  if ("visible" in chart.title) {
    chart.title.visible = true;
  }
  chart.title.text = title;
}

function applyExcelChartLegend(chart, legendPosition) {
  const legendConfig = getExcelChartLegendConfig(legendPosition);
  if (!legendConfig) {
    return;
  }

  if (!chart?.legend || typeof chart.legend !== "object") {
    throw new Error("Excel host does not support exact-safe chart legend positioning.");
  }

  if ("visible" in chart.legend) {
    chart.legend.visible = legendConfig.visible;
  }

  if (legendConfig.position) {
    if (!("position" in chart.legend)) {
      throw new Error("Excel host does not support exact-safe chart legend positioning.");
    }
    chart.legend.position = legendConfig.position;
  }
}

function applyExcelChartAxisTitle(axis, title, axisLabel) {
  if (typeof title !== "string" || title.trim().length === 0) {
    return;
  }

  if (!axis?.title || typeof axis.title !== "object" || !("text" in axis.title)) {
    throw new Error(`Excel host does not support exact-safe ${axisLabel} axis titles.`);
  }

  if ("visible" in axis.title) {
    axis.title.visible = true;
  }
  axis.title.text = title.trim();
}

function applyExcelChartAxisTitles(chart, plan) {
  if (!hasExcelChartAxisTitles(plan)) {
    return;
  }

  if (plan?.chartType === "pie") {
    throw new Error("Excel host does not support axis titles for pie charts.");
  }

  if (!chart?.axes || typeof chart.axes !== "object") {
    throw new Error("Excel host does not support exact-safe chart axis titles.");
  }

  applyExcelChartAxisTitle(chart.axes.categoryAxis, plan.horizontalAxisTitle, "horizontal");
  applyExcelChartAxisTitle(chart.axes.valueAxis, plan.verticalAxisTitle, "vertical");
}

function getExcelChartCustomSeriesLabels(plan) {
  if (!Array.isArray(plan?.series)) {
    return [];
  }

  return plan.series
    .map((series, index) => ({
      index,
      field: normalizeExcelChartField(series?.field),
      label: typeof series?.label === "string" ? series.label.trim() : ""
    }))
    .filter((series) => series.label && series.label !== series.field);
}

function applyExcelChartSeriesLabels(chart, plan) {
  const customLabels = getExcelChartCustomSeriesLabels(plan);
  if (customLabels.length === 0) {
    return;
  }

  if (!chart?.series || typeof chart.series.getItemAt !== "function") {
    throw new Error("Excel host does not support exact-safe chart series labels.");
  }

  customLabels.forEach((series) => {
    const chartSeries = chart.series.getItemAt(series.index);
    if (!chartSeries || typeof chartSeries !== "object" || !("name" in chartSeries)) {
      throw new Error("Excel host does not support exact-safe chart series labels.");
    }

    chartSeries.name = series.label;
  });
}

async function applyExcelChartPlan(context, worksheets, plan, platform) {
  assertExcelChartPlanSupport(plan);

  const sourceWorksheet = worksheets.getItem(plan.sourceSheet);
  const targetWorksheet = worksheets.getItem(plan.targetSheet);
  const sourceRange = sourceWorksheet.getRange(plan.sourceRange);
  const targetRange = targetWorksheet.getRange(plan.targetRange);

  sourceRange.load(["values", "rowCount", "columnCount"]);
  targetRange.load(["rowCount", "columnCount"]);
  await context.sync();

  validateExcelChartSourceLayout(sourceRange, plan);
  validateExcelChartTargetAnchor(targetRange);

  const charts = targetWorksheet.charts;
  if (!charts?.add) {
    throw new Error("Excel host does not support exact-safe chart creation yet.");
  }

  const chart = charts.add(
    getExcelChartTypeConfig(plan.chartType).chartType,
    sourceRange,
    Excel?.ChartSeriesBy?.columns || "Columns"
  );

  if (!chart || typeof chart.setPosition !== "function") {
    throw new Error("Excel host does not support exact-safe chart creation yet.");
  }

  chart.setPosition(targetRange);
  applyExcelChartTitle(chart, plan.title);
  applyExcelChartLegend(chart, plan.legendPosition);
  applyExcelChartAxisTitles(chart, plan);
  applyExcelChartSeriesLabels(chart, plan);
  await context.sync();

  return {
    kind: "chart_update",
    operation: "chart_update",
    hostPlatform: platform,
    sourceSheet: plan.sourceSheet,
    sourceRange: plan.sourceRange,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    chartType: plan.chartType,
    categoryField: plan.categoryField,
    series: Array.isArray(plan.series)
      ? plan.series.map((series) => {
          const nextSeries = { field: series.field };
          if (typeof series.label === "string" && series.label.trim().length > 0) {
            nextSeries.label = series.label.trim();
          }
          return nextSeries;
        })
      : [],
    title: plan.title,
    legendPosition: plan.legendPosition,
    horizontalAxisTitle: plan.horizontalAxisTitle,
    verticalAxisTitle: plan.verticalAxisTitle,
    explanation: plan.explanation,
    confidence: plan.confidence,
    requiresConfirmation: plan.requiresConfirmation,
    affectedRanges: plan.affectedRanges,
    overwriteRisk: plan.overwriteRisk,
    confirmationLevel: plan.confirmationLevel,
    summary: getChartStatusSummary(plan)
  };
}

function assertExcelTablePlanSupport(plan) {
  const supportError = getExcelPlanSupportError({
    kind: "table_plan",
    ...plan
  });
  if (supportError) {
    throw new Error(supportError);
  }
}

function setExcelTableProperty(table, propertyName, value, supportMessage) {
  if (value === undefined) {
    return;
  }

  if (!table || !(propertyName in table)) {
    throw new Error(supportMessage);
  }

  table[propertyName] = value;
}

async function applyExcelTablePlan(context, worksheets, plan, platform) {
  assertExcelTablePlanSupport(plan);

  const worksheet = worksheets.getItem(plan.targetSheet);
  const targetRange = worksheet.getRange(plan.targetRange);
  targetRange.load(["rowCount", "columnCount"]);
  await context.sync();

  if (targetRange.rowCount < 1 || targetRange.columnCount < 1) {
    throw new Error("Excel host requires a non-empty range for table creation.");
  }

  if (!worksheet.tables || typeof worksheet.tables.add !== "function") {
    throw new Error("Excel host does not support exact-safe native table creation.");
  }

  const table = worksheet.tables.add(targetRange, plan.hasHeaders);
  if (!table) {
    throw new Error("Excel host did not create a native table.");
  }

  if (plan.name) {
    setExcelTableProperty(
      table,
      "name",
      plan.name,
      "Excel host does not support exact-safe table naming."
    );
  }

  if (plan.styleName) {
    setExcelTableProperty(
      table,
      "style",
      plan.styleName,
      "Excel host does not support exact-safe table styles."
    );
  }

  setExcelTableProperty(
    table,
    "showBandedRows",
    plan.showBandedRows,
    "Excel host does not support exact-safe table banded rows."
  );
  setExcelTableProperty(
    table,
    "showBandedColumns",
    plan.showBandedColumns,
    "Excel host does not support exact-safe table banded columns."
  );
  setExcelTableProperty(
    table,
    "showFilterButton",
    plan.showFilterButton,
    "Excel host does not support exact-safe table filter buttons."
  );
  setExcelTableProperty(
    table,
    "showTotals",
    plan.showTotalsRow,
    "Excel host does not support exact-safe table totals rows."
  );

  await context.sync();

  return {
    kind: "table_update",
    operation: "table_update",
    hostPlatform: platform,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    name: plan.name,
    hasHeaders: plan.hasHeaders,
    styleName: plan.styleName,
    showBandedRows: plan.showBandedRows,
    showBandedColumns: plan.showBandedColumns,
    showFilterButton: plan.showFilterButton,
    showTotalsRow: plan.showTotalsRow,
    summary: getTableStatusSummary(plan)
  };
}

function getStructuredPreview(response) {
  if (!response || typeof response !== "object" || typeof response.type !== "string") {
    return null;
  }

  switch (response.type) {
    case "formula":
      return {
        kind: "formula",
        intent: response.data.intent,
        formula: response.data.formula,
        formulaLanguage: response.data.formulaLanguage,
        targetCell: response.data.targetCell,
        explanation: response.data.explanation,
        alternateFormulas: response.data.alternateFormulas || []
      };
    case "workbook_structure_update":
      return {
        kind: "workbook_structure_update",
        operation: response.data.operation,
        sheetName: response.data.sheetName,
        position: "position" in response.data ? response.data.position : undefined,
        newSheetName: "newSheetName" in response.data ? response.data.newSheetName : undefined,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        overwriteRisk: response.data.overwriteRisk
      };
    case "sheet_structure_update":
      return {
        kind: "sheet_structure_update",
        targetSheet: response.data.targetSheet,
        operation: response.data.operation,
        targetRange: response.data.targetRange,
        startIndex: response.data.startIndex,
        count: response.data.count,
        frozenRows: response.data.frozenRows,
        frozenColumns: response.data.frozenColumns,
        color: response.data.color,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        summary: getSheetStructureStatusSummary(response.data)
      };
    case "range_sort_plan":
      return {
        kind: "range_sort_plan",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        hasHeader: response.data.hasHeader,
        keys: response.data.keys,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        summary: getRangeSortStatusSummary(response.data)
      };
    case "range_filter_plan":
      return {
        kind: "range_filter_plan",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        hasHeader: response.data.hasHeader,
        conditions: response.data.conditions,
        combiner: response.data.combiner,
        clearExistingFilters: response.data.clearExistingFilters,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        summary: getRangeFilterStatusSummary(response.data)
      };
    case "range_format_update":
      return {
        kind: "range_format_update",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        format: response.data.format,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        overwriteRisk: response.data.overwriteRisk
      };
    case "conditional_format_plan":
      return {
        kind: "conditional_format_plan",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        managementMode: response.data.managementMode,
        ruleType: response.data.ruleType,
        comparator: response.data.comparator,
        value: response.data.value,
        value2: response.data.value2,
        text: response.data.text,
        formula: response.data.formula,
        rank: response.data.rank,
        direction: response.data.direction,
        points: response.data.points,
        style: response.data.style,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        replacesExistingRules: response.data.replacesExistingRules,
        summary: getConditionalFormatPreviewSummary(response.data)
      };
    case "data_validation_plan":
      return {
        kind: "data_validation_plan",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        ruleType: response.data.ruleType,
        comparator: response.data.comparator,
        value: response.data.value,
        value2: response.data.value2,
        showDropdown: response.data.showDropdown,
        formula: response.data.formula,
        checkedValue: response.data.checkedValue,
        uncheckedValue: response.data.uncheckedValue,
        values: response.data.values,
        sourceRange: response.data.sourceRange,
        namedRangeName: response.data.namedRangeName,
        allowBlank: response.data.allowBlank,
        invalidDataBehavior: response.data.invalidDataBehavior,
        helpText: response.data.helpText,
        inputTitle: response.data.inputTitle,
        inputMessage: response.data.inputMessage,
        errorTitle: response.data.errorTitle,
        errorMessage: response.data.errorMessage,
        replacesExistingValidation: response.data.replacesExistingValidation,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        summary: getDataValidationStatusSummary(response.data)
      };
    case "named_range_update":
      return {
        kind: "named_range_update",
        operation: response.data.operation,
        scope: response.data.scope,
        name: response.data.name,
        sheetName: response.data.sheetName,
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        newName: response.data.newName,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        summary: getNamedRangeStatusSummary(response.data)
      };
    case "range_transfer_plan":
      return {
        kind: "range_transfer_plan",
        sourceSheet: response.data.sourceSheet,
        sourceRange: response.data.sourceRange,
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        operation: response.data.operation,
        pasteMode: response.data.pasteMode,
        transpose: response.data.transpose,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        summary: getRangeTransferStatusSummary(response.data)
      };
    case "data_cleanup_plan":
      return {
        kind: "data_cleanup_plan",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        operation: response.data.operation,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        keyColumns: response.data.keyColumns,
        mode: response.data.mode,
        sourceColumn: response.data.sourceColumn,
        delimiter: response.data.delimiter,
        targetStartColumn: response.data.targetStartColumn,
        sourceColumns: response.data.sourceColumns,
        targetColumn: response.data.targetColumn,
        columns: response.data.columns,
        formatType: response.data.formatType,
        formatPattern: response.data.formatPattern,
        summary: getDataCleanupStatusSummary(response.data)
      };
    case "composite_plan":
      return {
        kind: "composite_plan",
        stepCount: response.data.steps.length,
        steps: response.data.steps.map((step) => ({
          ...buildCompositeStepPreview(step),
          plan: step.plan,
          supportError: getExcelPlanSupportError(step.plan)
        })),
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        reversible: response.data.reversible,
        dryRunRecommended: response.data.dryRunRecommended,
        dryRunRequired: response.data.dryRunRequired,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        summary: getCompositePreviewSummary(response.data)
      };
    case "analysis_report_plan":
      {
        const resolvedPlan = response.data.outputMode === "materialize_report"
          ? resolveMaterializedAnalysisReportPlan(response.data)
          : response.data;
      return {
        kind: "analysis_report_plan",
        sourceSheet: resolvedPlan.sourceSheet,
        sourceRange: resolvedPlan.sourceRange,
        outputMode: resolvedPlan.outputMode,
        targetSheet: resolvedPlan.targetSheet,
        targetRange: resolvedPlan.targetRange,
        sections: resolvedPlan.sections,
        explanation: resolvedPlan.explanation,
        confidence: resolvedPlan.confidence,
        requiresConfirmation: getRequiresConfirmation(response),
        affectedRanges: resolvedPlan.affectedRanges,
        overwriteRisk: resolvedPlan.overwriteRisk,
        confirmationLevel: resolvedPlan.confirmationLevel,
        summary: getAnalysisReportPreviewSummary(resolvedPlan),
        rows: buildAnalysisReportMatrix(resolvedPlan)
      };
      }
    case "pivot_table_plan":
      return {
        kind: "pivot_table_plan",
        sourceSheet: response.data.sourceSheet,
        sourceRange: response.data.sourceRange,
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        rowGroups: response.data.rowGroups,
        columnGroups: response.data.columnGroups || [],
        valueAggregations: response.data.valueAggregations,
        filters: response.data.filters || [],
        sort: response.data.sort,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        summary: getPivotTablePreviewSummary(response.data)
      };
    case "chart_plan":
      return {
        kind: "chart_plan",
        sourceSheet: response.data.sourceSheet,
        sourceRange: response.data.sourceRange,
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        chartType: response.data.chartType,
        categoryField: response.data.categoryField,
        series: response.data.series,
        title: response.data.title,
        legendPosition: response.data.legendPosition,
        horizontalAxisTitle: response.data.horizontalAxisTitle,
        verticalAxisTitle: response.data.verticalAxisTitle,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        summary: getChartPreviewSummary(response.data)
      };
    case "table_plan":
      return {
        kind: "table_plan",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        name: response.data.name,
        hasHeaders: response.data.hasHeaders,
        styleName: response.data.styleName,
        showBandedRows: response.data.showBandedRows,
        showBandedColumns: response.data.showBandedColumns,
        showFilterButton: response.data.showFilterButton,
        showTotalsRow: response.data.showTotalsRow,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        summary: getTablePreviewSummary(response.data)
      };
    case "external_data_plan":
      return {
        kind: "external_data_plan",
        sourceType: response.data.sourceType,
        provider: response.data.provider,
        query: response.data.query,
        sourceUrl: response.data.sourceUrl,
        selectorType: response.data.selectorType,
        selector: response.data.selector,
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        formula: response.data.formula,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        affectedRanges: response.data.affectedRanges,
        overwriteRisk: response.data.overwriteRisk,
        confirmationLevel: response.data.confirmationLevel,
        summary: response.data.sourceType === "market_data"
          ? `Would anchor market data for ${response.data.query.symbol} at ${response.data.targetSheet}!${response.data.targetRange}.`
          : `Would anchor a ${String(response.data.provider).toUpperCase()} import at ${response.data.targetSheet}!${response.data.targetRange}.`
      };
    case "sheet_update": {
      const matrices = [
        response.data.values ? "values" : null,
        response.data.formulas ? "formulas" : null,
        response.data.notes ? "notes" : null
      ].filter(Boolean);

      return {
        kind: "sheet_update",
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        shape: response.data.shape,
        operation: response.data.operation,
        explanation: response.data.explanation,
        confidence: response.data.confidence,
        requiresConfirmation: response.data.requiresConfirmation,
        overwriteRisk: response.data.overwriteRisk,
        matrixKind: matrices.length > 1 ? "mixed_update" : (matrices[0] || "values"),
        headers: [],
        rows: buildWriteMatrix(response.data)
      };
    }
    case "sheet_import_plan":
      return {
        kind: "sheet_import_plan",
        sourceAttachmentId: response.data.sourceAttachmentId,
        targetSheet: response.data.targetSheet,
        targetRange: response.data.targetRange,
        shape: response.data.shape,
        extractionMode: response.data.extractionMode,
        confidence: response.data.confidence,
        warnings: response.data.warnings,
        requiresConfirmation: response.data.requiresConfirmation,
        headers: response.data.headers,
        rows: response.data.values
      };
    case "attachment_analysis":
      return {
        kind: "attachment_analysis",
        sourceAttachmentId: response.data.sourceAttachmentId,
        contentKind: response.data.contentKind,
        extractionMode: response.data.extractionMode,
        summary: response.data.summary
      };
    case "extracted_table":
      return {
        kind: "extracted_table",
        sourceAttachmentId: response.data.sourceAttachmentId,
        extractionMode: response.data.extractionMode,
        headers: response.data.headers,
        rows: response.data.rows,
        shape: response.data.shape || {
          rows: response.data.rows.length,
          columns: response.data.headers.length > 0
            ? response.data.headers.length
            : Math.max(...response.data.rows.map((row) => row.length), 0)
        }
      };
    case "document_summary":
      return {
        kind: "document_summary",
        sourceAttachmentId: response.data.sourceAttachmentId,
        contentKind: response.data.contentKind,
        extractionMode: response.data.extractionMode,
        keyPoints: response.data.keyPoints || []
      };
    default:
      return null;
  }
}

async function parseJson(response) {
  if (!response.ok) {
    let message = `Request failed with ${response.status}`;
    const responseUrl = response.url || `${gatewayBaseUrl}/(unknown)`;
    try {
      const payload = await response.json();
      const payloadMessage = payload?.error?.message || payload?.error || payload?.message || message;
      const payloadUserAction = payload?.error?.userAction || payload?.userAction;
      const formattedPayloadMessage = formatUserFacingErrorText(payloadMessage, payloadUserAction);
      if (
        response.status === 404 &&
        typeof payloadMessage === "string" &&
        payloadMessage.trim().toLowerCase() === "the requested resource doesn't exist."
      ) {
        message = formatUserFacingErrorText(
          `The Hermes request was sent to a service path that does not exist (${responseUrl}, HTTP ${response.status}).`,
          "Close Excel, reload the Hermes add-in, and retry. If it keeps happening, reinstall the live manifest so Hermes uses the correct gateway path."
        );
      } else if (containsSensitiveGatewayErrorText(formattedPayloadMessage)) {
        message = formatRawGatewayTextFailure(response.status);
      } else {
        message = formattedPayloadMessage;
        message = appendGatewayIssueSummary(message, payload?.error?.issues);
      }
    } catch {
      const text = await response.text();
      if (text) {
        const trimmed = text.trim();
        if (trimmed.toLowerCase() === "the requested resource doesn't exist.") {
          message = formatUserFacingErrorText(
            `The Hermes request was sent to a service path that does not exist (${responseUrl}, HTTP ${response.status}).`,
            "Close Excel, reload the Hermes add-in, and retry. If it keeps happening, reinstall the live manifest so Hermes uses the correct gateway path."
          );
        } else if (/^<!doctype html/i.test(trimmed) || /^<html/i.test(trimmed)) {
          message = formatUserFacingErrorText(
            `The Hermes request was sent to a page that is not the Hermes gateway (${responseUrl}, HTTP ${response.status}).`,
            "Close Excel, reopen the add-in, and retry. If it keeps happening, clear the add-in cache and reload Hermes."
          );
        } else if (containsSensitiveGatewayErrorText(trimmed)) {
          message = formatRawGatewayTextFailure(response.status);
        } else {
          message = text;
        }
      }
    }
    throw new Error(message);
  }

  try {
    return await response.json();
  } catch {
    throw new Error(
      "The Hermes service returned a response Excel could not use.\n\n" +
      "Retry the request, then reload the client if it keeps happening."
    );
  }
}

function containsSensitiveGatewayErrorText(text) {
  const value = String(text || "");
  return (
    /(?:client_secret|refresh_token|access_token|authorization|api[_-]?key|approval_secret|HERMES_[A-Z0-9_]+)/i.test(value) ||
    /\bBearer\s+[A-Za-z0-9._~+/-]+=*/i.test(value) ||
    /\bat\s+(?:file:\/\/\/|\/|[A-Za-z]:\\)/i.test(value) ||
    /(?:^|\s)\/(?:srv|var|tmp|root|home|Users|opt|workspace|app|mnt)\/[^\s]+(?::\d+)?/i.test(value) ||
    /(?:^|\s)[A-Za-z]:\\[^\s]+/.test(value) ||
    /https?:\/\/(?:localhost|127(?:\.\d{1,3}){3}|0\.0\.0\.0|10\.|192\.168\.|172\.(?:1[6-9]|2\d|3[01])\.|[^/\s]*internal|[^/\s]*\.local)(?:[/:]|\s|$)/i.test(value) ||
    /\b(?:stack trace|traceback)\b/i.test(value)
  );
}

function formatRawGatewayTextFailure(status) {
  return `Hermes gateway request failed with HTTP ${status}.`;
}

function assertWritebackCompletionAck(payload) {
  if (!payload || payload.ok !== true) {
    throw new Error(
      "The Hermes service returned a writeback completion response the client could not use.\n\n" +
      "Retry the writeback completion, then reload the client if it keeps happening."
    );
  }
}

function assertWritebackApprovalPayload(payload) {
  if (
    !payload ||
    typeof payload.approvalToken !== "string" ||
    payload.approvalToken.trim().length === 0 ||
    typeof payload.planDigest !== "string" ||
    payload.planDigest.trim().length === 0 ||
    typeof payload.executionId !== "string" ||
    payload.executionId.trim().length === 0
  ) {
    throw new Error(
      "The Hermes service returned a writeback approval response the client could not use.\n\n" +
      "Retry the approval, then reload the client if it keeps happening."
    );
  }
}

function getGatewayUrl(path) {
  if (!gatewayBaseUrl) {
    throw new Error("Hermes gateway URL is not configured.");
  }

  return `${gatewayBaseUrl}${String(path).startsWith("/") ? path : `/${path}`}`;
}

function encodeGatewayPathSegment(value) {
  return encodeURIComponent(String(value));
}

const gateway = {
  async uploadImage(input) {
    const form = new FormData();
    form.set("file", input.file, input.fileName);
    form.set("source", input.source);
    form.set("sessionId", sessionId);
    form.set("workbookId", getWorkbookIdentity());
    const payload = await parseJson(await fetch(getGatewayUrl("/api/uploads/image"), {
      method: "POST",
      body: form
    }));
    return payload.attachment;
  },

  async startRun(request) {
    return parseJson(await fetch(getGatewayUrl("/api/requests"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(request)
    }));
  },

  async pollRun(runId, requestId) {
    const params = new URLSearchParams();
    if (requestId) {
      params.set("requestId", requestId);
    }
    params.set("sessionId", sessionId);

    return parseJson(await fetch(
      getGatewayUrl(`/api/requests/${encodeGatewayPathSegment(runId)}${params.size > 0 ? `?${params.toString()}` : ""}`)
    ));
  },

  async pollTrace(runId, after, requestId) {
    const params = new URLSearchParams({
      after: String(after)
    });
    if (requestId) {
      params.set("requestId", requestId);
    }
    params.set("sessionId", sessionId);

    return parseJson(await fetch(getGatewayUrl(`/api/trace/${encodeGatewayPathSegment(runId)}?${params.toString()}`)));
  },

  async approveWrite(input) {
    const payload = await parseJson(await fetch(getGatewayUrl("/api/writeback/approve"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(input)
    }));
    assertWritebackApprovalPayload(payload);
    return payload;
  },

  async completeWrite(input) {
    const payload = await parseJson(await fetch(getGatewayUrl("/api/writeback/complete"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(input)
    }));
    assertWritebackCompletionAck(payload);
    return payload;
  },

  async dryRunPlan(input) {
    return parseJson(await fetch(getGatewayUrl("/api/execution/dry-run"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(input)
    }));
  },

  async listPlanHistory(input) {
    const params = new URLSearchParams({
      workbookSessionKey: input.workbookSessionKey
    });
    if (input.sessionId) {
      params.set("sessionId", input.sessionId);
    }
    if (typeof input.limit === "number") {
      params.set("limit", String(input.limit));
    }
    if (input.cursor) {
      params.set("cursor", input.cursor);
    }
    return parseJson(await fetch(getGatewayUrl(`/api/execution/history?${params.toString()}`)));
  },

  async undoExecution(input) {
    return parseJson(await fetch(getGatewayUrl("/api/execution/undo"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(input)
    }));
  },

  async prepareUndoExecution(input) {
    return parseJson(await fetch(getGatewayUrl("/api/execution/undo/prepare"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(input)
    }));
  },

  async redoExecution(input) {
    return parseJson(await fetch(getGatewayUrl("/api/execution/redo"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(input)
    }));
  },

  async prepareRedoExecution(input) {
    return parseJson(await fetch(getGatewayUrl("/api/execution/redo/prepare"), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(input)
    }));
  }
};

async function dryRunCompositePlan(input) {
  return gateway.dryRunPlan({
    ...input,
    sessionId,
    workbookSessionKey: getWorkbookSessionKey()
  });
}

async function listExecutionHistory(input = {}) {
  return gateway.listPlanHistory({
    workbookSessionKey: getWorkbookSessionKey(),
    sessionId,
    cursor: input.cursor,
    limit: input.limit
  });
}

async function undoExecution(executionId) {
  const workbookSessionKey = getWorkbookSessionKey();
  const snapshot = getLocalExecutionSnapshot(workbookSessionKey, executionId);
  if (!snapshot) {
    throw new Error("That history entry no longer has an exact undo snapshot in this spreadsheet session.");
  }

  await validateLocalExecutionSnapshotForMode(snapshot, "undo");
  assertLocalExecutionSnapshotStoreWritable(workbookSessionKey, snapshot, "undo");
  const requestId = createRequestId();
  const request = {
    executionId,
    requestId,
    sessionId,
    workbookSessionKey
  };
  await gateway.prepareUndoExecution(request);
  await restoreLocalExecutionSnapshotForMode(snapshot, "undo");
  const result = await gateway.undoExecution(request);
  if (result?.executionId) {
    linkLocalExecutionSnapshot(workbookSessionKey, result.executionId, executionId);
  }
  return result;
}

async function redoExecution(executionId) {
  const workbookSessionKey = getWorkbookSessionKey();
  const snapshot = getLocalExecutionSnapshot(workbookSessionKey, executionId);
  if (!snapshot) {
    throw new Error("That history entry no longer has an exact redo snapshot in this spreadsheet session.");
  }

  await validateLocalExecutionSnapshotForMode(snapshot, "redo");
  assertLocalExecutionSnapshotStoreWritable(workbookSessionKey, snapshot, "redo");
  const requestId = createRequestId();
  const request = {
    executionId,
    requestId,
    sessionId,
    workbookSessionKey
  };
  await gateway.prepareRedoExecution(request);
  await restoreLocalExecutionSnapshotForMode(snapshot, "redo");
  const result = await gateway.redoExecution(request);
  if (result?.executionId) {
    linkLocalExecutionSnapshot(workbookSessionKey, result.executionId, executionId);
  }
  return result;
}

function getExecutionShortcutMode(prompt) {
  if (UNDO_PROMPT_PATTERN.test(prompt)) {
    return "undo";
  }

  if (REDO_PROMPT_PATTERN.test(prompt)) {
    return "redo";
  }

  return null;
}

function getLatestHistoryExecutionId(entries, mode) {
  if (!Array.isArray(entries)) {
    return null;
  }

  const matchingEntry = entries.find((entry) => mode === "undo"
    ? entry?.undoEligible === true
    : entry?.redoEligible === true);

  return typeof matchingEntry?.executionId === "string" && matchingEntry.executionId.trim().length > 0
    ? matchingEntry.executionId
    : null;
}

function buildUnderSpecifiedFollowUpMessage(messages) {
  const previousAssistantMessage = [...(messages || [])]
    .reverse()
    .find((message) => message?.role === "assistant");

  const followUpSuggestions = previousAssistantMessage?.response
    ? getFollowUpSuggestions(previousAssistantMessage.response)
    : [];

  if (followUpSuggestions.length > 0) {
    return `I still need a concrete next step, not just “yes”. Try one of these: ${followUpSuggestions.join(" | ")}`;
  }

  const recoveryAction = previousAssistantMessage?.response?.type === "error" &&
    typeof previousAssistantMessage.response?.data?.userAction === "string"
    ? previousAssistantMessage.response.data.userAction.trim()
    : "";

  if (recoveryAction) {
    return `I still need a concrete next step, not just “yes”. ${recoveryAction}`;
  }

  return "I still need the exact range, cell, sheet, or action you want me to apply.";
}

async function handleExecutionShortcut(mode, assistantMessage) {
  assistantMessage.statusLine = mode === "undo"
    ? "Looking up the latest reversible write..."
    : "Looking up the latest redo entry...";
  renderMessages();

  const history = await listExecutionHistory({ limit: EXECUTION_HISTORY_SHORTCUT_LIMIT });
  const executionId = getLatestHistoryExecutionId(history?.entries, mode);

  if (!executionId) {
    throw new Error(mode === "undo"
      ? "I can’t perform an exact undo for the latest write in this workbook session."
      : "I can’t perform an exact redo for the latest write in this workbook session.");
  }

  const result = mode === "undo"
    ? await undoExecution(executionId)
    : await redoExecution(executionId);

  applyWritebackResultToMessage(assistantMessage, result);
}

async function executeWritePlanMessage(message) {
  if (!message?.requestId || !message?.runId || !message?.response || !isWritePlanResponse(message.response)) {
    return false;
  }

  try {
    if (message.pendingCompletion) {
      await gateway.completeWrite(message.pendingCompletion);
      applyWritebackResultToMessage(message, message.pendingCompletion.result);
      return true;
    }

    const workbookSessionKey = getWorkbookSessionKey();
    const approvalRequest = buildWriteApprovalRequest({
      requestId: message.requestId,
      runId: message.runId,
      plan: message.response.data,
      workbookSessionKey
    });
    if (!approvalRequest) {
      message.statusLine = "Destructive write-back cancelled.";
      return true;
    }

    const approval = await gateway.approveWrite(approvalRequest);

    const result = await applyWritePlan({
      plan: approvalRequest.plan,
      requestId: message.requestId,
      runId: message.runId,
      approvalToken: approval.approvalToken,
      executionId: approval.executionId
    });
    const gatewayResult = prepareGatewayWritebackResult(
      result,
      approval.executionId,
      workbookSessionKey
    );
    message.pendingCompletion = buildPendingWritebackCompletionRequest(
      message,
      approval,
      gatewayResult,
      workbookSessionKey
    );

    await gateway.completeWrite(message.pendingCompletion);

    applyWritebackResultToMessage(message, gatewayResult);
  } catch (error) {
    message.statusLine = message.pendingCompletion
      ? getPendingCompletionRetryStatus()
      : sanitizeHostExecutionError(error, "Write-back failed.");
  }

  return true;
}

function renderAttachmentStrip() {
  elements.attachmentStrip.innerHTML = state.pendingAttachments.map((attachment, index) => `
    <div class="chip">
      <img src="${escapeHtmlAttribute(attachment.previewUrl)}" alt="" />
      <span>${escapeHtml(attachment.fileName)}</span>
      <span class="chip-status">${escapeHtml(attachment.status)}</span>
      <button type="button" data-remove-index="${index}">x</button>
    </div>
  `).join("");

  elements.attachmentStrip.querySelectorAll("[data-remove-index]").forEach((button) => {
    button.addEventListener("click", () => {
      const index = Number(button.getAttribute("data-remove-index"));
      const [removed] = state.pendingAttachments.splice(index, 1);
      if (removed?.previewUrl) {
        URL.revokeObjectURL(removed.previewUrl);
      }
      renderAttachmentStrip();
    });
  });
}

function renderWarnings(warnings) {
  return warnings.map((warning) => `
    <div class="warning-line">${escapeHtml(warning.message)}</div>
  `).join("");
}

function renderSuggestions(suggestions) {
  if (suggestions.length === 0) {
    return "";
  }

  return `
    <div class="suggestions">
      ${suggestions.map((suggestion, index) => `
        <button type="button" class="suggestion" data-suggestion-index="${index}">
          ${escapeHtml(suggestion)}
        </button>
      `).join("")}
    </div>
  `;
}

function renderTable(headers, rows) {
  const headMarkup = headers.length > 0 ? `
    <thead>
      <tr>${headers.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr>
    </thead>
  ` : "";

  return `
    <table>
      ${headMarkup}
      <tbody>
        ${rows.map((row) => `
          <tr>${row.map((cell) => `<td>${escapeHtml(cell == null ? "" : cell)}</td>`).join("")}</tr>
        `).join("")}
      </tbody>
    </table>
  `;
}

function formatValidationSource(preview) {
  if (typeof preview.namedRangeName === "string" && preview.namedRangeName.length > 0) {
    return `source named range ${preview.namedRangeName}`;
  }

  if (typeof preview.sourceRange === "string" && preview.sourceRange.length > 0) {
    return `source range ${preview.sourceRange}`;
  }

  if (Array.isArray(preview.values) && preview.values.length > 0) {
    return `values ${preview.values.join(", ")}`;
  }

  return "";
}

function formatValidationParameters(preview) {
  const parts = [];

  if (preview.showDropdown === true) {
    parts.push("dropdown shown");
  } else if (preview.showDropdown === false) {
    parts.push("dropdown hidden");
  }

  if (preview.value !== undefined) {
    parts.push(`value ${String(preview.value)}`);
  }

  if (preview.value2 !== undefined) {
    parts.push(`value2 ${String(preview.value2)}`);
  }

  return parts.join(" • ");
}

function formatCheckboxValues(preview) {
  if (preview.ruleType !== "checkbox") {
    return "";
  }

  const parts = [];

  if (preview.checkedValue !== undefined) {
    parts.push(`checked ${String(preview.checkedValue)}`);
  }

  if (preview.uncheckedValue !== undefined) {
    parts.push(`unchecked ${String(preview.uncheckedValue)}`);
  }

  return parts.join(" • ");
}

function renderStructuredPreview(response, message) {
  const preview = getStructuredPreview(response);
  if (!preview) {
    return "";
  }

  if (preview.kind === "formula") {
    return `
      <div class="preview">
        <div class="preview-meta">
          Formula${preview.intent ? ` • ${escapeHtml(preview.intent)}` : ""} • ${escapeHtml(preview.formulaLanguage)}
          ${preview.targetCell ? ` • ${escapeHtml(preview.targetCell)}` : ""}
        </div>
        <pre class="formula-block">${escapeHtml(preview.formula)}</pre>
      </div>
    `;
  }

  if (preview.kind === "workbook_structure_update") {
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.operation)}
          • ${escapeHtml(preview.sheetName)}
          ${preview.newSheetName ? ` • ${escapeHtml(preview.newSheetName)}` : ""}
          ${formatWorkbookPositionLabel(preview.position)}
          ${preview.overwriteRisk ? ` • overwrite ${escapeHtml(preview.overwriteRisk)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        ${getRequiresConfirmation(response) ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Workbook Update
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "sheet_structure_update") {
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.operation)}
          • ${escapeHtml(preview.targetSheet)}
          ${preview.targetRange ? ` • ${escapeHtml(preview.targetRange)}` : ""}
          ${preview.confirmationLevel ? ` • ${escapeHtml(preview.confirmationLevel)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${getRequiresConfirmation(response) ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Sheet Update
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "range_sort_plan") {
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${preview.hasHeader ? "header" : "no header"}
          • ${preview.keys.length} key(s)
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${getRequiresConfirmation(response) ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Sort
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "range_filter_plan") {
    const supportError = getExcelPlanSupportError(preview);

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${preview.hasHeader ? "header" : "no header"}
          • ${preview.conditions.length} condition(s)
          • ${escapeHtml(preview.combiner)}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Filter
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "range_format_update") {
    const formatFields = formatRangeFormatFields(preview.format);
    const supportError = getExcelPlanSupportError(preview);

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          ${formatFields ? ` • ${escapeHtml(formatFields)}` : ""}
          ${preview.overwriteRisk ? ` • overwrite ${escapeHtml(preview.overwriteRisk)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Format Update
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "conditional_format_plan") {
    const conditionalFields = formatConditionalFormatFields(preview);
    const conditionalDetails = formatConditionalFormatDetails(preview);
    const supportError = getExcelPlanSupportError(preview);

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          ${conditionalFields ? ` • ${escapeHtml(conditionalFields)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${conditionalDetails ? `<div class="preview-meta">${escapeHtml(conditionalDetails)}</div>` : ""}
        ${Array.isArray(preview.affectedRanges) && preview.affectedRanges.length > 0
          ? `<div class="preview-meta">${escapeHtml(preview.affectedRanges.join(", "))}</div>`
          : ""}
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Conditional Formatting
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "data_validation_plan") {
    const validationSource = formatValidationSource(preview);
    const validationParameters = formatValidationParameters(preview);
    const checkboxValues = formatCheckboxValues(preview);
    const supportError = getExcelPlanSupportError(preview);
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${escapeHtml(preview.ruleType)}
          ${preview.comparator ? ` • ${escapeHtml(preview.comparator)}` : ""}
          ${preview.allowBlank ? " • allow blank" : " • no blank"}
          • ${escapeHtml(preview.invalidDataBehavior)}
          ${preview.replacesExistingValidation ? " • replaces existing validation" : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        ${preview.formula ? `<div class="preview-meta">${escapeHtml(preview.formula)}</div>` : ""}
        ${validationParameters ? `<div class="preview-meta">${escapeHtml(validationParameters)}</div>` : ""}
        ${checkboxValues ? `<div class="preview-meta">${escapeHtml(checkboxValues)}</div>` : ""}
        ${preview.helpText ? `<div class="preview-meta">${escapeHtml(preview.helpText)}</div>` : ""}
        ${preview.inputTitle || preview.inputMessage
          ? `<div class="preview-meta">${escapeHtml([preview.inputTitle, preview.inputMessage].filter(Boolean).join(": "))}</div>`
          : ""}
        ${preview.errorTitle || preview.errorMessage
          ? `<div class="preview-meta">${escapeHtml([preview.errorTitle, preview.errorMessage].filter(Boolean).join(": "))}</div>`
          : ""}
        ${validationSource ? `<div class="preview-meta">${escapeHtml(validationSource)}</div>` : ""}
        ${Array.isArray(preview.affectedRanges) && preview.affectedRanges.length > 0
          ? `<div class="preview-meta">${escapeHtml(preview.affectedRanges.join(", "))}</div>`
          : ""}
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Validation
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "named_range_update") {
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.operation)}
          • ${escapeHtml(preview.name)}
          ${preview.scope ? ` • ${escapeHtml(preview.scope)}` : ""}
          ${preview.targetSheet && preview.targetRange ? ` • ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}` : ""}
          ${preview.newName ? ` • ${escapeHtml(preview.newName)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${getRequiresConfirmation(response) ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Named Range Update
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "range_transfer_plan") {
    const supportError = getExcelPlanSupportError(preview);
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.sourceSheet)}!${escapeHtml(preview.sourceRange)}
          → ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${escapeHtml(preview.operation)}
          • ${escapeHtml(preview.pasteMode)}
          • transpose ${preview.transpose ? "on" : "off"}
          ${preview.overwriteRisk ? ` • overwrite ${escapeHtml(preview.overwriteRisk)}` : ""}
          ${preview.confirmationLevel ? ` • ${escapeHtml(preview.confirmationLevel)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${Array.isArray(preview.affectedRanges) && preview.affectedRanges.length > 0
          ? `<div class="preview-meta">${escapeHtml(preview.affectedRanges.join(", "))}</div>`
          : ""}
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Transfer
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "data_cleanup_plan") {
    const cleanupDetails = [
      Array.isArray(preview.keyColumns) && preview.keyColumns.length > 0
        ? `keys ${preview.keyColumns.join(", ")}`
        : "",
      preview.mode ? `mode ${preview.mode}` : "",
      preview.sourceColumn ? `source ${preview.sourceColumn}` : "",
      preview.sourceColumns ? `sources ${preview.sourceColumns.join(", ")}` : "",
      preview.targetStartColumn ? `target start ${preview.targetStartColumn}` : "",
      preview.targetColumn ? `target ${preview.targetColumn}` : "",
      preview.columns ? `columns ${preview.columns.join(", ")}` : "",
      preview.delimiter ? `delimiter ${preview.delimiter}` : "",
      preview.formatType ? `format ${preview.formatType}` : "",
      preview.formatPattern ? `pattern ${preview.formatPattern}` : ""
    ].filter(Boolean).join(" • ");
    const supportError = getExcelPlanSupportError(preview);

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${escapeHtml(preview.operation)}
          ${preview.overwriteRisk ? ` • overwrite ${escapeHtml(preview.overwriteRisk)}` : ""}
          ${preview.confirmationLevel ? ` • ${escapeHtml(preview.confirmationLevel)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${cleanupDetails ? `<div class="preview-meta">${escapeHtml(cleanupDetails)}</div>` : ""}
        ${Array.isArray(preview.affectedRanges) && preview.affectedRanges.length > 0
          ? `<div class="preview-meta">${escapeHtml(preview.affectedRanges.join(", "))}</div>`
          : ""}
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Cleanup
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "composite_plan") {
    const workflowFlags = [
      preview.reversible ? "reversible" : "non-reversible",
      preview.dryRunRequired ? "dry run required" : "",
      preview.dryRunRecommended ? "dry run recommended" : "",
      preview.confirmationLevel ? preview.confirmationLevel : ""
    ].filter(Boolean).join(" • ");
    const unsupportedSteps = getCompositePreviewSupportErrors(preview.steps);
    const compositeSupportError = unsupportedSteps.length > 0
      ? "Some workflow steps can't run in this Excel runtime yet. Review the flagged steps and simplify the workflow before confirming."
      : "";

    return `
      <div class="preview">
        <div class="preview-meta">
          ${preview.stepCount} step(s)
          ${workflowFlags ? ` • ${escapeHtml(workflowFlags)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        <ul class="key-points">
          ${preview.steps.map((step) => {
            const stepFlags = [
              step.dependsOn.length > 0 ? `depends on ${step.dependsOn.join(", ")}` : "no dependencies",
              step.continueOnError ? "continue on error" : "stop on error",
              step.destructive ? "destructive" : "",
              step.reversible ? "reversible" : "non-reversible"
            ].filter(Boolean).join(" • ");
            return `<li><strong>${escapeHtml(step.stepId)}</strong> — ${escapeHtml(step.summary)}${stepFlags ? ` <span class="preview-meta">${escapeHtml(stepFlags)}</span>` : ""}${step.supportError ? ` <div class="warning-line">${escapeHtml(step.supportError)}</div>` : ""}</li>`;
          }).join("")}
        </ul>
        ${Array.isArray(preview.affectedRanges) && preview.affectedRanges.length > 0
          ? `<div class="preview-meta">${escapeHtml(preview.affectedRanges.join(", "))}</div>`
          : ""}
        ${compositeSupportError ? `<div class="warning-line">${escapeHtml(compositeSupportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !compositeSupportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Workflow
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "analysis_report_plan") {
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.sourceSheet)}!${escapeHtml(preview.sourceRange)}
          • ${escapeHtml(preview.outputMode)}
          • ${preview.sections.length} section(s)
          ${preview.targetSheet && preview.targetRange
            ? ` • ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}`
            : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${renderTable([], preview.rows)}
        ${getRequiresConfirmation(response) ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Analysis Report
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "pivot_table_plan") {
    const pivotDetails = [
      preview.columnGroups.length > 0 ? `columns ${preview.columnGroups.join(", ")}` : "",
      preview.filters.length > 0
        ? `filters ${preview.filters.map((filter) => filter.field).join(", ")}`
        : "",
      preview.sort
        ? `sort ${preview.sort.field} ${preview.sort.direction}`
        : ""
    ].filter(Boolean).join(" • ");
    const supportError = getExcelPlanSupportError(preview);

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.sourceSheet)}!${escapeHtml(preview.sourceRange)}
          • ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${preview.rowGroups.length} row group(s)
          • ${preview.valueAggregations.length} value(s)
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${pivotDetails ? `<div class="preview-meta">${escapeHtml(pivotDetails)}</div>` : ""}
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Pivot Table
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "chart_plan") {
    const chartDetails = [
      preview.categoryField ? `category ${preview.categoryField}` : "",
      preview.title ? `title ${preview.title}` : "",
      preview.legendPosition ? `legend ${preview.legendPosition}` : "",
      preview.horizontalAxisTitle ? `horizontal axis ${preview.horizontalAxisTitle}` : "",
      preview.verticalAxisTitle ? `vertical axis ${preview.verticalAxisTitle}` : ""
    ].filter(Boolean).join(" • ");
    const supportError = getExcelPlanSupportError(preview);

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.sourceSheet)}!${escapeHtml(preview.sourceRange)}
          • ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${escapeHtml(preview.chartType)}
          • ${preview.series.length} series
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${chartDetails ? `<div class="preview-meta">${escapeHtml(chartDetails)}</div>` : ""}
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Chart
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "table_plan") {
    const supportError = getExcelPlanSupportError(preview);
    const tableDetails = [
      preview.name ? `name ${preview.name}` : "",
      preview.styleName ? `style ${preview.styleName}` : "",
      preview.hasHeaders ? "headers on" : "headers off",
      preview.showBandedRows !== undefined ? `banded rows ${preview.showBandedRows ? "on" : "off"}` : "",
      preview.showBandedColumns !== undefined ? `banded columns ${preview.showBandedColumns ? "on" : "off"}` : "",
      preview.showFilterButton !== undefined ? `filters ${preview.showFilterButton ? "on" : "off"}` : "",
      preview.showTotalsRow !== undefined ? `totals ${preview.showTotalsRow ? "on" : "off"}` : ""
    ].filter(Boolean).join(" • ");

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          ${preview.overwriteRisk ? ` • overwrite ${escapeHtml(preview.overwriteRisk)}` : ""}
          ${preview.confirmationLevel ? ` • ${escapeHtml(preview.confirmationLevel)}` : ""}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${tableDetails ? `<div class="preview-meta">${escapeHtml(tableDetails)}</div>` : ""}
        ${Array.isArray(preview.affectedRanges) && preview.affectedRanges.length > 0
          ? `<div class="preview-meta">${escapeHtml(preview.affectedRanges.join(", "))}</div>`
          : ""}
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
        ${getRequiresConfirmation(response) && !supportError ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Table
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "external_data_plan") {
    const supportError = getExcelPlanSupportError(preview);
    const sourceDetails = preview.sourceType === "market_data"
      ? [
          preview.query?.symbol ? `symbol ${preview.query.symbol}` : "",
          preview.query?.attribute ? `attribute ${preview.query.attribute}` : ""
        ].filter(Boolean).join(" • ")
      : [
          preview.sourceUrl ? `url ${preview.sourceUrl}` : "",
          preview.selectorType ? `selector ${preview.selectorType}` : "",
          preview.selector !== undefined ? `${preview.selector}` : ""
        ].filter(Boolean).join(" • ");

    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${escapeHtml(preview.sourceType)}
          • ${escapeHtml(preview.provider)}
        </div>
        <div>${escapeHtml(preview.explanation)}</div>
        <div class="preview-meta">${escapeHtml(preview.summary)}</div>
        ${sourceDetails ? `<div class="preview-meta">${escapeHtml(sourceDetails)}</div>` : ""}
        <pre class="formula-block">${escapeHtml(preview.formula)}</pre>
        ${supportError ? `<div class="warning-line">${escapeHtml(supportError)}</div>` : ""}
      </div>
    `;
  }

  if (preview.kind === "sheet_update") {
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${preview.shape.rows}x${preview.shape.columns}
          • ${escapeHtml(preview.operation)}
          • ${escapeHtml(preview.matrixKind)}
          ${preview.overwriteRisk ? ` • overwrite ${escapeHtml(preview.overwriteRisk)}` : ""}
        </div>
        ${renderTable([], preview.rows)}
        ${getRequiresConfirmation(response) ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Update
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "sheet_import_plan") {
    return `
      <div class="preview">
        <div class="preview-meta">
          ${escapeHtml(preview.targetSheet)}!${escapeHtml(preview.targetRange)}
          • ${preview.shape.rows}x${preview.shape.columns}
          • extraction ${escapeHtml(preview.extractionMode)}
        </div>
        ${renderTable(preview.headers, preview.rows)}
        ${getRequiresConfirmation(response) ? `
          <div class="preview-actions">
            <button type="button" data-confirm-run="${escapeHtmlAttribute(message.runId || "")}" data-request="${escapeHtmlAttribute(message.requestId || "")}">
              Confirm Insert
            </button>
          </div>
        ` : ""}
      </div>
    `;
  }

  if (preview.kind === "extracted_table") {
    return `
      <div class="preview">
        <div class="preview-meta">
          Extracted table • ${preview.shape.rows}x${preview.shape.columns}
          • extraction ${escapeHtml(preview.extractionMode)}
        </div>
        ${renderTable(preview.headers, preview.rows)}
      </div>
    `;
  }

  if (preview.kind === "attachment_analysis") {
    return `
      <div class="preview">
        <div class="preview-meta">
          Attachment analysis • ${escapeHtml(preview.contentKind)}
          • extraction ${escapeHtml(preview.extractionMode)}
        </div>
      </div>
    `;
  }

  if (preview.kind === "document_summary") {
    return `
      <div class="preview">
        <div class="preview-meta">
          Document summary • ${escapeHtml(preview.contentKind)}
          • extraction ${escapeHtml(preview.extractionMode)}
        </div>
        ${preview.keyPoints.length > 0 ? `
          <ul class="key-points">
            ${preview.keyPoints.map((point) => `<li>${escapeHtml(point)}</li>`).join("")}
          </ul>
        ` : ""}
      </div>
    `;
  }

  return "";
}

function scrollMessagesToBottom() {
  const applyScroll = () => {
    elements.messages.scrollTop = elements.messages.scrollHeight;
    if (elements.messages.lastElementChild && typeof elements.messages.lastElementChild.scrollIntoView === "function") {
      elements.messages.lastElementChild.scrollIntoView({ block: "end" });
    }
  };

  applyScroll();

  if (typeof window.requestAnimationFrame === "function") {
    window.requestAnimationFrame(() => {
      applyScroll();
      if (typeof window.requestAnimationFrame === "function") {
        window.requestAnimationFrame(applyScroll);
      }
    });
  }
}

function getMessagesClientHeight() {
  const clientHeight = Number(elements.messages?.clientHeight ?? 0);
  return Number.isFinite(clientHeight) && clientHeight > 0 ? clientHeight : 0;
}

function isMessagesNearBottom() {
  const scrollTop = Number(elements.messages?.scrollTop ?? 0);
  const scrollHeight = Number(elements.messages?.scrollHeight ?? 0);
  const clientHeight = getMessagesClientHeight();

  if (scrollHeight <= 0) {
    return true;
  }

  if (clientHeight <= 0) {
    return scrollTop >= Math.max(0, scrollHeight - MESSAGE_SCROLL_BOTTOM_THRESHOLD_PX);
  }

  return scrollTop + clientHeight >= scrollHeight - MESSAGE_SCROLL_BOTTOM_THRESHOLD_PX;
}

function disconnectMessageAutoScrollObservers() {
  if (state.messageLayoutObserver && typeof state.messageLayoutObserver.disconnect === "function") {
    state.messageLayoutObserver.disconnect();
  }
  if (state.messageMutationObserver && typeof state.messageMutationObserver.disconnect === "function") {
    state.messageMutationObserver.disconnect();
  }

  state.messageLayoutObserver = null;
  state.messageMutationObserver = null;
}

function clearMessageScrollFollowUps() {
  if (typeof globalThis.clearTimeout === "function") {
    (state.messageScrollTimeoutIds || []).forEach((timeoutId) => {
      globalThis.clearTimeout(timeoutId);
    });
  }
  state.messageScrollTimeoutIds = [];
}

function scheduleMessagesAutoScroll(force = false) {
  clearMessageScrollFollowUps();

  if (force) {
    state.messageScrollPinned = true;
  }

  if (!force && !state.messageScrollPinned) {
    return;
  }

  scrollMessagesToBottom();

  if (typeof globalThis.setTimeout === "function") {
    for (const delay of MESSAGE_SCROLL_FOLLOWUP_DELAYS_MS) {
      const timeoutId = globalThis.setTimeout(() => {
        if (force || state.messageScrollPinned) {
          scrollMessagesToBottom();
        }
      }, delay);
      state.messageScrollTimeoutIds.push(timeoutId);
    }
  }
}

function bindMessageAutoScrollObservers() {
  if (!elements.messages || typeof elements.messages.addEventListener !== "function") {
    return;
  }

  if (!state.messageScrollListenersBound) {
    elements.messages.addEventListener("scroll", () => {
      state.messageScrollPinned = isMessagesNearBottom();
    });
    state.messageScrollListenersBound = true;
  }

  disconnectMessageAutoScrollObservers();

  if (typeof ResizeObserver === "function") {
    const resizeObserver = new ResizeObserver(() => {
      scheduleMessagesAutoScroll();
    });
    resizeObserver.observe(elements.messages);
    Array.from(elements.messages.children ?? []).forEach((child) => {
      resizeObserver.observe(child);
    });
    state.messageLayoutObserver = resizeObserver;
  }

  if (typeof MutationObserver === "function") {
    const mutationObserver = new MutationObserver(() => {
      bindMessageAutoScrollObservers();
      scheduleMessagesAutoScroll();
    });
    mutationObserver.observe(elements.messages, {
      childList: true,
      subtree: true,
      characterData: true
    });
    state.messageMutationObserver = mutationObserver;
  }
}

function renderMessages() {
  elements.messages.innerHTML = state.messages.map((message) => {
    const response = message.response;
    const content = response ? getResponseBodyText(response) : message.content;
    const warnings = response ? getResponseWarnings(response) : [];
    const suggestions = response ? getFollowUpSuggestions(response) : [];

    return `
      <section class="message ${message.role}">
        <div class="bubble">${escapeHtml(content)}</div>
        ${message.statusLine ? `<div class="status-line">${escapeHtml(message.statusLine)}</div>` : ""}
        ${warnings.length > 0 ? renderWarnings(warnings) : ""}
        ${renderStructuredPreview(response, message)}
        ${renderSuggestions(suggestions)}
      </section>
    `;
  }).join("");

  elements.messages.querySelectorAll("[data-confirm-run]").forEach((button) => {
    button.addEventListener("click", async () => {
      const runId = button.getAttribute("data-confirm-run");
      const requestId = button.getAttribute("data-request");
      const message = state.messages.find((entry) => entry.runId === runId);
      if (!runId || !requestId || !message?.response || !isWritePlanResponse(message.response)) {
        return;
      }

      button.setAttribute("disabled", "true");

      try {
        await executeWritePlanMessage(message);
      } finally {
        renderMessages();
      }
    });
  });

  elements.messages.querySelectorAll("[data-suggestion-index]").forEach((button) => {
    button.addEventListener("click", () => {
      const section = button.closest(".message");
      const index = Number(button.getAttribute("data-suggestion-index"));
      const messageIndex = Array.from(elements.messages.children).indexOf(section);
      const message = state.messages[messageIndex];
      const suggestions = message?.response ? getFollowUpSuggestions(message.response) : [];
      const suggestion = suggestions[index];
      if (suggestion) {
        elements.prompt.value = suggestion;
        elements.prompt.focus();
      }
    });
  });

  bindMessageAutoScrollObservers();
  scheduleMessagesAutoScroll();
}

async function getSpreadsheetSnapshot(prompt) {
  const runtimeConfig = getRuntimeConfig();
  const platform = detectExcelPlatform();

  return Excel.run(async (context) => {
    const isMissingResourceError = (error) =>
      /^The requested resource doesn't exist\.?$/i.test(
        String(error?.message || error || "").trim().replace(/^Error:\s*/i, "")
      );

    async function loadActiveSheet() {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
      return sheet;
    }

    async function loadSelectedRange(sheet) {
      let range = context.workbook.getSelectedRange();
      range.load(["address"]);
      try {
        await context.sync();
        return {
          range,
          usedFallback: false
        };
      } catch (error) {
        if (!isMissingResourceError(error)) {
          throw error;
        }

        range = sheet.getRange("A1");
        range.load(["address"]);
        await context.sync();
        return {
          range,
          usedFallback: true
        };
      }
    }

    async function loadActiveCell(sheet, selectedRange, preferSelectionAnchor = false) {
      const selectedRangeBounds = parseA1RangeReference(normalizeA1(selectedRange.address));
      const selectionAnchor = buildA1RangeFromBounds({
        startRow: selectedRangeBounds.startRow,
        endRow: selectedRangeBounds.startRow,
        startColumn: selectedRangeBounds.startColumn,
        endColumn: selectedRangeBounds.startColumn
      });

      let activeCell = preferSelectionAnchor
        ? sheet.getRange(selectionAnchor)
        : context.workbook.getActiveCell();
      activeCell.load(["address", "values", "formulas"]);
      try {
        await context.sync();
        return activeCell;
      } catch (error) {
        if (!isMissingResourceError(error)) {
          throw error;
        }

        activeCell = sheet.getRange(selectionAnchor);
        activeCell.load(["address", "values", "formulas"]);
        await context.sync();
        return activeCell;
      }
    }

    async function loadCurrentRegion(activeCell, selectedRange) {
      let currentRegion = selectedRange;
      try {
        if (typeof activeCell.getSurroundingRegion === "function") {
          currentRegion = activeCell.getSurroundingRegion();
        }
      } catch {
        currentRegion = selectedRange;
      }

      currentRegion.load(["address"]);
      try {
        await context.sync();
        return currentRegion;
      } catch (error) {
        if (!isMissingResourceError(error)) {
          throw error;
        }

        selectedRange.load(["address"]);
        await context.sync();
        return selectedRange;
      }
    }

    let sheet;
    try {
      sheet = await loadActiveSheet();
    } catch (error) {
      if (!isMissingResourceError(error)) {
        throw error;
      }
      sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();
    }

    const { range, usedFallback: usedSelectedRangeFallback } = await loadSelectedRange(sheet);
    const activeCell = await loadActiveCell(sheet, range, usedSelectedRangeFallback);
    const currentRegion = await loadCurrentRegion(activeCell, range);

    const workbookTitle = Office.context.document.url
      ? decodeURIComponent(Office.context.document.url.split("/").pop() || "Workbook")
      : "Workbook";

    const selectedRange = normalizeA1(range.address);
    const selectedRangeBounds = parseA1RangeReference(selectedRange);
    const includeSelectionMatrix = shouldIncludeRegionMatrix(selectedRangeBounds);
    let selectionHeaderRange;

    if (includeSelectionMatrix) {
      range.load(["values", "formulas"]);
    } else {
      selectionHeaderRange = sheet.getRange(buildA1RangeFromBounds({
        startRow: selectedRangeBounds.startRow,
        endRow: selectedRangeBounds.startRow,
        startColumn: selectedRangeBounds.startColumn,
        endColumn: selectedRangeBounds.endColumn
      }));
      selectionHeaderRange.load(["values"]);
    }

    const activeCellA1 = normalizeA1(activeCell.address);
    const firstCellValue = normalizeExcelCellValue(activeCell.values?.[0]?.[0]);
    const firstCellFormula = activeCell.formulas?.[0]?.[0];
    const currentRegionRange = normalizeA1(currentRegion.address);
    const currentRegionBounds = parseA1RangeReference(currentRegionRange);
    const includeCurrentRegionMatrix = shouldIncludeRegionMatrix(currentRegionBounds);
    let currentRegionHeaderRange;

    if (includeCurrentRegionMatrix) {
      currentRegion.load(["values", "formulas"]);
    } else {
      currentRegionHeaderRange = sheet.getRange(buildA1RangeFromBounds({
        startRow: currentRegionBounds.startRow,
        endRow: currentRegionBounds.startRow,
        startColumn: currentRegionBounds.startColumn,
        endColumn: currentRegionBounds.endColumn
      }));
      currentRegionHeaderRange.load(["values"]);
    }

    const referencedCellRanges = getPromptReferencedA1Notations(prompt, activeCellA1)
      .map((a1Notation) => {
        const cell = sheet.getRange(a1Notation);
        cell.load(["address", "values", "formulas"]);
        return cell;
      });

    if (
      includeSelectionMatrix ||
      selectionHeaderRange ||
      includeCurrentRegionMatrix ||
      currentRegionHeaderRange ||
      referencedCellRanges.length > 0
    ) {
      await context.sync();
    }

    const noteCollection = getExcelContextNoteCollection(context.workbook, sheet);
    let activeCellNote;
    let referencedCellNotes = [];
    if (noteCollection) {
      const activeCellNoteItem = queueExcelContextNoteRead(noteCollection, activeCell);
      const referencedCellNoteItems = referencedCellRanges.map((cell) =>
        queueExcelContextNoteRead(noteCollection, cell)
      );

      if (activeCellNoteItem || referencedCellNoteItems.some(Boolean)) {
        await context.sync();
      }

      activeCellNote = readExcelContextNote(activeCellNoteItem);
      referencedCellNotes = referencedCellNoteItems.map(readExcelContextNote);
    }

    const referencedCells = referencedCellRanges.map((cell, index) => {
      const cellValue = normalizeExcelCellValue(cell.values?.[0]?.[0]);
      const cellFormula = cell.formulas?.[0]?.[0];
      const note = referencedCellNotes[index];

      const cellContext = {
        a1Notation: normalizeA1(cell.address),
        displayValue: cellValue,
        value: cellValue,
        formula: normalizeExcelFormulaText(cellFormula) || undefined
      };

      if (note) {
        cellContext.note = note;
      }

      return cellContext;
    });

    const selectionHeaderValues = includeSelectionMatrix
      ? normalizeExcelMatrixValues(range.values).slice(0, 1)
      : normalizeExcelMatrixValues(selectionHeaderRange?.values || []);
    const currentRegionHeaderValues = includeCurrentRegionMatrix
      ? normalizeExcelMatrixValues(currentRegion.values).slice(0, 1)
      : normalizeExcelMatrixValues(currentRegionHeaderRange?.values || []);
    const selectionContext = {
      range: selectedRange,
      headers: getSelectionHeaders(selectionHeaderValues)
    };

    if (includeSelectionMatrix) {
      const normalizedSelectionValues = normalizeExcelMatrixValues(range.values);
      selectionContext.values = normalizedSelectionValues;
      selectionContext.formulas = normalizeFormulaMatrix(range.formulas);
    }

    const currentRegionContext = {
      range: currentRegionRange,
      headers: getSelectionHeaders(currentRegionHeaderValues)
    };

    if (includeCurrentRegionMatrix) {
      const normalizedCurrentRegionValues = normalizeExcelMatrixValues(currentRegion.values);
      currentRegionContext.values = normalizedCurrentRegionValues;
      currentRegionContext.formulas = normalizeFormulaMatrix(currentRegion.formulas);
    }

    const activeCellContext = {
      a1Notation: activeCellA1,
      displayValue: firstCellValue,
      value: firstCellValue,
      formula: normalizeExcelFormulaText(firstCellFormula) || undefined
    };

    if (activeCellNote) {
      activeCellContext.note = activeCellNote;
    }

    return {
      source: {
        channel: platform,
        clientVersion: runtimeConfig.clientVersion,
        sessionId
      },
      host: {
        platform,
        workbookTitle,
        workbookId: getWorkbookIdentity(),
        activeSheet: sheet.name,
        selectedRange,
        locale: Office.context.displayLanguage || navigator.language,
        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone
      },
      context: {
        selection: selectionContext,
        currentRegion: currentRegionContext,
        ...buildImplicitRegionTargets(currentRegionRange),
        activeCell: activeCellContext,
        referencedCells: referencedCells.length > 0 ? referencedCells : undefined
      }
    };
  });
}

function buildRequestEnvelope(input) {
  const runtimeConfig = getRuntimeConfig();
  return {
    schemaVersion: "1.0.0",
    requestId: createRequestId(),
    source: input.snapshot.source,
    host: input.snapshot.host,
    userMessage: truncateRequestText(input.userMessage),
    conversation: sanitizeConversation(input.conversation),
    context: {
      ...input.snapshot.context,
      attachments: input.attachments.length > 0 ? input.attachments : undefined
    },
    capabilities: {
      canRenderTrace: true,
      canRenderStructuredPreview: true,
      canConfirmWriteBack: true,
      supportsImageInputs: true,
      supportsWriteBackExecution: true,
      supportsNoteWrites: isExcelNoteWriteSupported()
    },
    reviewer: {
      reviewerSafeMode: runtimeConfig.reviewerSafeMode,
      forceExtractionMode: runtimeConfig.forceExtractionMode
    },
    confirmation: {
      state: "none"
    }
  };
}

async function uploadPendingAttachments() {
  const uploads = [];

  for (const attachment of state.pendingAttachments) {
    if (attachment.uploadedAttachment) {
      uploads.push(attachment.uploadedAttachment);
      continue;
    }

    attachment.status = "Uploading";
    renderAttachmentStrip();

    try {
      const uploaded = await gateway.uploadImage({
        file: attachment.file,
        fileName: attachment.fileName,
        source: attachment.source,
        sessionId
      });
      attachment.uploadedAttachment = uploaded;
      attachment.status = "Uploaded";
      uploads.push(uploaded);
      renderAttachmentStrip();
    } catch (error) {
      attachment.status = "Failed";
      renderAttachmentStrip();
      throw error;
    }
  }

  return uploads;
}

function cancelMessagePoll(message) {
  if (message && message.pollTimerId != null) {
    window.clearTimeout(message.pollTimerId);
    message.pollTimerId = null;
  }

  if (message) {
    message.pollingStopped = true;
  }
}

function scheduleMessagePoll(message, delayMs = MESSAGE_POLL_INTERVAL_MS) {
  if (!message || message.pollingStopped) {
    return;
  }

  message.pollTimerId = window.setTimeout(() => {
    message.pollTimerId = null;
    void runMessagePoll(message);
  }, delayMs);
}

async function runMessagePoll(message) {
  if (!message || message.pollingStopped || message.pollingInFlight) {
    return;
  }

  message.pollingInFlight = true;
  message.pollAttemptCount = Number(message.pollAttemptCount || 0) + 1;

  try {
    if (shouldPollTraceForMessage(message)) {
      try {
        const trace = await gateway.pollTrace(message.runId, message.traceIndex || 0, message.requestId);
        message.traceIndex = trace.nextIndex;
        if (trace.events.length > 0) {
          message.trace = trimMessageTraceEvents([...(message.trace || []), ...trace.events]);
          message.statusLine = summarizeLatestTrace(message.trace);
          renderMessages();
        }
      } catch (error) {
        if (isTraceUnavailablePollError(error) || isTraceBandwidthPollError(error)) {
          message.tracePollingDisabled = true;
        } else {
          throw error;
        }
      }
    }

    const run = await gateway.pollRun(message.runId, message.requestId);
    if (run.status === "completed" && run.response) {
      cancelMessagePoll(message);
      setMessageResponse(message, run.response);
      message.content = getResponseBodyText(message.response);
      message.statusLine = summarizeLatestTrace(message.response.trace);
      renderMessages();
      return;
    }

    if (run.status === "failed") {
      cancelMessagePoll(message);
      message.content = run.error || "Hermes failed to process the request.";
      message.statusLine = "Request failed";
      renderMessages();
      return;
    }
  } catch (error) {
    cancelMessagePoll(message);
    message.content = sanitizeHostExecutionError(error, "Failed while polling Hermes.");
    message.statusLine = "Request failed";
    renderMessages();
    return;
  } finally {
    message.pollingInFlight = false;
  }

  message.pollDelayMs = getNextMessagePollDelay(message);
  scheduleMessagePoll(message, message.pollDelayMs);
}

async function pollRun(message) {
  cancelMessagePoll(message);
  message.pollingInFlight = false;
  message.pollingStopped = false;
  message.pollAttemptCount = 0;
  message.pollDelayMs = MESSAGE_POLL_INTERVAL_MS;
  message.tracePollingDisabled = false;
  scheduleMessagePoll(message);
}

function clearPendingAttachments() {
  for (const attachment of state.pendingAttachments) {
    if (attachment.previewUrl) {
      URL.revokeObjectURL(attachment.previewUrl);
    }
  }
  state.pendingAttachments = [];
  renderAttachmentStrip();
}

function addPendingFiles(files, source) {
  const supportedFiles = filterSupportedImageFiles(files);
  for (const file of supportedFiles) {
    state.pendingAttachments.push({
      file,
      fileName: file.name || `image-${Date.now()}.png`,
      source,
      status: "Ready",
      previewUrl: URL.createObjectURL(file)
    });
  }
  renderAttachmentStrip();
}

function buildCompositeExecutionSummary(stepResults) {
  const completedCount = stepResults.filter((step) => step.status === "completed").length;
  const failedCount = stepResults.filter((step) => step.status === "failed").length;
  const skippedCount = stepResults.filter((step) => step.status === "skipped").length;
  const parts = [
    `${stepResults.length} step${stepResults.length === 1 ? "" : "s"}`
  ];

  if (completedCount > 0) {
    parts.push(`${completedCount} completed`);
  }
  if (failedCount > 0) {
    parts.push(`${failedCount} failed`);
  }
  if (skippedCount > 0) {
    parts.push(`${skippedCount} skipped`);
  }

  return `Workflow finished: ${parts.join(" • ")}.`;
}

async function applyCompositePlan({ plan, requestId, runId, approvalToken, executionId }) {
  if (typeof executionId !== "string" || executionId.trim().length === 0) {
    throw new Error("Composite workflow execution requires executionId.");
  }

  const stepResults = [];
  const localSnapshots = [];
  const completedSteps = new Set();
  const failedSteps = new Set();
  const skippedSteps = new Set();
  let halted = false;

  for (let stepIndex = 0; stepIndex < plan.steps.length; stepIndex += 1) {
    const step = plan.steps[stepIndex];
    if (halted) {
      stepResults.push({
        stepId: step.stepId,
        status: "skipped",
        summary: "Skipped because an earlier workflow step failed."
      });
      skippedSteps.add(step.stepId);
      continue;
    }

    if (step.dependsOn.some((dependency) => failedSteps.has(dependency) || skippedSteps.has(dependency))) {
      stepResults.push({
        stepId: step.stepId,
        status: "skipped",
        summary: "Skipped because a dependency failed or was skipped."
      });
      skippedSteps.add(step.stepId);
      continue;
    }

    const unresolvedDependency = step.dependsOn.find(
      (dependency) => !completedSteps.has(dependency) && !failedSteps.has(dependency) && !skippedSteps.has(dependency)
    );
    if (unresolvedDependency) {
      stepResults.push({
        stepId: step.stepId,
        status: "failed",
        summary: `Dependency ${unresolvedDependency} has not completed before this step.`
      });
      failedSteps.add(step.stepId);
      halted = true;
      continue;
    }

    try {
      const supportError = getExcelPlanSupportError(step.plan);
      if (supportError) {
        throw new Error(supportError);
      }

      const result = await applyWritePlan({
        plan: step.plan,
        requestId,
        runId,
        approvalToken,
        executionId: `${executionId}_${stepIndex + 1}_${step.stepId}`.slice(0, 160)
      });
      if (result?.__hermesLocalExecutionSnapshot) {
        localSnapshots.push(result.__hermesLocalExecutionSnapshot);
      }
      const gatewayResult = stripLocalExecutionSnapshot(result);
      stepResults.push({
        stepId: step.stepId,
        status: "completed",
        summary: getCompositeStepWritebackStatusLine(step.plan, result),
        result: gatewayResult
      });
      completedSteps.add(step.stepId);
    } catch (error) {
      const summary = sanitizeHostExecutionError(error, "Workflow step failed.");
      stepResults.push({
        stepId: step.stepId,
        status: "failed",
        summary
      });
      failedSteps.add(step.stepId);
      if (!step.continueOnError) {
        halted = true;
      }
    }
  }

  const result = {
    kind: "composite_update",
    operation: "composite_update",
    hostPlatform: detectExcelPlatform(),
    executionId,
    stepResults,
    summary: getCompositeStatusSummary({
      stepResults,
      summary: buildCompositeExecutionSummary(stepResults)
    })
  };
  const completedStepCount = stepResults.filter((stepResult) => stepResult.status === "completed").length;
  const canUndoComposite = plan.reversible === true &&
    completedStepCount > 0 &&
    completedStepCount === stepResults.length &&
    localSnapshots.length === completedStepCount;

  return canUndoComposite
    ? attachLocalExecutionSnapshot({
      ...result,
      undoReady: true
    }, createCompositeLocalExecutionSnapshot({
      executionId,
      entries: localSnapshots
    }))
    : result;
}

async function applyWritePlan({ plan, requestId, runId, approvalToken, executionId }) {
  if (isCompositePlan(plan)) {
    return applyCompositePlan({
      plan,
      requestId,
      runId,
      approvalToken,
      executionId
    });
  }

  return Excel.run(async (context) => {
    void requestId;
    void runId;
    void approvalToken;

    const resolvedPlan = isMaterializedAnalysisReportPlan(plan)
      ? resolveMaterializedAnalysisReportPlan(plan)
      : plan;
    const platform = detectExcelPlatform();
    const worksheets = context.workbook.worksheets;

    if (isWorkbookStructurePlan(resolvedPlan)) {
      worksheets.load("items/name,items/position,items/visibility");
      await context.sync();

      const orderedSheets = [...worksheets.items].sort((left, right) => left.position - right.position);
      const findSheet = (sheetName) => orderedSheets.find((sheet) => sheet.name === sheetName);
      const clampInsertPosition = (position, count) => {
        if (position === "start") {
          return 0;
        }

        if (position === "end" || position === undefined) {
          return count;
        }

        return Math.max(0, Math.min(position, count));
      };
      const clampExistingPosition = (position, count) => {
        if (position === "start") {
          return 0;
        }

        if (position === "end") {
          return Math.max(0, count - 1);
        }

        return Math.max(0, Math.min(position, Math.max(0, count - 1)));
      };

      switch (resolvedPlan.operation) {
        case "create_sheet": {
          const createdSheet = worksheets.add(resolvedPlan.sheetName);
          const desiredPosition = clampInsertPosition(resolvedPlan.position, orderedSheets.length);
          if (desiredPosition < orderedSheets.length) {
            createdSheet.position = desiredPosition;
          }
          await context.sync();
          return {
            kind: "workbook_structure_update",
            hostPlatform: platform,
            sheetName: resolvedPlan.sheetName,
            operation: resolvedPlan.operation,
            positionResolved: desiredPosition,
            sheetCount: orderedSheets.length + 1,
            summary: getWorkbookStructureStatusSummary(resolvedPlan)
          };
        }
        case "delete_sheet": {
          const sheet = findSheet(resolvedPlan.sheetName);
          if (!sheet) {
            throw new Error(`Target sheet not found: ${resolvedPlan.sheetName}`);
          }

          if (sheet.visibility === Excel.SheetVisibility.veryHidden) {
            sheet.visibility = Excel.SheetVisibility.hidden;
          }

          sheet.delete();
          await context.sync();
          return {
            kind: "workbook_structure_update",
            hostPlatform: platform,
            sheetName: resolvedPlan.sheetName,
            operation: resolvedPlan.operation,
            summary: getWorkbookStructureStatusSummary(resolvedPlan)
          };
        }
        case "rename_sheet": {
          const sheet = findSheet(resolvedPlan.sheetName);
          if (!sheet) {
            throw new Error(`Target sheet not found: ${resolvedPlan.sheetName}`);
          }

          sheet.name = resolvedPlan.newSheetName;
          await context.sync();
          return {
            kind: "workbook_structure_update",
            hostPlatform: platform,
            sheetName: resolvedPlan.sheetName,
            operation: resolvedPlan.operation,
            newSheetName: resolvedPlan.newSheetName,
            summary: getWorkbookStructureStatusSummary(resolvedPlan)
          };
        }
        case "duplicate_sheet": {
          const sourceSheet = findSheet(resolvedPlan.sheetName);
          if (!sourceSheet) {
            throw new Error(`Target sheet not found: ${resolvedPlan.sheetName}`);
          }

          const desiredPosition = clampInsertPosition(resolvedPlan.position, orderedSheets.length);
          let copiedSheet;

          if (desiredPosition <= 0) {
            copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.beginning);
          } else if (desiredPosition >= orderedSheets.length) {
            copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.end);
          } else {
            copiedSheet = sourceSheet.copy(
              Excel.WorksheetPositionType.before,
              orderedSheets[desiredPosition]
            );
          }

          if (resolvedPlan.newSheetName) {
            copiedSheet.name = resolvedPlan.newSheetName;
          }

          copiedSheet.load?.("name");
          await context.sync();
          const copiedSheetName = typeof copiedSheet.name === "string" &&
            copiedSheet.name.trim().length > 0
            ? copiedSheet.name
            : resolvedPlan.newSheetName || resolvedPlan.sheetName;

          return {
            kind: "workbook_structure_update",
            hostPlatform: platform,
            sheetName: resolvedPlan.sheetName,
            operation: resolvedPlan.operation,
            newSheetName: copiedSheetName,
            positionResolved: desiredPosition,
            sheetCount: orderedSheets.length + 1,
            summary: getWorkbookStructureStatusSummary(resolvedPlan)
          };
        }
        case "move_sheet": {
          const sheet = findSheet(resolvedPlan.sheetName);
          if (!sheet) {
            throw new Error(`Target sheet not found: ${resolvedPlan.sheetName}`);
          }

          sheet.position = clampExistingPosition(resolvedPlan.position, orderedSheets.length);
          await context.sync();
          return {
            kind: "workbook_structure_update",
            hostPlatform: platform,
            sheetName: resolvedPlan.sheetName,
            operation: resolvedPlan.operation,
            positionResolved: clampExistingPosition(resolvedPlan.position, orderedSheets.length),
            sheetCount: orderedSheets.length,
            summary: getWorkbookStructureStatusSummary(resolvedPlan)
          };
        }
        case "hide_sheet": {
          const sheet = findSheet(resolvedPlan.sheetName);
          if (!sheet) {
            throw new Error(`Target sheet not found: ${resolvedPlan.sheetName}`);
          }

          const visibleSheets = orderedSheets.filter((worksheet) =>
            worksheet.visibility === Excel.SheetVisibility.visible
          );
          if (visibleSheets.length <= 1 && sheet.visibility === Excel.SheetVisibility.visible) {
            throw new Error("Cannot hide the only visible worksheet.");
          }

          sheet.visibility = Excel.SheetVisibility.hidden;
          await context.sync();
          return {
            kind: "workbook_structure_update",
            hostPlatform: platform,
            sheetName: resolvedPlan.sheetName,
            operation: resolvedPlan.operation,
            summary: getWorkbookStructureStatusSummary(resolvedPlan)
          };
        }
        case "unhide_sheet": {
          const sheet = findSheet(resolvedPlan.sheetName);
          if (!sheet) {
            throw new Error(`Target sheet not found: ${resolvedPlan.sheetName}`);
          }

          sheet.visibility = Excel.SheetVisibility.visible;
          await context.sync();
          return {
            kind: "workbook_structure_update",
            hostPlatform: platform,
            sheetName: resolvedPlan.sheetName,
            operation: resolvedPlan.operation,
            summary: getWorkbookStructureStatusSummary(resolvedPlan)
          };
        }
        default:
          throw new Error("Unsupported workbook structure update.");
      }
    }

    if (isSheetStructurePlan(plan)) {
      const sheet = worksheets.getItem(plan.targetSheet);
      const getDimensionRange = (isRowOperation) =>
        getExcelFullRowOrColumnRange(sheet, plan, isRowOperation);

      switch (plan.operation) {
        case "insert_rows":
          getDimensionRange(true).insert(Excel.InsertShiftDirection.down);
          break;
        case "delete_rows":
          getDimensionRange(true).delete(Excel.DeleteShiftDirection.up);
          break;
        case "hide_rows":
          getDimensionRange(true).rowHidden = true;
          break;
        case "unhide_rows":
          getDimensionRange(true).rowHidden = false;
          break;
        case "group_rows":
          getDimensionRange(true).group(Excel.GroupOption.byRows);
          break;
        case "ungroup_rows":
          getDimensionRange(true).ungroup(Excel.GroupOption.byRows);
          break;
        case "insert_columns":
          getDimensionRange(false).insert(Excel.InsertShiftDirection.right);
          break;
        case "delete_columns":
          getDimensionRange(false).delete(Excel.DeleteShiftDirection.left);
          break;
        case "hide_columns":
          getDimensionRange(false).columnHidden = true;
          break;
        case "unhide_columns":
          getDimensionRange(false).columnHidden = false;
          break;
        case "group_columns":
          getDimensionRange(false).group(Excel.GroupOption.byColumns);
          break;
        case "ungroup_columns":
          getDimensionRange(false).ungroup(Excel.GroupOption.byColumns);
          break;
        case "merge_cells":
          sheet.getRange(plan.targetRange).merge(false);
          break;
        case "unmerge_cells":
          sheet.getRange(plan.targetRange).unmerge();
          break;
        case "freeze_panes": {
          const anchor = sheet.getRangeByIndexes(plan.frozenRows, plan.frozenColumns, 1, 1);
          if (typeof sheet.freezePanes?.freezeAt !== "function") {
            throw new Error("Excel host does not support freezing panes on this sheet.");
          }
          sheet.freezePanes.freezeAt(anchor);
          break;
        }
        case "unfreeze_panes":
          if (typeof sheet.freezePanes?.unfreeze !== "function") {
            throw new Error("Excel host does not support unfreezing panes on this sheet.");
          }
          sheet.freezePanes.unfreeze();
          break;
        case "autofit_rows":
          sheet.getRange(plan.targetRange).format.autofitRows();
          break;
        case "autofit_columns":
          sheet.getRange(plan.targetRange).format.autofitColumns();
          break;
        case "set_sheet_tab_color":
          sheet.tabColor = plan.color;
          break;
        default:
          throw new Error("Unsupported sheet structure update.");
      }

      await context.sync();
      const result = {
        kind: "sheet_structure_update",
        hostPlatform: platform,
        targetSheet: plan.targetSheet,
        operation: plan.operation,
        summary: getSheetStructureStatusSummary(plan)
      };

      switch (plan.operation) {
        case "insert_rows":
        case "delete_rows":
        case "hide_rows":
        case "unhide_rows":
        case "group_rows":
        case "ungroup_rows":
        case "insert_columns":
        case "delete_columns":
        case "hide_columns":
        case "unhide_columns":
        case "group_columns":
        case "ungroup_columns":
          result.startIndex = plan.startIndex;
          result.count = plan.count;
          break;
        case "merge_cells":
        case "unmerge_cells":
        case "autofit_rows":
        case "autofit_columns":
          result.targetRange = plan.targetRange;
          break;
        case "freeze_panes":
        case "unfreeze_panes":
          result.frozenRows = plan.frozenRows;
          result.frozenColumns = plan.frozenColumns;
          break;
        case "set_sheet_tab_color":
          result.color = plan.color;
          break;
      }

      return result;
    }

    function convertColumnLettersToNumber(value) {
      let total = 0;
      const text = String(value).trim().toUpperCase();
      for (const character of text) {
        total = (total * 26) + (character.charCodeAt(0) - 64);
      }
      return total;
    }

    function resolveColumnRef(columnRef, values, hasHeader, targetRangeA1) {
      const targetBounds = parseA1RangeReference(targetRangeA1);
      const assertInTarget = (offset) => {
        if (!Number.isInteger(offset) || offset < 0 || offset >= targetBounds.columnCount) {
          throw new Error(`Column ${columnRef} is outside ${targetRangeA1}.`);
        }

        return offset;
      };

      if (typeof columnRef === "number") {
        return assertInTarget(columnRef - 1);
      }

      if (typeof columnRef !== "string") {
        throw new Error(`Unsupported column reference: ${columnRef}`);
      }

      const trimmed = columnRef.trim();

      if (hasHeader && Array.isArray(values?.[0])) {
        const headerIndex = values[0].findIndex((value) => String(value).trim() === trimmed);
        if (headerIndex >= 0) {
          return assertInTarget(headerIndex);
        }
      }

      if (/^[A-Z]+$/i.test(trimmed)) {
        return assertInTarget(convertColumnLettersToNumber(trimmed) - targetBounds.startColumn);
      }

      if (/^\d+$/.test(trimmed)) {
        return assertInTarget(Number(trimmed) - 1);
      }

      throw new Error(`Column ${columnRef} is outside ${targetRangeA1}.`);
    }

    function buildFilterCriteria(condition) {
      switch (condition.operator) {
        case "equals":
          return { filterOn: "custom", criterion1: `=${condition.value}` };
        case "notEquals":
          return { filterOn: "custom", criterion1: `<>${condition.value}` };
        case "contains":
          return { filterOn: "custom", criterion1: `=*${condition.value}*` };
        case "startsWith":
          return { filterOn: "custom", criterion1: `=${condition.value}*` };
        case "endsWith":
          return { filterOn: "custom", criterion1: `=*${condition.value}` };
        case "greaterThan":
          return { filterOn: "custom", criterion1: `>${condition.value}` };
        case "greaterThanOrEqual":
          return { filterOn: "custom", criterion1: `>=${condition.value}` };
        case "lessThan":
          return { filterOn: "custom", criterion1: `<${condition.value}` };
        case "lessThanOrEqual":
          return { filterOn: "custom", criterion1: `<=${condition.value}` };
        case "isEmpty":
          return { filterOn: "custom", criterion1: "=" };
        case "isNotEmpty":
          return { filterOn: "custom", criterion1: "<>" };
        case "topN":
          return { filterOn: "topItems", criterion1: String(condition.value) };
        default:
          throw new Error(`Unsupported filter operator: ${condition.operator}`);
      }
    }

    function isComposableExcelFilterCriteria(criteria) {
      return criteria?.filterOn === "custom" &&
        typeof criteria.criterion1 === "string" &&
        typeof criteria.criterion2 === "undefined";
    }

    function getExcelAndFilterOperator() {
      if (typeof Excel !== "undefined" && Excel.FilterOperator?.and) {
        return Excel.FilterOperator.and;
      }

      return "and";
    }

    function buildExcelFilterCriteriaForColumn(conditions) {
      if (conditions.length === 1) {
        return buildFilterCriteria(conditions[0]);
      }

      if (conditions.length !== 2) {
        throw new Error("Excel host does not support multiple conditions for the same column.");
      }

      const firstCriteria = buildFilterCriteria(conditions[0]);
      const secondCriteria = buildFilterCriteria(conditions[1]);
      if (!isComposableExcelFilterCriteria(firstCriteria) || !isComposableExcelFilterCriteria(secondCriteria)) {
        throw new Error("Excel host does not support multiple conditions for the same column.");
      }

      return {
        filterOn: "custom",
        criterion1: firstCriteria.criterion1,
        criterion2: secondCriteria.criterion1,
        operator: getExcelAndFilterOperator()
      };
    }

    if (isRangeSortPlan(plan)) {
      const sheet = worksheets.getItem(plan.targetSheet);
      const target = sheet.getRange(plan.targetRange);
      target.load(["values", "formulas"]);
      await context.sync();
      const beforeValues = cloneMatrix(target.values);
      const beforeFormulas = cloneMatrix(target.formulas);

      const fields = buildExcelSortFields(plan).map((field) => ({
        ...field,
        key: resolveColumnRef(field.key, target.values, plan.hasHeader, plan.targetRange)
      }));
      const sort = typeof target.getSort === "function" ? target.getSort() : target.sort;
      if (!sort?.apply) {
        throw new Error("Excel host does not support range sort on this selection.");
      }

      sort.apply(fields, false, plan.hasHeader);
      target.load(["values", "formulas"]);
      await context.sync();
      return attachLocalExecutionSnapshot({
        kind: "range_sort",
        hostPlatform: platform,
        ...plan,
        summary: getRangeSortStatusSummary(plan)
      }, createLocalExecutionSnapshot({
        executionId,
        targetSheet: plan.targetSheet,
        targetRange: plan.targetRange,
        beforeValues,
        beforeFormulas,
        afterValues: target.values,
        afterFormulas: target.formulas
      }));
    }

    if (isRangeFilterPlan(plan)) {
      const sheet = worksheets.getItem(plan.targetSheet);
      const target = sheet.getRange(plan.targetRange);
      target.load(["values"]);
      await context.sync();

      const autoFilter = sheet.autoFilter;
      if (!autoFilter?.apply) {
        throw new Error("Excel host does not support range filters on this selection.");
      }

      if (plan.combiner !== "and") {
        throw new Error("Excel host does not support filter combiners other than and.");
      }

      const resolvedConditions = plan.conditions.map((condition) => ({
        ...condition,
        resolvedColumnRef: resolveColumnRef(condition.columnRef, target.values, plan.hasHeader, plan.targetRange)
      }));
      const conditionsByColumn = new Map();

      for (const condition of resolvedConditions) {
        const columnKey = String(condition.resolvedColumnRef);
        if (!conditionsByColumn.has(columnKey)) {
          conditionsByColumn.set(columnKey, {
            columnRef: condition.resolvedColumnRef,
            conditions: []
          });
        }

        conditionsByColumn.get(columnKey).conditions.push(condition);
      }

      const filterCriteriaByColumn = Array.from(conditionsByColumn.values()).map((columnGroup) => ({
        columnRef: columnGroup.columnRef,
        criteria: buildExcelFilterCriteriaForColumn(columnGroup.conditions)
      }));

      if (plan.clearExistingFilters && autoFilter.clearCriteria) {
        autoFilter.clearCriteria();
      }

      for (const filter of filterCriteriaByColumn) {
        autoFilter.apply(
          target,
          filter.columnRef,
          filter.criteria
        );
      }

      await context.sync();
      return {
        kind: "range_filter",
        hostPlatform: platform,
        ...plan,
        summary: getRangeFilterStatusSummary(plan)
      };
    }

    if (isConditionalFormatPlan(plan)) {
      const sheet = worksheets.getItem(plan.targetSheet);
      const target = sheet.getRange(plan.targetRange);
      applyExcelConditionalFormat(target, plan);
      await context.sync();
      return {
        kind: "conditional_format_update",
        hostPlatform: platform,
        ...plan,
        summary: getConditionalFormatStatusSummary(plan)
      };
    }

    if (isDataValidationPlan(plan)) {
      const sheet = worksheets.getItem(plan.targetSheet);
      const target = sheet.getRange(plan.targetRange);

      if (plan.ruleType === "checkbox") {
        if (hasCustomValidationPromptOrAlert(plan)) {
          throw new Error("Excel host cannot apply custom validation prompt or error text to checkbox controls.");
        }
        applyExcelCheckboxValidation(target, plan);
        await context.sync();
        return {
          kind: "data_validation_update",
          hostPlatform: platform,
          ...plan,
          summary: getDataValidationStatusSummary(plan)
        };
      }

      if (!target.dataValidation) {
        throw new Error("Excel host does not support data validation on this range.");
      }

      const validationRule = buildExcelValidationRule(plan);
      target.load(["rowCount", "columnCount"]);
      loadExcelDataValidationSnapshotTarget(target);
      await context.sync();
      const beforeValidation = readExcelDataValidationSnapshot(target);

      target.dataValidation.rule = validationRule;
      if (typeof plan.allowBlank === "boolean") {
        target.dataValidation.ignoreBlanks = plan.allowBlank;
      }
      target.dataValidation.prompt = {
        title: typeof plan.inputTitle === "string" && plan.inputTitle.trim().length > 0
          ? plan.inputTitle.trim()
          : "Validation",
        message: getValidationInputMessage(plan)
      };
      target.dataValidation.errorAlert = {
        title: typeof plan.errorTitle === "string" && plan.errorTitle.trim().length > 0
          ? plan.errorTitle.trim()
          : "Invalid data",
        message: typeof plan.errorMessage === "string" && plan.errorMessage.trim().length > 0
          ? plan.errorMessage.trim()
          : "Values must match the approved validation rule.",
        style: plan.invalidDataBehavior === "reject" ? "stop" : "warning",
        showAlert: true
      };
      loadExcelDataValidationSnapshotTarget(target);
      await context.sync();
      const afterValidation = readExcelDataValidationSnapshot(target);
      return attachLocalExecutionSnapshot({
        kind: "data_validation_update",
        hostPlatform: platform,
        ...plan,
        summary: getDataValidationStatusSummary(plan)
      }, createDataValidationLocalExecutionSnapshot({
        executionId,
        targetSheet: plan.targetSheet,
        targetRange: plan.targetRange,
        target,
        beforeValidation,
        afterValidation
      }));
    }

    if (isNamedRangeUpdatePlan(plan)) {
      const namedRangeWorksheetName = typeof plan.targetSheet === "string" && plan.targetSheet.length > 0
        ? plan.targetSheet
        : plan.scope === "sheet" && typeof plan.sheetName === "string" && plan.sheetName.length > 0
          ? plan.sheetName
          : undefined;
      const worksheet = namedRangeWorksheetName
        ? worksheets.getItem(namedRangeWorksheetName)
        : undefined;
      const namedRangeCollection = plan.scope === "sheet"
        ? worksheet?.names
        : context.workbook.names;
      if (plan.operation === "create") {
        await assertExcelNamedRangeCreateDoesNotCollide(context, namedRangeCollection, plan.name);
      }
      const target = plan.targetRange && worksheet ? worksheet.getRange(plan.targetRange) : undefined;
      const resolvedReference = target?.address ||
        (namedRangeWorksheetName && plan.targetRange ? `${namedRangeWorksheetName}!${plan.targetRange}` : "");
      const localSnapshot = await createNamedRangeLocalExecutionSnapshot({
        context,
        collection: namedRangeCollection,
        executionId,
        plan,
        resolvedReference
      });
      await applyExcelNamedRangeUpdate(
        context,
        context.workbook,
        worksheet || { names: undefined },
        plan,
        target || { address: namedRangeWorksheetName && plan.targetRange ? `${namedRangeWorksheetName}!${plan.targetRange}` : "" }
      );
      await context.sync();
      return attachLocalExecutionSnapshot({
        kind: "named_range_update",
        hostPlatform: platform,
        ...plan,
        summary: getNamedRangeStatusSummary(plan)
      }, localSnapshot);
    }

    if (isRangeTransferPlan(plan)) {
      const sourceWorksheet = worksheets.getItem(plan.sourceSheet);
      const targetWorksheet = worksheets.getItem(plan.targetSheet);
      const sourceRange = sourceWorksheet.getRange(plan.sourceRange);
      const targetRange = targetWorksheet.getRange(plan.targetRange);

      sourceRange.load(["rowCount", "columnCount", "values", "formulas"]);
      targetRange.load(["rowCount", "columnCount", "values", "formulas"]);
      await context.sync();

      const resolvedShape = getResolvedTransferShape(sourceRange, plan);
      const resolvedTargetRange = plan.operation === "append"
        ? resolveAppendMatrixTargetRange(
            targetRange,
            resolvedShape.rows,
            resolvedShape.columns
          )
        : resolveFullMatrixTargetRange(
            targetRange,
            resolvedShape.rows,
            resolvedShape.columns,
            "transfer shape"
          );
      const actualTargetRange = buildA1RangeFromBounds(
        deriveTransferTargetBounds(plan, resolvedTargetRange)
      );
      const canSnapshotCopyTransfer = plan.operation === "copy" && plan.pasteMode !== "formats";
      const beforeValues = canSnapshotCopyTransfer ? cloneMatrix(resolvedTargetRange.values) : null;
      const beforeFormulas = canSnapshotCopyTransfer ? cloneMatrix(resolvedTargetRange.formulas) : null;
      assertNonOverlappingTransfer(plan, resolvedTargetRange);
      writeTransferValues(resolvedTargetRange, sourceRange, plan);
      if (canSnapshotCopyTransfer) {
        resolvedTargetRange.load(["values", "formulas"]);
      }
      await context.sync();

      if (plan.operation === "move") {
        clearTransferredSource(sourceRange, plan);
        await context.sync();
      }

      const result = {
        kind: "range_transfer_update",
        hostPlatform: platform,
        operation: "range_transfer_update",
        sourceSheet: plan.sourceSheet,
        sourceRange: plan.sourceRange,
        targetSheet: plan.targetSheet,
        targetRange: actualTargetRange,
        transferOperation: plan.operation,
        pasteMode: plan.pasteMode,
        transpose: plan.transpose,
        summary: getRangeTransferStatusSummary({
          ...plan,
          targetRange: actualTargetRange
        })
      };

      return canSnapshotCopyTransfer
        ? attachLocalExecutionSnapshot(result, createLocalExecutionSnapshot({
            executionId,
            targetSheet: plan.targetSheet,
            targetRange: actualTargetRange,
            beforeValues,
            beforeFormulas,
            afterValues: resolvedTargetRange.values,
            afterFormulas: resolvedTargetRange.formulas
          }))
        : result;
    }

    if (isDataCleanupPlan(plan)) {
      const sheet = worksheets.getItem(plan.targetSheet);
      const target = sheet.getRange(plan.targetRange);
      target.load(["rowCount", "columnCount", "values", "formulas"]);
      await context.sync();
      const beforeValues = cloneMatrix(target.values);
      const beforeFormulas = cloneMatrix(target.formulas);

      if (!applyFormulaAwareFillDownCleanup(target, plan, target.values, target.formulas, "Excel host")) {
        const cleanupWrite = buildCleanupWriteMatrix(plan, target.values, target.formulas, "Excel host");
        if (cleanupWrite.kind === "formulas") {
          target.formulas = cleanupWrite.matrix;
        } else {
          target.values = cleanupWrite.matrix;
        }
      }
      target.load(["values", "formulas"]);
      await context.sync();

      return attachLocalExecutionSnapshot({
        kind: "data_cleanup_update",
        hostPlatform: platform,
        ...plan,
        summary: getDataCleanupStatusSummary(plan)
      }, createLocalExecutionSnapshot({
        executionId,
        targetSheet: plan.targetSheet,
        targetRange: plan.targetRange,
        beforeValues,
        beforeFormulas,
        afterValues: target.values,
        afterFormulas: target.formulas
      }));
    }

    if (isMaterializedAnalysisReportPlan(resolvedPlan)) {
      const sheet = worksheets.getItem(resolvedPlan.targetSheet);
      const target = sheet.getRange(resolvedPlan.targetRange);
      target.load(["rowCount", "columnCount"]);
      await context.sync();

      const reportMatrix = buildAnalysisReportMatrix(resolvedPlan);
      const resolvedTargetRange = resolveExactMatrixTargetRange(
        target,
        reportMatrix.length,
        reportMatrix[0]?.length || 1,
        "analysis report shape"
      );
      resolvedTargetRange.load(["values", "formulas"]);
      await context.sync();
      const beforeValues = cloneMatrix(resolvedTargetRange.values);
      const beforeFormulas = cloneMatrix(resolvedTargetRange.formulas);

      resolvedTargetRange.values = reportMatrix;
      resolvedTargetRange.load(["values", "formulas"]);
      await context.sync();
      const actualTargetRange = buildSizedA1RangeFromAnchor(
        resolvedPlan.targetRange,
        reportMatrix.length,
        reportMatrix[0]?.length || 1
      );

      return attachLocalExecutionSnapshot({
        kind: "analysis_report_update",
        hostPlatform: platform,
        ...resolvedPlan,
        summary: getAnalysisReportStatusSummary(resolvedPlan)
      }, createLocalExecutionSnapshot({
        executionId,
        targetSheet: resolvedPlan.targetSheet,
        targetRange: actualTargetRange,
        beforeValues,
        beforeFormulas,
        afterValues: resolvedTargetRange.values,
        afterFormulas: resolvedTargetRange.formulas
      }));
    }

    if (isPivotTablePlan(resolvedPlan)) {
      return applyExcelPivotTablePlan({
        context,
        worksheets,
        platform,
        plan: resolvedPlan
      });
    }

    if (isChartPlan(resolvedPlan)) {
      return applyExcelChartPlan(context, worksheets, resolvedPlan, platform);
    }

    if (isTablePlan(resolvedPlan)) {
      return applyExcelTablePlan(context, worksheets, resolvedPlan, platform);
    }

    const sheet = worksheets.getItem(plan.targetSheet);
    const target = sheet.getRange(plan.targetRange);
    target.load(["rowCount", "columnCount", "values", "formulas"]);
    await context.sync();
    const beforeValues = cloneMatrix(target.values);
    const beforeFormulas = cloneMatrix(target.formulas);

    if (isRangeFormatPlan(plan)) {
      const rangeFormatSupportError = getExcelRangeFormatSupportError(plan.format);
      if (rangeFormatSupportError) {
        throw new Error(rangeFormatSupportError);
      }

      const formatSnapshotTargets = prepareExcelRangeFormatSnapshotTargets(target, plan.format);
      loadExcelRangeFormatSnapshotTargets(formatSnapshotTargets);
      await context.sync();
      const beforeFormatCells = readExcelRangeFormatCellsSnapshot(formatSnapshotTargets);
      const beforeBorders = readExcelRangeFormatBorderSnapshot(formatSnapshotTargets);

      if (plan.format.backgroundColor) {
        target.format.fill.color = plan.format.backgroundColor;
      }

      if (plan.format.textColor) {
        target.format.font.color = plan.format.textColor;
      }

      if (plan.format.fontFamily) {
        target.format.font.name = plan.format.fontFamily;
      }

      if (typeof plan.format.fontSize === "number") {
        target.format.font.size = plan.format.fontSize;
      }

      if (typeof plan.format.bold === "boolean") {
        target.format.font.bold = plan.format.bold;
      }

      if (typeof plan.format.italic === "boolean") {
        target.format.font.italic = plan.format.italic;
      }

      if (typeof plan.format.underline === "boolean") {
        target.format.font.underline = plan.format.underline
          ? (Excel?.RangeUnderlineStyle?.single || "Single")
          : (Excel?.RangeUnderlineStyle?.none || "None");
      }

      if (typeof plan.format.strikethrough === "boolean") {
        target.format.font.strikethrough = plan.format.strikethrough;
      }

      const horizontalAlignment = mapHorizontalAlignmentToExcel(plan.format.horizontalAlignment);
      if (horizontalAlignment) {
        target.format.horizontalAlignment = horizontalAlignment;
      }

      const verticalAlignment = mapVerticalAlignmentToExcel(plan.format.verticalAlignment);
      if (verticalAlignment) {
        target.format.verticalAlignment = verticalAlignment;
      }

      const wrapText = mapWrapStrategyToExcel(plan.format.wrapStrategy);
      if (typeof wrapText === "boolean") {
        target.format.wrapText = wrapText;
      }

      if (plan.format.numberFormat) {
        target.numberFormat = Array.from({ length: target.rowCount }, () =>
          Array.from({ length: target.columnCount }, () => plan.format.numberFormat)
        );
      }

      if (plan.format.border) {
        applyExcelRangeBorder(target, plan.format.border);
      }

      if (typeof plan.format.columnWidth === "number") {
        target.format.columnWidth = plan.format.columnWidth;
      }

      if (typeof plan.format.rowHeight === "number") {
        target.format.rowHeight = plan.format.rowHeight;
      }

      loadExcelRangeFormatSnapshotTargets(formatSnapshotTargets);
      await context.sync();
      const afterFormatCells = readExcelRangeFormatCellsSnapshot(formatSnapshotTargets);
      const afterBorders = readExcelRangeFormatBorderSnapshot(formatSnapshotTargets);
      return attachLocalExecutionSnapshot({
        kind: "range_format_update",
        hostPlatform: platform,
        ...plan,
        summary: getRangeFormatStatusSummary(plan)
      }, createRangeFormatLocalExecutionSnapshot({
        executionId,
        targetSheet: plan.targetSheet,
        targetRange: plan.targetRange,
        target,
        beforeFormatCells,
        afterFormatCells,
        beforeBorders,
        afterBorders
      }));
    }

    if (target.rowCount !== plan.shape.rows || target.columnCount !== plan.shape.columns) {
      throw new Error("The approved targetRange does not match the proposed shape.");
    }

    const hasNoteMatrix = Array.isArray(plan.notes);
    let noteCollection = null;
    let noteTargets = null;
    let beforeNotes = null;
    let afterNotes = null;
    if (hasNoteMatrix) {
      noteCollection = getExcelNoteCollection(context.workbook, sheet);
      noteTargets = await prepareExcelNoteTargets(
        context,
        noteCollection,
        target,
        plan.shape.rows,
        plan.shape.columns
      );
      beforeNotes = readPreparedExcelNoteMatrix(noteTargets);
    }

    if (Array.isArray(plan.headers)) {
      if (rangeHasExistingContent(target.values) || hasAnyRealFormula(target.formulas)) {
        throw new Error("Target range already contains content. Clear it before confirming the import plan.");
      }

      target.values = [plan.headers, ...plan.values];
      target.load(["values", "formulas"]);
      await context.sync();
      return attachLocalExecutionSnapshot({
        kind: "range_write",
        hostPlatform: platform,
        ...plan,
        writtenRows: plan.shape.rows,
        writtenColumns: plan.shape.columns
      }, createLocalExecutionSnapshot({
        executionId,
        targetSheet: plan.targetSheet,
        targetRange: plan.targetRange,
        beforeValues,
        beforeFormulas,
        afterValues: target.values,
        afterFormulas: target.formulas
      }));
    }

    if (Array.isArray(plan.values) && !plan.formulas && !plan.notes) {
      target.values = plan.values;
      target.load(["values", "formulas"]);
      await context.sync();
    } else if (Array.isArray(plan.formulas) && !plan.values && !plan.notes) {
      target.formulas = plan.formulas.map((row) => row.map((cell) => cell ?? ""));
      target.load(["values", "formulas"]);
      await context.sync();
    } else if (Array.isArray(plan.notes) && !plan.values && !plan.formulas) {
      queueExcelNoteMatrixWrites(noteCollection, noteTargets, plan.notes, { clearBlankNotes: true });
      afterNotes = getExcelNoteMatrixAfterWrite(beforeNotes, plan.notes, { clearBlankNotes: true });
      target.load(["values", "formulas"]);
      await context.sync();
    } else {
      for (let rowIndex = 0; rowIndex < plan.shape.rows; rowIndex += 1) {
        for (let columnIndex = 0; columnIndex < plan.shape.columns; columnIndex += 1) {
          const cell = target.getCell(rowIndex, columnIndex);
          const formulaValue = plan.formulas?.[rowIndex]?.[columnIndex];
          const rawValue = plan.values?.[rowIndex]?.[columnIndex];

          if (typeof formulaValue === "string" && formulaValue.trim().length > 0) {
            cell.formulas = [[formulaValue]];
          } else if (rawValue !== undefined) {
            cell.values = [[rawValue]];
          } else {
            cell.values = [[""]];
          }
        }
      }
      if (hasNoteMatrix) {
        queueExcelNoteMatrixWrites(noteCollection, noteTargets, plan.notes, { clearBlankNotes: false });
        afterNotes = getExcelNoteMatrixAfterWrite(beforeNotes, plan.notes, { clearBlankNotes: false });
      }
      target.load(["values", "formulas"]);
      await context.sync();
    }

    return attachLocalExecutionSnapshot({
      kind: "range_write",
      hostPlatform: platform,
      ...plan,
      writtenRows: plan.shape.rows,
      writtenColumns: plan.shape.columns
    }, createLocalExecutionSnapshot({
      executionId,
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      beforeValues,
      beforeFormulas,
      beforeNotes,
      afterValues: target.values,
      afterFormulas: target.formulas,
      afterNotes
    }));
  });
}

async function sendPrompt() {
  const prompt = elements.prompt.value.trim();
  if (!prompt) {
    return;
  }
  const normalizedPrompt = truncateRequestText(prompt);

  const userMessage = {
    role: "user",
    content: prompt
  };
  appendStoredMessage(userMessage);
  state.messageScrollPinned = true;
  elements.prompt.value = "";

  if (UNDER_SPECIFIED_AFFIRMATION_PATTERN.test(normalizedPrompt)) {
    appendStoredMessage({
      role: "assistant",
      content: buildUnderSpecifiedFollowUpMessage(state.messages),
      statusLine: "Need more detail",
      requestId: "",
      runId: "",
      traceIndex: 0,
      trace: [],
      response: null
    });
    renderMessages();
    scheduleMessagesAutoScroll(true);
    return;
  }

  const executionShortcutMode = getExecutionShortcutMode(normalizedPrompt);

  const assistantMessage = {
    role: "assistant",
    content: executionShortcutMode === "undo"
      ? "Undoing the latest reversible write..."
      : executionShortcutMode === "redo"
        ? "Redoing the latest reversible write..."
        : "Thinking...",
    statusLine: executionShortcutMode
      ? "Sending execution-control request to Hermes"
      : "Sending request to Hermes",
    requestId: "",
    runId: "",
    traceIndex: 0,
    trace: [],
    response: null
  };

  appendStoredMessage(assistantMessage);
  renderMessages();
  scheduleMessagesAutoScroll(true);

  if (executionShortcutMode) {
    try {
      await handleExecutionShortcut(executionShortcutMode, assistantMessage);
    } catch (error) {
      assistantMessage.content = sanitizeHostExecutionError(
        error,
        executionShortcutMode === "undo" ? "Undo failed." : "Redo failed."
      );
      assistantMessage.statusLine = "Execution control failed";
    }

    renderMessages();
    return;
  }

  try {
    const [snapshot, attachments] = await Promise.all([
      getSpreadsheetSnapshot(normalizedPrompt),
      uploadPendingAttachments()
    ]);

    const request = buildRequestEnvelope({
      snapshot,
      userMessage: normalizedPrompt,
      conversation: state.messages,
      attachments
    });

    assistantMessage.requestId = request.requestId;
    const accepted = await gateway.startRun(request);
    assistantMessage.runId = accepted.runId;
    assistantMessage.statusLine = "Thinking...";
    clearPendingAttachments();
    renderMessages();
    await pollRun(assistantMessage);
  } catch (error) {
    assistantMessage.content = sanitizeHostExecutionError(error, "Failed to contact Hermes.");
    assistantMessage.statusLine = "Request failed before Hermes could process it";
    renderMessages();
  }
}

function hydratePrefill() {
  const prefill = Office.context.document.settings.get("hermesPrefillPrompt");
  if (prefill) {
    elements.prompt.value = prefill;
    Office.context.document.settings.remove("hermesPrefillPrompt");
    void saveDocumentSettingsAsync().catch((error) => {
      console.warn("Hermes could not clear the prefill prompt.", error);
    });
  }
}

function attachDragAndDrop() {
  const suppress = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  elements.app.addEventListener("dragover", suppress);
  elements.app.addEventListener("drop", (event) => {
    suppress(event);
    const files = filterSupportedImageFiles(event.dataTransfer?.files || []);
    if (files.length > 0) {
      addPendingFiles(files, "drag_drop");
    }
  });
}

Office.onReady(() => {
  void ensureDemoStartupDefaults();
  hydratePrefill();
  renderAttachmentStrip();
  elements.sendButton.addEventListener("click", () => void sendPrompt());
  elements.fileInput.addEventListener("change", (event) => {
    addPendingFiles(event.target.files || [], "upload");
    event.target.value = "";
  });
  elements.prompt.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void sendPrompt();
    }
  });
  window.addEventListener("paste", (event) => {
    const files = filterSupportedImageFiles(event.clipboardData?.files || []);
    if (files.length > 0) {
      addPendingFiles(files, "clipboard");
    }
  });
  attachDragAndDrop();
});

export {
  applyWritePlan,
  applyWritebackResultToMessage,
  appendStoredMessage,
  appendGatewayIssueSummary,
  buildRequestEnvelope,
  buildWriteApprovalRequest,
  dryRunCompositePlan,
  ensureDemoStartupDefaults,
  getSpreadsheetSnapshot,
  getExcelPlanSupportError,
  getRequiresConfirmation,
  getExecutionShortcutMode,
  getResponseBodyText,
  getResponseMetaLine,
  listExecutionHistory,
  redoExecution,
  bindMessageAutoScrollObservers,
  pollRun,
  renderStructuredPreview,
  scheduleMessagesAutoScroll,
  scrollMessagesToBottom,
  sendPrompt,
  sanitizeHostExecutionError,
  isTraceUnavailablePollError,
  sanitizeConversation,
  pruneStoredMessages,
  trimMessageTraceEvents,
  renderMessages,
  executeWritePlanMessage,
  parseJson as parseGatewayJsonResponse,
  undoExecution,
  getStructuredPreview,
  isWritePlanResponse
};
