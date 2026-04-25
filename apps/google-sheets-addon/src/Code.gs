const HERMES_SIDEBAR_TEMPLATE_PATH = 'html/Sidebar';

function onOpen() {
  removeLegacyMenus_();
  SpreadsheetApp.getUi()
    .createMenu('Hermes')
    .addItem('Open Hermes', 'showHermesSidebar')
    .addItem('Explain Selection', 'menuExplainSelection')
    .addItem('Generate Formula', 'menuGenerateFormula')
    .addItem('Clean Data', 'menuCleanData')
    .addItem('Summarize Sheet', 'menuSummarizeSheet')
    .addItem('Apply Suggested Update', 'menuApplySuggestedUpdate')
    .addItem('Insert From Image', 'menuInsertFromImage')
    .addToUi();
}

function removeLegacyMenus_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    spreadsheet.removeMenu('Hermes');
  } catch (_error) {
    // Ignore when the legacy menu does not exist or cannot be removed.
  }
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showHermesSidebar() {
  const html = HtmlService.createTemplateFromFile(HERMES_SIDEBAR_TEMPLATE_PATH)
    .evaluate()
    .setTitle('Hermes');
  SpreadsheetApp.getUi().showSidebar(html);
}

function setPrefillPrompt_(prompt) {
  PropertiesService.getDocumentProperties().setProperty('HERMES_PREFILL_PROMPT', prompt);
}

function consumePrefillPrompt() {
  const properties = PropertiesService.getDocumentProperties();
  const prompt = properties.getProperty('HERMES_PREFILL_PROMPT');
  if (prompt) {
    properties.deleteProperty('HERMES_PREFILL_PROMPT');
  }
  return prompt || '';
}

function menuExplainSelection() {
  setPrefillPrompt_('Explain the current selection.');
  showHermesSidebar();
}

function menuGenerateFormula() {
  setPrefillPrompt_('Suggest a formula for the current selection.');
  showHermesSidebar();
}

function menuCleanData() {
  setPrefillPrompt_('Suggest a safe cleanup plan for the current selection.');
  showHermesSidebar();
}

function menuSummarizeSheet() {
  setPrefillPrompt_('Summarize the active sheet.');
  showHermesSidebar();
}

function menuApplySuggestedUpdate() {
  setPrefillPrompt_('Review the last suggested update and prepare it for confirmation.');
  showHermesSidebar();
}

function menuInsertFromImage() {
  setPrefillPrompt_('Extract the attached table image and prepare an insert preview.');
  showHermesSidebar();
}

function getRuntimeConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const forceExtractionMode = scriptProperties.getProperty('HERMES_FORCE_EXTRACTION_MODE');
  const deploymentOverrides = resolveHermesDeploymentOverrides_();

  return {
    gatewayBaseUrl:
      deploymentOverrides.gatewayBaseUrl ||
      scriptProperties.getProperty('HERMES_GATEWAY_URL') ||
      '',
    clientVersion:
      deploymentOverrides.clientVersion ||
      scriptProperties.getProperty('HERMES_CLIENT_VERSION') ||
      'google-sheets-addon-live-demo',
    reviewerSafeMode:
      typeof deploymentOverrides.reviewerSafeMode === 'boolean'
        ? deploymentOverrides.reviewerSafeMode
        : scriptProperties.getProperty('HERMES_REVIEWER_SAFE_MODE') === 'true',
    forceExtractionMode:
      deploymentOverrides.forceExtractionMode === 'real' ||
      deploymentOverrides.forceExtractionMode === 'demo' ||
      deploymentOverrides.forceExtractionMode === 'unavailable'
        ? deploymentOverrides.forceExtractionMode
        : forceExtractionMode === 'real' ||
            forceExtractionMode === 'demo' ||
            forceExtractionMode === 'unavailable'
          ? forceExtractionMode
        : null
  };
}

function resolveHermesDeploymentOverrides_() {
  try {
    if (typeof getHermesDeploymentOverrides === 'function') {
      var overrides = getHermesDeploymentOverrides();
      if (overrides && typeof overrides === 'object' && !Array.isArray(overrides)) {
        return overrides;
      }
    }
  } catch (_error) {
    // Ignore invalid deployment override helpers and fall back to script properties.
  }

  return {};
}

function getWorkbookSessionKey() {
  const spreadsheet = SpreadsheetApp.getActive();
  return 'google_sheets::' + spreadsheet.getId();
}

function buildGatewayUrl_(path) {
  const baseUrl = String(getRuntimeConfig().gatewayBaseUrl || '').replace(/\/+$/, '');
  const normalizedPath = String(path || '').trim();

  if (!baseUrl) {
    throw new Error('Hermes gateway URL is not configured.');
  }

  if (!normalizedPath) {
    return baseUrl;
  }

  return normalizedPath.charAt(0) === '/'
    ? baseUrl + normalizedPath
    : baseUrl + '/' + normalizedPath;
}

function isAppsScriptReachableGatewayBaseUrl_(baseUrl) {
  var value = String(baseUrl || '').trim();
  var match = value.match(/^(https?):\/\/(\[[^\]]+\]|[^\/:?#]+)(?::\d+)?(?:[\/?#]|$)/i);
  if (!match) {
    return false;
  }

  var protocol = String(match[1] || '').toLowerCase();
  var host = String(match[2] || '').toLowerCase();
  if (protocol !== 'https') {
    return false;
  }
  if (!host) {
    return false;
  }

  if (host === 'localhost' || host === '127.0.0.1' || host === '0.0.0.0') {
    return false;
  }

  if (/^10\./.test(host) || /^192\.168\./.test(host) || /^172\.(1[6-9]|2\d|3[0-1])\./.test(host)) {
    return false;
  }

  var decodeMappedIpv4 = function(suffix) {
    if (/^\d{1,3}(?:\.\d{1,3}){3}$/.test(suffix)) {
      return suffix;
    }

    var hexMatch = suffix.match(/^([0-9a-f]{1,4}):([0-9a-f]{1,4})$/i);
    if (!hexMatch) {
      return null;
    }

    var high = parseInt(hexMatch[1], 16);
    var low = parseInt(hexMatch[2], 16);
    if (isNaN(high) || isNaN(low)) {
      return null;
    }

    return [
      (high >> 8) & 0xff,
      high & 0xff,
      (low >> 8) & 0xff,
      low & 0xff
    ].join('.');
  };

  var normalizedIpv6Host = host.replace(/^\[|\]$/g, '');
  if (normalizedIpv6Host.indexOf(':') !== -1) {
    if (normalizedIpv6Host === '::' || normalizedIpv6Host === '::1') {
      return false;
    }

    var mappedIpv4 = /^::ffff:/i.test(normalizedIpv6Host)
      ? decodeMappedIpv4(normalizedIpv6Host.replace(/^::ffff:/i, ''))
      : null;
    if (mappedIpv4) {
      if (
        mappedIpv4 === '127.0.0.1' ||
        mappedIpv4 === '0.0.0.0' ||
        /^10\./.test(mappedIpv4) ||
        /^192\.168\./.test(mappedIpv4) ||
        /^172\.(1[6-9]|2\d|3[0-1])\./.test(mappedIpv4)
      ) {
        return false;
      }
    }

    var firstHextet = normalizedIpv6Host
      .split(':')
      .find(function(segment) {
        return segment.length > 0;
      });
    if (firstHextet) {
      var firstValue = parseInt(firstHextet, 16);
      if (!isNaN(firstValue)) {
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

function extractGatewayErrorMessage_(statusCode, bodyText) {
  const fallback = 'Hermes gateway request failed with ' + statusCode + '.';
  if (!bodyText) {
    return fallback;
  }

  try {
    const parsed = JSON.parse(bodyText);
    const formatMessage = function(message, userAction) {
      if (typeof userAction === 'string' && userAction.trim() && userAction.trim() !== message) {
        return message + '\n\n' + userAction.trim();
      }
      return message;
    };
    if (parsed && parsed.error && typeof parsed.error.message === 'string' && parsed.error.message.trim()) {
      return formatMessage(parsed.error.message.trim(), parsed.error.userAction);
    }
    if (parsed && typeof parsed.error === 'string' && parsed.error.trim()) {
      return parsed.error;
    }
    if (parsed && typeof parsed.message === 'string' && parsed.message.trim()) {
      return formatMessage(parsed.message.trim(), parsed.userAction);
    }
  } catch (_error) {
    // Fall back to the raw text when the gateway does not return JSON.
  }

  return bodyText;
}

function formatUserFacingErrorText_(message, userAction) {
  const resolvedMessage = String(message || '').trim();
  const resolvedUserAction = typeof userAction === 'string' ? userAction.trim() : '';

  if (!resolvedUserAction || resolvedUserAction === resolvedMessage) {
    return resolvedMessage;
  }

  return resolvedMessage + '\n\n' + resolvedUserAction;
}

function sanitizeHostExecutionError_(error, fallbackMessage) {
  const rawMessage = error && error.message ? error.message : String(error || '');
  const message = String(rawMessage || '').trim().replace(/^Error:\s*/i, '');

  if (!message) {
    return fallbackMessage || 'Write-back failed.';
  }

  if (/Hermes gateway URL is not configured/i.test(message)) {
    return formatUserFacingErrorText_(
      'The Hermes connection is not configured for this sheet.',
      'Set the Hermes gateway URL, reload the sidebar, and retry.'
    );
  }

  if (
    /Hermes gateway returned invalid JSON/i.test(message) ||
    /structured gateway contract/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'The Hermes service returned a response the sidebar could not use.',
      'Retry the request. If it keeps happening, reload the sidebar or check the Hermes gateway.'
    );
  }

  if (/Hermes gateway proxy requires a request path/i.test(message)) {
    return formatUserFacingErrorText_(
      'The Hermes request could not be sent correctly.',
      'Retry the action. If it keeps happening, reload the sidebar and try again.'
    );
  }

  if (/Failed to fetch/i.test(message) || /NetworkError/i.test(message)) {
    return formatUserFacingErrorText_(
      'The Hermes service could not be reached.',
      'Check the network connection or Hermes gateway, then retry.'
    );
  }

  if (/Destructive confirmation is unavailable in this host/i.test(message)) {
    return formatUserFacingErrorText_(
      'This spreadsheet app cannot approve destructive changes inline.',
      'Ask Hermes for a safer alternative, or use a non-destructive step first.'
    );
  }

  const targetSheetMatch = message.match(/^Target sheet not found:\s*(.+)$/i);
  if (targetSheetMatch) {
    return formatUserFacingErrorText_(
      'Sheet "' + targetSheetMatch[1].trim() + '" was not found.',
      'Create or select that sheet, then retry.'
    );
  }

  const sourceSheetMatch = message.match(/^(?:Validation )?Source sheet not found:\s*(.+)$/i);
  if (sourceSheetMatch) {
    return formatUserFacingErrorText_(
      'Sheet "' + sourceSheetMatch[1].trim() + '" was not found.',
      'Select a valid source sheet, then retry.'
    );
  }

  const namedRangeMatch = message.match(/^Named range not found:\s*(.+)$/i);
  if (namedRangeMatch) {
    return formatUserFacingErrorText_(
      'Named range "' + namedRangeMatch[1].trim() + '" was not found.',
      'Check the range name or create it first, then retry.'
    );
  }

  const invalidRangeMatch = message.match(/^Unsupported A1 reference:\s*(.+)$/i);
  if (invalidRangeMatch) {
    return formatUserFacingErrorText_(
      'Range "' + invalidRangeMatch[1].trim() + '" is not a valid A1 reference.',
      'Use a valid cell or range address, then retry.'
    );
  }

  const duplicateHeaderMatch = message.match(/duplicate header:\s*(.+?)\.?$/i);
  if (duplicateHeaderMatch) {
    return formatUserFacingErrorText_(
      'Column "' + duplicateHeaderMatch[1].trim() + '" appears more than once in the header row.',
      'Rename duplicate columns or select a table with unique headers, then retry.'
    );
  }

  if (/requires a header row/i.test(message)) {
    return formatUserFacingErrorText_(
      'This action needs a table with a header row.',
      'Select or create a table with column headers, then retry.'
    );
  }

  const missingHeaderFieldMatch = message.match(/cannot find (?:pivot|chart) field in header row:\s*(.+?)\.?$/i);
  if (missingHeaderFieldMatch) {
    return formatUserFacingErrorText_(
      'Column "' + missingHeaderFieldMatch[1].trim() + '" was not found in the header row.',
      'Select the full table with headers, or use the exact column name in the request and retry.'
    );
  }

  if (
    /could not resolve any valid sort keys/i.test(message) ||
    /could not resolve a filter column inside the target range/i.test(message) ||
    /Column .* is outside /i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'The selected range does not include the columns this step needs.',
      'Select the full table, or update the request to use columns inside the chosen range.'
    );
  }

  const invalidDateMatch = message.match(/^Invalid date literal:\s*(.+)$/i);
  if (invalidDateMatch) {
    return formatUserFacingErrorText_(
      'The date "' + invalidDateMatch[1].trim() + '" is not valid.',
      'Use a real calendar date such as 2026-04-22, then retry.'
    );
  }

  if (
    /Unsupported filter operator/i.test(message) ||
    /grid filters cannot represent operator "topN" exactly/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This filter condition is not supported here.',
      'Try a simpler operator such as equals or contains, or ask Hermes for a different filter.'
    );
  }

  if (
    /Unsupported filter combiner/i.test(message) ||
    /filter combiners other than and/i.test(message) ||
    /multiple conditions for the same column/i.test(message) ||
    /cannot represent combiner "or" exactly/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This spreadsheet app cannot combine those filter conditions in one exact step.',
      'Use a single filter rule per column, or split the filter into smaller steps.'
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
    return formatUserFacingErrorText_(
      'This named range action is not supported in this spreadsheet app.',
      'Use a workbook-level named range or ask Hermes for a simpler named range update.'
    );
  }

  if (
    /Named range create and retarget require targetSheet and targetRange/i.test(message) ||
    /Named range rename requires newName/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This named range request is missing required details.',
      'Include the destination sheet and range, or provide the new name, then retry.'
    );
  }

  if (/approved targetRange does not match/i.test(message)) {
    return formatUserFacingErrorText_(
      'The spreadsheet changed, so the approved destination no longer matches the intended shape.',
      'Refresh the spreadsheet state and run the request again.'
    );
  }

  if (
    /cannot append exactly when the approved target range contains internal gaps/i.test(message) ||
    /cannot append exactly within the approved target range/i.test(message) ||
    /cannot split this column exactly within the approved target range/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'The chosen destination range cannot accept this write safely.',
      'Choose a clean target range or ask Hermes to write into a blank area.'
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
    return formatUserFacingErrorText_(
      'This transfer cannot be applied safely on the current source and destination ranges.',
      'Choose a simpler target range or ask Hermes for a different transfer plan.'
    );
  }

  if (/does not support exact-safe pivot table creation yet/i.test(message)) {
    return formatUserFacingErrorText_(
      'This spreadsheet app cannot create that pivot table safely yet.',
      'Ask for a preview only, or target a simpler transformation first.'
    );
  }

  if (/does not support exact-safe chart/i.test(message)) {
    return formatUserFacingErrorText_(
      'This spreadsheet app cannot create that chart safely yet.',
      'Ask for a preview only, or request a simpler supported chart.'
    );
  }

  if (
    /does not support exact-safe formula transfers on this range/i.test(message) ||
    /does not support exact-safe format transfers on this range/i.test(message) ||
    /does not support exact format append transfers on this range/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This spreadsheet app cannot apply that transfer safely on the current range.',
      'Try a simpler target range or ask Hermes for a direct cell update instead.'
    );
  }

  if (
    /does not support exact-safe cleanup semantics/i.test(message) ||
    /cannot apply cleanup plans exactly when the target range contains formulas/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This cleanup action cannot be applied safely on the current range.',
      'Try a narrower range or ask Hermes for a simpler cleanup step.'
    );
  }

  if (
    /Unsupported .*data validation rule type/i.test(message) ||
    /Unsupported .*validation comparator/i.test(message) ||
    /Unsupported invalidDataBehavior/i.test(message) ||
    /List validation requires values, sourceRange, or namedRangeName/i.test(message) ||
    /Custom formula validation requires/i.test(message) ||
    /checkbox .* only support boolean true\/false/i.test(message) ||
    /cannot represent allowBlank/i.test(message) ||
    /uncheckedValue without checkedValue/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This validation setup cannot be represented safely here.',
      'Try a simpler dropdown, checkbox, or date rule, then retry.'
    );
  }

  if (
    /requires a valid target range for /i.test(message) ||
    /requires a single-cell target anchor for /i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This action needs a valid destination cell or anchor.',
      'Choose a single target cell or a valid destination range, then retry.'
    );
  }

  if (
    /conditional-format/i.test(message) ||
    /conditional formatting/i.test(message) ||
    /text_contains conditional formatting/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This conditional formatting step is not supported here.',
      'Try a simpler highlight rule, or ask Hermes for a preview-only result first.'
    );
  }

  if (
    /does not support data validation on this range/i.test(message) ||
    /does not support checkbox cell controls on this range/i.test(message) ||
    /does not expose checkbox cell control support/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This validation action cannot run on the current range.',
      'Choose a standard editable cell range, then retry.'
    );
  }

  if (
    /does not support range sort on this selection/i.test(message) ||
    /does not support range filters on this selection/i.test(message) ||
    /does not support conditional formatting on this range/i.test(message) ||
    /does not support conditional formatting on this sheet/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This action cannot run on the current selection.',
      'Choose a standard table or cell range, then retry.'
    );
  }

  if (/Cannot hide the only visible worksheet/i.test(message)) {
    return formatUserFacingErrorText_(
      'At least one worksheet must stay visible.',
      'Keep another sheet visible or unhide one first, then retry.'
    );
  }

  if (
    /Unsupported workbook structure update/i.test(message) ||
    /Unsupported sheet structure update/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This sheet change is not supported in this spreadsheet app.',
      'Ask Hermes for a simpler sheet change, or try a different supported operation.'
    );
  }

  if (
    /pivot/i.test(message) &&
    (
      /does not support/i.test(message) ||
      /requires /i.test(message) ||
      /Unsupported pivot aggregation/i.test(message) ||
      /only supports equal_to pivot filters/i.test(message) ||
      /pivot filter criteria builders/i.test(message) ||
      /does not expose pivot creation/i.test(message)
    )
  ) {
    return formatUserFacingErrorText_(
      'This pivot configuration is not supported here yet.',
      'Try a simpler pivot, or ask Hermes for a preview-only result first.'
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
    return formatUserFacingErrorText_(
      'This chart configuration is not supported here yet.',
      'Try a simpler chart, or ask Hermes for a preview-only result first.'
    );
  }

  if (/Target range already contains content/i.test(message)) {
    return formatUserFacingErrorText_(
      'The destination already contains data.',
      'Clear that range or choose a blank destination, then retry.'
    );
  }

  if (/chat-only analysis reports are not writeback eligible/i.test(message)) {
    return formatUserFacingErrorText_(
      'This result is analysis only and cannot be applied directly.',
      'Ask Hermes to turn it into a specific writeback on a sheet or range.'
    );
  }

  if (
    /Composite workflow execution requires executionId/i.test(message) ||
    /Dependency .* has not completed before this step/i.test(message)
  ) {
    return formatUserFacingErrorText_(
      'This workflow is no longer valid for the current spreadsheet state.',
      'Run the request again so Hermes can rebuild the workflow before applying it.'
    );
  }

  return message;
}

function parseGatewayJsonResponse_(response) {
  const statusCode = response.getResponseCode();
  const bodyText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error(extractGatewayErrorMessage_(statusCode, bodyText));
  }

  if (!bodyText) {
    return {};
  }

  try {
    return JSON.parse(bodyText);
  } catch (_error) {
    throw new Error('Hermes gateway returned invalid JSON.');
  }
}

function proxyGatewayJson(input) {
  if (!input || typeof input.path !== 'string' || input.path.trim().length === 0) {
    throw new Error('Hermes gateway proxy requires a request path.');
  }

  const method = typeof input.method === 'string' && input.method.trim().length > 0
    ? input.method.toLowerCase()
    : 'get';
  const headers = {};
  const sourceHeaders = input.headers && typeof input.headers === 'object' ? input.headers : {};

  Object.keys(sourceHeaders).forEach(function(key) {
    const value = sourceHeaders[key];
    if (value !== undefined && value !== null) {
      headers[key] = String(value);
    }
  });

  const requestOptions = {
    method: method,
    muteHttpExceptions: true,
    headers: headers
  };

  if (input.body !== undefined && input.body !== null) {
    requestOptions.payload = typeof input.body === 'string'
      ? input.body
      : JSON.stringify(input.body);

    if (Object.prototype.hasOwnProperty.call(headers, 'content-type')) {
      requestOptions.contentType = headers['content-type'];
      delete headers['content-type'];
    } else if (Object.prototype.hasOwnProperty.call(headers, 'Content-Type')) {
      requestOptions.contentType = headers['Content-Type'];
      delete headers['Content-Type'];
    } else if (typeof input.body !== 'string') {
      requestOptions.contentType = 'application/json';
    }
  }

  return parseGatewayJsonResponse_(UrlFetchApp.fetch(buildGatewayUrl_(input.path), requestOptions));
}

function uploadGatewayImageAttachment(input) {
  if (!input || typeof input.fileName !== 'string' || input.fileName.trim().length === 0) {
    throw new Error('Hermes image upload requires a file name.');
  }
  if (!input.base64Data || typeof input.base64Data !== 'string') {
    throw new Error('Hermes image upload requires base64 image data.');
  }
  if (typeof input.sessionId !== 'string' || input.sessionId.trim().length === 0) {
    throw new Error('Hermes image upload requires a session id.');
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var workbookId = spreadsheet && typeof spreadsheet.getId === 'function'
    ? spreadsheet.getId()
    : '';
  if (typeof workbookId !== 'string' || workbookId.trim().length === 0) {
    throw new Error('Hermes image upload requires a workbook id.');
  }

  var runtimeConfig = getRuntimeConfig();
  if (!isAppsScriptReachableGatewayBaseUrl_(runtimeConfig.gatewayBaseUrl)) {
    throw new Error(
      'Google Sheets image upload requires a public Hermes gateway URL.\n\n' +
      'Set HERMES_GATEWAY_URL to a reachable HTTPS or public address, then retry the upload.'
    );
  }

  const mimeType = typeof input.mimeType === 'string' && input.mimeType.trim().length > 0
    ? input.mimeType
    : 'application/octet-stream';
  const normalizedBase64 = input.base64Data.replace(/^data:[^;]+;base64,/, '');
  const blob = Utilities.newBlob(
    Utilities.base64Decode(normalizedBase64),
    mimeType,
    input.fileName
  );

  return parseGatewayJsonResponse_(UrlFetchApp.fetch(buildGatewayUrl_('/api/uploads/image'), {
    method: 'post',
    muteHttpExceptions: true,
    payload: {
      file: blob,
      source: typeof input.source === 'string' && input.source.trim().length > 0
        ? input.source
        : 'upload',
      sessionId: input.sessionId.trim(),
      workbookId: workbookId.trim()
    }
  }));
}

function normalizeA1_(address) {
  return String(address || '').replace(/\$/g, '');
}

const MAX_CONTEXT_CELL_TEXT_LENGTH_ = 4000;
const MAX_CONTEXT_HEADER_TEXT_LENGTH_ = 256;
const MAX_CONTEXT_FORMULA_LENGTH_ = 16000;
const MAX_CONTEXT_NOTE_LENGTH_ = 4000;

function truncateContextText_(value, maxLength) {
  const text = String(value == null ? '' : value);
  if (text.length <= maxLength) {
    return text;
  }

  return text.slice(0, maxLength - 1) + '…';
}

function normalizeFormulas_(formulas) {
  return formulas.map(function(row) {
    return row.map(function(cell) {
      return typeof cell === 'string' && cell.trim().length > 0
        ? truncateContextText_(cell, MAX_CONTEXT_FORMULA_LENGTH_)
        : null;
    });
  });
}

function normalizeCellValue_(value, displayValue) {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return typeof value === 'string'
      ? truncateContextText_(value, MAX_CONTEXT_CELL_TEXT_LENGTH_)
      : value;
  }

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return truncateContextText_(displayValue || String(value), MAX_CONTEXT_CELL_TEXT_LENGTH_);
  }

  return truncateContextText_(displayValue || String(value), MAX_CONTEXT_CELL_TEXT_LENGTH_);
}

function normalizeMatrixValues_(values, displayValues) {
  return values.map(function(row, rowIndex) {
    const displayRow = displayValues[rowIndex] || [];

  return row.map(function(cell, columnIndex) {
      return normalizeCellValue_(cell, displayRow[columnIndex] || '');
    });
  });
}

function getHeaders_(values) {
  const firstRow = values[0] || [];
  if (firstRow.length === 0) {
    return null;
  }

  const areHeaders = firstRow.every(function(cell) {
    return typeof cell === 'string' && cell.trim().length > 0;
  });

  return areHeaders ? firstRow.map(function(cell) {
    return truncateContextText_(cell, MAX_CONTEXT_HEADER_TEXT_LENGTH_);
  }) : null;
}

const ANALYSIS_OUTPUT_MODES_ = [
  'chat_only',
  'materialize_report'
];

const CHART_TYPES_ = [
  'bar',
  'column',
  'stacked_bar',
  'stacked_column',
  'line',
  'area',
  'pie',
  'scatter'
];

function joinList_(values) {
  return Array.isArray(values) && values.length > 0 ? values.join(', ') : '';
}

function padMatrixRows_(rows) {
  const width = Math.max(1, ...rows.map(function(row) {
    return row.length;
  }));

  return rows.map(function(row) {
    return Array.from({ length: width }, function(_unused, index) {
      return row[index] !== undefined ? row[index] : '';
    });
  });
}

function isAnalysisReportPlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === 'string' &&
    typeof plan.sourceRange === 'string' &&
    ANALYSIS_OUTPUT_MODES_.indexOf(plan.outputMode) !== -1 &&
    Array.isArray(plan.sections) &&
    plan.sections.length > 0
  );
}

function isMaterializedAnalysisReportPlan_(plan) {
  return isAnalysisReportPlan_(plan) &&
    plan.outputMode === 'materialize_report' &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string';
}

function isPivotTablePlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === 'string' &&
    typeof plan.sourceRange === 'string' &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string' &&
    Array.isArray(plan.rowGroups) &&
    plan.rowGroups.length > 0 &&
    Array.isArray(plan.valueAggregations) &&
    plan.valueAggregations.length > 0
  );
}

function isChartPlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === 'string' &&
    typeof plan.sourceRange === 'string' &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string' &&
    CHART_TYPES_.indexOf(plan.chartType) !== -1 &&
    Array.isArray(plan.series) &&
    plan.series.length > 0
  );
}

function isExternalDataPlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceType === 'string' &&
    typeof plan.provider === 'string' &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string' &&
    typeof plan.formula === 'string' &&
    plan.formula.trim().indexOf('=') === 0
  );
}

function isCompositePlan_(plan) {
  return Boolean(
    plan &&
    Array.isArray(plan.steps) &&
    plan.steps.length > 0 &&
    typeof plan.explanation === 'string'
  );
}

function getCompositeStatusSummary_(result) {
  const resolvedSummary = result && typeof result.summary === 'string' ? result.summary.trim() : '';
  const stepResults = Array.isArray(result && result.stepResults) ? result.stepResults : [];
  const count = stepResults.length;
  const baseSummary = resolvedSummary || ('Completed workflow with ' + count + ' step' + (count === 1 ? '' : 's') + '.');
  const detailSummary = buildCompositeResultDetailSummary_(stepResults);

  if (!detailSummary) {
    return baseSummary;
  }

  if (isGenericCompositeSummary_(baseSummary)) {
    return (baseSummary + ' ' + detailSummary).trim();
  }

  return baseSummary;
}

function isGenericCompositeSummary_(summary) {
  const normalized = String(summary || '').trim();
  if (!normalized) {
    return true;
  }

  return /^Workflow finished:/i.test(normalized) ||
    /^Completed workflow with \d+ step/i.test(normalized);
}

function normalizeCompositeStepSummary_(summary) {
  return String(summary || '').trim().replace(/\s+/g, ' ').replace(/[.]+$/g, '').trim();
}

function joinCompositeSummaryParts_(parts, limit) {
  if (parts.length === 0) {
    return '';
  }

  const visible = parts.slice(0, limit);
  const hiddenCount = parts.length - visible.length;
  const joined = visible.join('; ');
  return hiddenCount > 0 ? joined + '; +' + hiddenCount + ' more' : joined;
}

function buildCompositeResultDetailSummary_(stepResults) {
  const completed = [];
  const failed = [];
  const skipped = [];

  stepResults.forEach(function(step) {
    const normalized = normalizeCompositeStepSummary_(step && step.summary);
    if (!normalized) {
      return;
    }

    if (step.status === 'completed') {
      completed.push(normalized);
      return;
    }

    if (step.status === 'failed') {
      failed.push(normalized);
      return;
    }

    if (step.status === 'skipped') {
      skipped.push(normalized);
    }
  });

  const details = [];
  if (completed.length > 0) {
    details.push('Completed: ' + joinCompositeSummaryParts_(completed, 3) + '.');
  }
  if (failed.length > 0) {
    details.push('Failed: ' + joinCompositeSummaryParts_(failed, 2) + '.');
  }
  if (skipped.length > 0) {
    details.push('Skipped: ' + joinCompositeSummaryParts_(skipped, 2) + '.');
  }

  return details.join(' ');
}

function buildCompositeExecutionSummary_(stepResults) {
  const completedCount = stepResults.filter(function(step) {
    return step.status === 'completed';
  }).length;
  const failedCount = stepResults.filter(function(step) {
    return step.status === 'failed';
  }).length;
  const skippedCount = stepResults.filter(function(step) {
    return step.status === 'skipped';
  }).length;
  const parts = [String(stepResults.length) + ' step' + (stepResults.length === 1 ? '' : 's')];

  if (completedCount > 0) {
    parts.push(String(completedCount) + ' completed');
  }
  if (failedCount > 0) {
    parts.push(String(failedCount) + ' failed');
  }
  if (skippedCount > 0) {
    parts.push(String(skippedCount) + ' skipped');
  }

  return 'Workflow finished: ' + parts.join(' • ') + '.';
}

function getCompositeStepWritebackStatusLine_(plan, result) {
  if (result && result.kind === 'range_write') {
    const targetSheet = plan && plan.targetSheet ? plan.targetSheet : result.targetSheet;
    const targetRange = plan && plan.targetRange ? plan.targetRange : result.targetRange;
    const target = targetSheet && targetRange ? targetSheet + '!' + targetRange : '';
    const hasValues = plan && Array.isArray(plan.values);
    const hasFormulas = plan && Array.isArray(plan.formulas);
    const hasNotes = plan && Array.isArray(plan.notes);

    if (plan && plan.sourceAttachmentId && target) {
      return 'Inserted imported data into ' + target + '.';
    }

    if (hasFormulas && !hasValues && !hasNotes && target) {
      return result.writtenRows === 1 && result.writtenColumns === 1
        ? 'Set a formula in ' + target + '.'
        : 'Set formulas in ' + target + '.';
    }

    if (hasValues && !hasFormulas && !hasNotes && target) {
      return result.writtenRows === 1 && result.writtenColumns === 1
        ? 'Wrote a value to ' + target + '.'
        : 'Wrote values to ' + target + '.';
    }

    if (hasNotes && !hasValues && !hasFormulas && target) {
      return 'Updated notes in ' + target + '.';
    }

    if (target && (hasValues || hasFormulas || hasNotes)) {
      return 'Updated cells in ' + target + '.';
    }
  }

  return result && result.summary ? result.summary : 'Completed workflow step.';
}

function buildAnalysisReportMatrix_(plan) {
  const rows = [
    ['Analysis report'],
    ['Source sheet', plan.sourceSheet],
    ['Source range', plan.sourceRange],
    ['Section', 'Title', 'Summary', 'Source ranges']
  ].concat(plan.sections.map(function(section) {
    return [
      section.type,
      section.title,
      section.summary,
      joinList_(section.sourceRanges)
    ];
  }));

  return padMatrixRows_(rows);
}

function normalizeAnalysisReportAffectedRanges_(plan, resolvedTargetRange) {
  const resolvedTargetRef = plan.targetSheet + '!' + resolvedTargetRange;
  const anchorTargetRef = plan.targetSheet + '!' + plan.targetRange;
  const normalizedRanges = Array.isArray(plan.affectedRanges)
    ? plan.affectedRanges.map(function(range) {
        return range === anchorTargetRef ? resolvedTargetRef : range;
      })
    : [];

  if (normalizedRanges.indexOf(resolvedTargetRef) === -1) {
    normalizedRanges.push(resolvedTargetRef);
  }

  return normalizedRanges;
}

function buildHeaderMap_(sourceRange) {
  const headerRow = sourceRange && typeof sourceRange.getDisplayValues === 'function'
    ? sourceRange.getDisplayValues()[0] || []
    : [];
  const headerMap = {};

  headerRow.forEach(function(value, index) {
    const key = String(value || '').trim();
    if (!key) {
      return;
    }

    if (Object.prototype.hasOwnProperty.call(headerMap, key)) {
      throw new Error('Google Sheets host requires unique headers, but found duplicate header: ' + key + '.');
    }

    headerMap[key] = index + 1;
  });

  if (Object.keys(headerMap).length === 0) {
    throw new Error('Google Sheets host requires a header row for pivot tables.');
  }

  return headerMap;
}

function requireSingleCellAnchor_(range, kind) {
  if (!range || typeof range.getNumRows !== 'function' || typeof range.getNumColumns !== 'function') {
    throw new Error('Google Sheets host requires a valid target range for ' + kind + '.');
  }

  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    throw new Error('Google Sheets host requires a single-cell target anchor for ' + kind + '.');
  }
}

function getPivotSummarizeFunction_(aggregation) {
  switch (aggregation) {
    case 'sum':
      return SpreadsheetApp.PivotTableSummarizeFunction.SUM;
    case 'count':
      return SpreadsheetApp.PivotTableSummarizeFunction.COUNTA;
    case 'average':
      return SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE;
    case 'min':
      return SpreadsheetApp.PivotTableSummarizeFunction.MIN;
    case 'max':
      return SpreadsheetApp.PivotTableSummarizeFunction.MAX;
    default:
      throw new Error('Unsupported pivot aggregation: ' + aggregation);
  }
}

function getPivotFilterOperatorMethods_(operator, value) {
  const isNumber = typeof value === 'number';

  switch (operator) {
    case 'equal_to':
      return isNumber
        ? { primary: 'whenNumberEqualTo' }
        : { primary: 'whenTextEqualTo' };
    case 'not_equal_to':
      return isNumber
        ? { primary: 'whenNumberNotEqualTo' }
        : { primary: 'whenTextNotEqualTo' };
    case 'greater_than':
      return { primary: 'whenNumberGreaterThan', requiresNumber: true };
    case 'greater_than_or_equal_to':
      return { primary: 'whenNumberGreaterThanOrEqualTo', requiresNumber: true };
    case 'less_than':
      return { primary: 'whenNumberLessThan', requiresNumber: true };
    case 'less_than_or_equal_to':
      return { primary: 'whenNumberLessThanOrEqualTo', requiresNumber: true };
    default:
      throw new Error('Unsupported pivot filter operator: ' + operator);
  }
}

function coercePivotFilterNumber_(value, operator) {
  if (typeof value === 'number') {
    return value;
  }

  if (typeof value === 'string' && value.trim() !== '' && !Number.isNaN(Number(value))) {
    return Number(value);
  }

  throw new Error('Google Sheets host requires numeric pivot filter values for operator ' + operator + '.');
}

function buildPivotFilterCriteria_(filter) {
  const builder = SpreadsheetApp.newFilterCriteria();
  const methods = getPivotFilterOperatorMethods_(filter && filter.operator, filter && filter.value);
  const filterValue = methods.requiresNumber
    ? coercePivotFilterNumber_(filter && filter.value, filter && filter.operator)
    : filter && filter.value;

  if (!builder || typeof builder[methods.primary] !== 'function') {
    throw new Error('Google Sheets host does not support pivot filter criteria builders.');
  }

  builder[methods.primary](filterValue);
  return builder.build();
}

function getPivotHeaderIndex_(headerMap, field) {
  const key = String(field || '').trim();
  const columnIndex = headerMap[key];

  if (!columnIndex) {
    throw new Error('Google Sheets host cannot find pivot field in header row: ' + field + '.');
  }

  return columnIndex;
}

function preflightPivotTableStructure_(plan) {
  const rowGroups = Array.isArray(plan.rowGroups) ? plan.rowGroups : [];
  const columnGroups = Array.isArray(plan.columnGroups) ? plan.columnGroups : [];
  const valueAggregations = Array.isArray(plan.valueAggregations) ? plan.valueAggregations : [];
  const filters = Array.isArray(plan.filters) ? plan.filters : [];

  if (plan.filters && !Array.isArray(plan.filters)) {
    throw new Error('Google Sheets host requires filters to be an array when present.');
  }

  if (plan.columnGroups && !Array.isArray(plan.columnGroups)) {
    throw new Error('Google Sheets host requires columnGroups to be an array when present.');
  }

  valueAggregations.forEach(function(aggregation) {
    getPivotSummarizeFunction_(aggregation && aggregation.aggregation);
  });

  filters.forEach(function(filter) {
    getPivotFilterOperatorMethods_(filter && filter.operator, filter && filter.value);
  });

  if (plan.sort) {
    const groupedFields = rowGroups.concat(columnGroups).map(function(field) {
      return String(field || '').trim();
    });
    const aggregatedFields = valueAggregations.map(function(aggregation) {
      return String(aggregation && aggregation.field || '').trim();
    });
    const sortField = String(plan.sort.field || '').trim();

    if (plan.sort.sortOn === 'group_field') {
      if (groupedFields.indexOf(sortField) === -1) {
        throw new Error('Google Sheets host can only sort an existing pivot row or column group.');
      }
    } else if (plan.sort.sortOn === 'aggregated_value') {
      if (aggregatedFields.indexOf(sortField) === -1) {
        throw new Error('Google Sheets host can only sort by an existing pivot value field.');
      }

      if (rowGroups.length > 0 && columnGroups.length > 0) {
        throw new Error('Google Sheets host cannot sort pivot values when both row and column groups are present exactly.');
      }
    }
  }

  return {
    rowGroups: rowGroups,
    columnGroups: columnGroups,
    valueAggregations: valueAggregations,
    filters: filters,
    sort: plan.sort
  };
}

function validatePivotTableFieldReferences_(headerMap, planState) {
  planState.rowGroups.forEach(function(field) {
    getPivotHeaderIndex_(headerMap, field);
  });

  planState.columnGroups.forEach(function(field) {
    getPivotHeaderIndex_(headerMap, field);
  });

  planState.valueAggregations.forEach(function(aggregation) {
    getPivotHeaderIndex_(headerMap, aggregation && aggregation.field);
    getPivotSummarizeFunction_(aggregation && aggregation.aggregation);
  });

  planState.filters.forEach(function(filter) {
    getPivotHeaderIndex_(headerMap, filter.field);
  });
}

function buildPivotFilterCriteriaList_(planState) {
  return planState.filters.map(function(filter) {
    return {
      field: filter.field,
      fieldIndex: null,
      criteria: buildPivotFilterCriteria_(filter)
    };
  });
}

function getRequiredPivotMethods_(planState) {
  const requiredMethods = ['addRowGroup', 'addPivotValue'];

  if (planState.columnGroups.length > 0) {
    requiredMethods.push('addColumnGroup');
  }

  if (planState.filters.length > 0) {
    requiredMethods.push('addFilter');
  }

  if (planState.sort) {
    requiredMethods.push('getRowGroups');
    requiredMethods.push('getColumnGroups');
  }

  return requiredMethods;
}

function validatePivotTableCapabilities_(planState) {
  const requiredMethods = getRequiredPivotMethods_(planState);

  if (planState.filters.length > 0) {
    if (typeof SpreadsheetApp.newFilterCriteria !== 'function') {
      throw new Error('Google Sheets host does not support pivot filter criteria builders.');
    }

    const criteriaBuilder = SpreadsheetApp.newFilterCriteria();
    if (!criteriaBuilder || typeof criteriaBuilder.whenTextEqualTo !== 'function') {
      throw new Error('Google Sheets host does not support pivot filter criteria builders.');
    }
  }

  return requiredMethods;
}

function resolvePivotSortState_(planState) {
  if (!planState.sort) {
    return null;
  }

  if (planState.sort.sortOn === 'group_field') {
    return {
      mode: 'group_field',
      field: String(planState.sort.field || '').trim(),
      direction: planState.sort.direction
    };
  }

  return {
    mode: 'aggregated_value',
    valueField: String(planState.sort.field || '').trim(),
    targetField: planState.rowGroups.length > 0
      ? String(planState.rowGroups[planState.rowGroups.length - 1] || '').trim()
      : String(planState.columnGroups[planState.columnGroups.length - 1] || '').trim(),
    direction: planState.sort.direction
  };
}

function applyPivotSort_(pivotTable, planState, createdGroups, createdValues) {
  const sortState = resolvePivotSortState_(planState);
  if (!sortState) {
    return;
  }

  if (sortState.mode === 'group_field') {
    const group = createdGroups[sortState.field];
    if (!group) {
      throw new Error('Google Sheets host cannot sort pivot group: ' + sortState.field + '.');
    }

    if (sortState.direction === 'desc') {
      if (typeof group.sortDescending !== 'function') {
        throw new Error('Google Sheets host does not expose descending pivot sort here.');
      }
      group.sortDescending();
    } else {
      if (typeof group.sortAscending !== 'function') {
        throw new Error('Google Sheets host does not expose ascending pivot sort here.');
      }
      group.sortAscending();
    }
    return;
  }

  const group = createdGroups[sortState.targetField];
  const pivotValue = createdValues[sortState.valueField];

  if (!group || !pivotValue) {
    throw new Error('Google Sheets host cannot sort this pivot value exactly.');
  }

  if (typeof group.sortBy !== 'function') {
    throw new Error('Google Sheets host does not expose pivot value sorting here.');
  }

  group.sortBy(pivotValue, []);
  if (sortState.direction === 'desc') {
    if (typeof group.sortDescending !== 'function') {
      throw new Error('Google Sheets host does not expose descending pivot sort here.');
    }
    group.sortDescending();
  } else {
    if (typeof group.sortAscending !== 'function') {
      throw new Error('Google Sheets host does not expose ascending pivot sort here.');
    }
    group.sortAscending();
  }
}

function applyPivotTablePlan_(spreadsheet, plan) {
  const planState = preflightPivotTableStructure_(plan);
  const sourceSheet = spreadsheet.getSheetByName(plan.sourceSheet);
  const targetSheet = spreadsheet.getSheetByName(plan.targetSheet);

  if (!sourceSheet) {
    throw new Error('Source sheet not found: ' + plan.sourceSheet);
  }

  if (!targetSheet) {
    throw new Error('Target sheet not found: ' + plan.targetSheet);
  }

  const sourceRange = sourceSheet.getRange(plan.sourceRange);
  const anchorRange = targetSheet.getRange(plan.targetRange);
  requireSingleCellAnchor_(anchorRange, 'pivot tables');

  if (typeof anchorRange.createPivotTable !== 'function') {
    throw new Error('Google Sheets host does not expose pivot creation on this range.');
  }

  const headerMap = buildHeaderMap_(sourceRange);
  validatePivotTableFieldReferences_(headerMap, planState);
  validatePivotTableCapabilities_(planState);
  const pivotFilterCriteriaList = buildPivotFilterCriteriaList_(planState);
  const pivotTable = anchorRange.createPivotTable(sourceRange);
  const createdGroups = {};
  const createdValues = {};
  planState.rowGroups.forEach(function(field) {
    createdGroups[String(field || '').trim()] = pivotTable.addRowGroup(getPivotHeaderIndex_(headerMap, field));
  });

  planState.columnGroups.forEach(function(field) {
    createdGroups[String(field || '').trim()] = pivotTable.addColumnGroup(getPivotHeaderIndex_(headerMap, field));
  });

  planState.valueAggregations.forEach(function(aggregation) {
    createdValues[String(aggregation.field || '').trim()] = pivotTable.addPivotValue(
      getPivotHeaderIndex_(headerMap, aggregation.field),
      getPivotSummarizeFunction_(aggregation.aggregation)
    );
  });

  pivotFilterCriteriaList.forEach(function(filterCriteria) {
    pivotTable.addFilter(
      getPivotHeaderIndex_(headerMap, filterCriteria.field),
      filterCriteria.criteria
    );
  });

  applyPivotSort_(pivotTable, planState, createdGroups, createdValues);

  SpreadsheetApp.flush();

  return {
    kind: 'pivot_table_update',
    operation: 'pivot_table_update',
    hostPlatform: 'google_sheets',
    sourceSheet: plan.sourceSheet,
    sourceRange: plan.sourceRange,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    rowGroups: plan.rowGroups,
    columnGroups: plan.columnGroups,
    valueAggregations: plan.valueAggregations,
    filters: plan.filters,
    sort: plan.sort,
    explanation: plan.explanation,
    confidence: plan.confidence,
    requiresConfirmation: plan.requiresConfirmation,
    affectedRanges: plan.affectedRanges,
    overwriteRisk: plan.overwriteRisk,
    confirmationLevel: plan.confirmationLevel,
    summary: 'Created pivot table on ' + plan.targetSheet + '!' + normalizeA1_(plan.targetRange) + '.'
  };
}

const CHART_TYPE_CONFIGS_ = {
  bar: {
    chartType: 'BAR',
    stacked: false
  },
  column: {
    chartType: 'COLUMN',
    stacked: false
  },
  stacked_bar: {
    chartType: 'BAR',
    stacked: true
  },
  stacked_column: {
    chartType: 'COLUMN',
    stacked: true
  },
  line: {
    chartType: 'LINE',
    stacked: false
  },
  area: {
    chartType: 'AREA',
    stacked: false
  },
  pie: {
    chartType: 'PIE',
    stacked: false
  },
  scatter: {
    chartType: 'SCATTER',
    stacked: false
  }
};

const CHART_LEGEND_POSITIONS_ = ['bottom', 'left', 'right', 'top', 'none'];
const CHART_LEGEND_POSITION_MAP_ = {
  hidden: 'none'
};

function getChartTypeConfig_(chartType) {
  const config = CHART_TYPE_CONFIGS_[chartType];

  if (!config) {
    throw new Error('Google Sheets host does not support chart type: ' + chartType + '.');
  }

  if (typeof Charts === 'undefined' ||
    !Charts.ChartType ||
    !Charts.ChartType[config.chartType]) {
    throw new Error('Google Sheets host does not support exact-safe chart creation yet.');
  }

  return {
    chartType: Charts.ChartType[config.chartType],
    stacked: config.stacked
  };
}

function normalizeChartHeaderCell_(value) {
  return String(value || '').trim();
}

function getChartFieldSequence_(plan) {
  if (typeof plan.categoryField !== 'string' || plan.categoryField.trim().length === 0) {
    throw new Error('Google Sheets host requires categoryField for exact-safe chart creation.');
  }

  if (!Array.isArray(plan.series) || plan.series.length === 0) {
    throw new Error('Google Sheets host requires at least one series for exact-safe chart creation.');
  }

  const fields = [normalizeChartHeaderCell_(plan.categoryField)];
  const seenFields = {};
  seenFields[fields[0]] = true;

  plan.series.forEach(function(series) {
    const field = normalizeChartHeaderCell_(series && series.field);
    if (!field) {
      throw new Error('Google Sheets host requires exact-safe chart series fields.');
    }

    if (seenFields[field]) {
      throw new Error('Google Sheets host requires chart fields to be unique.');
    }

    seenFields[field] = true;
    fields.push(field);
  });

  return fields;
}

function getChartLegendPosition_(legendPosition) {
  if (legendPosition === undefined || legendPosition === null || legendPosition === '') {
    return null;
  }

  const normalizedLegendPosition = String(legendPosition).trim();
  const mappedLegendPosition = CHART_LEGEND_POSITION_MAP_[normalizedLegendPosition] || normalizedLegendPosition;

  if (CHART_LEGEND_POSITIONS_.indexOf(mappedLegendPosition) === -1) {
    throw new Error('Google Sheets host does not support exact-safe chart legend positioning for ' + legendPosition + '.');
  }

  return mappedLegendPosition;
}

function getChartSeriesLabel_(series) {
  if (!series || typeof series !== 'object') {
    return null;
  }

  if (series.label === undefined || series.label === null || String(series.label).trim().length === 0) {
    return null;
  }

  const label = String(series.label).trim();
  const field = normalizeChartHeaderCell_(series.field);

  if (!field) {
    throw new Error('Google Sheets host requires exact-safe chart series fields.');
  }

  return label;
}

function validateChartSeriesLabels_(plan) {
  if (!Array.isArray(plan.series)) {
    return;
  }

  plan.series.forEach(function(series) {
    getChartSeriesLabel_(series);
  });
}

function buildChartSeriesOptions_(plan) {
  const options = {};

  (plan.series || []).forEach(function(series, index) {
    const label = getChartSeriesLabel_(series);
    const field = normalizeChartHeaderCell_(series && series.field);
    if (!label || !field || label === field) {
      return;
    }

    options[index] = {
      labelInLegend: label
    };
  });

  return options;
}

function validateChartSourceLayout_(sourceRange, headerMap, plan) {
  if (!sourceRange || typeof sourceRange.getNumRows !== 'function') {
    throw new Error('Google Sheets host requires a valid source range for charts.');
  }

  if (sourceRange.getNumRows() < 2) {
    throw new Error('Google Sheets host requires at least one data row for exact-safe chart creation.');
  }

  getChartFieldSequence_(plan).forEach(function(field) {
    if (!Object.prototype.hasOwnProperty.call(headerMap, field)) {
      throw new Error('Google Sheets host cannot find chart field in header row: ' + field + '.');
    }
  });
}

function getChartFieldColumnIndex_(headerMap, field) {
  const key = normalizeChartHeaderCell_(field);
  const columnIndex = headerMap[key];

  if (!columnIndex) {
    throw new Error('Google Sheets host cannot find chart field in header row: ' + field + '.');
  }

  return columnIndex;
}

function buildChartSourceRanges_(sourceSheet, sourceRange, headerMap, plan) {
  const startRow = sourceRange.getRow();
  const startColumn = sourceRange.getColumn();
  const numRows = sourceRange.getNumRows();
  const fieldSequence = getChartFieldSequence_(plan);

  return fieldSequence.map(function(field) {
    const columnIndex = getChartFieldColumnIndex_(headerMap, field);
    return sourceSheet.getRange(startRow, startColumn + columnIndex - 1, numRows, 1);
  });
}

function applyChartOptionOrFail_(chartBuilder, optionName, optionValue, failureLabel) {
  if (typeof chartBuilder.setOption !== 'function') {
    throw new Error('Google Sheets host does not support exact-safe chart creation yet.');
  }

  try {
    chartBuilder.setOption(optionName, optionValue);
  } catch (error) {
    throw new Error('Google Sheets host does not support exact-safe chart ' + failureLabel + '.');
  }
}

function applyChartPlan_(spreadsheet, plan) {
  const sourceSheet = spreadsheet.getSheetByName(plan.sourceSheet);
  const targetSheet = spreadsheet.getSheetByName(plan.targetSheet);

  if (!sourceSheet) {
    throw new Error('Source sheet not found: ' + plan.sourceSheet);
  }

  if (!targetSheet) {
    throw new Error('Target sheet not found: ' + plan.targetSheet);
  }

  if (typeof targetSheet.newChart !== 'function' || typeof targetSheet.insertChart !== 'function') {
    throw new Error('Google Sheets host does not support exact-safe chart creation yet.');
  }

  const sourceRange = sourceSheet.getRange(plan.sourceRange);
  const headerMap = buildHeaderMap_(sourceRange);
  validateChartSourceLayout_(sourceRange, headerMap, plan);
  validateChartSeriesLabels_(plan);
  const targetRange = targetSheet.getRange(plan.targetRange);
  requireSingleCellAnchor_(targetRange, 'charts');

  const chartConfig = getChartTypeConfig_(plan.chartType);
  const legendPosition = getChartLegendPosition_(plan.legendPosition);
  const chartBuilder = targetSheet.newChart();

  if (!chartBuilder ||
    typeof chartBuilder.addRange !== 'function' ||
    typeof chartBuilder.setChartType !== 'function' ||
    typeof chartBuilder.setPosition !== 'function' ||
    typeof chartBuilder.build !== 'function') {
    throw new Error('Google Sheets host does not support exact-safe chart creation yet.');
  }

  const chartRanges = buildChartSourceRanges_(sourceSheet, sourceRange, headerMap, plan);
  chartRanges.forEach(function(range) {
    chartBuilder.addRange(range);
  });
  chartBuilder.setChartType(chartConfig.chartType);
  chartBuilder.setPosition(targetRange.getRow(), targetRange.getColumn(), 0, 0);

  if ((plan.title || legendPosition || chartConfig.stacked) && typeof chartBuilder.setOption !== 'function') {
    throw new Error('Google Sheets host does not support exact-safe chart creation yet.');
  }

  if (plan.title) {
    applyChartOptionOrFail_(chartBuilder, 'title', plan.title, 'title options');
  }

  if (legendPosition) {
    applyChartOptionOrFail_(chartBuilder, 'legend', { position: legendPosition }, 'legend positioning');
  }

  if (chartConfig.stacked) {
    applyChartOptionOrFail_(chartBuilder, 'isStacked', true, 'stacking');
  }

  const seriesOptions = buildChartSeriesOptions_(plan);
  if (Object.keys(seriesOptions).length > 0) {
    applyChartOptionOrFail_(chartBuilder, 'series', seriesOptions, 'series labels');
  }

  const chart = chartBuilder.build();
  targetSheet.insertChart(chart);
  SpreadsheetApp.flush();

  return {
    kind: 'chart_update',
    operation: 'chart_update',
    hostPlatform: 'google_sheets',
    sourceSheet: plan.sourceSheet,
    sourceRange: plan.sourceRange,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    chartType: plan.chartType,
    categoryField: plan.categoryField,
    series: plan.series,
    title: plan.title,
    legendPosition: plan.legendPosition,
    explanation: plan.explanation,
    confidence: plan.confidence,
    requiresConfirmation: plan.requiresConfirmation,
    affectedRanges: plan.affectedRanges,
    overwriteRisk: plan.overwriteRisk,
    confirmationLevel: plan.confirmationLevel,
    summary: 'Created ' + plan.chartType + ' chart on ' + plan.targetSheet + '!' + normalizeA1_(plan.targetRange) + '.'
  };
}

function buildCellContext_(cell) {
  if (!cell) {
    return undefined;
  }

  const displayValue = cell.getDisplayValue();
  const value = normalizeCellValue_(cell.getValue(), displayValue);
  const formulaRaw = cell.getFormula();
  const noteRaw = cell.getNote();
  const formula = typeof formulaRaw === 'string'
    ? truncateContextText_(formulaRaw, MAX_CONTEXT_FORMULA_LENGTH_)
    : '';
  const note = typeof noteRaw === 'string'
    ? truncateContextText_(noteRaw, MAX_CONTEXT_NOTE_LENGTH_)
    : '';
  const context = {
    a1Notation: normalizeA1_(cell.getA1Notation()),
    displayValue: truncateContextText_(displayValue, MAX_CONTEXT_CELL_TEXT_LENGTH_),
    value: value
  };

  if (formula) {
    context.formula = formula;
  }

  if (note) {
    context.note = note;
  }

  return context;
}

function buildReferencedCellsContext_(sheet, prompt, activeCellContext) {
  if (!sheet || typeof extractReferencedA1Notations_ !== 'function') {
    return undefined;
  }

  const activeCellA1 = activeCellContext ? activeCellContext.a1Notation : '';
  const referencedA1Notations = extractReferencedA1Notations_(prompt).filter(function(a1Notation) {
    return a1Notation && a1Notation !== activeCellA1;
  });

  if (referencedA1Notations.length === 0) {
    return undefined;
  }

  const contexts = referencedA1Notations.map(function(a1Notation) {
    try {
      return buildCellContext_(sheet.getRange(a1Notation));
    } catch (_error) {
      return null;
    }
  }).filter(function(context) {
    return Boolean(context);
  });

  return contexts.length > 0 ? contexts : undefined;
}

function shouldIncludeRegionMatrix_(rangeA1) {
  if (!rangeA1) {
    return false;
  }

  const bounds = parseA1RangeReference_(rangeA1);
  return bounds.rowCount * bounds.columnCount <= 400;
}

function buildImplicitRegionTargets_(rangeA1) {
  if (!rangeA1) {
    return {};
  }

  const bounds = parseA1RangeReference_(rangeA1);
  return {
    currentRegionArtifactTarget: buildA1RangeFromBounds_({
      startRow: bounds.endRow + 2,
      endRow: bounds.endRow + 2,
      startColumn: bounds.startColumn,
      endColumn: bounds.startColumn
    }),
    currentRegionAppendTarget: buildA1RangeFromBounds_({
      startRow: bounds.endRow + 1,
      endRow: bounds.endRow + 1,
      startColumn: bounds.startColumn,
      endColumn: bounds.endColumn
    })
  };
}

function rangeHasExistingContent_(values) {
  return Array.isArray(values) && values.some(function(row) {
    return Array.isArray(row) && row.some(function(cell) {
      return cell !== null && cell !== undefined && cell !== '';
    });
  });
}

function isWorkbookStructurePlan_(plan) {
  return [
    'create_sheet',
    'delete_sheet',
    'rename_sheet',
    'duplicate_sheet',
    'move_sheet',
    'hide_sheet',
    'unhide_sheet'
  ].indexOf(plan && plan.operation) !== -1;
}

function isRangeFormatPlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string' &&
    plan.format &&
    typeof plan.format === 'object'
  );
}

function isDataValidationPlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string' &&
    typeof plan.ruleType === 'string' &&
    typeof plan.managementMode !== 'string'
  );
}

function isConditionalFormatPlan_(plan) {
  if (!plan ||
    typeof plan.targetSheet !== 'string' ||
    typeof plan.targetRange !== 'string' ||
    typeof plan.managementMode !== 'string') {
    return false;
  }

  if (plan.managementMode === 'clear_on_target') {
    return true;
  }

  return typeof plan.ruleType === 'string';
}

function isSupportedConditionalFormatManagementMode_(managementMode) {
  return managementMode === 'add' ||
    managementMode === 'replace_all_on_target' ||
    managementMode === 'clear_on_target';
}

function getConditionalFormatStatusSummary_(plan) {
  if (!plan || !plan.targetSheet || !plan.targetRange) {
    return 'Conditional formatting updated.';
  }

  switch (plan.managementMode) {
    case 'add':
      return 'Added conditional formatting on ' + plan.targetSheet + '!' + plan.targetRange + '.';
    case 'replace_all_on_target':
      return 'Replaced conditional formatting on ' + plan.targetSheet + '!' + plan.targetRange + '.';
    case 'clear_on_target':
      return 'Cleared conditional formatting on ' + plan.targetSheet + '!' + plan.targetRange + '.';
    default:
      return 'Conditional formatting updated on ' + plan.targetSheet + '!' + plan.targetRange + '.';
  }
}

function getSupportedDateTextPatternSpec_(formatPattern) {
  if (typeof formatPattern !== 'string') {
    return null;
  }

  const trimmed = formatPattern.trim();
  const match = trimmed.match(/^[Yy]{4}([\-/.])[Mm]{2}\1[Dd]{2}$/);
  if (!match) {
    return null;
  }

  return {
    formatType: 'date_text',
    separator: match[1],
    formatPattern: trimmed
  };
}

function getSupportedNumberTextPatternSpec_(formatPattern) {
  if (typeof formatPattern !== 'string') {
    return null;
  }

  const trimmed = formatPattern.trim();
  const match = trimmed.match(/^(#,##0|0)(?:\.(0+))?$/);
  if (!match) {
    return null;
  }

  return {
    formatType: 'number_text',
    useGrouping: match[1] === '#,##0',
    decimals: match[2] ? match[2].length : 0,
    formatPattern: trimmed
  };
}

function getSupportedStandardizeFormatSpec_(formatType, formatPattern) {
  if (formatType === 'date_text') {
    return getSupportedDateTextPatternSpec_(formatPattern);
  }

  if (formatType === 'number_text') {
    return getSupportedNumberTextPatternSpec_(formatPattern);
  }

  return null;
}

function getStandardizeFormatSupportError_(formatType, formatPattern, hostLabel) {
  const resolvedHostLabel = hostLabel || 'This host';

  if (
    typeof formatType !== 'string' ||
    !formatType.trim() ||
    typeof formatPattern !== 'string' ||
    !formatPattern.trim()
  ) {
    return resolvedHostLabel + ' requires an exact formatType and formatPattern for standardize_format.';
  }

  if (getSupportedStandardizeFormatSpec_(formatType, formatPattern)) {
    return '';
  }

  if (formatType === 'date_text') {
    return resolvedHostLabel + ' only supports exact year-first date text patterns like YYYY-MM-DD, YYYY/MM/DD, or YYYY.MM.DD.';
  }

  if (formatType === 'number_text') {
    return resolvedHostLabel + ' only supports simple fixed-decimal number text patterns like #,##0.00 or 0.00.';
  }

  return resolvedHostLabel + ' can\'t standardize ' + formatType + ' with pattern ' + formatPattern + ' exactly.';
}

function isDateObject_(value) {
  return Object.prototype.toString.call(value) === '[object Date]' &&
    typeof value.getTime === 'function';
}

function isValidDateParts_(year, month, day) {
  const candidate = new Date(year, month - 1, day);
  return candidate.getFullYear() === year &&
    candidate.getMonth() === month - 1 &&
    candidate.getDate() === day;
}

function normalizeIntegerDigits_(integerDigits) {
  const normalized = String(integerDigits || '').replace(/^0+(?=\d)/, '');
  return normalized.length > 0 ? normalized : '0';
}

function formatGroupedIntegerDigits_(integerDigits) {
  return normalizeIntegerDigits_(integerDigits).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function parseExactNumericParts_(value, hostLabel) {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new Error(hostLabel + ' cannot standardize non-finite numbers exactly.');
    }

    const serialized = String(value);
    if (/e/i.test(serialized)) {
      throw new Error(hostLabel + ' cannot standardize scientific-notation numbers exactly.');
    }

    return parseExactNumericParts_(serialized, hostLabel);
  }

  if (typeof value !== 'string' || value !== value.trim()) {
    throw new Error(hostLabel + ' cannot standardize non-text numeric values exactly.');
  }

  const match = value.match(/^([+-]?)(?:(\d{1,3}(?:,\d{3})+)|(\d+))(?:\.(\d+))?$/);
  if (!match) {
    throw new Error(hostLabel + ' cannot standardize numeric text exactly for value ' + JSON.stringify(value) + '.');
  }

  return {
    sign: match[1] === '-' ? '-' : '',
    integerDigits: normalizeIntegerDigits_((match[2] || match[3] || '').replace(/,/g, '')),
    fractionDigits: match[4] || ''
  };
}

function standardizeDateTextValueExact_(value, spec, hostLabel) {
  if (isBlankCellValue_(value)) {
    return value;
  }

  let year;
  let month;
  let day;

  if (isDateObject_(value)) {
    if (Number.isNaN(value.getTime())) {
      throw new Error(hostLabel + ' cannot standardize invalid dates exactly.');
    }

    if (
      value.getHours() !== 0 ||
      value.getMinutes() !== 0 ||
      value.getSeconds() !== 0 ||
      value.getMilliseconds() !== 0
    ) {
      throw new Error(hostLabel + ' cannot rewrite date-time values as date_text without losing precision.');
    }

    year = value.getFullYear();
    month = value.getMonth() + 1;
    day = value.getDate();
  } else if (typeof value === 'string' && value === value.trim()) {
    const match = value.match(/^(\d{4})([-/.])(\d{1,2})\2(\d{1,2})$/);
    if (!match) {
      throw new Error(hostLabel + ' cannot standardize date text exactly for value ' + JSON.stringify(value) + '.');
    }

    year = Number(match[1]);
    month = Number(match[3]);
    day = Number(match[4]);
  } else {
    throw new Error(hostLabel + ' cannot standardize non-date values as date_text exactly.');
  }

  if (!isValidDateParts_(year, month, day)) {
    throw new Error(hostLabel + ' cannot standardize invalid calendar dates exactly.');
  }

  return String(year).padStart(4, '0') +
    spec.separator +
    String(month).padStart(2, '0') +
    spec.separator +
    String(day).padStart(2, '0');
}

function standardizeNumberTextValueExact_(value, spec, hostLabel) {
  if (isBlankCellValue_(value)) {
    return value;
  }

  const parsed = parseExactNumericParts_(value, hostLabel);
  const discardedFraction = parsed.fractionDigits.slice(spec.decimals);
  if (discardedFraction.replace(/0/g, '').length > 0) {
    throw new Error(hostLabel + ' cannot standardize numeric text exactly without rounding.');
  }

  const integerDigits = spec.useGrouping
    ? formatGroupedIntegerDigits_(parsed.integerDigits)
    : normalizeIntegerDigits_(parsed.integerDigits);
  const fractionDigits = parsed.fractionDigits.slice(0, spec.decimals).padEnd(spec.decimals, '0');
  return parsed.sign + integerDigits + (spec.decimals > 0 ? '.' + fractionDigits : '');
}

function standardizeFormatMatrixExact_(plan, values, hostLabel) {
  const spec = getSupportedStandardizeFormatSpec_(plan && plan.formatType, plan && plan.formatPattern);
  if (!spec) {
    throw new Error(
      getStandardizeFormatSupportError_(plan && plan.formatType, plan && plan.formatPattern, hostLabel)
    );
  }

  return values.map(function(row) {
    return row.map(function(value) {
      return spec.formatType === 'date_text'
        ? standardizeDateTextValueExact_(value, spec, hostLabel)
        : standardizeNumberTextValueExact_(value, spec, hostLabel);
    });
  });
}

function getUnsupportedConditionalFormatStyleFields_(style) {
  if (!style || typeof style !== 'object') {
    return [];
  }

  return ['underline', 'strikethrough', 'numberFormat'].filter(function(field) {
    return style[field] !== undefined;
  });
}

function validateConditionalFormatStyle_(style) {
  const unsupportedFields = getUnsupportedConditionalFormatStyleFields_(style);

  if (unsupportedFields.length > 0) {
    throw new Error(
      'Google Sheets host does not support exact conditional-format style mapping for fields: ' +
      unsupportedFields.join(', ') +
      '.'
    );
  }
}

function requiresSecondConditionalComparatorValue_(comparator) {
  return comparator === 'between' || comparator === 'not_between';
}

function validateConditionalComparatorPlan_(plan, hostLabel) {
  if (typeof plan.comparator !== 'string' || plan.comparator.length === 0) {
    throw new Error(hostLabel + ' requires a comparator for ruleType ' + plan.ruleType + '.');
  }

  if (plan.value === undefined) {
    throw new Error(hostLabel + ' requires value for comparator ' + plan.comparator + '.');
  }

  if (requiresSecondConditionalComparatorValue_(plan.comparator) && plan.value2 === undefined) {
    throw new Error(hostLabel + ' requires value2 for comparator ' + plan.comparator + '.');
  }
}

function parseConditionalNumberValue_(value, hostLabel) {
  const parsed = parseExactNumericParts_(value, hostLabel);
  const serialized = parsed.sign +
    normalizeIntegerDigits_(parsed.integerDigits) +
    (parsed.fractionDigits ? '.' + parsed.fractionDigits : '');
  const numeric = Number(serialized);

  if (!Number.isFinite(numeric)) {
    throw new Error(hostLabel + ' cannot use non-finite number comparisons exactly.');
  }

  return numeric;
}

function tryParseConditionalNumberValue_(value) {
  try {
    return parseConditionalNumberValue_(value, 'Google Sheets host');
  } catch {
    return null;
  }
}

function parseConditionalDateParts_(value, hostLabel) {
  let year;
  let month;
  let day;

  if (isDateObject_(value)) {
    if (Number.isNaN(value.getTime())) {
      throw new Error(hostLabel + ' cannot use invalid dates in conditional formatting.');
    }

    if (
      value.getHours() !== 0 ||
      value.getMinutes() !== 0 ||
      value.getSeconds() !== 0 ||
      value.getMilliseconds() !== 0
    ) {
      throw new Error(hostLabel + ' cannot compare date-time values exactly as date_compare.');
    }

    year = value.getFullYear();
    month = value.getMonth() + 1;
    day = value.getDate();
  } else if (typeof value === 'string' && value === value.trim()) {
    const match = value.match(/^(\d{4})([-/.])(\d{1,2})\2(\d{1,2})$/);
    if (!match) {
      throw new Error(hostLabel + ' requires exact year-first date literals for date_compare.');
    }

    year = Number(match[1]);
    month = Number(match[3]);
    day = Number(match[4]);
  } else {
    throw new Error(hostLabel + ' requires exact year-first date literals for date_compare.');
  }

  if (!isValidDateParts_(year, month, day)) {
    throw new Error(hostLabel + ' cannot use invalid calendar dates in conditional formatting.');
  }

  return {
    year: year,
    month: month,
    day: day
  };
}

function buildConditionalDateLiteral_(value, hostLabel) {
  const parts = parseConditionalDateParts_(value, hostLabel);
  return 'DATE(' + parts.year + ',' + parts.month + ',' + parts.day + ')';
}

function tryBuildConditionalDateLiteral_(value) {
  try {
    return buildConditionalDateLiteral_(value, 'Google Sheets host');
  } catch {
    return null;
  }
}

function buildFormulaStringLiteral_(value) {
  return '"' + String(value).replace(/"/g, '""') + '"';
}

function buildRelativeTargetCellReference_(target) {
  return convertColumnNumberToLetters_(target.getColumn()) + target.getRow();
}

function buildAbsoluteTargetRangeReference_(target) {
  const startColumn = convertColumnNumberToLetters_(target.getColumn());
  const endColumn = convertColumnNumberToLetters_(target.getColumn() + target.getNumColumns() - 1);
  const startRow = target.getRow();
  const endRow = target.getRow() + target.getNumRows() - 1;
  const startCell = '$' + startColumn + '$' + startRow;
  const endCell = '$' + endColumn + '$' + endRow;
  return startCell === endCell ? startCell : startCell + ':' + endCell;
}

function buildComparatorFormula_(cellReference, comparator, valueLiteral, value2Literal) {
  switch (comparator) {
    case 'between':
      return '=AND(' + cellReference + '>=' + valueLiteral + ',' + cellReference + '<=' + value2Literal + ')';
    case 'not_between':
      return '=OR(' + cellReference + '<' + valueLiteral + ',' + cellReference + '>' + value2Literal + ')';
    case 'equal_to':
      return '=' + cellReference + '=' + valueLiteral;
    case 'not_equal_to':
      return '=' + cellReference + '<>' + valueLiteral;
    case 'greater_than':
      return '=' + cellReference + '>' + valueLiteral;
    case 'greater_than_or_equal_to':
      return '=' + cellReference + '>=' + valueLiteral;
    case 'less_than':
      return '=' + cellReference + '<' + valueLiteral;
    case 'less_than_or_equal_to':
      return '=' + cellReference + '<=' + valueLiteral;
    default:
      throw new Error('Google Sheets host does not support exact conditional-format comparator ' + comparator + '.');
  }
}

function resolveSingleColorComparisonLiterals_(plan, hostLabel) {
  validateConditionalComparatorPlan_(plan, hostLabel);

  const numericValue = tryParseConditionalNumberValue_(plan.value);
  const numericValue2 = plan.value2 === undefined ? null : tryParseConditionalNumberValue_(plan.value2);
  if (numericValue !== null && (plan.value2 === undefined || numericValue2 !== null)) {
    return {
      valueLiteral: String(numericValue),
      value2Literal: numericValue2 === null ? undefined : String(numericValue2)
    };
  }

  const dateLiteral = tryBuildConditionalDateLiteral_(plan.value);
  const dateLiteral2 = plan.value2 === undefined ? null : tryBuildConditionalDateLiteral_(plan.value2);
  if (dateLiteral && (plan.value2 === undefined || dateLiteral2)) {
    return {
      valueLiteral: dateLiteral,
      value2Literal: dateLiteral2 || undefined
    };
  }

  if (plan.comparator !== 'equal_to' && plan.comparator !== 'not_equal_to') {
    throw new Error(
      'Google Sheets host only supports string equality checks for single_color conditional formatting.'
    );
  }

  if (typeof plan.value !== 'string' || valueHasLeadingOrTrailingWhitespace_(plan.value)) {
    throw new Error('Google Sheets host requires exact string literals for single_color text comparisons.');
  }

  return {
    valueLiteral: buildFormulaStringLiteral_(plan.value),
    value2Literal: undefined
  };
}

function valueHasLeadingOrTrailingWhitespace_(value) {
  return typeof value === 'string' && value !== value.trim();
}

function buildSingleColorFormula_(target, plan, hostLabel) {
  const literals = resolveSingleColorComparisonLiterals_(plan, hostLabel);
  return buildComparatorFormula_(
    buildRelativeTargetCellReference_(target),
    plan.comparator,
    literals.valueLiteral,
    literals.value2Literal
  );
}

function buildDateCompareFormula_(target, plan, hostLabel) {
  validateConditionalComparatorPlan_(plan, hostLabel);
  return buildComparatorFormula_(
    buildRelativeTargetCellReference_(target),
    plan.comparator,
    buildConditionalDateLiteral_(plan.value, hostLabel),
    plan.value2 === undefined ? undefined : buildConditionalDateLiteral_(plan.value2, hostLabel)
  );
}

function buildDuplicateValuesFormula_(target) {
  const cellReference = buildRelativeTargetCellReference_(target);
  const rangeReference = buildAbsoluteTargetRangeReference_(target);
  return '=COUNTIF(' + rangeReference + ',' + cellReference + ')>1';
}

function buildTopNFormula_(target, plan) {
  const cellReference = buildRelativeTargetCellReference_(target);
  const rangeReference = buildAbsoluteTargetRangeReference_(target);
  const rankSuffix = plan.direction === 'bottom' ? ',TRUE' : '';
  return '=RANK(' + cellReference + ',' + rangeReference + rankSuffix + ')<=' + plan.rank;
}

function buildAverageCompareFormula_(target, plan) {
  const cellReference = buildRelativeTargetCellReference_(target);
  const comparison = plan.direction === 'below' ? '<' : '>';
  return '=' + cellReference + comparison + 'AVERAGE(' + buildAbsoluteTargetRangeReference_(target) + ')';
}

function getNumberCompareBuilderMethodName_(comparator) {
  switch (comparator) {
    case 'between':
      return 'whenNumberBetween';
    case 'not_between':
      return 'whenNumberNotBetween';
    case 'equal_to':
      return 'whenNumberEqualTo';
    case 'not_equal_to':
      return 'whenNumberNotEqualTo';
    case 'greater_than':
      return 'whenNumberGreaterThan';
    case 'greater_than_or_equal_to':
      return 'whenNumberGreaterThanOrEqualTo';
    case 'less_than':
      return 'whenNumberLessThan';
    case 'less_than_or_equal_to':
      return 'whenNumberLessThanOrEqualTo';
    default:
      return '';
  }
}

function applyConditionalFormulaRule_(builder, formula, ruleLabel) {
  if (typeof builder.whenFormulaSatisfied !== 'function') {
    throw new Error('Google Sheets host does not support ' + ruleLabel + ' conditional formatting.');
  }

  builder.whenFormulaSatisfied(formula);
}

function applyConditionalNumberCompareRule_(builder, plan) {
  validateConditionalComparatorPlan_(plan, 'Google Sheets host');

  const methodName = getNumberCompareBuilderMethodName_(plan.comparator);
  if (!methodName || typeof builder[methodName] !== 'function') {
    throw new Error('Google Sheets host does not support number_compare conditional formatting.');
  }

  const value = parseConditionalNumberValue_(plan.value, 'Google Sheets host');
  if (requiresSecondConditionalComparatorValue_(plan.comparator)) {
    builder[methodName](value, parseConditionalNumberValue_(plan.value2, 'Google Sheets host'));
    return;
  }

  builder[methodName](value);
}

function getInterpolationTypeEnum_(type) {
  const interpolationType = SpreadsheetApp && SpreadsheetApp.InterpolationType
    ? SpreadsheetApp.InterpolationType
    : {};

  switch (type) {
    case 'number':
      return interpolationType.NUMBER || 'NUMBER';
    case 'percent':
      return interpolationType.PERCENT || 'PERCENT';
    case 'percentile':
      return interpolationType.PERCENTILE || 'PERCENTILE';
    default:
      return '';
  }
}

function applyConditionalColorScalePoint_(builder, position, point) {
  const capitalizedPosition = position.charAt(0).toUpperCase() + position.slice(1);

  if (point.type === 'min' || point.type === 'max') {
    const methodName = 'setGradient' + capitalizedPosition + 'point';
    if (typeof builder[methodName] !== 'function') {
      throw new Error('Google Sheets host does not support color_scale conditional formatting.');
    }
    builder[methodName](point.color);
    return;
  }

  const interpolationType = getInterpolationTypeEnum_(point.type);
  if (!interpolationType) {
    throw new Error('Google Sheets host does not support color_scale point type ' + point.type + '.');
  }

  const methodName = 'setGradient' + capitalizedPosition + 'pointWithValue';
  if (typeof builder[methodName] !== 'function') {
    throw new Error('Google Sheets host does not support color_scale conditional formatting.');
  }

  builder[methodName](point.color, interpolationType, String(point.value));
}

function applyConditionalColorScaleRule_(builder, plan) {
  if (!Array.isArray(plan.points) || plan.points.length < 2 || plan.points.length > 3) {
    throw new Error('Google Sheets host requires 2 or 3 color_scale points.');
  }

  const positions = plan.points.length === 2
    ? ['min', 'max']
    : ['min', 'mid', 'max'];

  plan.points.forEach(function(point, index) {
    applyConditionalColorScalePoint_(builder, positions[index], point);
  });
}

function applyConditionalFormatStyleToBuilder_(builder, style) {
  if (style && style.backgroundColor !== undefined) {
    if (typeof builder.setBackground !== 'function') {
      throw new Error('Google Sheets host does not support exact conditional-format style mapping for fields: backgroundColor.');
    }
    builder.setBackground(style.backgroundColor);
  }

  if (style && style.textColor !== undefined) {
    if (typeof builder.setFontColor !== 'function') {
      throw new Error('Google Sheets host does not support exact conditional-format style mapping for fields: textColor.');
    }
    builder.setFontColor(style.textColor);
  }

  if (style && style.bold !== undefined) {
    if (typeof builder.setBold !== 'function') {
      throw new Error('Google Sheets host does not support exact conditional-format style mapping for fields: bold.');
    }
    builder.setBold(style.bold);
  }

  if (style && style.italic !== undefined) {
    if (typeof builder.setItalic !== 'function') {
      throw new Error('Google Sheets host does not support exact conditional-format style mapping for fields: italic.');
    }
    builder.setItalic(style.italic);
  }
}

function validateConditionalFormatPlanSupport_(plan) {
  if (!isSupportedConditionalFormatManagementMode_(plan && plan.managementMode)) {
    throw new Error(
      'Google Sheets host does not support exact conditional-format managementMode ' +
      plan.managementMode +
      '.'
    );
  }

  validateConditionalFormatStyle_(plan && plan.style);

  if (!plan || plan.managementMode === 'clear_on_target') {
    return;
  }

  switch (plan.ruleType) {
    case 'single_color':
      resolveSingleColorComparisonLiterals_(plan, 'Google Sheets host');
      return;
    case 'text_contains':
      if (typeof plan.text !== 'string' || plan.text.length === 0) {
        throw new Error('Google Sheets host requires text for ruleType text_contains.');
      }
      return;
    case 'number_compare':
      validateConditionalComparatorPlan_(plan, 'Google Sheets host');
      parseConditionalNumberValue_(plan.value, 'Google Sheets host');
      if (requiresSecondConditionalComparatorValue_(plan.comparator)) {
        parseConditionalNumberValue_(plan.value2, 'Google Sheets host');
      }
      return;
    case 'date_compare':
      validateConditionalComparatorPlan_(plan, 'Google Sheets host');
      buildConditionalDateLiteral_(plan.value, 'Google Sheets host');
      if (requiresSecondConditionalComparatorValue_(plan.comparator)) {
        buildConditionalDateLiteral_(plan.value2, 'Google Sheets host');
      }
      return;
    case 'duplicate_values':
      return;
    case 'custom_formula':
      if (typeof plan.formula !== 'string' || plan.formula.trim().length === 0) {
        throw new Error('Google Sheets host requires formula for ruleType custom_formula.');
      }
      return;
    case 'top_n':
      if (!Number.isInteger(plan.rank) || plan.rank <= 0) {
        throw new Error('Google Sheets host requires a positive rank for ruleType top_n.');
      }
      if (plan.direction !== 'top' && plan.direction !== 'bottom') {
        throw new Error('Google Sheets host requires direction top or bottom for ruleType top_n.');
      }
      return;
    case 'average_compare':
      if (plan.direction !== 'above' && plan.direction !== 'below') {
        throw new Error('Google Sheets host requires direction above or below for ruleType average_compare.');
      }
      return;
    case 'color_scale':
      if (!Array.isArray(plan.points) || plan.points.length < 2 || plan.points.length > 3) {
        throw new Error('Google Sheets host requires 2 or 3 color_scale points.');
      }
      return;
    default:
      throw new Error(
        'Google Sheets host does not support exact conditional-format mapping for ruleType ' +
        plan.ruleType +
        '.'
      );
  }
}

function rangeMatchesExactly_(left, right) {
  return left &&
    right &&
    typeof left.getRow === 'function' &&
    typeof right.getRow === 'function' &&
    left.getRow() === right.getRow() &&
    left.getColumn() === right.getColumn() &&
    left.getNumRows() === right.getNumRows() &&
    left.getNumColumns() === right.getNumColumns();
}

function rangesOverlap_(left, right) {
  if (!left || !right ||
    typeof left.getRow !== 'function' ||
    typeof right.getRow !== 'function') {
    return false;
  }

  const leftStartRow = left.getRow();
  const leftEndRow = leftStartRow + left.getNumRows() - 1;
  const leftStartColumn = left.getColumn();
  const leftEndColumn = leftStartColumn + left.getNumColumns() - 1;
  const rightStartRow = right.getRow();
  const rightEndRow = rightStartRow + right.getNumRows() - 1;
  const rightStartColumn = right.getColumn();
  const rightEndColumn = rightStartColumn + right.getNumColumns() - 1;

  return !(
    leftEndRow < rightStartRow ||
    rightEndRow < leftStartRow ||
    leftEndColumn < rightStartColumn ||
    rightEndColumn < leftStartColumn
  );
}

function partitionConditionalFormatRules_(sheet, target) {
  if (!sheet || typeof sheet.getConditionalFormatRules !== 'function') {
    throw new Error('Google Sheets host does not support conditional formatting on this sheet.');
  }

  const existingRules = sheet.getConditionalFormatRules() || [];
  const preservedRules = [];

  existingRules.forEach(function(rule) {
    const ranges = rule && typeof rule.getRanges === 'function' ? rule.getRanges() || [] : [];
    const overlapsTarget = ranges.some(function(range) {
      return rangesOverlap_(range, target);
    });

    if (!overlapsTarget) {
      preservedRules.push(rule);
      return;
    }

    if (ranges.length === 1 && rangeMatchesExactly_(ranges[0], target)) {
      return;
    }

    throw new Error(
      'Google Sheets host cannot modify conditional formatting exactly when an existing rule overlaps the target range without matching it exactly.'
    );
  });

  return preservedRules;
}

function buildConditionalFormatRule_(target, plan) {
  validateConditionalFormatPlanSupport_(plan);

  if (typeof SpreadsheetApp.newConditionalFormatRule !== 'function') {
    throw new Error('Google Sheets host does not support creating conditional formatting rules.');
  }

  const builder = SpreadsheetApp.newConditionalFormatRule();

  switch (plan.ruleType) {
    case 'single_color':
      applyConditionalFormulaRule_(
        builder,
        buildSingleColorFormula_(target, plan, 'Google Sheets host'),
        'single_color'
      );
      break;
    case 'text_contains':
      if (typeof builder.whenTextContains !== 'function') {
        throw new Error('Google Sheets host does not support text_contains conditional formatting.');
      }
      builder.whenTextContains(plan.text);
      break;
    case 'number_compare':
      applyConditionalNumberCompareRule_(builder, plan);
      break;
    case 'date_compare':
      applyConditionalFormulaRule_(
        builder,
        buildDateCompareFormula_(target, plan, 'Google Sheets host'),
        'date_compare'
      );
      break;
    case 'duplicate_values':
      applyConditionalFormulaRule_(builder, buildDuplicateValuesFormula_(target), 'duplicate_values');
      break;
    case 'custom_formula':
      applyConditionalFormulaRule_(builder, plan.formula, 'custom_formula');
      break;
    case 'top_n':
      applyConditionalFormulaRule_(builder, buildTopNFormula_(target, plan), 'top_n');
      break;
    case 'average_compare':
      applyConditionalFormulaRule_(builder, buildAverageCompareFormula_(target, plan), 'average_compare');
      break;
    case 'color_scale':
      applyConditionalColorScaleRule_(builder, plan);
      break;
    default:
      throw new Error(
        'Google Sheets host does not support exact conditional-format mapping for ruleType ' +
        plan.ruleType +
        '.'
      );
  }

  applyConditionalFormatStyleToBuilder_(builder, plan.style);

  if (typeof builder.setRanges !== 'function') {
    throw new Error('Google Sheets host does not support targeting conditional-format ranges.');
  }

  return builder.setRanges([target]).build();
}

function isNamedRangeUpdatePlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.operation === 'string' &&
    typeof plan.scope === 'string' &&
    typeof plan.name === 'string'
  );
}

function getWorkbookStructureStatusSummary_(plan) {
  switch (plan && plan.operation) {
    case 'create_sheet':
      return 'Created sheet ' + plan.sheetName + '.';
    case 'delete_sheet':
      return 'Deleted sheet ' + plan.sheetName + '.';
    case 'rename_sheet':
      return 'Renamed sheet ' + plan.sheetName + ' to ' + plan.newSheetName + '.';
    case 'duplicate_sheet':
      return 'Duplicated sheet ' + plan.sheetName +
        (plan.newSheetName ? ' as ' + plan.newSheetName : '') + '.';
    case 'move_sheet':
      return 'Moved sheet ' + plan.sheetName + '.';
    case 'hide_sheet':
      return 'Hid sheet ' + plan.sheetName + '.';
    case 'unhide_sheet':
      return 'Unhid sheet ' + plan.sheetName + '.';
    default:
      return 'Workbook update applied.';
  }
}

function getDataValidationStatusSummary_(plan) {
  if (plan && plan.targetSheet && plan.targetRange) {
    return 'Applied validation to ' + plan.targetSheet + '!' + plan.targetRange + '.';
  }

  return 'Applied validation.';
}

function getNamedRangeStatusSummary_(plan) {
  if (plan && plan.operation === 'delete') {
    return 'Deleted named range ' + plan.name + '.';
  }

  if (plan && plan.operation === 'rename' && plan.newName) {
    return 'Renamed named range ' + plan.name + ' to ' + plan.newName + '.';
  }

  if (plan && plan.operation === 'create' && plan.targetSheet && plan.targetRange) {
    return 'Created named range ' + plan.name + ' at ' + plan.targetSheet + '!' + plan.targetRange + '.';
  }

  if (plan && plan.targetSheet && plan.targetRange) {
    return 'Retargeted ' + plan.name + ' to ' + plan.targetSheet + '!' + plan.targetRange + '.';
  }

  if (plan && plan.name) {
    return 'Updated named range ' + plan.name + '.';
  }

  return 'Updated named range.';
}

function isRangeTransferPlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === 'string' &&
    typeof plan.sourceRange === 'string' &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string' &&
    (plan.operation === 'copy' || plan.operation === 'move' || plan.operation === 'append')
  );
}

function isDataCleanupPlan_(plan) {
  return Boolean(
    plan &&
    typeof plan.targetSheet === 'string' &&
    typeof plan.targetRange === 'string' &&
    typeof plan.operation === 'string' &&
    [
      'trim_whitespace',
      'remove_blank_rows',
      'remove_duplicate_rows',
      'normalize_case',
      'split_column',
      'join_columns',
      'fill_down',
      'standardize_format'
    ].indexOf(plan.operation) !== -1
  );
}

function getRangeTransferStatusSummary_(plan) {
  const operation = plan && (plan.transferOperation || plan.operation);
  const source = plan && plan.sourceSheet && plan.sourceRange
    ? plan.sourceSheet + '!' + plan.sourceRange
    : 'the source range';
  const target = plan && plan.targetSheet && plan.targetRange
    ? plan.targetSheet + '!' + plan.targetRange
    : 'the target range';

  switch (operation) {
    case 'copy':
      return 'Copied ' + source + ' to ' + target + '.';
    case 'move':
      return 'Moved ' + source + ' to ' + target + '.';
    case 'append':
      return 'Appended ' + source + ' into ' + target + '.';
    default:
      return 'Transferred ' + source + ' to ' + target + '.';
  }
}

function getDataCleanupStatusSummary_(plan) {
  const operation = plan && (plan.cleanupOperation || plan.operation);
  const target = plan && plan.targetSheet && plan.targetRange
    ? plan.targetSheet + '!' + plan.targetRange
    : 'the target range';

  switch (operation) {
    case 'trim_whitespace':
      return 'Trimmed whitespace in ' + target + '.';
    case 'remove_blank_rows':
      return 'Removed blank rows from ' + target + '.';
    case 'remove_duplicate_rows':
      return 'Removed duplicate rows from ' + target + '.';
    case 'normalize_case':
      return 'Normalized case in ' + target + '.';
    case 'split_column':
      return 'Split column values in ' + target + '.';
    case 'join_columns':
      return 'Joined column values in ' + target + '.';
    case 'fill_down':
      return 'Filled down values in ' + target + '.';
    case 'standardize_format':
      return 'Standardized format in ' + target + '.';
    default:
      return 'Applied cleanup in ' + target + '.';
  }
}

function getExternalDataStatusSummary_(plan) {
  const target = plan && plan.targetSheet && plan.targetRange
    ? plan.targetSheet + '!' + plan.targetRange
    : 'the target cell';

  if (plan && plan.sourceType === 'market_data') {
    const symbol = plan.query && typeof plan.query.symbol === 'string' && plan.query.symbol
      ? plan.query.symbol
      : 'the requested symbol';
    return 'Applied market data formula for ' + symbol + ' at ' + target + '.';
  }

  const provider = plan && typeof plan.provider === 'string' && plan.provider
    ? String(plan.provider).toUpperCase()
    : 'external';
  return 'Applied ' + provider + ' import formula at ' + target + '.';
}

function getRangeFormatStatusSummary_(plan) {
  const target = plan && plan.targetSheet && plan.targetRange
    ? plan.targetSheet + '!' + plan.targetRange
    : 'the target range';
  return 'Applied formatting to ' + target + '.';
}

function cloneMatrix_(matrix) {
  return (matrix || []).map(function(row) {
    return Array.isArray(row) ? row.slice() : [];
  });
}

function transposeMatrix_(matrix) {
  const rowCount = matrix ? matrix.length : 0;
  const columnCount = Math.max(0, ...((matrix || []).map(function(row) {
    return row ? row.length : 0;
  })));

  return Array.from({ length: columnCount }, function(_column, columnIndex) {
    return Array.from({ length: rowCount }, function(_row, rowIndex) {
      return matrix && matrix[rowIndex] ? matrix[rowIndex][columnIndex] : null;
    });
  });
}

function normalizeTransferValues_(values, plan) {
  const base = cloneMatrix_(values);
  return plan && plan.transpose ? transposeMatrix_(base) : base;
}

function normalizeTransferFormulas_(formulas, values, plan) {
  const formulaMatrix = (formulas || []).map(function(row, rowIndex) {
    return (row || []).map(function(cell, columnIndex) {
      if (typeof cell === 'string' && cell.trim().length > 0) {
        return {
          kind: 'formula',
          value: cell
        };
      }

      return {
        kind: 'value',
        value: values && values[rowIndex] ? values[rowIndex][columnIndex] : ''
      };
    });
  });

  return plan && plan.transpose ? transposeMatrix_(formulaMatrix) : formulaMatrix;
}

function parseA1CellReference_(reference) {
  const match = String(reference || '').trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error('Unsupported A1 reference: ' + reference);
  }

  return {
    row: Number(match[2]),
    column: convertColumnLettersToNumber_(match[1])
  };
}

function parseA1RangeReference_(reference) {
  const normalized = normalizeA1_(reference).trim().toUpperCase();
  const parts = normalized.split(':');
  const start = parseA1CellReference_(parts[0]);
  const end = parseA1CellReference_(parts[1] || parts[0]);

  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
    rowCount: Math.abs(end.row - start.row) + 1,
    columnCount: Math.abs(end.column - start.column) + 1
  };
}

function convertColumnNumberToLetters_(columnNumber) {
  let value = Number(columnNumber);
  let letters = '';

  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }

  return letters;
}

function buildA1RangeFromBounds_(bounds) {
  const startCell = convertColumnNumberToLetters_(bounds.startColumn) + bounds.startRow;
  const endCell = convertColumnNumberToLetters_(bounds.endColumn) + bounds.endRow;
  return startCell === endCell ? startCell : startCell + ':' + endCell;
}

function buildSizedA1RangeFromAnchor_(anchorRange, rowCount, columnCount) {
  const bounds = parseA1RangeReference_(anchorRange);
  return buildA1RangeFromBounds_({
    startRow: bounds.startRow,
    endRow: bounds.startRow + rowCount - 1,
    startColumn: bounds.startColumn,
    endColumn: bounds.startColumn + columnCount - 1
  });
}

function serializeExecutionSnapshotScalar_(value) {
  if (isDateObject_(value)) {
    return {
      type: 'date',
      value: value.toISOString()
    };
  }

  if (value === null) {
    return {
      type: 'null'
    };
  }

  if (value === undefined) {
    return {
      type: 'blank'
    };
  }

  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
    return {
      type: typeof value,
      value: value
    };
  }

  return {
    type: 'string',
    value: String(value)
  };
}

function deserializeExecutionSnapshotScalar_(serialized) {
  if (!serialized || typeof serialized !== 'object') {
    return '';
  }

  switch (serialized.type) {
    case 'date':
      return typeof serialized.value === 'string' ? new Date(serialized.value) : '';
    case 'null':
      return null;
    case 'number':
    case 'string':
    case 'boolean':
      return serialized.value;
    case 'blank':
    default:
      return '';
  }
}

function buildExecutionSnapshotCellMatrix_(values, formulas) {
  return (values || []).map(function(row, rowIndex) {
    return (row || []).map(function(value, columnIndex) {
      const formulaValue = formulas && formulas[rowIndex] && typeof formulas[rowIndex][columnIndex] === 'string'
        ? formulas[rowIndex][columnIndex].trim()
        : '';
      if (typeof formulaValue === 'string' && formulaValue.trim().startsWith('=')) {
        return {
          kind: 'formula',
          formula: formulaValue
        };
      }

      return {
        kind: 'value',
        value: serializeExecutionSnapshotScalar_(value)
      };
    });
  });
}

function createLocalExecutionSnapshot_(input) {
  if (!input || !input.executionId || !input.targetSheet || !input.targetRange) {
    return null;
  }

  return {
    baseExecutionId: input.executionId,
    targetSheet: input.targetSheet,
    targetRange: input.targetRange,
    beforeCells: buildExecutionSnapshotCellMatrix_(input.beforeValues, input.beforeFormulas),
    afterCells: buildExecutionSnapshotCellMatrix_(input.afterValues, input.afterFormulas)
  };
}

function attachLocalExecutionSnapshot_(result, snapshot) {
  if (!snapshot) {
    return result;
  }

  const nextResult = {};
  Object.keys(result || {}).forEach(function(key) {
    nextResult[key] = result[key];
  });
  nextResult.__hermesLocalExecutionSnapshot = snapshot;
  return nextResult;
}

function stripLocalExecutionSnapshot_(result) {
  if (!result || typeof result !== 'object' || Array.isArray(result)) {
    return result;
  }

  const nextResult = {};
  Object.keys(result).forEach(function(key) {
    if (key !== '__hermesLocalExecutionSnapshot') {
      nextResult[key] = result[key];
    }
  });
  return nextResult;
}

function resolveExecutionCellSnapshot_(input) {
  if (!input || typeof input !== 'object') {
    throw new Error('Execution snapshot payload is required.');
  }

  const cells = Array.isArray(input.cells) ? input.cells : [];
  if (!input.targetSheet || !input.targetRange || cells.length === 0) {
    throw new Error('That history entry is no longer available for exact undo or redo in this sheet session.');
  }

  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(input.targetSheet);
  if (!sheet) {
    throw new Error('Target sheet not found: ' + input.targetSheet);
  }

  const target = sheet.getRange(input.targetRange);
  if (
    target.getNumRows() !== cells.length ||
    target.getNumColumns() !== ((cells[0] && cells[0].length) || 0)
  ) {
    throw new Error('The saved undo snapshot no longer matches the current range shape.');
  }

  return {
    spreadsheet: spreadsheet,
    target: target,
    cells: cells
  };
}

function validateExecutionCellSnapshot(input) {
  const validated = resolveExecutionCellSnapshot_(input);
  return {
    ok: true,
    targetSheet: input.targetSheet,
    targetRange: input.targetRange,
    rowCount: validated.target.getNumRows(),
    columnCount: validated.target.getNumColumns()
  };
}

function applyExecutionCellSnapshot(input) {
  const validated = resolveExecutionCellSnapshot_(input);
  const target = validated.target;
  const cells = validated.cells;

  for (let rowIndex = 0; rowIndex < cells.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < (cells[rowIndex] || []).length; columnIndex += 1) {
      const cell = target.getCell(rowIndex + 1, columnIndex + 1);
      const snapshotCell = cells[rowIndex][columnIndex];
      if (snapshotCell && snapshotCell.kind === 'formula' && typeof snapshotCell.formula === 'string') {
        cell.setFormula(snapshotCell.formula);
      } else {
        cell.setValue(deserializeExecutionSnapshotScalar_(snapshotCell && snapshotCell.value));
      }
    }
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    targetSheet: input.targetSheet,
    targetRange: input.targetRange
  };
}

function validateExecutionCellSnapshot(input) {
  if (!input || typeof input !== 'object') {
    throw new Error('Execution snapshot payload is required.');
  }

  const cells = Array.isArray(input.cells) ? input.cells : [];
  if (!input.targetSheet || !input.targetRange || cells.length === 0) {
    throw new Error('That history entry is no longer available for exact undo or redo in this sheet session.');
  }

  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(input.targetSheet);
  if (!sheet) {
    throw new Error('Target sheet not found: ' + input.targetSheet);
  }

  const target = sheet.getRange(input.targetRange);
  if (
    target.getNumRows() !== cells.length ||
    target.getNumColumns() !== ((cells[0] && cells[0].length) || 0)
  ) {
    throw new Error('The saved undo snapshot no longer matches the current range shape.');
  }

  return {
    ok: true,
    targetSheet: input.targetSheet,
    targetRange: input.targetRange
  };
}

function rangesOverlapBounds_(left, right) {
  return !(
    left.endRow < right.startRow ||
    right.endRow < left.startRow ||
    left.endColumn < right.startColumn ||
    right.endColumn < left.startColumn
  );
}

function resolveColumnOffsetWithinRange_(columnRef, targetRangeA1) {
  const bounds = parseA1RangeReference_(targetRangeA1);
  const trimmed = String(columnRef || '').trim().toUpperCase();
  let column;

  if (/^\d+$/.test(trimmed)) {
    column = Number(trimmed);
  } else {
    column = parseA1CellReference_(trimmed + '1').column;
  }

  const offset = column - bounds.startColumn;
  if (offset < 0 || offset >= bounds.columnCount) {
    throw new Error('Column ' + columnRef + ' is outside ' + targetRangeA1 + '.');
  }

  return offset;
}

function getResolvedTransferShape_(sourceRange, plan) {
  return {
    rows: plan && plan.transpose ? sourceRange.getNumColumns() : sourceRange.getNumRows(),
    columns: plan && plan.transpose ? sourceRange.getNumRows() : sourceRange.getNumColumns()
  };
}

function resolveTransferTargetRange_(targetRange, expectedRows, expectedColumns) {
  if (targetRange.getNumRows() === expectedRows && targetRange.getNumColumns() === expectedColumns) {
    return targetRange;
  }

  if (targetRange.getNumRows() === 1 &&
    targetRange.getNumColumns() === 1 &&
    typeof targetRange.getResizedRange === 'function') {
    return targetRange.getResizedRange(expectedRows - 1, expectedColumns - 1);
  }

  throw new Error('The approved targetRange does not match the transfer shape.');
}

function deriveTransferTargetRangeA1_(plan, targetRange) {
  const planBounds = parseA1RangeReference_(plan.targetRange);
  if (planBounds.rowCount === targetRange.getNumRows() &&
    planBounds.columnCount === targetRange.getNumColumns()) {
    return normalizeA1_(plan.targetRange);
  }

  if (planBounds.rowCount === 1 && planBounds.columnCount === 1) {
    return buildA1RangeFromBounds_({
      startRow: planBounds.startRow,
      endRow: planBounds.startRow + targetRange.getNumRows() - 1,
      startColumn: planBounds.startColumn,
      endColumn: planBounds.startColumn + targetRange.getNumColumns() - 1
    });
  }

  return normalizeA1_(targetRange.getA1Notation());
}

function getActualAppendTargetRange_(targetRangeA1, startRowOffset, rowCount, columnCount) {
  const bounds = parseA1RangeReference_(normalizeA1_(targetRangeA1));
  return buildA1RangeFromBounds_({
    startRow: bounds.startRow + startRowOffset,
    endRow: bounds.startRow + startRowOffset + rowCount - 1,
    startColumn: bounds.startColumn,
    endColumn: bounds.startColumn + columnCount - 1
  });
}

function assertNonOverlappingTransfer_(plan, targetRangeA1) {
  if (!plan || plan.sourceSheet !== plan.targetSheet) {
    return;
  }

  if (plan.operation === 'copy') {
    return;
  }

  if (rangesOverlapBounds_(
    parseA1RangeReference_(plan.sourceRange),
    parseA1RangeReference_(targetRangeA1)
  )) {
    throw new Error('Google Sheets host cannot apply an overlapping ' + plan.operation + ' transfer exactly.');
  }
}

function clearTransferredSource_(sourceRange, plan) {
  if (plan.pasteMode === 'formats') {
    if (typeof sourceRange.clearFormat === 'function') {
      sourceRange.clearFormat();
      return;
    }

    throw new Error('Google Sheets host cannot clear the source formatting for this move.');
  }

  if (typeof sourceRange.clearContent === 'function') {
    sourceRange.clearContent();
    return;
  }

  if (typeof sourceRange.setValues === 'function') {
    sourceRange.setValues(Array.from({ length: sourceRange.getNumRows() }, function() {
      return Array.from({ length: sourceRange.getNumColumns() }, function() {
        return '';
      });
    }));
    return;
  }

  throw new Error('Google Sheets host cannot clear the source range for this move.');
}

function writeTransferToTarget_(targetRange, sourceRange, plan) {
  if (plan.pasteMode === 'formats') {
    if (typeof sourceRange.copyTo !== 'function' ||
      !SpreadsheetApp.CopyPasteType ||
      !SpreadsheetApp.CopyPasteType.PASTE_FORMAT) {
      throw new Error('Google Sheets host does not support exact-safe format transfers on this range.');
    }

    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, Boolean(plan && plan.transpose));
    return;
  }

  if (plan.pasteMode === 'values') {
    targetRange.setValues(normalizeTransferValues_(sourceRange.getValues(), plan));
    return;
  }

  if (plan.pasteMode === 'formulas') {
    if (typeof sourceRange.copyTo !== 'function' ||
      !SpreadsheetApp.CopyPasteType ||
      !SpreadsheetApp.CopyPasteType.PASTE_FORMULA) {
      throw new Error('Google Sheets host does not support exact-safe formula transfers on this range.');
    }

    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, Boolean(plan && plan.transpose));
    return;
  }

  throw new Error('Google Sheets host does not support exact-safe transfer pasteMode ' + plan.pasteMode + '.');
}

function getRangeOccupancyMatrix_(targetRange) {
  const values = targetRange.getValues();
  const formulas = typeof targetRange.getFormulas === 'function'
    ? targetRange.getFormulas()
    : [];

  return values.map(function(row, rowIndex) {
    const formulaRow = formulas[rowIndex] || [];
    return row.map(function(value, columnIndex) {
      const formulaValue = formulaRow[columnIndex];
      return (typeof formulaValue === 'string' && formulaValue.trim().length > 0) ||
        !isBlankCellValue_(value);
    });
  });
}

function getInsertedTransferMatrix_(sourceRange, plan) {
  if (plan.pasteMode === 'values') {
    return normalizeTransferValues_(sourceRange.getValues(), plan);
  }

  if (plan.pasteMode === 'formulas') {
    return normalizeTransferFormulas_(sourceRange.getFormulas(), sourceRange.getValues(), plan);
  }

  throw new Error('Google Sheets host does not support exact-safe transfer pasteMode ' + plan.pasteMode + '.');
}

function resolveAppendTransferTarget_(targetSheet, targetRange, sourceRange, plan) {
  const resolvedShape = getResolvedTransferShape_(sourceRange, plan);
  const insertedMatrix = plan.pasteMode === 'formats'
    ? null
    : getInsertedTransferMatrix_(sourceRange, plan);

  if (targetRange.getNumColumns() !== resolvedShape.columns) {
    throw new Error('Google Sheets host cannot append when the approved target range width does not match the transfer width.');
  }

  const occupancyMatrix = getRangeOccupancyMatrix_(targetRange);
  const anchorOnlyRange =
    targetRange.getNumRows() === 1 &&
    targetRange.getNumRows() < resolvedShape.rows &&
    occupancyMatrix.every(function(row) {
      return row.every(function(isOccupied) {
        return !isOccupied;
      });
    });

  if (anchorOnlyRange) {
    if (typeof targetSheet.getRange !== 'function') {
      throw new Error('Google Sheets host cannot expand the approved append anchor exactly.');
    }

    const expandedTargetRange = targetSheet.getRange(
      targetRange.getRow(),
      targetRange.getColumn(),
      resolvedShape.rows,
      resolvedShape.columns
    );

    if (expandedTargetRange.getNumRows() !== resolvedShape.rows ||
      expandedTargetRange.getNumColumns() !== resolvedShape.columns) {
      throw new Error('Google Sheets host cannot expand the approved append anchor exactly.');
    }

    return {
      writeRange: expandedTargetRange,
      insertedMatrix: insertedMatrix,
      actualTargetRange: getActualAppendTargetRange_(
        targetRange.getA1Notation(),
        0,
        resolvedShape.rows,
        resolvedShape.columns
      )
    };
  }

  let firstEmptyRow = occupancyMatrix.length;
  let seenGap = false;

  for (let rowIndex = 0; rowIndex < occupancyMatrix.length; rowIndex += 1) {
    const isEmptyRow = occupancyMatrix[rowIndex].every(function(isOccupied) {
      return !isOccupied;
    });

    if (isEmptyRow) {
      if (!seenGap) {
        firstEmptyRow = rowIndex;
        seenGap = true;
      }
      continue;
    }

    if (seenGap) {
      throw new Error('Google Sheets host cannot append exactly when the approved target range contains internal gaps.');
    }
  }

  if (firstEmptyRow + resolvedShape.rows > targetRange.getNumRows()) {
    throw new Error('Google Sheets host cannot append exactly within the approved target range.');
  }

  return {
    writeRange: targetSheet.getRange(
      targetRange.getRow() + firstEmptyRow,
      targetRange.getColumn(),
      resolvedShape.rows,
      resolvedShape.columns
    ),
    insertedMatrix: insertedMatrix,
    actualTargetRange: getActualAppendTargetRange_(
      targetRange.getA1Notation(),
      firstEmptyRow,
      resolvedShape.rows,
      resolvedShape.columns
    )
  };
}

function buildRangeTransferResult_(plan, actualTargetRange) {
  return {
    kind: 'range_transfer_update',
    hostPlatform: 'google_sheets',
    operation: 'range_transfer_update',
    sourceSheet: plan.sourceSheet,
    sourceRange: plan.sourceRange,
    targetSheet: plan.targetSheet,
    targetRange: actualTargetRange,
    transferOperation: plan.operation,
    pasteMode: plan.pasteMode,
    transpose: Boolean(plan.transpose),
    summary: getRangeTransferStatusSummary_({
      sourceSheet: plan.sourceSheet,
      sourceRange: plan.sourceRange,
      targetSheet: plan.targetSheet,
      targetRange: actualTargetRange,
      operation: plan.operation
    })
  };
}

function isBlankCellValue_(value) {
  return value === null || value === undefined || value === '';
}

function getCleanupColumnOffsets_(plan) {
  switch (plan.operation) {
    case 'remove_blank_rows':
    case 'remove_duplicate_rows':
      return (plan.keyColumns || []).map(function(columnRef) {
        return resolveColumnOffsetWithinRange_(columnRef, plan.targetRange);
      });
    case 'split_column':
      return {
        source: resolveColumnOffsetWithinRange_(plan.sourceColumn, plan.targetRange),
        targetStart: resolveColumnOffsetWithinRange_(plan.targetStartColumn, plan.targetRange)
      };
    case 'join_columns':
      return {
        source: (plan.sourceColumns || []).map(function(columnRef) {
          return resolveColumnOffsetWithinRange_(columnRef, plan.targetRange);
        }),
        target: resolveColumnOffsetWithinRange_(plan.targetColumn, plan.targetRange)
      };
    case 'fill_down':
      return (plan.columns || []).map(function(columnRef) {
        return resolveColumnOffsetWithinRange_(columnRef, plan.targetRange);
      });
    default:
      return [];
  }
}

function fillTrailingBlankRows_(rows, targetColumnCount, targetRowCount) {
  const paddedRows = rows.map(function(row) {
    return Array.from({ length: targetColumnCount }, function(_value, index) {
      return row[index] !== undefined ? row[index] : '';
    });
  });

  while (paddedRows.length < targetRowCount) {
    paddedRows.push(Array.from({ length: targetColumnCount }, function() {
      return '';
    }));
  }

  return paddedRows.slice(0, targetRowCount);
}

function hasAnyRealFormula_(formulas) {
  return (formulas || []).some(function(row) {
    return (row || []).some(function(value) {
      return typeof value === 'string' && value.trim().startsWith('=');
    });
  });
}

function toTitleCaseText_(value) {
  const lowerCased = String(value == null ? '' : value).toLocaleLowerCase();
  return lowerCased.replace(/(^|[^A-Za-z0-9])([A-Za-z0-9])/g, function(_match, prefix, character) {
    return prefix + character.toLocaleUpperCase();
  });
}

function getFormulaAwareCleanupTransform_(plan, hostLabel) {
  switch (plan.operation) {
    case 'trim_whitespace':
      return {
        applyToValue(value) {
          return typeof value === 'string' ? value.trim() : value;
        },
        formulaFunction: 'TRIM'
      };
    case 'normalize_case':
      switch (plan.mode) {
        case 'upper':
          return {
            applyToValue(value) {
              return typeof value === 'string' ? value.toLocaleUpperCase() : value;
            },
            formulaFunction: 'UPPER'
          };
        case 'lower':
          return {
            applyToValue(value) {
              return typeof value === 'string' ? value.toLocaleLowerCase() : value;
            },
            formulaFunction: 'LOWER'
          };
        case 'title':
          return {
            applyToValue(value) {
              return typeof value === 'string' ? toTitleCaseText_(value) : value;
            },
            formulaFunction: 'PROPER'
          };
        default:
          throw new Error(
            hostLabel + ' does not support exact-safe cleanup semantics for normalize_case mode ' +
            plan.mode +
            '.'
          );
      }
    default:
      return null;
  }
}

function wrapFormulaWithCleanupTransform_(formula, formulaFunction) {
  const normalizedFormula = typeof formula === 'string' ? formula.trim() : '';
  if (!normalizedFormula || !normalizedFormula.startsWith('=')) {
    return formula;
  }

  const expression = normalizedFormula.slice(1);
  return '=LET(_hermes_value, ' + expression + ', IF(ISTEXT(_hermes_value), ' +
    formulaFunction + '(_hermes_value), _hermes_value))';
}

function buildCleanupWriteMatrix_(plan, inputValues, inputFormulas, hostLabel) {
  const values = cloneMatrix_(inputValues);
  const formulas = cloneMatrix_(inputFormulas);
  const formulaAwareTransform = getFormulaAwareCleanupTransform_(plan, hostLabel);

  if (!hasAnyRealFormula_(formulas)) {
    return applyCleanupTransform_(plan, values, hostLabel);
  }

  if (!formulaAwareTransform) {
    throw new Error(hostLabel + ' cannot apply cleanup plans exactly when the target range contains formulas.');
  }

  return values.map(function(row, rowIndex) {
    return (row || []).map(function(value, columnIndex) {
      const formulaValue = formulas[rowIndex] && formulas[rowIndex][columnIndex];
      if (typeof formulaValue === 'string' && formulaValue.trim().startsWith('=')) {
        return wrapFormulaWithCleanupTransform_(formulaValue, formulaAwareTransform.formulaFunction);
      }

      return formulaAwareTransform.applyToValue(value);
    });
  });
}

function applyCleanupTransform_(plan, inputValues, hostLabel) {
  const values = cloneMatrix_(inputValues);
  const formulaAwareTransform = getFormulaAwareCleanupTransform_(plan, hostLabel || 'Google Sheets host');

  if (formulaAwareTransform) {
    return values.map(function(row) {
      return row.map(function(value) {
        return formulaAwareTransform.applyToValue(value);
      });
    });
  }

  switch (plan.operation) {
    case 'remove_blank_rows': {
      const keyOffsets = getCleanupColumnOffsets_(plan);
      const retainedRows = values.filter(function(row) {
        const candidateValues = keyOffsets.length > 0
          ? keyOffsets.map(function(index) {
            return row[index];
          })
          : row;

        return candidateValues.some(function(value) {
          return !isBlankCellValue_(value);
        });
      });

      return fillTrailingBlankRows_(retainedRows, values[0] ? values[0].length : 0, values.length);
    }
    case 'remove_duplicate_rows': {
      const keyOffsets = getCleanupColumnOffsets_(plan);
      const retainedRows = [];
      const seen = {};

      values.forEach(function(row) {
        const keyValues = keyOffsets.length > 0
          ? keyOffsets.map(function(index) {
            return row[index];
          })
          : row;
        const digest = JSON.stringify(keyValues);

        if (seen[digest]) {
          return;
        }

        seen[digest] = true;
        retainedRows.push(row);
      });

      return fillTrailingBlankRows_(retainedRows, values[0] ? values[0].length : 0, values.length);
    }
    case 'split_column': {
      const offsets = getCleanupColumnOffsets_(plan);
      const targetCapacity = (values[0] ? values[0].length : 0) - offsets.targetStart;

      return values.map(function(row) {
        const parts = String(row[offsets.source] == null ? '' : row[offsets.source]).split(plan.delimiter);
        if (parts.length > targetCapacity) {
          throw new Error('Google Sheets host cannot split this column exactly within the approved target range.');
        }

        const nextRow = row.slice();
        for (let offset = 0; offset < targetCapacity; offset += 1) {
          nextRow[offsets.targetStart + offset] = parts[offset] !== undefined ? parts[offset] : '';
        }
        return nextRow;
      });
    }
    case 'join_columns': {
      const offsets = getCleanupColumnOffsets_(plan);

      return values.map(function(row) {
        const nextRow = row.slice();
        nextRow[offsets.target] = offsets.source.map(function(index) {
          return String(row[index] == null ? '' : row[index]);
        }).join(plan.delimiter);
        return nextRow;
      });
    }
    case 'fill_down': {
      const explicitOffsets = getCleanupColumnOffsets_(plan);
      const targetOffsets = explicitOffsets.length > 0
        ? explicitOffsets
        : Array.from({ length: values[0] ? values[0].length : 0 }, function(_value, index) {
          return index;
        });
      const nextValues = cloneMatrix_(values);

      targetOffsets.forEach(function(columnIndex) {
        let lastSeen = null;

        for (let rowIndex = 0; rowIndex < nextValues.length; rowIndex += 1) {
          const currentValue = nextValues[rowIndex][columnIndex];
          if (isBlankCellValue_(currentValue)) {
            if (lastSeen !== null) {
              nextValues[rowIndex][columnIndex] = lastSeen;
            }
          } else {
            lastSeen = currentValue;
          }
        }
      });

      return nextValues;
    }
    case 'standardize_format':
      return standardizeFormatMatrixExact_(plan, values, 'Google Sheets host');
    default:
      throw new Error(
        'Google Sheets host does not support exact-safe cleanup semantics for ' + plan.operation + '.'
      );
  }
}

function getInsertSheetIndex_(spreadsheet, position) {
  const sheetCount = spreadsheet.getSheets().length;
  if (position === 'start') {
    return 0;
  }

  if (position === 'end' || position === undefined) {
    return sheetCount;
  }

  return Math.max(0, Math.min(position, sheetCount));
}

function getMoveSheetPosition_(spreadsheet, position) {
  const sheetCount = spreadsheet.getSheets().length;
  if (position === 'start') {
    return 1;
  }

  if (position === 'end' || position === undefined) {
    return sheetCount;
  }

  return Math.max(1, Math.min(position + 1, sheetCount));
}

function getWrapStrategyEnum_(strategy) {
  switch (strategy) {
    case 'wrap':
      return SpreadsheetApp.WrapStrategy.WRAP;
    case 'clip':
      return SpreadsheetApp.WrapStrategy.CLIP;
    case 'overflow':
      return SpreadsheetApp.WrapStrategy.OVERFLOW;
    default:
      return null;
  }
}

function getRowSliceRange_(sheet, startIndex, count) {
  return sheet.getRange(startIndex + 1, 1, count, 1);
}

function getColumnSliceRange_(sheet, startIndex, count) {
  return sheet.getRange(1, startIndex + 1, 1, count);
}

function convertColumnLettersToNumber_(value) {
  let total = 0;
  const text = String(value || '').trim().toUpperCase();

  for (let index = 0; index < text.length; index += 1) {
    total = (total * 26) + (text.charCodeAt(index) - 64);
  }

  return total;
}

function resolveRelativeColumnRef_(columnRef, target, hasHeader) {
  const width = target.getNumColumns();

  if (typeof columnRef === 'number') {
    return columnRef >= 1 && columnRef <= width ? columnRef : null;
  }

  if (typeof columnRef !== 'string') {
    return null;
  }

  const trimmed = columnRef.trim();
  if (!trimmed) {
    return null;
  }

  if (hasHeader) {
    const headerRow = target.getDisplayValues()[0] || [];
    for (let index = 0; index < headerRow.length; index += 1) {
      if (String(headerRow[index]).trim() === trimmed) {
        return index + 1;
      }
    }
  }

  if (/^[A-Z]+$/i.test(trimmed)) {
    const absoluteColumn = convertColumnLettersToNumber_(trimmed);
    const relativeColumn = absoluteColumn - target.getColumn() + 1;
    return relativeColumn >= 1 && relativeColumn <= width ? relativeColumn : null;
  }

  const numericColumn = Number(trimmed);
  if (Number.isInteger(numericColumn)) {
    return resolveRelativeColumnRef_(numericColumn, target, hasHeader);
  }

  return null;
}

function buildGoogleSheetsSortSpecs_(plan, target) {
  return buildSortSpec_({
    keys: plan.keys.map(function(key) {
      return {
        columnRef: resolveRelativeColumnRef_(key && key.columnRef, target, plan.hasHeader),
        direction: key && key.direction
      };
    })
  }).map(function(spec) {
    return {
      column: spec.dimensionIndex + 1,
      ascending: spec.sortOrder !== 'DESCENDING'
    };
  });
}

function getOrCreateFilter_(sheet, target, clearExistingFilters) {
  let filter = sheet.getFilter();

  if (filter) {
    const existingRange = normalizeA1_(filter.getRange().getA1Notation());
    const targetRange = normalizeA1_(target.getA1Notation());

    if (existingRange !== targetRange) {
      if (!clearExistingFilters) {
        throw new Error(
          'Existing sheet filter applies to ' + existingRange + '. Rejecting this plan without clearExistingFilters=true.'
        );
      }

      filter.remove();
      filter = null;
    }
  }

  if (!filter) {
    filter = target.createFilter();
  }

  if (clearExistingFilters) {
    for (let columnIndex = 1; columnIndex <= target.getNumColumns(); columnIndex += 1) {
      filter.removeColumnFilterCriteria(columnIndex);
    }
  }

  return filter;
}

function buildFilterCriteria_(condition) {
  const builder = SpreadsheetApp.newFilterCriteria();

  switch (condition.operator) {
    case 'equals':
      if (typeof condition.value === 'number') {
        return builder.whenNumberEqualTo(condition.value).build();
      }

      return builder.whenTextEqualTo(String(condition.value)).build();
    case 'notEquals':
      if (typeof condition.value === 'number') {
        return builder.whenNumberNotEqualTo(condition.value).build();
      }

      return builder.setHiddenValues([String(condition.value)]).build();
    case 'contains':
      return builder.whenTextContains(String(condition.value)).build();
    case 'startsWith':
      return builder.whenTextStartsWith(String(condition.value)).build();
    case 'endsWith':
      return builder.whenTextEndsWith(String(condition.value)).build();
    case 'greaterThan':
      return builder.whenNumberGreaterThan(condition.value).build();
    case 'greaterThanOrEqual':
      return builder.whenNumberGreaterThanOrEqualTo(condition.value).build();
    case 'lessThan':
      return builder.whenNumberLessThan(condition.value).build();
    case 'lessThanOrEqual':
      return builder.whenNumberLessThanOrEqualTo(condition.value).build();
    case 'isEmpty':
      return builder.whenCellEmpty().build();
    case 'isNotEmpty':
      return builder.whenCellNotEmpty().build();
    case 'topN':
      throw new Error('Google Sheets grid filters cannot represent operator "topN" exactly.');
    default:
      throw new Error('Unsupported filter operator: ' + condition.operator);
  }
}

function splitSheetAndRangeA1_(value) {
  const match = String(value || '').match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
  if (!match) {
    return null;
  }

  const quotedSheetName = match[1];
  const plainSheetName = match[2];
  const rangeA1 = match[3];

  return {
    sheetName: quotedSheetName
      ? quotedSheetName.replace(/''/g, "'")
      : plainSheetName,
    rangeA1: rangeA1
  };
}

function resolveSourceRange_(spreadsheet, fallbackSheet, sourceRange) {
  const parsed = splitSheetAndRangeA1_(sourceRange);
  const sheet = parsed
    ? spreadsheet.getSheetByName(parsed.sheetName)
    : fallbackSheet;
  if (!sheet) {
    throw new Error('Validation source sheet not found: ' + (parsed ? parsed.sheetName : fallbackSheet.getName()));
  }

  return sheet.getRange(parsed ? parsed.rangeA1 : sourceRange);
}

function findNamedRange_(spreadsheet, name) {
  const namedRanges = spreadsheet.getNamedRanges ? spreadsheet.getNamedRanges() : [];

  for (let index = 0; index < namedRanges.length; index += 1) {
    const namedRange = namedRanges[index];
    const candidateName = typeof namedRange.getName === 'function'
      ? namedRange.getName()
      : namedRange.name;
    if (candidateName === name) {
      return namedRange;
    }
  }

  return null;
}

function applyValidationBuilderOptions_(builder, plan) {
  if (plan.invalidDataBehavior !== 'reject' && plan.invalidDataBehavior !== 'warn') {
    throw new Error('Unsupported invalidDataBehavior: ' + plan.invalidDataBehavior);
  }

  builder.setAllowInvalid(plan.invalidDataBehavior !== 'reject');

  if (plan.helpText) {
    builder.setHelpText(plan.helpText);
  }

  return builder;
}

function extractTopLeftCellA1_(targetRangeA1) {
  return String(targetRangeA1 || '').split(':')[0];
}

function parseDateLiteral_(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    throw new Error('Invalid date literal: ' + value);
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const parsed = new Date(year, month - 1, day);

  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) {
    throw new Error('Invalid date literal: ' + value);
  }

  return parsed;
}

function buildComparisonExpression_(cellA1, comparator, value, value2) {
  switch (comparator) {
    case 'between':
      return 'AND(' + cellA1 + '>=' + value + ',' + cellA1 + '<=' + value2 + ')';
    case 'not_between':
      return 'OR(' + cellA1 + '<' + value + ',' + cellA1 + '>' + value2 + ')';
    case 'equal_to':
      return cellA1 + '=' + value;
    case 'not_equal_to':
      return cellA1 + '<>' + value;
    case 'greater_than':
      return cellA1 + '>' + value;
    case 'greater_than_or_equal_to':
      return cellA1 + '>=' + value;
    case 'less_than':
      return cellA1 + '<' + value;
    case 'less_than_or_equal_to':
      return cellA1 + '<=' + value;
    default:
      throw new Error('Unsupported validation comparator: ' + comparator);
  }
}

function buildBlankAwareValidationFormula_(cellA1, innerExpression, allowBlank) {
  if (allowBlank) {
    return '=OR(ISBLANK(' + cellA1 + '),' + innerExpression + ')';
  }

  return '=AND(NOT(ISBLANK(' + cellA1 + ')),' + innerExpression + ')';
}

function buildWholeNumberValidationFormula_(plan, targetRangeA1) {
  const cellA1 = extractTopLeftCellA1_(targetRangeA1);
  const condition = buildComparisonExpression_(cellA1, plan.comparator, plan.value, plan.value2);
  const integerConstraint = 'AND(ISNUMBER(' + cellA1 + '),' +
    cellA1 + '=INT(' + cellA1 + '),' +
    condition +
    ')';

  return buildBlankAwareValidationFormula_(cellA1, integerConstraint, plan.allowBlank);
}

function buildDecimalValidationFormula_(plan, targetRangeA1) {
  const cellA1 = extractTopLeftCellA1_(targetRangeA1);
  const comparison = buildComparisonExpression_(cellA1, plan.comparator, plan.value, plan.value2);
  const numericConstraint = 'AND(ISNUMBER(' + cellA1 + '),' + comparison + ')';

  return buildBlankAwareValidationFormula_(cellA1, numericConstraint, plan.allowBlank);
}

function formatDateLiteralFormula_(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    throw new Error('Invalid date literal: ' + value);
  }

  return 'DATE(' + Number(match[1]) + ',' + Number(match[2]) + ',' + Number(match[3]) + ')';
}

function buildDateValidationFormula_(plan, targetRangeA1) {
  const cellA1 = extractTopLeftCellA1_(targetRangeA1);
  const comparison = buildComparisonExpression_(
    cellA1,
    plan.comparator,
    formatDateLiteralFormula_(plan.value),
    plan.value2 ? formatDateLiteralFormula_(plan.value2) : undefined
  );
  const dateConstraint = 'AND(ISNUMBER(' + cellA1 + '),' + comparison + ')';

  return buildBlankAwareValidationFormula_(cellA1, dateConstraint, plan.allowBlank);
}

function buildTextLengthValidationFormula_(plan, targetRangeA1) {
  const cellA1 = extractTopLeftCellA1_(targetRangeA1);
  const comparison = buildComparisonExpression_('LEN(' + cellA1 + ')', plan.comparator, plan.value, plan.value2);
  const textLengthConstraint = 'AND(LEN(' + cellA1 + ')>=0,' + comparison + ')';

  return buildBlankAwareValidationFormula_(cellA1, textLengthConstraint, plan.allowBlank);
}

function normalizeCustomValidationFormula_(formula) {
  return String(formula || '').replace(/^=/, '');
}

function buildCustomValidationFormula_(plan, targetRangeA1) {
  const cellA1 = extractTopLeftCellA1_(targetRangeA1);
  const expression = normalizeCustomValidationFormula_(plan.formula);

  if (!expression) {
    throw new Error('Custom formula validation requires a non-empty formula.');
  }

  return buildBlankAwareValidationFormula_(cellA1, '(' + expression + ')', plan.allowBlank);
}

function buildDataValidationRule_(spreadsheet, sheet, target, plan) {
  const builder = applyValidationBuilderOptions_(SpreadsheetApp.newDataValidation(), plan);

  switch (plan.ruleType) {
    case 'list':
      if (plan.allowBlank) {
        throw new Error('Google Sheets host cannot represent allowBlank=true exactly for list validation.');
      }

      if (Array.isArray(plan.values)) {
        return builder.requireValueInList(
          plan.values,
          typeof plan.showDropdown === 'boolean' ? plan.showDropdown : undefined
        ).build();
      }

      if (typeof plan.sourceRange === 'string' && plan.sourceRange.length > 0) {
        return builder.requireValueInRange(
          resolveSourceRange_(spreadsheet, sheet, plan.sourceRange),
          typeof plan.showDropdown === 'boolean' ? plan.showDropdown : undefined
        ).build();
      }

      if (typeof plan.namedRangeName === 'string' && plan.namedRangeName.length > 0) {
        const namedRange = findNamedRange_(spreadsheet, plan.namedRangeName);
        if (!namedRange || typeof namedRange.getRange !== 'function') {
          throw new Error('Named range not found: ' + plan.namedRangeName);
        }

        return builder.requireValueInRange(
          namedRange.getRange(),
          typeof plan.showDropdown === 'boolean' ? plan.showDropdown : undefined
        ).build();
      }

      throw new Error('List validation requires values, sourceRange, or namedRangeName.');
    case 'checkbox':
      if (plan.allowBlank && !(plan.checkedValue !== undefined && plan.uncheckedValue === undefined)) {
        throw new Error('Google Sheets host cannot represent allowBlank=true exactly for this checkbox configuration.');
      }

      if (!plan.allowBlank && plan.checkedValue !== undefined && plan.uncheckedValue === undefined) {
        throw new Error('Google Sheets host cannot represent allowBlank=false exactly for single-value checkbox validation.');
      }

      if (plan.checkedValue === undefined && plan.uncheckedValue !== undefined) {
        throw new Error('Google Sheets checkbox validation cannot set uncheckedValue without checkedValue.');
      }

      if (plan.checkedValue !== undefined && plan.uncheckedValue !== undefined) {
        return builder.requireCheckbox(plan.checkedValue, plan.uncheckedValue).build();
      }

      if (plan.checkedValue !== undefined) {
        return builder.requireCheckbox(plan.checkedValue).build();
      }

      return builder.requireCheckbox().build();
    case 'whole_number':
      return builder.requireFormulaSatisfied(
        buildWholeNumberValidationFormula_(plan, target.getA1Notation())
      ).build();
    case 'decimal':
      if (plan.allowBlank) {
        return builder.requireFormulaSatisfied(
          buildDecimalValidationFormula_(plan, target.getA1Notation())
        ).build();
      }

      switch (plan.comparator) {
        case 'between':
          return builder.requireNumberBetween(plan.value, plan.value2).build();
        case 'not_between':
          return builder.requireNumberNotBetween(plan.value, plan.value2).build();
        case 'equal_to':
          return builder.requireNumberEqualTo(plan.value).build();
        case 'not_equal_to':
          return builder.requireNumberNotEqualTo(plan.value).build();
        case 'greater_than':
          return builder.requireNumberGreaterThan(plan.value).build();
        case 'greater_than_or_equal_to':
          return builder.requireNumberGreaterThanOrEqualTo(plan.value).build();
        case 'less_than':
          return builder.requireNumberLessThan(plan.value).build();
        case 'less_than_or_equal_to':
          return builder.requireNumberLessThanOrEqualTo(plan.value).build();
        default:
          throw new Error('Unsupported decimal validation comparator: ' + plan.comparator);
      }
    case 'date': {
      if (plan.allowBlank) {
        return builder.requireFormulaSatisfied(
          buildDateValidationFormula_(plan, target.getA1Notation())
        ).build();
      }

      const dateValue = parseDateLiteral_(plan.value);
      const secondDateValue = plan.value2 ? parseDateLiteral_(plan.value2) : null;

      switch (plan.comparator) {
        case 'between':
          return builder.requireDateBetween(dateValue, secondDateValue).build();
        case 'not_between':
          return builder.requireDateNotBetween(dateValue, secondDateValue).build();
        case 'equal_to':
          return builder.requireDateEqualTo(dateValue).build();
        case 'greater_than':
          return builder.requireDateAfter(dateValue).build();
        case 'greater_than_or_equal_to':
          return builder.requireDateOnOrAfter(dateValue).build();
        case 'less_than':
          return builder.requireDateBefore(dateValue).build();
        case 'less_than_or_equal_to':
          return builder.requireDateOnOrBefore(dateValue).build();
        default:
          throw new Error('Unsupported date validation comparator: ' + plan.comparator);
      }
    }
    case 'text_length':
      if (plan.allowBlank) {
        return builder.requireFormulaSatisfied(
          buildTextLengthValidationFormula_(plan, target.getA1Notation())
        ).build();
      }

      switch (plan.comparator) {
        case 'between':
          return builder.requireTextLengthBetween(plan.value, plan.value2).build();
        case 'not_between':
          return builder.requireTextLengthNotBetween(plan.value, plan.value2).build();
        case 'equal_to':
          return builder.requireTextLengthEqualTo(plan.value).build();
        case 'greater_than':
          return builder.requireTextLengthGreaterThan(plan.value).build();
        case 'greater_than_or_equal_to':
          return builder.requireTextLengthGreaterThanOrEqualTo(plan.value).build();
        case 'less_than':
          return builder.requireTextLengthLessThan(plan.value).build();
        case 'less_than_or_equal_to':
          return builder.requireTextLengthLessThanOrEqualTo(plan.value).build();
        default:
          throw new Error('Unsupported text length validation comparator: ' + plan.comparator);
      }
    case 'custom_formula':
      return builder.requireFormulaSatisfied(
        buildCustomValidationFormula_(plan, target.getA1Notation())
      ).build();
    default:
      throw new Error('Unsupported Google Sheets data validation rule type.');
  }
}

function getSpreadsheetSnapshot(prompt, sessionId) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const range = spreadsheet.getActiveRange();
  const activeCell = spreadsheet.getCurrentCell();
  let currentRegion = range;
  try {
    if (typeof activeCell.getDataRegion === 'function') {
      currentRegion = activeCell.getDataRegion();
    }
  } catch (_error) {
    currentRegion = range;
  }
  const selectedRange = normalizeA1_(range.getA1Notation());
  const includeSelectionMatrix = shouldIncludeRegionMatrix_(selectedRange);
  let selectionHeaders = null;
  const currentRegionRange = normalizeA1_(currentRegion.getA1Notation());
  const includeCurrentRegionMatrix = shouldIncludeRegionMatrix_(currentRegionRange);
  let currentRegionHeaders = null;
  const activeCellContext = buildCellContext_(activeCell);
  const referencedCellsContext = buildReferencedCellsContext_(sheet, prompt, activeCellContext);

  const selectionContext = {
    range: selectedRange
  };

  if (includeSelectionMatrix) {
    const values = range.getValues();
    const displayValues = range.getDisplayValues();
    const formulas = range.getFormulas();
    const normalizedValues = normalizeMatrixValues_(values, displayValues);
    selectionHeaders = getHeaders_(normalizedValues.slice(0, 1));
    selectionContext.values = normalizedValues;
    selectionContext.formulas = normalizeFormulas_(formulas);
  } else {
    const selectionHeaderRange = range.offset(
      0,
      0,
      Math.min(1, range.getNumRows()),
      range.getNumColumns()
    );
    const selectionHeaderValues = normalizeMatrixValues_(
      selectionHeaderRange.getValues(),
      selectionHeaderRange.getDisplayValues()
    );
    selectionHeaders = getHeaders_(selectionHeaderValues);
  }

  if (Array.isArray(selectionHeaders) && selectionHeaders.length > 0) {
    selectionContext.headers = selectionHeaders;
  }

  const context = {
    selection: selectionContext,
    currentRegion: {
      range: currentRegionRange
    },
    activeCell: activeCellContext
  };

  if (includeCurrentRegionMatrix) {
    const currentRegionValues = currentRegion.getValues();
    const currentRegionDisplayValues = currentRegion.getDisplayValues();
    const currentRegionFormulas = currentRegion.getFormulas();
    const normalizedCurrentRegionValues = normalizeMatrixValues_(currentRegionValues, currentRegionDisplayValues);
    currentRegionHeaders = getHeaders_(normalizedCurrentRegionValues.slice(0, 1));
    context.currentRegion.values = normalizedCurrentRegionValues;
    context.currentRegion.formulas = normalizeFormulas_(currentRegionFormulas);
  } else {
    const currentRegionHeaderRange = currentRegion.offset(
      0,
      0,
      Math.min(1, currentRegion.getNumRows()),
      currentRegion.getNumColumns()
    );
    const currentRegionHeaderValues = normalizeMatrixValues_(
      currentRegionHeaderRange.getValues(),
      currentRegionHeaderRange.getDisplayValues()
    );
    currentRegionHeaders = getHeaders_(currentRegionHeaderValues);
  }

  if (Array.isArray(currentRegionHeaders) && currentRegionHeaders.length > 0) {
    context.currentRegion.headers = currentRegionHeaders;
  }

  if (Array.isArray(referencedCellsContext) && referencedCellsContext.length > 0) {
    context.referencedCells = referencedCellsContext;
  }

  const implicitTargets = buildImplicitRegionTargets_(currentRegionRange);
  if (implicitTargets.currentRegionArtifactTarget) {
    context.currentRegionArtifactTarget = implicitTargets.currentRegionArtifactTarget;
  }
  if (implicitTargets.currentRegionAppendTarget) {
    context.currentRegionAppendTarget = implicitTargets.currentRegionAppendTarget;
  }

  const normalizedSessionId = typeof sessionId === 'string' && sessionId.trim().length > 0
    ? sessionId.trim()
    : PropertiesService.getUserProperties().getProperty('HERMES_SESSION_ID') || '';
  if (normalizedSessionId) {
    PropertiesService.getUserProperties().setProperty('HERMES_SESSION_ID', normalizedSessionId);
  }

  return {
    source: {
      channel: 'google_sheets',
      clientVersion: getRuntimeConfig().clientVersion,
      sessionId: normalizedSessionId
    },
    host: {
      platform: 'google_sheets',
      workbookTitle: spreadsheet.getName(),
      workbookId: spreadsheet.getId(),
      activeSheet: sheet.getName(),
      selectedRange: normalizeA1_(range.getA1Notation()),
      locale: Session.getActiveUserLocale(),
      timeZone: Session.getScriptTimeZone()
    },
    context: context
  };
}

function applyCompositePlan_(input) {
  if (typeof input.executionId !== 'string' || input.executionId.trim().length === 0) {
    throw new Error('Composite workflow execution requires executionId.');
  }

  const stepResults = [];
  const completedSteps = {};
  const failedSteps = {};
  const skippedSteps = {};
  let halted = false;

  input.plan.steps.forEach(function(step) {
    if (halted) {
      stepResults.push({
        stepId: step.stepId,
        status: 'skipped',
        summary: 'Skipped because an earlier workflow step failed.'
      });
      skippedSteps[step.stepId] = true;
      return;
    }

    if ((step.dependsOn || []).some(function(dependency) {
      return failedSteps[dependency] || skippedSteps[dependency];
    })) {
      stepResults.push({
        stepId: step.stepId,
        status: 'skipped',
        summary: 'Skipped because a dependency failed or was skipped.'
      });
      skippedSteps[step.stepId] = true;
      return;
    }

    const unresolvedDependency = (step.dependsOn || []).find(function(dependency) {
      return !completedSteps[dependency] && !failedSteps[dependency] && !skippedSteps[dependency];
    });
    if (unresolvedDependency) {
      stepResults.push({
        stepId: step.stepId,
        status: 'failed',
        summary: 'Dependency ' + unresolvedDependency + ' has not completed before this step.'
      });
      failedSteps[step.stepId] = true;
      if (!step.continueOnError) {
        halted = true;
      }
      return;
    }

    try {
      const result = applyWritePlan({
        requestId: input.requestId,
        runId: input.runId,
        approvalToken: input.approvalToken,
        plan: step.plan
      });
      const gatewayResult = stripLocalExecutionSnapshot_(result);
      stepResults.push({
        stepId: step.stepId,
        status: 'completed',
        summary: getCompositeStepWritebackStatusLine_(step.plan, result),
        result: gatewayResult
      });
      completedSteps[step.stepId] = true;
    } catch (error) {
      const summary = sanitizeHostExecutionError_(error, 'Workflow step failed.');
      stepResults.push({
        stepId: step.stepId,
        status: 'failed',
        summary: summary
      });
      failedSteps[step.stepId] = true;
      if (!step.continueOnError) {
        halted = true;
      }
    }
  });

  return {
    kind: 'composite_update',
    operation: 'composite_update',
    hostPlatform: 'google_sheets',
    executionId: input.executionId,
    stepResults: stepResults,
    summary: getCompositeStatusSummary_({
      stepResults: stepResults,
      summary: buildCompositeExecutionSummary_(stepResults)
    })
  };
}

function applyWritePlan(input) {
  const plan = input.plan;
  if (isCompositePlan_(plan)) {
    return applyCompositePlan_(input);
  }
  const spreadsheet = SpreadsheetApp.getActive();

  if (isAnalysisReportPlan_(plan)) {
    if (plan.outputMode !== 'materialize_report') {
      throw new Error('Chat-only analysis reports are not writeback eligible.');
    }

    const reportSheet = spreadsheet.getSheetByName(plan.targetSheet);
    if (!reportSheet) {
      throw new Error('Target sheet not found: ' + plan.targetSheet);
    }

    const reportMatrix = buildAnalysisReportMatrix_(plan);
    const targetAnchor = reportSheet.getRange(plan.targetRange);
    const resolvedTargetRange = resolveTransferTargetRange_(
      targetAnchor,
      reportMatrix.length,
      reportMatrix[0] ? reportMatrix[0].length : 1
    );
    const actualTargetRange = deriveTransferTargetRangeA1_(plan, resolvedTargetRange);
    const beforeValues = resolvedTargetRange.getValues();
    const beforeFormulas = resolvedTargetRange.getFormulas();

    resolvedTargetRange.setValues(reportMatrix);
    SpreadsheetApp.flush();

    return attachLocalExecutionSnapshot_({
      kind: 'analysis_report_update',
      hostPlatform: 'google_sheets',
      sourceSheet: plan.sourceSheet,
      sourceRange: plan.sourceRange,
      outputMode: 'materialize_report',
      targetSheet: plan.targetSheet,
      targetRange: actualTargetRange,
      sections: plan.sections,
      explanation: plan.explanation,
      confidence: plan.confidence,
      requiresConfirmation: plan.requiresConfirmation,
      affectedRanges: normalizeAnalysisReportAffectedRanges_(plan, actualTargetRange),
      overwriteRisk: plan.overwriteRisk,
      confirmationLevel: plan.confirmationLevel,
      summary: 'Created analysis report on ' + plan.targetSheet + '!' + actualTargetRange + '.'
    }, createLocalExecutionSnapshot_({
      executionId: input.executionId,
      targetSheet: plan.targetSheet,
      targetRange: actualTargetRange,
      beforeValues: beforeValues,
      beforeFormulas: beforeFormulas,
      afterValues: resolvedTargetRange.getValues(),
      afterFormulas: resolvedTargetRange.getFormulas()
    }));
  }

  if (isPivotTablePlan_(plan)) {
    return applyPivotTablePlan_(spreadsheet, plan);
  }

  if (isChartPlan_(plan)) {
    return applyChartPlan_(spreadsheet, plan);
  }

  if (isWorkbookStructurePlan_(plan)) {
    switch (plan.operation) {
      case 'create_sheet': {
        const createdSheet = spreadsheet.insertSheet(plan.sheetName, getInsertSheetIndex_(spreadsheet, plan.position));
        SpreadsheetApp.flush();
        return {
          kind: 'workbook_structure_update',
          hostPlatform: 'google_sheets',
          sheetName: createdSheet.getName(),
          operation: plan.operation,
          positionResolved: createdSheet.getIndex() - 1,
          sheetCount: spreadsheet.getSheets().length,
          summary: getWorkbookStructureStatusSummary_(plan)
        };
      }
      case 'delete_sheet': {
        const sheet = spreadsheet.getSheetByName(plan.sheetName);
        if (!sheet) {
          throw new Error('Target sheet not found: ' + plan.sheetName);
        }

        spreadsheet.deleteSheet(sheet);
        SpreadsheetApp.flush();
        return {
          kind: 'workbook_structure_update',
          hostPlatform: 'google_sheets',
          sheetName: plan.sheetName,
          operation: plan.operation,
          summary: getWorkbookStructureStatusSummary_(plan)
        };
      }
      case 'rename_sheet': {
        const sheet = spreadsheet.getSheetByName(plan.sheetName);
        if (!sheet) {
          throw new Error('Target sheet not found: ' + plan.sheetName);
        }

        sheet.setName(plan.newSheetName);
        SpreadsheetApp.flush();
        return {
          kind: 'workbook_structure_update',
          hostPlatform: 'google_sheets',
          sheetName: plan.sheetName,
          operation: plan.operation,
          newSheetName: sheet.getName(),
          summary: getWorkbookStructureStatusSummary_(plan)
        };
      }
      case 'duplicate_sheet': {
        const sourceSheet = spreadsheet.getSheetByName(plan.sheetName);
        if (!sourceSheet) {
          throw new Error('Target sheet not found: ' + plan.sheetName);
        }

        const copiedSheet = sourceSheet.copyTo(spreadsheet);
        if (plan.newSheetName) {
          copiedSheet.setName(plan.newSheetName);
        }

        spreadsheet.setActiveSheet(copiedSheet);
        spreadsheet.moveActiveSheet(getMoveSheetPosition_(spreadsheet, plan.position));
        SpreadsheetApp.flush();
        return {
          kind: 'workbook_structure_update',
          hostPlatform: 'google_sheets',
          sheetName: plan.sheetName,
          operation: plan.operation,
          newSheetName: copiedSheet.getName(),
          positionResolved: copiedSheet.getIndex() - 1,
          sheetCount: spreadsheet.getSheets().length,
          summary: getWorkbookStructureStatusSummary_(plan)
        };
      }
      case 'move_sheet': {
        const sheet = spreadsheet.getSheetByName(plan.sheetName);
        if (!sheet) {
          throw new Error('Target sheet not found: ' + plan.sheetName);
        }

        spreadsheet.setActiveSheet(sheet);
        spreadsheet.moveActiveSheet(getMoveSheetPosition_(spreadsheet, plan.position));
        SpreadsheetApp.flush();
        return {
          kind: 'workbook_structure_update',
          hostPlatform: 'google_sheets',
          sheetName: plan.sheetName,
          operation: plan.operation,
          positionResolved: sheet.getIndex() - 1,
          sheetCount: spreadsheet.getSheets().length,
          summary: getWorkbookStructureStatusSummary_(plan)
        };
      }
      case 'hide_sheet': {
        const sheet = spreadsheet.getSheetByName(plan.sheetName);
        if (!sheet) {
          throw new Error('Target sheet not found: ' + plan.sheetName);
        }

        const visibleSheets = spreadsheet.getSheets().filter(function(candidate) {
          return !candidate.isSheetHidden();
        });
        if (visibleSheets.length <= 1 && !sheet.isSheetHidden()) {
          throw new Error('Cannot hide the only visible worksheet.');
        }

        sheet.hideSheet();
        SpreadsheetApp.flush();
        return {
          kind: 'workbook_structure_update',
          hostPlatform: 'google_sheets',
          sheetName: plan.sheetName,
          operation: plan.operation,
          summary: getWorkbookStructureStatusSummary_(plan)
        };
      }
      case 'unhide_sheet': {
        const sheet = spreadsheet.getSheetByName(plan.sheetName);
        if (!sheet) {
          throw new Error('Target sheet not found: ' + plan.sheetName);
        }

        sheet.showSheet();
        SpreadsheetApp.flush();
        return {
          kind: 'workbook_structure_update',
          hostPlatform: 'google_sheets',
          sheetName: plan.sheetName,
          operation: plan.operation,
          summary: getWorkbookStructureStatusSummary_(plan)
        };
      }
      default:
        throw new Error('Unsupported workbook structure update.');
    }
  }

  if (isNamedRangeUpdatePlan_(plan)) {
    if (plan.scope !== 'workbook') {
      throw new Error('Google Sheets host does not support sheet-scoped named ranges.');
    }

    if (plan.operation === 'create' || plan.operation === 'retarget') {
      if (!plan.targetSheet || !plan.targetRange) {
        throw new Error('Named range create and retarget require targetSheet and targetRange.');
      }
    }

    switch (plan.operation) {
      case 'create': {
        const namedRangeSheet = spreadsheet.getSheetByName(plan.targetSheet);
        if (!namedRangeSheet) {
          throw new Error('Target sheet not found: ' + plan.targetSheet);
        }

        if (typeof spreadsheet.setNamedRange !== 'function') {
          throw new Error('Google Sheets host does not support creating named ranges.');
        }

        spreadsheet.setNamedRange(plan.name, namedRangeSheet.getRange(plan.targetRange));
        SpreadsheetApp.flush();
        return {
          kind: 'named_range_update',
          hostPlatform: 'google_sheets',
          ...plan,
          summary: getNamedRangeStatusSummary_(plan)
        };
      }
      case 'rename': {
        const namedRange = findNamedRange_(spreadsheet, plan.name);
        if (!namedRange) {
          throw new Error('Named range not found: ' + plan.name);
        }

        if (typeof namedRange.setName !== 'function') {
          throw new Error('Google Sheets host does not support renaming named ranges.');
        }

        namedRange.setName(plan.newName);
        SpreadsheetApp.flush();
        return {
          kind: 'named_range_update',
          hostPlatform: 'google_sheets',
          ...plan,
          summary: getNamedRangeStatusSummary_(plan)
        };
      }
      case 'delete':
        if (typeof spreadsheet.removeNamedRange !== 'function') {
          throw new Error('Google Sheets host does not support deleting named ranges.');
        }

        spreadsheet.removeNamedRange(plan.name);
        SpreadsheetApp.flush();
        return {
          kind: 'named_range_update',
          hostPlatform: 'google_sheets',
          ...plan,
          summary: getNamedRangeStatusSummary_(plan)
        };
      case 'retarget': {
        const namedRange = findNamedRange_(spreadsheet, plan.name);
        const namedRangeSheet = spreadsheet.getSheetByName(plan.targetSheet);
        if (!namedRange) {
          throw new Error('Named range not found: ' + plan.name);
        }

        if (!namedRangeSheet) {
          throw new Error('Target sheet not found: ' + plan.targetSheet);
        }

        if (typeof namedRange.setRange !== 'function') {
          throw new Error('Google Sheets host does not support retargeting named ranges.');
        }

        namedRange.setRange(namedRangeSheet.getRange(plan.targetRange));
        SpreadsheetApp.flush();
        return {
          kind: 'named_range_update',
          hostPlatform: 'google_sheets',
          ...plan,
          summary: getNamedRangeStatusSummary_(plan)
        };
      }
      default:
        throw new Error('Unsupported named range update.');
    }
  }

  if (isRangeTransferPlan_(plan)) {
    const sourceSheet = spreadsheet.getSheetByName(plan.sourceSheet);
    const targetSheet = spreadsheet.getSheetByName(plan.targetSheet);
    if (!sourceSheet) {
      throw new Error('Source sheet not found: ' + plan.sourceSheet);
    }

    if (!targetSheet) {
      throw new Error('Target sheet not found: ' + plan.targetSheet);
    }

    const sourceRange = sourceSheet.getRange(plan.sourceRange);
    const targetAnchor = targetSheet.getRange(plan.targetRange);

    if (plan.operation === 'append') {
      const appendTarget = resolveAppendTransferTarget_(targetSheet, targetAnchor, sourceRange, plan);
      assertNonOverlappingTransfer_(plan, appendTarget.actualTargetRange);

      if (plan.pasteMode === 'values') {
        appendTarget.writeRange.setValues(appendTarget.insertedMatrix);
      } else {
        writeTransferToTarget_(appendTarget.writeRange, sourceRange, plan);
      }

      SpreadsheetApp.flush();
      return buildRangeTransferResult_(plan, appendTarget.actualTargetRange);
    }

    const transferShape = getResolvedTransferShape_(sourceRange, plan);
    const resolvedTargetRange = resolveTransferTargetRange_(
      targetAnchor,
      transferShape.rows,
      transferShape.columns
    );
    const actualTargetRange = deriveTransferTargetRangeA1_(plan, resolvedTargetRange);

    assertNonOverlappingTransfer_(plan, actualTargetRange);
    writeTransferToTarget_(resolvedTargetRange, sourceRange, plan);

    if (plan.operation === 'move') {
      clearTransferredSource_(sourceRange, plan);
    }

    SpreadsheetApp.flush();
    return buildRangeTransferResult_(plan, actualTargetRange);
  }

  const sheet = spreadsheet.getSheetByName(plan.targetSheet);
  if (!sheet) {
    throw new Error('Target sheet not found: ' + plan.targetSheet);
  }

  if (isSheetStructurePlan_(plan)) {
    switch (plan.operation) {
      case 'insert_rows':
        sheet.insertRowsBefore(plan.startIndex + 1, plan.count);
        break;
      case 'delete_rows':
        sheet.deleteRows(plan.startIndex + 1, plan.count);
        break;
      case 'hide_rows':
        sheet.hideRows(plan.startIndex + 1, plan.count);
        break;
      case 'unhide_rows':
        sheet.showRows(plan.startIndex + 1, plan.count);
        break;
      case 'group_rows':
        getRowSliceRange_(sheet, plan.startIndex, plan.count).shiftRowGroupDepth(1);
        break;
      case 'ungroup_rows':
        getRowSliceRange_(sheet, plan.startIndex, plan.count).shiftRowGroupDepth(-1);
        break;
      case 'insert_columns':
        sheet.insertColumnsBefore(plan.startIndex + 1, plan.count);
        break;
      case 'delete_columns':
        sheet.deleteColumns(plan.startIndex + 1, plan.count);
        break;
      case 'hide_columns':
        sheet.hideColumns(plan.startIndex + 1, plan.count);
        break;
      case 'unhide_columns':
        sheet.showColumns(plan.startIndex + 1, plan.count);
        break;
      case 'group_columns':
        getColumnSliceRange_(sheet, plan.startIndex, plan.count).shiftColumnGroupDepth(1);
        break;
      case 'ungroup_columns':
        getColumnSliceRange_(sheet, plan.startIndex, plan.count).shiftColumnGroupDepth(-1);
        break;
      case 'merge_cells':
        sheet.getRange(plan.targetRange).merge();
        break;
      case 'unmerge_cells':
        sheet.getRange(plan.targetRange).breakApart();
        break;
      case 'freeze_panes':
      case 'unfreeze_panes':
        sheet.setFrozenRows(plan.frozenRows);
        sheet.setFrozenColumns(plan.frozenColumns);
        break;
      case 'autofit_rows': {
        const rowTarget = sheet.getRange(plan.targetRange);
        sheet.autoResizeRows(rowTarget.getRow(), rowTarget.getNumRows());
        break;
      }
      case 'autofit_columns': {
        const columnTarget = sheet.getRange(plan.targetRange);
        sheet.autoResizeColumns(columnTarget.getColumn(), columnTarget.getNumColumns());
        break;
      }
      case 'set_sheet_tab_color':
        sheet.setTabColor(plan.color);
        break;
      default:
        throw new Error('Unsupported sheet structure update.');
    }

    SpreadsheetApp.flush();
    const result = {
      kind: 'sheet_structure_update',
      hostPlatform: 'google_sheets',
      operation: plan.operation,
      targetSheet: plan.targetSheet,
      summary: getSheetStructureStatusSummary_(plan)
    };

    switch (plan.operation) {
      case 'insert_rows':
      case 'delete_rows':
      case 'hide_rows':
      case 'unhide_rows':
      case 'group_rows':
      case 'ungroup_rows':
      case 'insert_columns':
      case 'delete_columns':
      case 'hide_columns':
      case 'unhide_columns':
      case 'group_columns':
      case 'ungroup_columns':
        result.startIndex = plan.startIndex;
        result.count = plan.count;
        break;
      case 'merge_cells':
      case 'unmerge_cells':
      case 'autofit_rows':
      case 'autofit_columns':
        result.targetRange = plan.targetRange;
        break;
      case 'freeze_panes':
      case 'unfreeze_panes':
        result.frozenRows = plan.frozenRows;
        result.frozenColumns = plan.frozenColumns;
        break;
      case 'set_sheet_tab_color':
        result.color = plan.color;
        break;
    }

    return result;
  }

  const target = sheet.getRange(plan.targetRange);

  if (isExternalDataPlan_(plan)) {
    if (target.getNumRows() !== 1 || target.getNumColumns() !== 1) {
      throw new Error('Google Sheets host requires a single-cell target anchor for external data formulas.');
    }

    target.setFormula(plan.formula);
    SpreadsheetApp.flush();
    return {
      kind: 'external_data_update',
      hostPlatform: 'google_sheets',
      sourceType: plan.sourceType,
      provider: plan.provider,
      query: plan.query,
      sourceUrl: plan.sourceUrl,
      selectorType: plan.selectorType,
      selector: plan.selector,
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      formula: plan.formula,
      explanation: plan.explanation,
      confidence: plan.confidence,
      requiresConfirmation: plan.requiresConfirmation,
      affectedRanges: plan.affectedRanges,
      overwriteRisk: plan.overwriteRisk,
      confirmationLevel: plan.confirmationLevel,
      summary: getExternalDataStatusSummary_(plan)
    };
  }

  if (isConditionalFormatPlan_(plan)) {
    validateConditionalFormatPlanSupport_(plan);

    const preservedRules = (
      plan.managementMode === 'replace_all_on_target' || plan.managementMode === 'clear_on_target'
    )
      ? partitionConditionalFormatRules_(sheet, target)
      : (typeof sheet.getConditionalFormatRules === 'function' ? (sheet.getConditionalFormatRules() || []) : []);

    if (typeof sheet.setConditionalFormatRules !== 'function') {
      throw new Error('Google Sheets host does not support updating conditional formatting on this sheet.');
    }

    if (plan.managementMode === 'clear_on_target') {
      sheet.setConditionalFormatRules(preservedRules);
    } else {
      const newRule = buildConditionalFormatRule_(target, plan);
      sheet.setConditionalFormatRules(preservedRules.concat([newRule]));
    }

    SpreadsheetApp.flush();
    return {
      kind: 'conditional_format_update',
      hostPlatform: 'google_sheets',
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      explanation: plan.explanation,
      confidence: plan.confidence,
      requiresConfirmation: plan.requiresConfirmation,
      affectedRanges: plan.affectedRanges,
      replacesExistingRules: plan.replacesExistingRules,
      managementMode: plan.managementMode,
      ruleType: plan.ruleType,
      comparator: plan.comparator,
      value: plan.value,
      value2: plan.value2,
      text: plan.text,
      formula: plan.formula,
      rank: plan.rank,
      direction: plan.direction,
      points: plan.points,
      style: plan.style,
      summary: getConditionalFormatStatusSummary_(plan)
    };
  }

  if (isDataValidationPlan_(plan)) {
    target.setDataValidation(buildDataValidationRule_(spreadsheet, sheet, target, plan));
    SpreadsheetApp.flush();
    return {
      kind: 'data_validation_update',
      hostPlatform: 'google_sheets',
      operation: plan.operation,
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      allowBlank: plan.allowBlank,
      invalidDataBehavior: plan.invalidDataBehavior,
      helpText: plan.helpText,
      explanation: plan.explanation,
      confidence: plan.confidence,
      requiresConfirmation: plan.requiresConfirmation,
      affectedRanges: plan.affectedRanges,
      replacesExistingValidation: plan.replacesExistingValidation,
      ruleType: plan.ruleType,
      values: plan.values,
      sourceRange: plan.sourceRange,
      namedRangeName: plan.namedRangeName,
      showDropdown: plan.showDropdown,
      checkedValue: plan.checkedValue,
      uncheckedValue: plan.uncheckedValue,
      comparator: plan.comparator,
      value: plan.value,
      value2: plan.value2,
      formula: plan.formula,
      summary: getDataValidationStatusSummary_(plan)
    };
  }

  if (isRangeSortPlan_(plan)) {
    const beforeValues = target.getValues();
    const beforeFormulas = target.getFormulas();
    const sortSpecs = buildGoogleSheetsSortSpecs_(plan, target);
    if (sortSpecs.length === 0) {
      throw new Error('Google Sheets host could not resolve any valid sort keys for this range.');
    }

    const headerOffset = plan.hasHeader ? 1 : 0;
    const sortableRowCount = target.getNumRows() - headerOffset;
    if (sortableRowCount > 0) {
      sheet.getRange(
        target.getRow() + headerOffset,
        target.getColumn(),
        sortableRowCount,
        target.getNumColumns()
      ).sort(sortSpecs);
    }

    SpreadsheetApp.flush();
    return attachLocalExecutionSnapshot_({
      kind: 'range_sort',
      hostPlatform: 'google_sheets',
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      hasHeader: plan.hasHeader,
      keys: plan.keys,
      explanation: plan.explanation,
      confidence: plan.confidence,
      requiresConfirmation: plan.requiresConfirmation,
      affectedRanges: plan.affectedRanges,
      summary: getRangeSortStatusSummary_(plan)
    }, createLocalExecutionSnapshot_({
      executionId: input.executionId,
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      beforeValues: beforeValues,
      beforeFormulas: beforeFormulas,
      afterValues: target.getValues(),
      afterFormulas: target.getFormulas()
    }));
  }

  if (isRangeFilterPlan_(plan)) {
    if (!plan.hasHeader) {
      throw new Error('Google Sheets host requires hasHeader=true for filter plans.');
    }

    if (plan.combiner === 'or' && plan.conditions.length > 1) {
      throw new Error('Google Sheets grid filters cannot represent combiner "or" exactly for multiple conditions.');
    }

    if (plan.combiner !== 'and' && plan.combiner !== 'or') {
      throw new Error('Unsupported filter combiner: ' + plan.combiner);
    }

    const resolvedConditions = plan.conditions.map(function(condition) {
      return {
        columnPosition: resolveRelativeColumnRef_(condition.columnRef, target, plan.hasHeader),
        criteria: buildFilterCriteria_(condition)
      };
    });
    const seenColumns = {};

    resolvedConditions.forEach(function(condition) {
      if (!condition.columnPosition) {
        throw new Error('Google Sheets host could not resolve a filter column inside the target range.');
      }

      if (seenColumns[condition.columnPosition]) {
        throw new Error('Google Sheets host does not support multiple filter conditions for the same column.');
      }

      seenColumns[condition.columnPosition] = true;
    });

    const filter = getOrCreateFilter_(sheet, target, plan.clearExistingFilters);
    resolvedConditions.forEach(function(condition) {
      filter.setColumnFilterCriteria(condition.columnPosition, condition.criteria);
    });

    SpreadsheetApp.flush();
    return {
      kind: 'range_filter',
      hostPlatform: 'google_sheets',
      ...plan,
      summary: getRangeFilterStatusSummary_(plan)
    };
  }

  if (isRangeFormatPlan_(plan)) {
    if (plan.format.backgroundColor) {
      target.setBackground(plan.format.backgroundColor);
    }

    if (plan.format.textColor) {
      target.setFontColor(plan.format.textColor);
    }

    if (typeof plan.format.bold === 'boolean') {
      target.setFontWeight(plan.format.bold ? 'bold' : 'normal');
    }

    if (typeof plan.format.italic === 'boolean') {
      target.setFontStyle(plan.format.italic ? 'italic' : 'normal');
    }

    if (plan.format.horizontalAlignment) {
      target.setHorizontalAlignment(
        plan.format.horizontalAlignment === 'general'
          ? null
          : plan.format.horizontalAlignment
      );
    }

    if (plan.format.verticalAlignment) {
      target.setVerticalAlignment(plan.format.verticalAlignment);
    }

    if (plan.format.wrapStrategy) {
      const wrapStrategy = getWrapStrategyEnum_(plan.format.wrapStrategy);
      if (wrapStrategy) {
        target.setWrapStrategy(wrapStrategy);
      }
    }

    if (plan.format.numberFormat) {
      target.setNumberFormat(plan.format.numberFormat);
    }

    if (typeof plan.format.columnWidth === 'number') {
      sheet.setColumnWidths(target.getColumn(), target.getNumColumns(), plan.format.columnWidth);
    }

    if (typeof plan.format.rowHeight === 'number') {
      sheet.setRowHeights(target.getRow(), target.getNumRows(), plan.format.rowHeight);
    }

    SpreadsheetApp.flush();
    return {
      kind: 'range_format_update',
      hostPlatform: 'google_sheets',
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      format: plan.format,
      explanation: plan.explanation,
      confidence: plan.confidence,
      requiresConfirmation: plan.requiresConfirmation,
      overwriteRisk: plan.overwriteRisk,
      summary: getRangeFormatStatusSummary_(plan)
    };
  }

  if (isDataCleanupPlan_(plan)) {
    const beforeValues = target.getValues();
    const beforeFormulas = target.getFormulas();
    const nextValues = buildCleanupWriteMatrix_(plan, beforeValues, beforeFormulas, 'Google Sheets host');
    target.setValues(nextValues);
    SpreadsheetApp.flush();
    return attachLocalExecutionSnapshot_({
      kind: 'data_cleanup_update',
      hostPlatform: 'google_sheets',
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      operation: plan.operation,
      keyColumns: plan.keyColumns,
      mode: plan.mode,
      sourceColumn: plan.sourceColumn,
      targetColumns: plan.targetColumns,
      delimiter: plan.delimiter,
      sourceColumns: plan.sourceColumns,
      destinationColumn: plan.destinationColumn,
      formatType: plan.formatType,
      formatPattern: plan.formatPattern,
      explanation: plan.explanation,
      confidence: plan.confidence,
      requiresConfirmation: plan.requiresConfirmation,
      affectedRanges: plan.affectedRanges,
      overwriteRisk: plan.overwriteRisk,
      confirmationLevel: plan.confirmationLevel,
      summary: getDataCleanupStatusSummary_(plan)
    }, createLocalExecutionSnapshot_({
      executionId: input.executionId,
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      beforeValues: beforeValues,
      beforeFormulas: beforeFormulas,
      afterValues: target.getValues(),
      afterFormulas: target.getFormulas()
    }));
  }

  if (target.getNumRows() !== plan.shape.rows || target.getNumColumns() !== plan.shape.columns) {
    throw new Error('The approved targetRange does not match the proposed shape.');
  }

  if (Array.isArray(plan.headers)) {
    if (rangeHasExistingContent_(target.getDisplayValues()) || hasAnyRealFormula_(target.getFormulas())) {
      throw new Error('Target range already contains content. Clear it before confirming the import plan.');
    }

    const beforeValues = target.getValues();
    const beforeFormulas = target.getFormulas();
    target.setValues([plan.headers].concat(plan.values));
    SpreadsheetApp.flush();
    return attachLocalExecutionSnapshot_({
      kind: 'range_write',
      hostPlatform: 'google_sheets',
      ...plan,
      writtenRows: plan.shape.rows,
      writtenColumns: plan.shape.columns
    }, createLocalExecutionSnapshot_({
      executionId: input.executionId,
      targetSheet: plan.targetSheet,
      targetRange: plan.targetRange,
      beforeValues: beforeValues,
      beforeFormulas: beforeFormulas,
      afterValues: target.getValues(),
      afterFormulas: target.getFormulas()
    }));
  }

  const beforeValues = target.getValues();
  const beforeFormulas = target.getFormulas();
  if (plan.values && !plan.formulas && !plan.notes) {
    target.setValues(plan.values);
    SpreadsheetApp.flush();
  } else if (plan.formulas && !plan.values && !plan.notes) {
    target.setFormulas(plan.formulas.map(function(row) {
      return row.map(function(cell) {
        return cell || '';
      });
    }));
    SpreadsheetApp.flush();
  } else if (plan.notes && !plan.values && !plan.formulas) {
    target.setNotes(plan.notes.map(function(row) {
      return row.map(function(cell) {
        return cell == null ? '' : String(cell);
      });
    }));
    SpreadsheetApp.flush();
  } else {
    for (let rowIndex = 0; rowIndex < plan.shape.rows; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < plan.shape.columns; columnIndex += 1) {
        const cell = target.getCell(rowIndex + 1, columnIndex + 1);
        const noteValue = plan.notes && plan.notes[rowIndex] ? plan.notes[rowIndex][columnIndex] : null;
        const formulaValue = plan.formulas && plan.formulas[rowIndex] ? plan.formulas[rowIndex][columnIndex] : null;
        const rawValue = plan.values && plan.values[rowIndex] ? plan.values[rowIndex][columnIndex] : null;

        if (typeof formulaValue === 'string' && formulaValue.trim().length > 0) {
          cell.setFormula(formulaValue);
        } else if (rawValue !== null && rawValue !== undefined) {
          cell.setValue(rawValue);
        } else if (formulaValue === null || formulaValue === '') {
          cell.setValue('');
        }

        if (noteValue !== null && noteValue !== undefined && noteValue !== '') {
          cell.setNote(String(noteValue));
        }
      }
    }
    SpreadsheetApp.flush();
  }

  return attachLocalExecutionSnapshot_({
    kind: 'range_write',
    hostPlatform: 'google_sheets',
    ...plan,
    writtenRows: plan.shape.rows,
    writtenColumns: plan.shape.columns
  }, createLocalExecutionSnapshot_({
    executionId: input.executionId,
    targetSheet: plan.targetSheet,
    targetRange: plan.targetRange,
    beforeValues: beforeValues,
    beforeFormulas: beforeFormulas,
    afterValues: target.getValues(),
    afterFormulas: target.getFormulas()
  }));
}

if (typeof module !== 'undefined') {
  module.exports = {
    applyWritePlan,
    applyExecutionCellSnapshot,
    validateExecutionCellSnapshot,
    buildFilterCriteria_,
    buildDataValidationRule_,
    buildGoogleSheetsSortSpecs_,
    findNamedRange_,
    getConditionalFormatStatusSummary_,
    getOrCreateFilter_,
    isConditionalFormatPlan_,
    isSupportedConditionalFormatManagementMode_,
    isDataValidationPlan_,
    isNamedRangeUpdatePlan_,
    validateConditionalFormatPlanSupport_,
    resolveRelativeColumnRef_
  };
}
