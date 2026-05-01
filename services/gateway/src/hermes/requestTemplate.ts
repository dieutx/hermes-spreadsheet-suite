import type { HermesRequest } from "@hermes/contracts";

const STRUCTURED_BODY_TYPES = [
  "chat",
  "formula",
  "composite_plan",
  "workbook_structure_update",
  "range_format_update",
  "conditional_format_plan",
  "sheet_structure_update",
  "range_sort_plan",
  "range_filter_plan",
  "data_validation_plan",
  "analysis_report_plan",
  "pivot_table_plan",
  "chart_plan",
  "table_plan",
  "named_range_update",
  "range_transfer_plan",
  "data_cleanup_plan",
  "analysis_report_update",
  "pivot_table_update",
  "chart_update",
  "table_update",
  "sheet_update",
  "sheet_import_plan",
  "external_data_plan",
  "error",
  "attachment_analysis",
  "extracted_table",
  "document_summary"
] as const;

const REQUIRED_STRUCTURED_BODY_FIELDS = [
  "type",
  "data"
] as const;

export type PreferredResponseType =
  | "chat"
  | "formula"
  | "error"
  | "composite_plan"
  | "sheet_update"
  | "workbook_structure_update"
  | "range_format_update"
  | "sheet_structure_update"
  | "range_sort_plan"
  | "range_filter_plan"
  | "data_validation_plan"
  | "analysis_report_plan"
  | "pivot_table_plan"
  | "chart_plan"
  | "table_plan"
  | "named_range_update"
  | "conditional_format_plan"
  | "range_transfer_plan"
  | "data_cleanup_plan"
  | "external_data_plan";

export type SpreadsheetRoutingHints = {
  preferredResponseType: PreferredResponseType;
  generatedDataRequest: boolean;
  explicitWriteIntent: boolean;
  mixedAdvisoryAndWriteRequest: boolean;
  toolScaffoldingOpportunity: boolean;
  inputLayoutConflictRisk: boolean;
};

const A1_REFERENCE_PATTERN = /(?:^|[^a-z])\$?[a-z]{1,3}\$?\d+(?:\:\$?[a-z]{1,3}\$?\d+)?(?:[^a-z]|$)/i;
const VALIDATION_KEYWORD_PATTERN = /\b(dropdown|checkbox|validation|data validation|allow only|only allow|reject invalid|pick list|list from range|whole number|whole numbers|integer|integers|decimal|decimals|date|dates|text length|custom formula|chi cho phep|hop le|xac thuc)\b/;
const CONDITIONAL_FORMAT_KEYWORD_PATTERN = /\b(conditional formatting|conditional format|highlight|highlights|highlighting|mark duplicates|duplicate values|duplicates|color scale|3-color scale|2-color scale|to mau|danh dau trung lap)\b/;
const RANGE_TRANSFER_KEYWORD_PATTERN = /\b(copy|move|append|transpose|transposed|sao chep|di chuyen|them vao cuoi)\b/;
const DATA_CLEANUP_KEYWORD_PATTERN = /\b(cleanup|clean up|reshape|trim whitespace|remove blank rows|remove duplicate rows|normalize case|split column|join columns|fill down|standardize|lam sach|xoa dong trong|xoa trung lap|tach cot|gop cot|dien xuong|chuan hoa)\b/;
const ANALYSIS_REPORT_KEYWORD_PATTERN = /\b(analy[sz]e|analysis|report|insights|findings|anomal(?:y|ies)|trend|trends|phan tich|bao cao|bat thuong|xu huong)\b/;
const PIVOT_TABLE_KEYWORD_PATTERN = /\b(pivot table|pivot|bang tong hop)\b/;
const CHART_KEYWORD_PATTERN = /\b(chart|graph|plot|bieu do|do thi)\b/;
const TABLE_PLAN_KEYWORD_PATTERN = /\b(format as table|format .* as a table|convert .* to (?:a )?table|make .* (?:a )?table|create (?:a )?table from|filterable table|banded rows?|native table|excel table|structured table)\b/;
const STATIC_RANGE_FORMAT_KEYWORD_PATTERN = /\b(format|formatting|bold|italic|underline|strikethrough|font|font size|fill|background|text color|font color|number format|currency|percent|percentage|decimal places?|wrap|wrapped|center|align|alignment|vertical align|border|borders|row height|column width|shade|color)\b/;
const STATIC_RANGE_FORMAT_ACTION_PATTERN = /\b(make|set|apply|format|change|turn|wrap|center|align|color|shade|bold|italic|underline|strikethrough|resize)\b/;
const STATIC_RANGE_FORMAT_TARGET_PATTERN = /\b(selected|selection|range|ranges|cell|cells|row|rows|column|columns|header|headers|table|this|current|vung|o|dong|cot)\b/;
const EXTERNAL_DATA_PROVIDER_PATTERN = /\b(googlefinance|importhtml|importxml|importdata)\b/;
const MARKET_DATA_KEYWORD_PATTERN = /\b(stock|stocks|ticker|tickers|quote|quotes|share price|share prices|market data|price history|crypto|coin|coins|googlefinance|chung khoan|co phieu|gia co phieu|gia crypto|btc|eth)\b/;
const WEB_IMPORT_KEYWORD_PATTERN = /\b(website|web page|web table|html table|table from a website|public url|url|importhtml|importxml|importdata|scrape table|lay bang tu website|lay du lieu tu website)\b/;
const EXTERNAL_DATA_ACTION_PATTERN = /\b(fetch|get|pull|import|insert|write|put|fill|load|show|latest|current|history|lay|nhap|chen|dien)\b/;
const SORT_KEYWORD_PATTERN = /\b(sort|sap xep)\b/;
const FILTER_KEYWORD_PATTERN = /\b(filter|loc)\b/;
const COMPOSITE_CONNECTOR_PATTERN = /\b(and then|then|after that|afterwards|next|followed by|roi|sau do|tiep theo)\b/;
const FORMULA_REFERENCE_PATTERN = /\b(formula|cong thuc)\b/;
const FORMULA_APPLY_PATTERN = /(fix|apply|set|sua|ap dung|dat).*(formula|cell|cong thuc)|(?:formula|cong thuc).*(fix|apply|set|sua|ap dung|dat)/;
const FORMULA_DEBUG_KEYWORD_PATTERN = /\b(debug|diagnos(?:e|ing|is)?|broken|wrong|fix|explain|translate|why|sua|giai thich|tai sao|loi|sai|cong thuc)\b/;
const FORMULA_ERROR_TOKEN_PATTERN = /#(REF!|N\/A|VALUE!|NAME\?|DIV\/0!|NUM!|NULL!|SPILL!|CALC!)/i;
const GENERATED_DATA_KEYWORD_PATTERN = /\b(random|sample|mock|dummy|seed(?:ed)?|populate|fill|generate|ngau nhien|du lieu mau|gia lap|dien)\b/;
const GENERATED_DATA_DOMAIN_PATTERN = /\b(data|dataset|rows?|records?|table|sales|revenue|orders?|customers?|inventory|du lieu|dong|ban ghi|doanh thu|don hang|khach hang|ton kho|ban hang)\b/;
const TOTAL_ROW_KEYWORD_PATTERN = /\b(total row|totals row|grand total|subtotal|totals|hang tong|dong tong|tong cong)\b/;
const TOTAL_ROW_VERB_PATTERN = /\b(add|insert|append|create|them|chen|tao)\b/;
const SELECTION_EXPLANATION_KEYWORD_PATTERN = /\b(explain|describe|summari[sz]e|interpret|walk me through|giai thich|mo ta|tom tat)\b/;
const SELECTION_REFERENCE_PATTERN = /\b(selection|selected|current|this|range|table|data|sheet|formula|sheet nay|range nay|du lieu nay|vung nay)\b/;
const LOCATION_REFERENCE_PATTERN = /\b(sheet|tab|worksheet|range|cell|selection|table|this|current|here|sheet nay|range nay|cell nay|o nay|vung nay|du lieu nay|hien tai)\b/;
const IMPLICIT_TARGET_PATTERN = /\b(table|range|this|it|selection|current|here|sheet nay|range nay|du lieu nay|vung nay)\b/;
const CURRENT_CELL_TARGET_PATTERN = /\b(?:current|active|selected|this)\s+cell\b/;
const MATERIALIZE_VERB_PATTERN = /\b(write|put|place|insert|save|materiali[sz]e|output|export|ghi|dat|dua|xuat|tao)\b/;
const MATERIALIZE_TARGET_PATTERN = /\b(new sheet|sheet|tab|worksheet|range|cell|summary sheet|report sheet|sheet moi)\b/;
const TOOL_FLOW_INTENT_PATTERN = /\b(build|create|make|set up|setup|scaffold|tao|dung|lap)\b/;
const TOOL_FLOW_ENTITY_PATTERN = /\b(tool|tracker|helper sheet|input sheet|output sheet|control sheet|parameter sheet|lookup tool|search tool|calculator|template)\b/;
const LOOKUP_FLOW_PATTERN = /\b(lookup|xlookup|vlookup|index match|index\/match|search)\b/;
const COLUMN_FILL_KEYWORD_PATTERN = /\b(fill(?:\s+in)?|populate|complete|set|update)\b/;
const COLUMN_FILL_CONTEXT_PATTERN = /\b(column|field|lookup table|based on|using|from|character|prefix|item code|item name)\b/;
const TOOL_FLOW_LAYOUT_PATTERN = /\b(input|inputs|output|outputs|result|results|criteria|parameter|parameters|control cell|control cells|helper sheet)\b/;
const CONTROL_CELL_ROLE_PATTERN = /\b(?:with|using|use|treat)\b[\s\S]{0,120}?\bcell\s+\$?[A-Za-z]{1,3}\$?\d+\b/i;
const CONTROL_CELL_AS_PATTERN = /\b(?:use|using|treat)\s+\$?[A-Za-z]{1,3}\$?\d+\s+as\b/i;
const ADVISORY_ARTIFACT_QUESTION_PATTERN = /\b(explain|describe|what is|what are|how do i|how can i|how should i|how to|why|diagnos(?:e|ing|is)?|troubleshoot|help me understand|giai thich|tai sao|huong dan)\b/;
const DIRECT_ARTIFACT_ACTION_PATTERN = /^\s*(?:please\s+)?(?:create|make|build|insert|add|apply|set up|setup|tao|them|chen)\b/;
const WHOLE_MESSAGE_HOW_TO_PATTERN = /^\s*(?:please\s+)?(?:explain\s+how\s+to|describe\s+how\s+to|help me understand\s+how\s+to|how\s+(?:do|can|should)\s+i|how\s+to|what\s+(?:is|are)|huong dan)\b/;
const WRITE_ACTION_AFTER_CONNECTOR_PATTERN = /\b(?:and then|then|after that|afterwards|next|followed by|roi|sau do|tiep theo)\b[\s\S]{0,160}?\b(?:sort|filter|add|create|make|insert|apply|set|copy|move|append|transpose|remove|delete|fill|standardize|trim|split|join|rename|retarget|freeze|unfreeze|autofit|group|ungroup|merge|unmerge|hide|unhide|format|highlight|chart|pivot|import|get|pull|fetch|sap xep|loc|them|tao|chen|ap dung|xoa|doi ten)\b/;
const WRITE_CAPABLE_RESPONSE_TYPES = new Set<PreferredResponseType>([
  "composite_plan",
  "sheet_update",
  "workbook_structure_update",
  "range_format_update",
  "sheet_structure_update",
  "range_sort_plan",
  "range_filter_plan",
  "data_validation_plan",
  "pivot_table_plan",
  "chart_plan",
  "table_plan",
  "named_range_update",
  "conditional_format_plan",
  "range_transfer_plan",
  "data_cleanup_plan",
  "external_data_plan"
]);

function normalizeNaturalLanguage(value: string): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[đĐ]/g, "d")
    .toLowerCase();
}

type ParsedCellRef = {
  row: number;
  column: number;
};

type ParsedA1Bounds = {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
};

function stripSheetPrefix(value: string): string {
  const bangIndex = value.lastIndexOf("!");
  return bangIndex >= 0 ? value.slice(bangIndex + 1) : value;
}

function columnLettersToNumber(value: string): number {
  let total = 0;
  for (const char of value.toUpperCase()) {
    total = total * 26 + (char.charCodeAt(0) - 64);
  }
  return total;
}

function parseSingleCellReference(value: string): ParsedCellRef | null {
  const match = stripSheetPrefix(value).replace(/\$/g, "").match(/^([A-Za-z]{1,3})(\d+)$/);
  if (!match) {
    return null;
  }

  return {
    column: columnLettersToNumber(match[1]),
    row: Number.parseInt(match[2], 10)
  };
}

function parseA1Bounds(value: string | undefined): ParsedA1Bounds | null {
  if (typeof value !== "string" || value.trim().length === 0) {
    return null;
  }

  const normalized = stripSheetPrefix(value).replace(/\$/g, "");
  const [startToken, endToken = normalized] = normalized.split(":");
  const start = parseSingleCellReference(startToken);
  const end = parseSingleCellReference(endToken);

  if (!start || !end) {
    return null;
  }

  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column)
  };
}

function rangeContainsCellReference(range: string | undefined, cellRef: string): boolean {
  const bounds = parseA1Bounds(range);
  const cell = parseSingleCellReference(cellRef);
  if (!bounds || !cell) {
    return false;
  }

  return cell.row >= bounds.startRow &&
    cell.row <= bounds.endRow &&
    cell.column >= bounds.startColumn &&
    cell.column <= bounds.endColumn;
}

function hasTabularContext(selection: HermesRequest["context"]["selection"] | undefined): boolean {
  if (!selection) {
    return false;
  }

  if (Array.isArray(selection.headers) && selection.headers.length >= 2) {
    return true;
  }

  if (Array.isArray(selection.values) && selection.values.length > 0 && selection.values[0]?.length >= 2) {
    return true;
  }

  if (Array.isArray(selection.formulas) && selection.formulas.length > 0 && selection.formulas[0]?.length >= 2) {
    return true;
  }

  const bounds = parseA1Bounds(selection.range);
  if (!bounds) {
    return false;
  }

  return (bounds.endColumn - bounds.startColumn + 1) >= 2;
}

function getLikelySourceTableRange(request: HermesRequest): string | undefined {
  if (hasTabularContext(request.context.currentRegion)) {
    return request.context.currentRegion?.range;
  }

  if (hasTabularContext(request.context.selection)) {
    return request.context.selection?.range;
  }

  return undefined;
}

function extractExplicitSingleCellReferences(rawMessage: string): string[] {
  const matches = [...rawMessage.matchAll(/(?:^|[^A-Za-z0-9$])(\$?[A-Za-z]{1,3}\$?\d+)(?!\s*:)(?=$|[^A-Za-z0-9$])/g)];
  return Array.from(new Set(matches.map((match) => match[1].replace(/\$/g, "").toUpperCase())));
}

function hasControlCellRoleLanguage(rawMessage: string, normalizedMessage: string): boolean {
  return CONTROL_CELL_ROLE_PATTERN.test(rawMessage) ||
    CONTROL_CELL_AS_PATTERN.test(rawMessage) ||
    /\b(input|inputs|output|outputs|result|results|criteria|parameter|parameters|lookup|search tool|control cells?)\b/.test(normalizedMessage);
}

function hasInputLayoutConflictRisk(
  request: HermesRequest,
  rawMessage: string,
  normalizedMessage: string
): boolean {
  const explicitCellRefs = extractExplicitSingleCellReferences(rawMessage);
  if (explicitCellRefs.length < 2) {
    return false;
  }

  if (!hasControlCellRoleLanguage(rawMessage, normalizedMessage)) {
    return false;
  }

  const sourceTableRange = getLikelySourceTableRange(request);
  if (!sourceTableRange) {
    return false;
  }

  const overlappingRefs = explicitCellRefs.filter((cellRef) =>
    rangeContainsCellReference(sourceTableRange, cellRef)
  );

  return overlappingRefs.length >= 2;
}

function isLikelyToolScaffoldingRequest(
  request: HermesRequest,
  rawMessage: string,
  normalizedMessage: string
): boolean {
  if (hasInputLayoutConflictRisk(request, rawMessage, normalizedMessage)) {
    return true;
  }

  if (/\bhelper sheet|input sheet|output sheet|control sheet|parameter sheet\b/.test(normalizedMessage)) {
    return true;
  }

  if (!TOOL_FLOW_INTENT_PATTERN.test(normalizedMessage)) {
    return false;
  }

  if (TOOL_FLOW_ENTITY_PATTERN.test(normalizedMessage)) {
    return true;
  }

  return LOOKUP_FLOW_PATTERN.test(normalizedMessage) &&
    TOOL_FLOW_LAYOUT_PATTERN.test(normalizedMessage);
}

function isLikelyWorkbookStructureRequest(userMessage: string): boolean {
  const createSheetDirect = /\b(create|add|insert|tao|them|chen)\s+(?:a\s+new\s+|new\s+|1\s+|mot\s+|moi\s+)?(?:sheet|tab|worksheet)\b/.test(userMessage);
  const createSheetIndirect = /\b(create|add|insert|tao|them|chen)\b/.test(userMessage) &&
    /\b(?:sheet|tab|worksheet)\s+(?:moi|new)\b/.test(userMessage);
  const mutateSheet = /\b(delete|remove|rename|duplicate|copy|move|hide|unhide|xoa|doi ten|nhan ban|sao chep|di chuyen|bo an|hien lai)\b.*\b(sheet|tab|worksheet)\b/.test(userMessage);

  return (
    createSheetDirect ||
    createSheetIndirect ||
    mutateSheet
  );
}

function isLikelyConditionalFormatRequest(userMessage: string): boolean {
  if (CONDITIONAL_FORMAT_KEYWORD_PATTERN.test(userMessage)) {
    return true;
  }

  if (/\boverdue\b/.test(userMessage) && /\b(highlight|mark|flag|format|dates?)\b/.test(userMessage)) {
    return true;
  }

  if (/\b(color|highlight|mark)\b/.test(userMessage) &&
    /\bvalues?\b/.test(userMessage) &&
    /\b(above|below|greater than|less than|equal to|between)\b/.test(userMessage)) {
    return true;
  }

  if (/\bclear\b/.test(userMessage) && /\bconditional formatting\b/.test(userMessage)) {
    return true;
  }

  return false;
}

function isLikelyRangeTransferRequest(userMessage: string): boolean {
  if (!RANGE_TRANSFER_KEYWORD_PATTERN.test(userMessage)) {
    return false;
  }

  if (/\b(copy|move|sao chep|di chuyen)\b/.test(userMessage) && /\b(to|into|onto|vao|sang)\b/.test(userMessage)) {
    return true;
  }

  if (/\bappend\b/.test(userMessage) && /\b(to|into|onto|end of|vao|cuoi)\b/.test(userMessage)) {
    return true;
  }

  if (/\btranspose|transposed\b/.test(userMessage)) {
    return true;
  }

  return false;
}

function isLikelyDataCleanupRequest(userMessage: string): boolean {
  return DATA_CLEANUP_KEYWORD_PATTERN.test(userMessage);
}

function isLikelyAnalysisReportRequest(userMessage: string): boolean {
  if (!ANALYSIS_REPORT_KEYWORD_PATTERN.test(userMessage)) {
    return false;
  }

  return (
    /\b(analy[sz]e|analysis|phan tich)\b/.test(userMessage) ||
    /\b(report|insights|findings|anomal(?:y|ies)|trend|trends|bao cao|bat thuong|xu huong)\b/.test(userMessage)
  );
}

function isLikelyPivotTableRequest(userMessage: string): boolean {
  if (!PIVOT_TABLE_KEYWORD_PATTERN.test(userMessage)) {
    return false;
  }

  return true;
}

function isLikelyChartRequest(userMessage: string): boolean {
  if (!CHART_KEYWORD_PATTERN.test(userMessage)) {
    return false;
  }

  return (
    /\b(line|bar|column|area|pie|scatter)\b/.test(userMessage) ||
    /\b(?:create|make|build|insert|add|show|tao|them|chen|hien)\b/.test(userMessage)
  );
}

function isLikelyTablePlanRequest(userMessage: string): boolean {
  if (TABLE_PLAN_KEYWORD_PATTERN.test(userMessage)) {
    return true;
  }

  return /\b(add|create|enable|turn on|apply|them|tao|bat|ap dung)\b/.test(userMessage) &&
    /\b(table filters?|filter buttons?|banded rows?|totals row|native table|excel table)\b/.test(userMessage);
}

function isLikelyRangeFormatRequest(rawMessage: string, userMessage: string): boolean {
  if (isLikelyConditionalFormatRequest(userMessage) || isLikelyTablePlanRequest(userMessage)) {
    return false;
  }

  if (!STATIC_RANGE_FORMAT_KEYWORD_PATTERN.test(userMessage)) {
    return false;
  }

  if (!STATIC_RANGE_FORMAT_ACTION_PATTERN.test(userMessage)) {
    return false;
  }

  return A1_REFERENCE_PATTERN.test(rawMessage) ||
    STATIC_RANGE_FORMAT_TARGET_PATTERN.test(userMessage);
}

function isLikelySheetStructureRequest(userMessage: string): boolean {
  return /\b(insert|delete|hide|unhide|merge|unmerge|freeze|unfreeze|group|ungroup|autofit)\b/.test(userMessage) ||
    /\b(?:sheet\s+)?tab\s+color\b/.test(userMessage);
}

function isLikelyGeneratedDataRequest(rawMessage: string, lowerMessage: string): boolean {
  if (!GENERATED_DATA_KEYWORD_PATTERN.test(lowerMessage)) {
    return false;
  }

  if (!GENERATED_DATA_DOMAIN_PATTERN.test(lowerMessage)) {
    return false;
  }

  return (
    LOCATION_REFERENCE_PATTERN.test(lowerMessage) ||
    A1_REFERENCE_PATTERN.test(rawMessage)
  );
}

function isLikelyCurrentTableFormulaFillRequest(
  request: HermesRequest,
  normalizedMessage: string
): boolean {
  if (!hasCurrentRegionContext(request)) {
    return false;
  }

  if (!COLUMN_FILL_KEYWORD_PATTERN.test(normalizedMessage)) {
    return false;
  }

  if (!/\bcolumn\b/.test(normalizedMessage)) {
    return false;
  }

  return LOOKUP_FLOW_PATTERN.test(normalizedMessage) ||
    COLUMN_FILL_CONTEXT_PATTERN.test(normalizedMessage);
}

function isLikelyCompositeRequest(
  request: HermesRequest,
  rawMessage: string,
  lowerMessage: string
): boolean {
  if (isLikelyWorkbookStructureRequest(lowerMessage) && isLikelyGeneratedDataRequest(rawMessage, lowerMessage)) {
    return true;
  }

  if (isLikelyTotalsRowRequest(request, rawMessage, lowerMessage)) {
    return true;
  }

  if (!COMPOSITE_CONNECTOR_PATTERN.test(lowerMessage)) {
    return false;
  }

  const stepSignals = [
    isLikelyWorkbookStructureRequest(lowerMessage),
    SORT_KEYWORD_PATTERN.test(lowerMessage) &&
      (IMPLICIT_TARGET_PATTERN.test(lowerMessage) ||
        /(?:^|[^a-z])\$?[a-z]{1,3}\$?\d+(?:\:\$?[a-z]{1,3}\$?\d+)?(?:[^a-z]|$)/i.test(lowerMessage)),
    FILTER_KEYWORD_PATTERN.test(lowerMessage),
    /\b(create|rename|delete|retarget)\b/.test(lowerMessage) && isLikelyNamedRangeRequest(rawMessage, lowerMessage),
    VALIDATION_KEYWORD_PATTERN.test(lowerMessage) &&
      (/\b(add|set|apply|use|restrict|limit|validate|require)\b/.test(lowerMessage) ||
        /\bnamed range\b/.test(lowerMessage) ||
        A1_REFERENCE_PATTERN.test(lowerMessage)),
    isLikelyConditionalFormatRequest(lowerMessage),
    isLikelyRangeTransferRequest(lowerMessage),
    isLikelyDataCleanupRequest(lowerMessage),
    isLikelyTotalsRowRequest(request, rawMessage, lowerMessage),
    isLikelyPivotTableRequest(lowerMessage),
    isLikelyChartRequest(lowerMessage),
    isLikelyTablePlanRequest(lowerMessage),
    isLikelyRangeFormatRequest(rawMessage, lowerMessage),
    isLikelyAnalysisReportRequest(lowerMessage),
    /\b(insert|delete|hide|unhide|merge|unmerge|freeze|unfreeze|group|ungroup|autofit)\b/.test(lowerMessage) ||
      /\b(?:sheet\s+)?tab\s+color\b/.test(lowerMessage),
    !/\bvalidation\b/.test(lowerMessage) &&
      FORMULA_APPLY_PATTERN.test(lowerMessage) &&
      A1_REFERENCE_PATTERN.test(lowerMessage)
  ];

  return stepSignals.filter(Boolean).length >= 2;
}

function looksLikeNamedRangeIdentifier(value: string): boolean {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(value);
}

function isDisallowedNamedRangeToken(value: string): boolean {
  return /^(sheet|tab|worksheet|row|rows|column|columns)\d*$/i.test(value);
}

function isLikelyNamedRangeRequest(rawMessage: string, lowerMessage: string): boolean {
  if (/\bnamed range\b/.test(lowerMessage)) {
    return true;
  }

  const renameMatch = rawMessage.match(/\brename\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s+on\s+[A-Za-z0-9_ ]+?)?\s+to\s+([A-Za-z_][A-Za-z0-9_]*)\b/i);
  if (renameMatch &&
    !isDisallowedNamedRangeToken(renameMatch[1]) &&
    looksLikeNamedRangeIdentifier(renameMatch[1]) &&
    looksLikeNamedRangeIdentifier(renameMatch[2])) {
    return true;
  }

  const retargetMatch = rawMessage.match(/\bretarget\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s+on\s+[A-Za-z0-9_ ]+?)?\s+to\s+\$?[A-Za-z]{1,3}\$?\d+(?:\:\$?[A-Za-z]{1,3}\$?\d+)?\b/i);
  if (retargetMatch &&
    !isDisallowedNamedRangeToken(retargetMatch[1]) &&
    looksLikeNamedRangeIdentifier(retargetMatch[1])) {
    return true;
  }

  const createMatch = rawMessage.match(/\bcreate\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s+on\s+[A-Za-z0-9_ ]+?)?\s+(?:for|to)\s+\$?[A-Za-z]{1,3}\$?\d+(?:\:\$?[A-Za-z]{1,3}\$?\d+)?\b/i);
  if (createMatch &&
    !isDisallowedNamedRangeToken(createMatch[1]) &&
    looksLikeNamedRangeIdentifier(createMatch[1])) {
    return true;
  }

  const deleteMatch = rawMessage.match(/\bdelete\s+([A-Za-z_][A-Za-z0-9_]*)(?:\s+on\s+[A-Za-z0-9_ ]+?)?\b/i);
  if (deleteMatch &&
    !isDisallowedNamedRangeToken(deleteMatch[1]) &&
    looksLikeNamedRangeIdentifier(deleteMatch[1])) {
    return true;
  }

  return false;
}

function isLikelyAdvisoryOnlyArtifactQuestion(rawMessage: string, userMessage: string): boolean {
  if (!ADVISORY_ARTIFACT_QUESTION_PATTERN.test(userMessage)) {
    return false;
  }

  if (DIRECT_ARTIFACT_ACTION_PATTERN.test(userMessage)) {
    return false;
  }

  if (WRITE_ACTION_AFTER_CONNECTOR_PATTERN.test(userMessage) &&
    !WHOLE_MESSAGE_HOW_TO_PATTERN.test(userMessage)) {
    return false;
  }

  return isLikelyPivotTableRequest(userMessage) ||
    isLikelyChartRequest(userMessage) ||
    isLikelyTablePlanRequest(userMessage) ||
    isLikelyWorkbookStructureRequest(userMessage) ||
    isLikelySheetStructureRequest(userMessage) ||
    (
      SORT_KEYWORD_PATTERN.test(userMessage) &&
      (IMPLICIT_TARGET_PATTERN.test(userMessage) ||
        /(?:^|[^a-z])\$?[a-z]{1,3}\$?\d+(?:\:\$?[a-z]{1,3}\$?\d+)?(?:[^a-z]|$)/i.test(userMessage))
    ) ||
    FILTER_KEYWORD_PATTERN.test(userMessage) ||
    isLikelyNamedRangeRequest(rawMessage, userMessage) ||
    (
      VALIDATION_KEYWORD_PATTERN.test(userMessage) &&
      (
        /\b(add|set|apply|use|restrict|limit|validate|require)\b/.test(userMessage) ||
        /\bnamed range\b/.test(userMessage) ||
        A1_REFERENCE_PATTERN.test(userMessage)
      )
    ) ||
    isLikelyRangeFormatRequest(rawMessage, userMessage) ||
    isLikelyConditionalFormatRequest(userMessage) ||
    isLikelyRangeTransferRequest(userMessage) ||
    isLikelyDataCleanupRequest(userMessage) ||
    isLikelyExternalDataRequest(userMessage);
}

function hasFormulaContext(request: HermesRequest): boolean {
  if (typeof request.context.activeCell?.formula === "string" && request.context.activeCell.formula.trim()) {
    return true;
  }

  if ((request.context.referencedCells ?? []).some((cell) =>
    typeof cell.formula === "string" && cell.formula.trim().length > 0
  )) {
    return true;
  }

  return Boolean(request.context.selection?.formulas?.some((row) =>
    row.some((cell) => typeof cell === "string" && cell.trim().length > 0)
  ));
}

function buildPromptSafeRequest(request: HermesRequest): HermesRequest {
  if (!request.context.attachments || request.context.attachments.length === 0) {
    return request;
  }

  return {
    ...request,
    context: {
      ...request.context,
      attachments: request.context.attachments.map((attachment) => ({
        id: attachment.id,
        type: attachment.type,
        mimeType: attachment.mimeType,
        ...(attachment.fileName ? { fileName: attachment.fileName } : {}),
        ...(typeof attachment.size === "number" ? { size: attachment.size } : {}),
        source: attachment.source
      }))
    }
  };
}

function hasCurrentRegionContext(request: HermesRequest): boolean {
  return typeof request.context.currentRegion?.range === "string" &&
    request.context.currentRegion.range.trim().length > 0;
}

function isLikelyTotalsRowRequest(
  request: HermesRequest,
  rawMessage: string,
  normalizedMessage: string
): boolean {
  if (!TOTAL_ROW_KEYWORD_PATTERN.test(normalizedMessage)) {
    return false;
  }

  if (!TOTAL_ROW_VERB_PATTERN.test(normalizedMessage)) {
    return false;
  }

  return hasCurrentRegionContext(request) ||
    IMPLICIT_TARGET_PATTERN.test(normalizedMessage) ||
    A1_REFERENCE_PATTERN.test(rawMessage);
}

function isLikelySelectionExplanationRequest(normalizedMessage: string): boolean {
  return SELECTION_EXPLANATION_KEYWORD_PATTERN.test(normalizedMessage) &&
    SELECTION_REFERENCE_PATTERN.test(normalizedMessage);
}

function isLikelyMaterializedAnalysisRequest(rawMessage: string, normalizedMessage: string): boolean {
  if (!isLikelyAnalysisReportRequest(normalizedMessage)) {
    return false;
  }

  return (
    (MATERIALIZE_VERB_PATTERN.test(normalizedMessage) &&
      (MATERIALIZE_TARGET_PATTERN.test(normalizedMessage) || A1_REFERENCE_PATTERN.test(rawMessage))) ||
    /\b(to|into|onto|in|on|vao|tren)\b/.test(normalizedMessage) &&
      (MATERIALIZE_TARGET_PATTERN.test(normalizedMessage) || A1_REFERENCE_PATTERN.test(rawMessage))
  );
}

function isLikelyFormulaDebugRequest(
  request: HermesRequest,
  rawMessage: string,
  lowerMessage: string
): boolean {
  const formulaContextAvailable = hasFormulaContext(request);

  if (FORMULA_ERROR_TOKEN_PATTERN.test(rawMessage)) {
    return FORMULA_REFERENCE_PATTERN.test(lowerMessage) || formulaContextAvailable;
  }

  if (
    FORMULA_DEBUG_KEYWORD_PATTERN.test(lowerMessage) &&
    (FORMULA_REFERENCE_PATTERN.test(lowerMessage) || formulaContextAvailable)
  ) {
    return true;
  }

  if (
    formulaContextAvailable &&
    /\b(this|current|selected)\b/.test(lowerMessage) &&
    /\b(broken|wrong|fix|debug|explain|why)\b/.test(lowerMessage)
  ) {
    return true;
  }

  return false;
}

function isLikelyCurrentCellFormulaWriteTarget(
  request: HermesRequest,
  normalizedMessage: string
): boolean {
  return hasFormulaContext(request) && CURRENT_CELL_TARGET_PATTERN.test(normalizedMessage);
}

function isLikelyExternalDataRequest(
  normalizedMessage: string
): boolean {
  if (EXTERNAL_DATA_PROVIDER_PATTERN.test(normalizedMessage)) {
    return true;
  }

  if (MARKET_DATA_KEYWORD_PATTERN.test(normalizedMessage) &&
    EXTERNAL_DATA_ACTION_PATTERN.test(normalizedMessage)) {
    return true;
  }

  return WEB_IMPORT_KEYWORD_PATTERN.test(normalizedMessage) &&
    EXTERNAL_DATA_ACTION_PATTERN.test(normalizedMessage);
}

function isExcelHost(request: HermesRequest): boolean {
  return request.host.platform === "excel_windows" || request.host.platform === "excel_macos";
}

function getPreferredResponseType(
  request: HermesRequest,
  rawMessage: string,
  userMessage: string
): PreferredResponseType {
  if (isLikelyCompositeRequest(request, rawMessage, userMessage)) {
    return "composite_plan";
  }
  if (isLikelyToolScaffoldingRequest(request, rawMessage, userMessage)) {
    return "composite_plan";
  }
  if (hasInputLayoutConflictRisk(request, rawMessage, userMessage)) {
    return "composite_plan";
  }
  if (isLikelyAdvisoryOnlyArtifactQuestion(rawMessage, userMessage)) {
    return "chat";
  }
  if (isLikelyExternalDataRequest(userMessage)) {
    if (isExcelHost(request)) {
      return "error";
    }
    return "external_data_plan";
  }
  if (!/\bvalidation\b/.test(userMessage) &&
    FORMULA_APPLY_PATTERN.test(userMessage) &&
    (
      A1_REFERENCE_PATTERN.test(userMessage) ||
      isLikelyCurrentCellFormulaWriteTarget(request, userMessage)
    )) {
    return "sheet_update";
  }
  if (isLikelyCurrentTableFormulaFillRequest(request, userMessage)) {
    return "sheet_update";
  }
  if (isLikelyGeneratedDataRequest(rawMessage, userMessage)) {
    return "sheet_update";
  }
  if (isLikelyFormulaDebugRequest(request, rawMessage, userMessage)) {
    return "formula";
  }
  if (isLikelyTablePlanRequest(userMessage)) {
    return "table_plan";
  }
  if (SORT_KEYWORD_PATTERN.test(userMessage) &&
    (IMPLICIT_TARGET_PATTERN.test(userMessage) ||
      /(?:^|[^a-z])\$?[a-z]{1,3}\$?\d+(?:\:\$?[a-z]{1,3}\$?\d+)?(?:[^a-z]|$)/i.test(userMessage))) {
    return "range_sort_plan";
  }
  if (FILTER_KEYWORD_PATTERN.test(userMessage)) {
    return "range_filter_plan";
  }
  if (/\b(create|rename|delete|retarget)\b/.test(userMessage) && isLikelyNamedRangeRequest(rawMessage, userMessage)) {
    return "named_range_update";
  }
  if (
    VALIDATION_KEYWORD_PATTERN.test(userMessage) &&
    (
      /\b(add|set|apply|use|restrict|limit|validate|require)\b/.test(userMessage) ||
      /\bnamed range\b/.test(userMessage) ||
      A1_REFERENCE_PATTERN.test(userMessage)
    )
  ) {
    return "data_validation_plan";
  }
  if (
    isLikelyConditionalFormatRequest(userMessage)
  ) {
    return "conditional_format_plan";
  }
  if (isLikelyRangeFormatRequest(rawMessage, userMessage)) {
    return "range_format_update";
  }
  if (isLikelyPivotTableRequest(userMessage)) {
    return "pivot_table_plan";
  }
  if (isLikelyChartRequest(userMessage)) {
    return "chart_plan";
  }
  if (isLikelyAnalysisReportRequest(userMessage)) {
    return "analysis_report_plan";
  }
  if (isLikelyWorkbookStructureRequest(userMessage)) {
    return "workbook_structure_update";
  }
  if (isLikelyRangeTransferRequest(userMessage)) {
    return "range_transfer_plan";
  }
  if (isLikelyDataCleanupRequest(userMessage)) {
    return "data_cleanup_plan";
  }
  if (
    isLikelySheetStructureRequest(userMessage)
  ) {
    return "sheet_structure_update";
  }
  return /formula|cong thuc|sum|average|subtotal|fix .*formula|suggest .*formula|^=/.test(userMessage)
    ? "formula"
    : "chat";
}

function isLikelyExplicitWriteIntent(
  request: HermesRequest,
  rawMessage: string,
  normalizedMessage: string,
  preferredResponseType: PreferredResponseType
): boolean {
  if (preferredResponseType === "error") {
    return true;
  }

  if (WRITE_CAPABLE_RESPONSE_TYPES.has(preferredResponseType)) {
    return true;
  }

  return isLikelyMaterializedAnalysisRequest(rawMessage, normalizedMessage);
}

function buildHostCapabilityMatrixLines(request: HermesRequest): string[] {
  const supportsNoteWrites = request.capabilities.supportsNoteWrites === true;
  const helperSheetLine = [
    "- helper_sheet_scaffolding: supported.",
    "Use composite_plan for lookup tools, trackers, helper-sheet flows, and other user-facing input/output workflows.",
    "Keep the source table intact, put control cells and formulas outside the source table, and seed visible input labels, visible output labels, and a short guidance block."
  ].join(" ");
  const noteWriteLine = supportsNoteWrites
    ? "- note_writes: supported. sheet_update operations may use notes when they help the user-facing workflow."
    : "- note_writes: unsupported. Do not propose sheet_update operations that depend on notes.";

  switch (request.host.platform) {
    case "google_sheets":
      return [
        "Host capability matrix for google_sheets. Use this as the planning source of truth for host-exact routing:",
        helperSheetLine,
        "- pivot_table_plan: limited. Require a single-cell target anchor. Supported value aggregations: sum, count, average, min, max. Supported pivot sorting: group_field on an existing row or column group, and aggregated_value only when the pivot does not mix row and column groups. Supported pivot filters use existing pivot fields with operators equal_to, not_equal_to, greater_than, greater_than_or_equal_to, less_than, less_than_or_equal_to, between, or not_between. Every pivot filter requires value; between and not_between also require value2.",
        "- chart_plan: limited. Require a single-cell target anchor and categoryField. Supported chart types: bar, column, stacked_bar, stacked_column, line, area, pie, scatter. Series labels may use explicit custom legend text when provided. Supported legend positions: bottom, left, right, top, hidden. Optional horizontalAxisTitle and verticalAxisTitle are supported for charts with axes, but not pie charts. Pie charts support exactly one series. If you would naturally say none, emit hidden.",
        "- table_plan: limited. Format an existing range as a table-like range using row banding and optional filter buttons. Do not emit styleName or showTotalsRow=true on Google Sheets because those are not exact native table features.",
        "- external_data_plan: limited. Require a single-cell target anchor. Supported sourceType/provider pairs: market_data/googlefinance and web_table_import/importhtml, importxml, or importdata. Formula text must match the provider exactly. Web imports require data.sourceUrl to be a public HTTP(S) URL and must not assume authenticated or JavaScript-rendered pages. Web import formulas must not reference private or internal URLs.",
        "- range_format_update: supported. Supported static formatting fields: numberFormat, backgroundColor, textColor, fontFamily, fontSize, bold, italic, underline, strikethrough, horizontalAlignment, verticalAlignment, wrapStrategy, border, columnWidth, and rowHeight. Use conditional_format_plan for rule-based highlighting.",
        "- range_transfer_plan: limited. Supported pasteMode values: values, formulas, formats. Do not plan overlapping move or append destinations on the same sheet.",
        "- data_cleanup_plan: limited. normalize_case only supports upper, lower, title, and sentence. standardize_format only supports exact year-first date text patterns and fixed-decimal number text patterns.",
        "- range_filter_plan: limited. Do not combine multiple OR conditions in one exact step. topN filters are supported when display values cleanly separate visible and hidden rows. Use positive whole-number topN values only. Duplicate display values crossing the top-N cutoff are unsupported. Repeated conditions on the same column are unsupported.",
        "- named_range_update: limited. Only workbook-scoped named ranges are exact-safe.",
        "- data_validation_plan: limited. inputMessage maps to Google Sheets validation help text. inputTitle, errorTitle, and errorMessage are unsupported on Google Sheets. List validation cannot preserve allowBlank=true exactly. Single-value checkbox validation cannot preserve allowBlank=false exactly.",
        noteWriteLine
      ];
    case "excel_windows":
    case "excel_macos":
      return [
        `Host capability matrix for ${request.host.platform}. Use this as the planning source of truth for host-exact routing:`,
        helperSheetLine,
        "- pivot_table_plan: limited. Require a single-cell target anchor. Supported value aggregations: sum, count, average, min, max. Supported pivot sorting: group_field on an existing row or column group, and aggregated_value only when the pivot does not mix row and column groups. Supported pivot filters use existing pivot fields with operators equal_to, not_equal_to, greater_than, greater_than_or_equal_to, less_than, less_than_or_equal_to, between, or not_between. Every pivot filter requires value; between and not_between also require value2.",
        "- chart_plan: limited. Require a single-cell target anchor, categoryField, and at least one unique series field. Supported chart types: bar, column, stacked_bar, stacked_column, line, area, pie, scatter. Series labels may use explicit custom legend text when provided. Supported legend positions: top, bottom, left, right, hidden. Optional horizontalAxisTitle and verticalAxisTitle are supported for charts with axes, but not pie charts. Pie charts support exactly one series. sourceRange must expose a header row whose field order matches categoryField followed by the requested series fields exactly.",
        "- table_plan: limited. Create a native Excel table over an existing range. Supported options: name, styleName, hasHeaders, showBandedRows, showBandedColumns, showFilterButton, and showTotalsRow.",
        "- external_data_plan: unsupported. Do not emit external_data_plan on Excel hosts. Return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\" and suggest a simpler alternative such as a preview, a plain formula explanation, or moving the task to Google Sheets.",
        "- range_format_update: supported. Supported static formatting fields: numberFormat, backgroundColor, textColor, fontFamily, fontSize, bold, italic, underline, strikethrough, horizontalAlignment, verticalAlignment, wrapStrategy, border, columnWidth, and rowHeight. Use conditional_format_plan for rule-based highlighting.",
        "- range_transfer_plan: limited. Supported pasteMode values: values, formulas, formats.",
        "- data_cleanup_plan: limited. normalize_case only supports upper, lower, title, and sentence. standardize_format only supports exact year-first date text patterns and fixed-decimal number text patterns.",
        "- range_filter_plan: limited. Use combiner=and only. topN filters are supported through native Excel top item filters. Use positive whole-number topN values only. Repeated conditions on the same column are exact-safe only when exactly two custom criteria can be combined with AND.",
        "- data_validation_plan: limited. Checkbox values must stay on true and false. inputTitle, inputMessage, errorTitle, and errorMessage are supported for non-checkbox validation rules.",
        noteWriteLine
      ];
    default:
      return [noteWriteLine];
  }
}

export function getSpreadsheetRoutingHints(request: HermesRequest): SpreadsheetRoutingHints {
  const rawMessage = request.userMessage;
  const normalizedMessage = normalizeNaturalLanguage(rawMessage);
  const preferredResponseType = getPreferredResponseType(request, rawMessage, normalizedMessage);
  const generatedDataRequest = isLikelyGeneratedDataRequest(rawMessage, normalizedMessage);
  const toolScaffoldingOpportunity = isLikelyToolScaffoldingRequest(request, rawMessage, normalizedMessage);
  const inputLayoutConflictRisk = hasInputLayoutConflictRisk(request, rawMessage, normalizedMessage);
  const explicitWriteIntent = isLikelyExplicitWriteIntent(
    request,
    rawMessage,
    normalizedMessage,
    preferredResponseType
  );
  const mixedAdvisoryAndWriteRequest =
    COMPOSITE_CONNECTOR_PATTERN.test(normalizedMessage) &&
    explicitWriteIntent &&
    (isLikelySelectionExplanationRequest(normalizedMessage) ||
      isLikelyFormulaDebugRequest(request, rawMessage, normalizedMessage));

  return {
    preferredResponseType,
    generatedDataRequest,
    explicitWriteIntent,
    mixedAdvisoryAndWriteRequest,
    toolScaffoldingOpportunity,
    inputLayoutConflictRisk
  };
}

export function buildHermesSpreadsheetRequestPrompt(request: HermesRequest): string {
  const routingHints = getSpreadsheetRoutingHints(request);
  const preferredResponseType = routingHints.preferredResponseType;
  const reviewerUnavailable = request.reviewer.forceExtractionMode === "unavailable";
  const generatedDataRequest = routingHints.generatedDataRequest;
  const normalizedUserMessage = normalizeNaturalLanguage(request.userMessage);
  const currentTableFormulaFillRequest = isLikelyCurrentTableFormulaFillRequest(request, normalizedUserMessage);
  const formulaWriteNeedsContext =
    preferredResponseType === "sheet_update" &&
    hasFormulaContext(request) &&
    FORMULA_APPLY_PATTERN.test(normalizedUserMessage);

  return [
    "Return JSON only.",
    "Return exactly one JSON object and nothing else.",
    "Do not wrap the JSON in markdown fences.",
    "Do not include prose before or after the JSON object.",
    "Do not expose chain-of-thought, hidden reasoning, internal prompts, secrets, or raw stack traces.",
    "The final object must validate against exactly one internal Hermes structured-body schema.",
    `Choose exactly one response type from: ${STRUCTURED_BODY_TYPES.join(", ")}.`,
    `Required structured body fields: ${REQUIRED_STRUCTURED_BODY_FIELDS.join(", ")}.`,
    "If type=\"chat\", data.message is required. data.followUpSuggestions and data.confidence are optional. Do not include any other keys inside chat data.",
    "If type=\"formula\", data.intent, data.formula, data.formulaLanguage, data.explanation, and data.confidence are required. Use data.formulaLanguage=\"excel\" for excel_windows/excel_macos and \"google_sheets\" for google_sheets.",
    "If type=\"composite_plan\", data.steps, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.overwriteRisk, data.confirmationLevel, data.reversible, data.dryRunRecommended, and data.dryRunRequired are required. composite_plan always requires confirmation. Each step must include stepId, dependsOn, continueOnError, and plan. Step plans must be contract-valid executable plans only. Do not nest composite_plan inside a composite_plan. Do not include analysis_report_plan with outputMode=chat_only inside a composite_plan.",
    "If type=\"workbook_structure_update\", data.operation, data.sheetName, data.explanation, data.confidence, and data.requiresConfirmation must be true. Supported operations are create_sheet, delete_sheet, rename_sheet, duplicate_sheet, move_sheet, hide_sheet, and unhide_sheet.",
    "If type=\"range_format_update\", data.targetSheet, data.targetRange, data.format, data.explanation, data.confidence, and data.requiresConfirmation must be true. data.format may only contain supported formatting fields: numberFormat, backgroundColor, textColor, fontFamily, fontSize, bold, italic, underline, strikethrough, horizontalAlignment, verticalAlignment, wrapStrategy, border, columnWidth, rowHeight. data.format.border may include all, outer, inner, top, bottom, left, right, innerHorizontal, or innerVertical with style none, solid, dashed, dotted, double, medium, or thick and optional color.",
    "If type=\"conditional_format_plan\", data.targetSheet, data.targetRange, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.replacesExistingRules, and data.managementMode are required. conditional_format_plan is distinct from range_format_update. clear_on_target must contain no rule payload. replace_all_on_target removes existing target rules before applying the new rule. number_compare values must be finite numbers. date_compare values must be valid YYYY-MM-DD literals. Do not return a vague highlight plan with only explanation and ranges. Include a full contract-valid rule payload. For row-highlighting logic driven by a status/breach/overdue column or by comparisons between columns, prefer ruleType=\"custom_formula\" with an exact formula instead of static row-by-row formatting. If the requested conditional formatting cannot be mapped exactly on the current host, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"sheet_structure_update\", data.targetSheet, data.operation, data.explanation, data.confidence, data.requiresConfirmation, and data.confirmationLevel are required. Use sheet_structure_update for insert/delete/hide/unhide/merge/unmerge/freeze/unfreeze/group/ungroup/autofit/tab color style changes. delete_rows and delete_columns require data.confirmationLevel=\"destructive\". all other sheet_structure_update operations require data.confirmationLevel=\"standard\". unfreeze_panes must resolve to data.frozenRows=0 and data.frozenColumns=0. Provide only the operation-specific contract fields.",
    "If type=\"range_sort_plan\", data.targetSheet, data.targetRange, data.hasHeader, data.keys, data.explanation, data.confidence, and data.requiresConfirmation must be true. data.keys must include one or more sort keys.",
    "If type=\"range_filter_plan\", data.targetSheet, data.targetRange, data.hasHeader, data.conditions, data.combiner, data.clearExistingFilters, data.explanation, data.confidence, and data.requiresConfirmation must be true. data.conditions must include one or more filter conditions.",
    "If type=\"data_validation_plan\", data.targetSheet, data.targetRange, data.ruleType, data.allowBlank, data.invalidDataBehavior, data.explanation, data.confidence, and data.requiresConfirmation must be true. For list validation, use exactly one source: values, sourceRange, or namedRangeName. Optional inputTitle, inputMessage, errorTitle, and errorMessage customize validation prompts and invalid-entry alerts when the current host can apply them exactly.",
    "If type=\"analysis_report_plan\", data.sourceSheet, data.sourceRange, data.outputMode, data.sections, data.explanation, data.confidence, data.affectedRanges, data.overwriteRisk, and data.confirmationLevel are required. data.sections must be an array of objects with type, title, summary, and sourceRanges. Do not use plain strings or slugs for report sections. data.outputMode must be chat_only or materialize_report. chat_only requires data.requiresConfirmation=false and must not ask for confirmation. materialize_report requires data.targetSheet, data.targetRange, and data.requiresConfirmation=true. For materialize_report, data.targetRange must be the full 4-column destination rectangle for the report matrix, never just the anchor cell. If the requested report artifact cannot be mapped exactly on the current host, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"pivot_table_plan\", data.sourceSheet, data.sourceRange, data.targetSheet, data.targetRange, data.rowGroups, data.valueAggregations, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.overwriteRisk, and data.confirmationLevel are required. If the requested pivot table cannot be mapped exactly on the current host, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"chart_plan\", data.sourceSheet, data.sourceRange, data.targetSheet, data.targetRange, data.chartType, data.categoryField, data.series, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.overwriteRisk, and data.confirmationLevel are required. chartType must be bar, column, stacked_bar, stacked_column, line, area, pie, or scatter. categoryField must reference the source header used for the category axis. Each series item must use field to reference a source header name. Do not use A1 ranges or name/range objects inside data.series. Use optional horizontalAxisTitle and verticalAxisTitle for axis titles only when the selected chart type has axes; do not use axis titles with pie charts. If the requested chart cannot be mapped exactly on the current host, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"table_plan\", data.targetSheet, data.targetRange, data.hasHeaders, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.overwriteRisk, and data.confirmationLevel are required. data.name, data.styleName, data.showBandedRows, data.showBandedColumns, data.showFilterButton, and data.showTotalsRow are optional. If the requested table behavior cannot be mapped exactly on the current host, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"named_range_update\", data.operation, data.scope, data.name, data.explanation, data.confidence, and data.requiresConfirmation must be true. create and retarget must include targetSheet and targetRange. rename must include newName. delete must not invent a target range.",
    "If type=\"range_transfer_plan\", data.sourceSheet, data.sourceRange, data.targetSheet, data.targetRange, data.operation, data.pasteMode, data.transpose, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.overwriteRisk, and data.confirmationLevel are required. data.targetRange is required and must be the full destination rectangle, never just an anchor cell. range_transfer_plan is distinct from sheet_update. move requires data.confirmationLevel=\"destructive\". copy and append require data.confirmationLevel=\"standard\". If the target sheet is known but the full destination rectangle cannot be resolved, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\"; do not default data.targetRange to A1. If overlap ambiguity cannot be resolved exactly, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"data_cleanup_plan\", data.targetSheet, data.targetRange, data.operation, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.overwriteRisk, and data.confirmationLevel are required. data_cleanup_plan is distinct from sheet_update. Destructive cleanup operations require data.confirmationLevel=\"destructive\". Non-destructive cleanup operations require data.confirmationLevel=\"standard\". Do not compress multiple cleanup transforms into one broad cleanup step. If the request mixes trim, casing, duplicate removal, fill-down, split/join, or standardize-format work, prefer type=\"composite_plan\" with one exact cleanup step per transform. For operation=\"standardize_format\", target the specific column or range being normalized and include exactly one formatType and one formatPattern per step. Unsupported fuzzy or heuristic cleanup requests must return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"analysis_report_update\", data.operation, data.targetSheet, data.targetRange, and data.summary are required.",
    "If type=\"pivot_table_update\", data.operation, data.targetSheet, data.targetRange, and data.summary are required.",
    "If type=\"chart_update\", data.operation, data.targetSheet, data.targetRange, data.chartType, and data.summary are required.",
    "If type=\"table_update\", data.operation, data.targetSheet, data.targetRange, data.hasHeaders, and data.summary are required.",
    "If type=\"sheet_update\", data.targetSheet, data.targetRange, data.operation, data.explanation, data.confidence, data.shape, and data.requiresConfirmation must be true. Any provided values, formulas, or notes matrix must match data.shape, and targetRange must match the same rectangle.",
    "If type=\"sheet_import_plan\", data.sourceAttachmentId, data.targetSheet, data.targetRange, data.headers, data.values, data.confidence, data.extractionMode, data.shape, and data.requiresConfirmation must be true. shape.rows includes the header row and targetRange must match data.shape.",
    "If type=\"external_data_plan\", data.targetSheet, data.targetRange, data.sourceType, data.provider, data.formula, data.explanation, data.confidence, data.requiresConfirmation, data.affectedRanges, data.overwriteRisk, and data.confirmationLevel are required. targetRange must be a single-cell anchor. For sourceType=\"market_data\", provider must be googlefinance and data.query.symbol is required. For sourceType=\"web_table_import\", provider must be importhtml, importxml, or importdata, data.sourceUrl is required and must be a public HTTP(S) URL, web import formulas must not reference private or internal URLs, and selector requirements must match the provider exactly: importhtml uses selectorType table or list with a positive numeric selector, importxml uses selectorType xpath with a string selector, and importdata uses selectorType direct with no selector. If the requested external data flow cannot be mapped exactly on the current host, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    "If type=\"error\", data.code, data.message, and data.retryable are required. Prefer type=\"error\" with data.code=\"EXTRACTION_UNAVAILABLE\" when reviewer-safe unavailable mode blocks extraction. data.message must be user-facing. If extra guidance helps, include data.userAction. Do not mention internal contracts, schema names, or validation failures in data.message.",
    "If the user asks to generate, seed, or populate sample/random/mock data into an existing sheet or range, prefer type=\"sheet_update\" with operation=\"replace_range\" and concrete values.",
    "If the user asks to create a new sheet and then generate, seed, or populate sample/random/mock data, prefer type=\"composite_plan\" with a create_sheet step followed by a sheet_update step that uses operation=\"replace_range\".",
    "Use the host capability matrix below as the planning source of truth for current host exact-safe support. If a capability family is marked unsupported, do not emit that plan type. If it is marked limited, stay within the listed safe subset or return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\".",
    ...buildHostCapabilityMatrixLines(request),
    "You may receive context.currentRegion, context.currentRegionArtifactTarget, and context.currentRegionAppendTarget.",
    "For large selections or large current tables, full values/formulas matrices may be omitted to keep the payload bounded. Use the provided range, headers, activeCell, referencedCells, and any available currentRegion context instead of treating omitted large matrices as missing context by default.",
    "If the user refers to the current table, current data, current range, this table, this data, or this range and context.currentRegion is present, use context.currentRegion as the implicit table/range instead of asking the user to reselect it.",
    "If the user asks for a chart, pivot table, table formatting, sort, filter, cleanup, or analysis artifact on the current table/range/data and context.currentRegion is present, use host.activeSheet plus context.currentRegion.range as the implicit source or target range when no explicit A1 range is provided.",
    "If the user asks for a chart, pivot table, or materialized report on the current table/range/data and no explicit artifact anchor is provided, use context.currentRegionArtifactTarget as the default artifact anchor when it is present. For chart and pivot plans, that anchor is the targetRange; for materialized analysis reports, expand from that anchor to the full 4-column report destination rectangle.",
    "When an explicit pivot request leaves some layout choices unspecified and the source table is identifiable, do not degrade to chat-only just because the user did not name every pivot field.",
    "For under-specified pivot requests, infer a conservative default pivot: prefer one categorical row group from a header like Category, Region, Department, Type, or Status; add one or two numeric valueAggregations using sum when clear numeric measures exist; if no numeric measure is obvious, use count on a stable identifier-like field.",
    "For under-specified pivot requests without an explicit artifact anchor, use context.currentRegionArtifactTarget when it is present; otherwise choose a safe nearby artifact anchor or a dedicated report sheet instead of blocking on clarification.",
    "If the user asks to add a totals row, subtotal row, or grand total for the current table/range/data and context.currentRegionAppendTarget is present, prefer type=\"composite_plan\" with an insert_rows step followed by a sheet_update step that uses operation=\"set_formulas\" on the append target range.",
    "Do not ask the user to select the whole table again when context.currentRegion already identifies it.",
    currentTableFormulaFillRequest
      ? "This request fills or populates an existing column in the current table based on other columns or a lookup table."
      : "",
    currentTableFormulaFillRequest
      ? "Prefer type=\"sheet_update\" with operation=\"set_formulas\" for the target data rows, using currentRegion headers to infer the source and target columns when the user names them by header text."
      : "",
    currentTableFormulaFillRequest
      ? "Do not ask the user to reselect the table when context.currentRegion already identifies the current table."
      : "",
    routingHints.inputLayoutConflictRisk
      ? "The request names explicit input/output control cells that appear to overlap the current source table or header region."
      : "",
    routingHints.inputLayoutConflictRisk
      ? "Do not reject the request solely because those control cells overlap the current source table."
      : "",
    routingHints.inputLayoutConflictRisk
      ? "If the source table is identifiable from the current sheet or context.currentRegion, prefer a safe write-capable response that preserves the source table, usually type=\"composite_plan\" that creates a helper sheet (for example Lookup_Demo) and places the control cells and formula there."
      : "",
    routingHints.inputLayoutConflictRisk
      ? "Only ask for clarification if you still cannot identify both a source table and a safe target layout."
      : "",
    routingHints.toolScaffoldingOpportunity
      ? "This request is a tool-like spreadsheet flow with user-facing inputs and outputs."
      : "",
    routingHints.toolScaffoldingOpportunity
      ? "Prefer type=\"composite_plan\" that creates or reuses a dedicated helper sheet instead of silently dropping a single formula into the source table."
      : "",
    routingHints.toolScaffoldingOpportunity
      ? "When scaffolding that helper sheet, seed clear labels for inputs and outputs, keep the source table unchanged, and add a short guidance row or section so the sheet is usable without extra chat context."
      : "",
    routingHints.toolScaffoldingOpportunity
      ? "Use workbook_structure_update/create_sheet plus one or more sheet_update steps for labels, formulas, and any safe example or placeholder values."
      : "",
    "If the user asks for an unsupported workbook or formatting action that cannot be represented by the contract, return type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\" and a clear user-facing message about the unsupported workbook or formatting action. Do not mention internal contracts, schema names, or validation failures. Briefly explain the limitation, then either ask one concise clarifying question or suggest up to three closest supported alternatives.",
    routingHints.explicitWriteIntent
      ? "This request explicitly asks for a spreadsheet change. Do not return type=\"chat\" or any other non-applying advisory-only response. Return the closest valid write-capable plan, or return a user-facing error if exact execution is not possible."
      : "",
    routingHints.mixedAdvisoryAndWriteRequest
      ? "This request mixes an advisory explanation/debug step with a spreadsheet write action."
      : "",
    routingHints.mixedAdvisoryAndWriteRequest
      ? "If one write-capable plan can satisfy the write and the explanation fits naturally in data.explanation, prefer that write-capable plan instead of a separate chat-only step."
      : "",
    routingHints.mixedAdvisoryAndWriteRequest
      ? "If the user is explicitly asking for a separate chat-only step before the write and that exact sequence cannot be represented by one contract-valid response or composite_plan, return type=\"error\" with a user-facing message asking the user to split the analysis and writeback into separate steps."
      : "",
    "Optional structured body fields are: warnings, skillsUsed, downstreamProvider.",
    "If you include warnings, warnings must be an array of objects. Each warning object must include code and message, and may include severity and field.",
    "Do not answer explicit confirmation phrasing with a chat acknowledgement when the user is still naming a workbook or spreadsheet action. If the message says things like \"confirm create sheet ...\" or \"confirm delete rows ...\" and the action maps to a contract plan type, return that plan type instead of chat.",
    "For the chosen response type, data must contain only contract-defined fields. Do not add extra keys.",
    "If capabilities.supportsNoteWrites is not true, do not propose note-based sheet updates.",
    "Do not include external Step 1 envelope fields such as schemaVersion, requestId, hermesRunId, processedBy, serviceLabel, environmentLabel, startedAt, completedAt, durationMs, trace, or ui.",
    "Preserve contract-safe public fields only in data, warnings, skillsUsed, and downstreamProvider.",
    preferredResponseType === "composite_plan"
      ? "This request is an explicit multi-step spreadsheet workflow. Prefer type=\"composite_plan\" over a single plan when the user clearly asks for multiple ordered actions."
      : preferredResponseType === "error"
      ? "This request asks for a spreadsheet flow that the current host capability matrix marks unsupported. Prefer type=\"error\" with data.code=\"UNSUPPORTED_OPERATION\" instead of emitting a host-unsupported write plan."
      : preferredResponseType === "sheet_update"
      ? generatedDataRequest
        ? "This request is an explicit data-population flow. Prefer type=\"sheet_update\" with concrete values when the target sheet or range already exists."
        : "This request targets a specific cell or range for formula application. Prefer type=\"sheet_update\" with operation=\"set_formulas\" over a formula-only advisory response when you can identify the targetCell or targetRange."
      : preferredResponseType === "range_sort_plan"
      ? "This request is an explicit spreadsheet sort flow. Prefer type=\"range_sort_plan\" when you can identify the target table or range."
      : preferredResponseType === "range_filter_plan"
      ? "This request is an explicit spreadsheet filter flow. Prefer type=\"range_filter_plan\" when you can identify the target table or range."
      : preferredResponseType === "data_validation_plan"
      ? "This request is an explicit validation flow. Prefer type=\"data_validation_plan\" for dropdowns, checkboxes, validation rules, allow-only rules, reject-invalid rules, and named-range-backed validation."
      : preferredResponseType === "analysis_report_plan"
      ? "This request is an explicit spreadsheet analysis flow. Prefer type=\"analysis_report_plan\" for structured analysis reports, with chat_only for non-write summaries and materialize_report for confirmable report artifacts."
      : preferredResponseType === "pivot_table_plan"
      ? "This request is an explicit spreadsheet pivot flow. Prefer type=\"pivot_table_plan\" for native pivot table artifacts."
      : preferredResponseType === "chart_plan"
      ? "This request is an explicit spreadsheet chart flow. Prefer type=\"chart_plan\" for native chart artifacts."
      : preferredResponseType === "table_plan"
      ? "This request is an explicit format-as-table flow. Prefer type=\"table_plan\" for native Excel tables or exact-safe Google Sheets table-like formatting."
      : preferredResponseType === "conditional_format_plan"
      ? "This request is an explicit conditional-formatting flow. Prefer type=\"conditional_format_plan\" for highlight, duplicate-marking, threshold-coloring, color-scale, and clear-conditional-format asks."
      : preferredResponseType === "range_format_update"
      ? "This request is an explicit static range-formatting flow. Prefer type=\"range_format_update\" for number formats, fill/text color, font style, alignment, wrapping, borders, row height, and column width."
      : preferredResponseType === "named_range_update"
      ? "This request is an explicit named-range flow. Prefer type=\"named_range_update\" for create, rename, delete, or retarget named range requests."
      : preferredResponseType === "workbook_structure_update"
      ? "This request is an explicit workbook sheet-structure flow. Prefer type=\"workbook_structure_update\" for create, delete, rename, duplicate, move, hide, or unhide sheet/tab/worksheet actions."
      : preferredResponseType === "range_transfer_plan"
      ? "This request is an explicit spreadsheet transfer flow. Prefer type=\"range_transfer_plan\" for copy, move, append, and transpose asks."
      : preferredResponseType === "data_cleanup_plan"
      ? "This request is an explicit spreadsheet cleanup flow. Prefer type=\"data_cleanup_plan\" for trim, duplicate removal, split, join, fill-down, and standardize-format asks."
      : preferredResponseType === "external_data_plan"
      ? "This request is an explicit external-data spreadsheet flow. Prefer type=\"external_data_plan\" for stock/crypto market data or public website-table imports into Google Sheets."
      : preferredResponseType === "sheet_structure_update"
      ? "This request is an explicit spreadsheet structure flow. Prefer type=\"sheet_structure_update\" for insert, delete, hide, unhide, merge, unmerge, freeze, unfreeze, group, ungroup, autofit, or tab color operations."
      : preferredResponseType === "chat"
      ? "This request is a spreadsheet selection explanation flow. Prefer type=\"chat\" unless another valid structured body type is strictly more correct."
      : "This request is a spreadsheet formula help flow. Prefer type=\"formula\" unless another valid structured body type is strictly more correct.",
    preferredResponseType === "formula"
      ? "For formula debugging or explanation asks, prefer data.intent=\"explain\"."
      : "",
    preferredResponseType === "formula"
      ? "For formula correction asks that do not explicitly apply a write to a target cell or range, prefer data.intent=\"fix\"."
      : "",
    preferredResponseType === "formula" || formulaWriteNeedsContext
      ? "When available, inspect context.activeCell.formula, context.selection.formulas, and context.referencedCells formulas before answering."
      : "",
    "If the request cannot be completed, return a valid structured error body instead of prose.",
    reviewerUnavailable
      ? "Reviewer mode is forcing extractionMode=\"unavailable\". Never fabricate extracted_table or sheet_import_plan output. Prefer type=\"error\" with data.code=\"EXTRACTION_UNAVAILABLE\"."
      : "Never fabricate extracted_table or sheet_import_plan output when reviewer-safe unavailable mode is active.",
    "Serialized Step 2 backend request envelope JSON follows.",
    JSON.stringify(buildPromptSafeRequest(request), null, 2)
  ].join("\n");
}
