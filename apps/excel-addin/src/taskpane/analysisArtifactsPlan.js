const ANALYSIS_OUTPUT_MODES = new Set([
  "chat_only",
  "materialize_report"
]);

const CHART_TYPES = new Set([
  "bar",
  "column",
  "stacked_bar",
  "stacked_column",
  "line",
  "area",
  "pie",
  "scatter"
]);

function joinList(values) {
  return Array.isArray(values) && values.length > 0 ? values.join(", ") : "";
}

function columnLettersToNumber(columnLetters) {
  let column = 0;
  for (const character of String(columnLetters || "").trim().toUpperCase()) {
    column = (column * 26) + (character.charCodeAt(0) - 64);
  }
  return column;
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

function parseA1Anchor(range) {
  const normalized = String(range || "").trim().toUpperCase().replaceAll("$", "");
  const withoutSheet = normalized.includes("!") ? normalized.split("!").pop() : normalized;
  const anchor = withoutSheet.split(":")[0];
  const match = anchor.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    throw new Error(`Unsupported analysis report targetRange anchor: ${range}`);
  }

  return {
    startColumn: columnLettersToNumber(match[1]),
    startRow: Number(match[2])
  };
}

function parseA1Range(range) {
  const normalized = String(range || "").trim().toUpperCase().replaceAll("$", "");
  const withoutSheet = normalized.includes("!") ? normalized.split("!").pop() : normalized;
  const [startRef, endRef = startRef] = withoutSheet.split(":");
  const start = parseA1Anchor(startRef);
  const end = parseA1Anchor(endRef);

  return {
    startRow: Math.min(start.startRow, end.startRow),
    endRow: Math.max(start.startRow, end.startRow),
    startColumn: Math.min(start.startColumn, end.startColumn),
    endColumn: Math.max(start.startColumn, end.startColumn)
  };
}

function a1RangeReferencesMatch(left, right) {
  try {
    const leftBounds = parseA1Range(left);
    const rightBounds = parseA1Range(right);
    return leftBounds.startRow === rightBounds.startRow &&
      leftBounds.endRow === rightBounds.endRow &&
      leftBounds.startColumn === rightBounds.startColumn &&
      leftBounds.endColumn === rightBounds.endColumn;
  } catch {
    return false;
  }
}

function padMatrixRows(rows) {
  const width = Math.max(1, ...rows.map((row) => row.length));
  return rows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ""));
}

export function isAnalysisReportPlan(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === "string" &&
    typeof plan.sourceRange === "string" &&
    ANALYSIS_OUTPUT_MODES.has(plan.outputMode) &&
    Array.isArray(plan.sections) &&
    plan.sections.length > 0
  );
}

export function isMaterializedAnalysisReportPlan(plan) {
  return isAnalysisReportPlan(plan) &&
    plan.outputMode === "materialize_report" &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string";
}

export function isPivotTablePlan(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === "string" &&
    typeof plan.sourceRange === "string" &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    Array.isArray(plan.rowGroups) &&
    plan.rowGroups.length > 0 &&
    Array.isArray(plan.valueAggregations) &&
    plan.valueAggregations.length > 0
  );
}

export function isChartPlan(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === "string" &&
    typeof plan.sourceRange === "string" &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    CHART_TYPES.has(plan.chartType) &&
    Array.isArray(plan.series) &&
    plan.series.length > 0
  );
}

export function buildAnalysisReportMatrix(plan) {
  const rows = [
    ["Analysis report"],
    ["Source sheet", plan.sourceSheet],
    ["Source range", plan.sourceRange],
    ["Section", "Title", "Summary", "Source ranges"],
    ...plan.sections.map((section) => [
      section.type,
      section.title,
      section.summary,
      joinList(section.sourceRanges)
    ])
  ];

  return padMatrixRows(rows);
}

export function resolveAnalysisReportTargetRange(plan) {
  const matrix = buildAnalysisReportMatrix(plan);
  const anchor = parseA1Anchor(plan.targetRange);
  const endRow = anchor.startRow + matrix.length - 1;
  const endColumn = anchor.startColumn + (matrix[0]?.length || 1) - 1;
  const startCell = `${columnNumberToLetters(anchor.startColumn)}${anchor.startRow}`;
  const endCell = `${columnNumberToLetters(endColumn)}${endRow}`;

  return startCell === endCell ? startCell : `${startCell}:${endCell}`;
}

function normalizeAnalysisReportAffectedRanges(plan, resolvedTargetRange) {
  const resolvedTargetRef = `${plan.targetSheet}!${resolvedTargetRange}`;
  const anchorTargetRef = `${plan.targetSheet}!${plan.targetRange}`;
  const normalizedRanges = Array.isArray(plan.affectedRanges)
    ? plan.affectedRanges.map((range) => range === anchorTargetRef ? resolvedTargetRef : range)
    : [];

  if (!normalizedRanges.includes(resolvedTargetRef)) {
    normalizedRanges.push(resolvedTargetRef);
  }

  return normalizedRanges;
}

export function getAnalysisReportTargetRangeSupportError(plan) {
  if (!isMaterializedAnalysisReportPlan(plan)) {
    return "";
  }

  let expectedTargetRange = "";
  try {
    expectedTargetRange = resolveAnalysisReportTargetRange(plan);
  } catch {
    return "Excel host requires analysis report targetRange to match the full destination rectangle.";
  }

  if (!a1RangeReferencesMatch(plan.targetRange, expectedTargetRange)) {
    return "Excel host requires analysis report targetRange to match the full destination rectangle.";
  }

  return "";
}

export function resolveMaterializedAnalysisReportPlan(plan) {
  if (!isMaterializedAnalysisReportPlan(plan)) {
    return plan;
  }

  if (getAnalysisReportTargetRangeSupportError(plan)) {
    return plan;
  }

  const resolvedTargetRange = resolveAnalysisReportTargetRange(plan);

  return {
    ...plan,
    targetRange: resolvedTargetRange,
    affectedRanges: normalizeAnalysisReportAffectedRanges(plan, resolvedTargetRange)
  };
}

export function getAnalysisReportPreviewSummary(plan) {
  if (plan.outputMode === "chat_only") {
    return `Will answer with a chat-only analysis of ${plan.sourceSheet}!${plan.sourceRange}.`;
  }

  const resolvedPlan = resolveMaterializedAnalysisReportPlan(plan);
  return `Will materialize an analysis report on ${resolvedPlan.targetSheet}!${resolvedPlan.targetRange}.`;
}

export function getPivotTablePreviewSummary(plan) {
  return `Will create a pivot table on ${plan.targetSheet}!${plan.targetRange}.`;
}

export function getChartPreviewSummary(plan) {
  return `Will create a ${plan.chartType} chart on ${plan.targetSheet}!${plan.targetRange}.`;
}

export function getAnalysisReportStatusSummary(plan) {
  const resolvedPlan = resolveMaterializedAnalysisReportPlan(plan);
  return `Created analysis report on ${resolvedPlan.targetSheet}!${resolvedPlan.targetRange}.`;
}

export function getPivotTableStatusSummary(plan) {
  return `Created pivot table on ${plan.targetSheet}!${plan.targetRange}.`;
}

export function getChartStatusSummary(plan) {
  const chartType = typeof plan?.chartType === "string" && plan.chartType.length > 0
    ? `${plan.chartType} `
    : "";
  return `Created ${chartType}chart on ${plan.targetSheet}!${plan.targetRange}.`;
}
