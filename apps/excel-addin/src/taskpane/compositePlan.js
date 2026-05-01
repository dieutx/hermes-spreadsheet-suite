function isPivotTablePlanLike(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === "string" &&
    typeof plan.sourceRange === "string" &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    Array.isArray(plan.rowGroups) &&
    Array.isArray(plan.valueAggregations)
  );
}

function isChartPlanLike(plan) {
  return Boolean(
    plan &&
    typeof plan.sourceSheet === "string" &&
    typeof plan.sourceRange === "string" &&
    typeof plan.targetSheet === "string" &&
    typeof plan.targetRange === "string" &&
    typeof plan.chartType === "string" &&
    Array.isArray(plan.series)
  );
}

function isDestructiveCompositeStep(step) {
  return step?.plan?.confirmationLevel === "destructive";
}

function isLikelyReversibleCompositeStep(step) {
  if (!step?.plan) {
    return false;
  }

  if (isPivotTablePlanLike(step.plan) || isChartPlanLike(step.plan)) {
    return step.plan.confirmationLevel !== "destructive";
  }

  if (step.plan.confirmationLevel === "destructive" ||
    step.plan.operation === "move" ||
    step.plan.operation === "remove_blank_rows" ||
    step.plan.operation === "remove_duplicate_rows" ||
    step.plan.operation === "delete_rows" ||
    step.plan.operation === "delete_columns" ||
    step.plan.operation === "delete_sheet") {
    return false;
  }

  return true;
}

export function getCompositePreviewSummary(plan) {
  return `Will run ${plan.steps.length} workflow step${plan.steps.length === 1 ? "" : "s"}.`;
}

export function isCompositePlan(plan) {
  return Boolean(
    plan &&
    Array.isArray(plan.steps) &&
    plan.steps.length > 0 &&
    typeof plan.explanation === "string"
  );
}

export function buildCompositeStepSummary(step) {
  if (typeof step?.plan?.explanation === "string" && step.plan.explanation.trim().length > 0) {
    return step.plan.explanation;
  }

  if (typeof step?.plan?.operation === "string" && step.plan.operation.trim().length > 0) {
    return `Execute ${step.plan.operation}.`;
  }

  return `Execute ${step?.stepId || "workflow step"}.`;
}

export function buildCompositeStepPreview(step) {
  return {
    stepId: step.stepId,
    dependsOn: step.dependsOn || [],
    continueOnError: Boolean(step.continueOnError),
    destructive: isDestructiveCompositeStep(step),
    reversible: isLikelyReversibleCompositeStep(step),
    summary: buildCompositeStepSummary(step)
  };
}

export function getCompositeStatusSummary(result) {
  const resolvedSummary = typeof result?.summary === "string" ? result.summary.trim() : "";
  const stepResults = Array.isArray(result?.stepResults) ? result.stepResults : [];
  const count = stepResults.length;
  const baseSummary = resolvedSummary || `Completed workflow with ${count} step${count === 1 ? "" : "s"}.`;
  const detailSummary = buildCompositeResultDetailSummary(stepResults);

  if (!detailSummary) {
    return baseSummary;
  }

  if (isGenericCompositeSummary(baseSummary)) {
    return `${baseSummary} ${detailSummary}`.trim();
  }

  return baseSummary;
}

function isGenericCompositeSummary(summary) {
  const normalized = String(summary || "").trim();
  if (!normalized) {
    return true;
  }

  return /^Workflow finished:/i.test(normalized) ||
    /^Completed workflow with \d+ step/i.test(normalized);
}

function normalizeCompositeStepSummary(summary) {
  const normalized = String(summary || "").trim().replace(/\s+/g, " ");
  return normalized.replace(/[.]+$/g, "").trim();
}

function joinCompositeSummaryParts(parts, limit) {
  if (parts.length === 0) {
    return "";
  }

  const visible = parts.slice(0, limit);
  const hiddenCount = parts.length - visible.length;
  const joined = visible.join("; ");
  return hiddenCount > 0 ? `${joined}; +${hiddenCount} more` : joined;
}

function buildCompositeResultDetailSummary(stepResults) {
  const completed = [];
  const failed = [];
  const skipped = [];

  for (const step of stepResults) {
    const normalized = normalizeCompositeStepSummary(step?.summary);
    if (!normalized) {
      continue;
    }

    if (step.status === "completed") {
      completed.push(normalized);
      continue;
    }

    if (step.status === "failed") {
      failed.push(normalized);
      continue;
    }

    if (step.status === "skipped") {
      skipped.push(normalized);
    }
  }

  const details = [];
  if (completed.length > 0) {
    details.push(`Completed: ${joinCompositeSummaryParts(completed, 3)}.`);
  }
  if (failed.length > 0) {
    details.push(`Failed: ${joinCompositeSummaryParts(failed, 2)}.`);
  }
  if (skipped.length > 0) {
    details.push(`Skipped: ${joinCompositeSummaryParts(skipped, 2)}.`);
  }

  return details.join(" ");
}
