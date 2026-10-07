// The validator is the safety gate between "whatever the LLM said" and the
// user's workbook. It runs in two layers:
//
//   1. Shape  — does the JSON match the Zod schema? (types, required keys, no extras)
//   2. Meaning — do the addresses exist, do array sizes match, are the formulas
//                safe, is the request a sane size?
//
// It never throws on bad input: it returns a list of issues so the UI can show
// them and the orchestrator can send them back to the LLM for a repair attempt.

import {
  cellCount,
  formatRange,
  isSingleCell,
  parseRange,
  rangeDims,
  rangesOverlap,
  type RangeRef,
} from "./a1";
import { scanFormula, unknownPlaceholders } from "./formula";
import { CopilotResponseSchema, type Action, type CopilotResponse, type Style } from "./schema";
import type { WorkbookContext } from "./types";

export interface ValidationLimits {
  /** Max cells a single action may touch. */
  maxCellsPerAction: number;
  /** Max cells all actions together may touch. */
  maxTotalCells: number;
  /** Max number of actions in one response. */
  maxActions: number;
}

export const DEFAULT_LIMITS: ValidationLimits = {
  maxCellsPerAction: 2_000,
  maxTotalCells: 10_000,
  maxActions: 25,
};

export interface ValidationIssue {
  severity: "error" | "warning";
  /** Index into response.actions, or null for response-level problems. */
  actionIndex: number | null;
  message: string;
}

/** An action that passed validation, with its target resolved to a concrete sheet. */
export interface ValidatedAction {
  index: number;
  action: Action;
  /** Resolved target (sheet always set). Explain actions without a target have none. */
  range: RangeRef | null;
}

export interface ValidationResult {
  /** True when there are no errors (warnings are allowed). */
  ok: boolean;
  response: CopilotResponse | null;
  actions: ValidatedAction[];
  issues: ValidationIssue[];
}

const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;

export function validateResponse(
  raw: unknown,
  ctx: Pick<WorkbookContext, "sheetNames" | "activeSheet">,
  limits: ValidationLimits = DEFAULT_LIMITS,
): ValidationResult {
  const issues: ValidationIssue[] = [];

  // ---- Layer 1: shape ------------------------------------------------------
  const parsed = CopilotResponseSchema.safeParse(raw);
  if (!parsed.success) {
    for (const zi of parsed.error.issues) {
      const [first, idx] = zi.path;
      const actionIndex = first === "actions" && typeof idx === "number" ? idx : null;
      const where = zi.path.length ? zi.path.join(".") : "(root)";
      issues.push({ severity: "error", actionIndex, message: `${where}: ${zi.message}` });
    }
    return { ok: false, response: null, actions: [], issues };
  }
  const response = parsed.data;

  // ---- Layer 2: meaning ----------------------------------------------------
  if (response.actions.length > limits.maxActions) {
    issues.push({
      severity: "error",
      actionIndex: null,
      message: `Too many actions (${response.actions.length}); the limit is ${limits.maxActions}`,
    });
  }

  const validated: ValidatedAction[] = [];
  response.actions.forEach((action, index) => {
    const add = (severity: ValidationIssue["severity"], message: string) =>
      issues.push({ severity, actionIndex: index, message });
    const range = resolveTarget(action, ctx, add);
    if (range === undefined) return; // address error already reported
    checkAction(action, range, limits, add);
    validated.push({ index, action, range });
  });

  // Cross-action checks: total size and overlapping writes.
  const writes = validated.filter((v) => v.range && v.action.type !== "explain");
  const total = writes.reduce((n, v) => n + cellCount(v.range!), 0);
  if (total > limits.maxTotalCells) {
    issues.push({
      severity: "error",
      actionIndex: null,
      message: `Actions touch ${total} cells in total; the limit is ${limits.maxTotalCells}`,
    });
  }
  for (let i = 0; i < writes.length; i++) {
    for (let j = i + 1; j < writes.length; j++) {
      const a = writes[i]!;
      const b = writes[j]!;
      const bothContent = a.action.type !== "format" && b.action.type !== "format";
      if (bothContent && rangesOverlap(a.range!, b.range!)) {
        issues.push({
          severity: "warning",
          actionIndex: b.index,
          message: `Overwrites cells already written by action ${a.index + 1}; the later action wins`,
        });
      }
    }
  }

  const ok = !issues.some((i) => i.severity === "error");
  return { ok, response, actions: ok ? validated : [], issues };
}

/** Returns the resolved range, null (explain with no target) or undefined (error). */
function resolveTarget(
  action: Action,
  ctx: Pick<WorkbookContext, "sheetNames" | "activeSheet">,
  add: (s: ValidationIssue["severity"], m: string) => void,
): RangeRef | null | undefined {
  const address =
    action.type === "write_formula" ? action.cell : action.type === "explain" ? action.target : action.range;
  if (address === undefined) return null;
  let range: RangeRef;
  try {
    range = parseRange(address);
  } catch (e) {
    add("error", `Invalid address "${address}": ${(e as Error).message}`);
    return undefined;
  }
  const sheet = range.sheet ?? ctx.activeSheet;
  const match = ctx.sheetNames.find((s) => s.toLowerCase() === sheet.toLowerCase());
  if (!match) {
    add("error", `Sheet "${sheet}" does not exist (sheets: ${ctx.sheetNames.join(", ")})`);
    return undefined;
  }
  return { ...range, sheet: match };
}

function checkAction(
  action: Action,
  range: RangeRef | null,
  limits: ValidationLimits,
  add: (s: ValidationIssue["severity"], m: string) => void,
): void {
  if (range && action.type !== "explain" && cellCount(range) > limits.maxCellsPerAction) {
    add("error", `Range ${formatRange(range)} has ${cellCount(range)} cells; the limit is ${limits.maxCellsPerAction}`);
  }

  switch (action.type) {
    case "write_formula": {
      if (range && !isSingleCell(range)) {
        add("error", `write_formula targets one cell, but "${action.cell}" is a range; use fill_range instead`);
      }
      checkFormula(action.formula, range, add);
      break;
    }

    case "fill_range": {
      const hasValues = action.values !== undefined;
      const hasTemplate = action.formulaTemplate !== undefined;
      if (hasValues === hasTemplate) {
        add("error", "fill_range needs exactly one of `values` or `formulaTemplate`");
        break;
      }
      if (hasValues && range) {
        const { rows, cols } = rangeDims(range);
        const v = action.values!;
        if (v.length !== rows || v.some((row) => row.length !== cols)) {
          const got = `${v.length}x${v[0]?.length ?? 0}`;
          add("error", `values is ${got} but range ${formatRange(range)} is ${rows}x${cols}`);
        }
        v.flat().forEach((cell) => {
          if (typeof cell === "string" && cell.startsWith("=")) checkFormula(cell, null, add);
          if (typeof cell === "number" && !Number.isFinite(cell)) add("error", "Numbers must be finite");
        });
      }
      if (hasTemplate) {
        const t = action.formulaTemplate!;
        if (!t.startsWith("=")) add("error", 'formulaTemplate must start with "="');
        const unknown = unknownPlaceholders(t);
        if (unknown.length) add("error", `Unknown placeholder(s) in formulaTemplate: ${unknown.map((u) => `{${u}}`).join(", ")}`);
        // Check a concrete expansion so the formula checks see real addresses.
        checkFormula(t.replace(/\{row\}/g, "1").replace(/\{col\}/g, "A"), null, add);
      }
      break;
    }

    case "format":
      checkStyle(action.style, add);
      break;

    case "explain":
      if (!action.text.trim()) add("error", "Explanation text is empty");
      break;
  }
}

function checkFormula(
  formula: string,
  target: RangeRef | null,
  add: (s: ValidationIssue["severity"], m: string) => void,
): void {
  const scan = scanFormula(formula);
  scan.errors.forEach((e) => add("error", `${formula}: ${e}`));
  scan.warnings.forEach((w) => add("warning", `${formula}: ${w}`));
  if (target) {
    const self = scan.references.some((ref) =>
      rangesOverlap({ ...ref, sheet: ref.sheet ?? target.sheet }, target),
    );
    if (self) add("warning", `${formula} refers to its own cell (circular reference)`);
  }
}

function checkStyle(style: Style, add: (s: ValidationIssue["severity"], m: string) => void): void {
  if (Object.keys(style).length === 0) add("error", "format.style is empty");
  for (const key of ["fontColor", "fillColor"] as const) {
    const c = style[key];
    if (c !== undefined && !HEX_COLOR.test(c)) add("error", `${key} must be a hex colour like #1F4E79, got "${c}"`);
  }
  if (style.numberFormat !== undefined && (style.numberFormat.length === 0 || style.numberFormat.length > 255)) {
    add("error", "numberFormat must be 1-255 characters");
  }
}

/** Turn issues into a short bullet list (used in the repair prompt and the UI). */
export function describeIssues(issues: ValidationIssue[]): string {
  return issues
    .map((i) => `- [${i.severity}]${i.actionIndex === null ? "" : ` action ${i.actionIndex + 1}:`} ${i.message}`)
    .join("\n");
}
