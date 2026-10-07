// The orchestrator: glues provider, validator, planner and workbook together.
//
//   propose()  request -> LLM -> parse -> validate (-> repair) -> snapshot -> plan
//   apply()    re-read cells, refuse if stale, snapshot for undo, write
//   undo()     check for later edits (drift), then restore the snapshot
//
// Nothing here imports Office.js or React, so the whole flow runs in Vitest.

import type { WorkbookGateway } from "./gateway";
import { buildPlan, buildUndoOps, rangesToSnapshot, snapshotDifferences, type Plan } from "./planner";
import { buildExplainRequest, buildRepairMessage, buildUserMessage, SYSTEM_PROMPT } from "./prompt";
import { responseJsonSchema } from "./schema";
import type { Snapshot, WriteOp } from "./types";
import { describeIssues, validateResponse, type ValidationIssue, type ValidationLimits, DEFAULT_LIMITS } from "./validator";
import type { ChatMessage, LLMProvider } from "../providers/types";

export interface Proposal {
  plan: Plan;
  /** Cell states when the preview was built. Used to detect a stale preview. */
  snapshot: Snapshot;
  /** Warnings that survived validation (errors would have stopped us). */
  warnings: ValidationIssue[];
  /** How many LLM calls it took (1 = first try was valid). */
  attempts: number;
}

export class ValidationFailedError extends Error {
  constructor(
    readonly issues: ValidationIssue[],
    readonly attempts: number,
    readonly lastReply: string,
  ) {
    super(`The model's reply failed validation after ${attempts} attempt(s):\n${describeIssues(issues)}`);
    this.name = "ValidationFailedError";
  }
}

export class StalePreviewError extends Error {
  constructor(readonly cells: string[]) {
    super(`These cells changed after the preview was made: ${cells.slice(0, 10).join(", ")}${cells.length > 10 ? "…" : ""}. Refresh the preview.`);
    this.name = "StalePreviewError";
  }
}

export interface AppliedPlan {
  plan: Plan;
  /** Writes that restore the snapshot taken immediately before applying. */
  undoOps: WriteOp[];
  /**
   * The cells as Excel reports them right after applying. Excel normalises what
   * we write (e.g. "=sum(a1)" becomes "=SUM(A1)"), so drift is measured against
   * this read, not against what the plan intended to write.
   */
  afterSnapshot: Snapshot;
  appliedAt: Date;
}

export interface ProposeOptions {
  /** Extra LLM calls allowed to fix an invalid reply. */
  maxRepairs?: number;
  limits?: ValidationLimits;
  signal?: AbortSignal;
}

/** Pull a JSON object out of a reply, tolerating ```json fences or stray prose. */
export function extractJson(text: string): unknown {
  const trimmed = text.trim();
  try {
    return JSON.parse(trimmed);
  } catch {
    const fenced = /```(?:json)?\s*([\s\S]*?)```/.exec(trimmed);
    const candidate = fenced ? fenced[1]! : trimmed.slice(trimmed.indexOf("{"), trimmed.lastIndexOf("}") + 1);
    try {
      return JSON.parse(candidate);
    } catch {
      return undefined;
    }
  }
}

export class Copilot {
  constructor(
    private provider: LLMProvider,
    private workbook: WorkbookGateway,
  ) {}

  async propose(request: string, opts: ProposeOptions = {}): Promise<Proposal> {
    const maxRepairs = opts.maxRepairs ?? 1;
    const limits = opts.limits ?? DEFAULT_LIMITS;
    const ctx = await this.workbook.getContext();
    const messages: ChatMessage[] = [{ role: "user", content: buildUserMessage(request, ctx) }];

    for (let attempt = 1; ; attempt++) {
      const reply = await this.provider.complete({
        system: SYSTEM_PROMPT,
        messages,
        jsonSchema: responseJsonSchema(),
        ...(opts.signal ? { signal: opts.signal } : {}),
      });
      const json = extractJson(reply);
      const result =
        json === undefined
          ? {
              ok: false,
              response: null,
              actions: [],
              issues: [{ severity: "error" as const, actionIndex: null, message: "Reply was not valid JSON" }],
            }
          : validateResponse(json, ctx, limits);

      if (result.ok && result.response) {
        const snapshot = await this.workbook.readCells(rangesToSnapshot(result.actions));
        const plan = buildPlan(result.response.summary, result.actions, snapshot);
        return { plan, snapshot, warnings: result.issues, attempts: attempt };
      }

      if (attempt > maxRepairs) throw new ValidationFailedError(result.issues, attempt, reply);
      // Show the model its own reply and what was wrong with it, then ask again.
      messages.push({ role: "assistant", content: reply });
      messages.push({ role: "user", content: buildRepairMessage(describeIssues(result.issues)) });
    }
  }

  async apply(proposal: Proposal): Promise<AppliedPlan> {
    const { plan } = proposal;
    // Fresh read: this is both the stale-preview check and the undo snapshot.
    const now = await this.workbook.readCells(plan.touchedRanges);
    const changed = snapshotDifferences(proposal.snapshot, now);
    if (changed.length) throw new StalePreviewError(changed);
    const undoOps = buildUndoOps(plan.touchedRanges, now);
    await this.workbook.applyOps(plan.ops);
    const afterSnapshot = await this.workbook.readCells(plan.touchedRanges);
    return { plan, undoOps, afterSnapshot, appliedAt: new Date() };
  }

  /** Keys ("Sheet1!B2") of cells edited after `applied` — undo would overwrite these. */
  async checkDrift(applied: AppliedPlan): Promise<string[]> {
    const current = await this.workbook.readCells(applied.plan.touchedRanges);
    return snapshotDifferences(applied.afterSnapshot, current);
  }

  async undo(applied: AppliedPlan): Promise<void> {
    await this.workbook.applyOps(applied.undoOps);
  }

  /**
   * "Explain this formula" for the selected cell. Returns the explanation text,
   * or a friendly message when there is nothing to explain.
   */
  async explainSelection(opts: { signal?: AbortSignal } = {}): Promise<{ address: string; formula: string; text: string }> {
    const ctx = await this.workbook.getContext();
    const sel = ctx.selection;
    if (!sel) throw new Error("Select a cell first.");
    const cell = sel.cells[0]?.[0];
    const address = `${sel.sheet}!${sel.address.split(":")[0]}`;
    if (!cell?.formula) throw new Error(`${address} doesn't contain a formula.`);

    const proposal = await this.propose(buildExplainRequest(address, cell.formula), {
      maxRepairs: 1,
      ...(opts.signal ? { signal: opts.signal } : {}),
    });
    const text = proposal.plan.explanations.map((e) => e.text).join("\n\n");
    if (!text) throw new Error("The model did not return an explanation.");
    // Ignore any write actions the model added despite instructions: explain never edits.
    return { address, formula: cell.formula, text };
  }
}
