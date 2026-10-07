import { useMemo, useRef, useState } from "react";
import { Copilot, StalePreviewError, ValidationFailedError, type AppliedPlan, type Proposal } from "../core/copilot";
import type { WorkbookGateway } from "../core/gateway";
import type { MemoryWorkbook } from "../core/memoryWorkbook";
import type { CellChange } from "../core/planner";
import type { ValidationIssue } from "../core/validator";
import { DEMO_PROMPTS } from "../demo/demoProvider";
import { ProviderError } from "../providers/types";
import { DemoGrid } from "./components/DemoGrid";
import { IssueList } from "./components/IssueList";
import { PreviewTable } from "./components/PreviewTable";
import { SettingsPanel } from "./components/SettingsPanel";
import { createProvider } from "./providerFactory";
import { loadSettings, saveSettings, settingsProblem, type Settings } from "./settings";

export interface AppProps {
  workbook: WorkbookGateway;
  /** Set in browser demo mode so we can draw the fake sheet. */
  demoWorkbook?: MemoryWorkbook;
}

type Busy = null | "thinking" | "applying" | "undoing" | "explaining";

interface Explanation {
  address: string;
  formula: string;
  text: string;
}

export function App({ workbook, demoWorkbook }: AppProps) {
  const [settings, setSettings] = useState<Settings>(() => loadSettings());
  const [showSettings, setShowSettings] = useState(false);
  const [request, setRequest] = useState("");
  const [busy, setBusy] = useState<Busy>(null);
  const [error, setError] = useState<string | null>(null);
  const [failedIssues, setFailedIssues] = useState<ValidationIssue[]>([]);
  const [proposal, setProposal] = useState<Proposal | null>(null);
  const [history, setHistory] = useState<AppliedPlan[]>([]);
  const [drift, setDrift] = useState<CellChange[] | null>(null);
  const [explanation, setExplanation] = useState<Explanation | null>(null);
  const [, forceRender] = useState(0); // re-draw the demo grid after writes
  const abortRef = useRef<AbortController | null>(null);

  const copilot = useMemo(() => new Copilot(createProvider(settings), workbook), [settings, workbook]);
  const problem = settingsProblem(settings);

  /** Run an async step with shared busy/error handling. */
  async function run(kind: Exclude<Busy, null>, step: (signal: AbortSignal) => Promise<void>) {
    const controller = new AbortController();
    abortRef.current = controller;
    setBusy(kind);
    setError(null);
    setFailedIssues([]);
    try {
      await step(controller.signal);
    } catch (e) {
      if (controller.signal.aborted) return;
      if (e instanceof ValidationFailedError) {
        setError("The model's answer didn't pass validation, so nothing was changed.");
        setFailedIssues(e.issues);
      } else if (e instanceof ProviderError || e instanceof StalePreviewError) {
        setError(e.message);
      } else {
        setError((e as Error)?.message ?? String(e));
      }
    } finally {
      setBusy(null);
      abortRef.current = null;
      forceRender((n) => n + 1);
    }
  }

  const onPropose = () =>
    run("thinking", async (signal) => {
      setProposal(null);
      setExplanation(null);
      setProposal(await copilot.propose(request, { signal }));
    });

  const onApply = () =>
    run("applying", async () => {
      if (!proposal) return;
      try {
        const applied = await copilot.apply(proposal);
        setHistory((h) => [applied, ...h]);
        setProposal(null);
        setRequest("");
      } catch (e) {
        if (e instanceof StalePreviewError) setProposal(null);
        throw e;
      }
    });

  const onUndo = (force: boolean) =>
    run("undoing", async () => {
      const latest = history[0];
      if (!latest) return;
      if (!force) {
        const edited = await copilot.checkDrift(latest);
        if (edited.length) {
          setDrift(edited);
          return;
        }
      }
      await copilot.undo(latest);
      setHistory((h) => h.slice(1));
      setDrift(null);
    });

  const onExplain = () =>
    run("explaining", async (signal) => {
      setExplanation(null);
      setExplanation(await copilot.explainSelection({ signal }));
    });

  const onSaveSettings = (s: Settings) => {
    saveSettings(s);
    setSettings(s);
    setShowSettings(false);
  };

  const plan = proposal?.plan;
  const multiSheet = !!plan && new Set(plan.changes.map((c) => c.sheet)).size > 1;
  const pending = new Set(plan?.changes.map((c) => c.key) ?? []);

  return (
    <main>
      <header className="top">
        <h1>Excel Copilot</h1>
        <span className="muted small">{providerLabel(settings)}</span>
        <button className="link" onClick={() => setShowSettings((v) => !v)}>Settings</button>
      </header>

      {showSettings && (
        <SettingsPanel settings={settings} onSave={onSaveSettings} onClose={() => setShowSettings(false)} />
      )}

      {demoWorkbook && (
        <DemoGrid
          workbook={demoWorkbook}
          pending={pending}
          onSelect={(addr) => {
            demoWorkbook.selection = addr;
            forceRender((n) => n + 1);
          }}
        />
      )}

      <section className="card">
        <label htmlFor="request" className="label">What should I do?</label>
        <textarea
          id="request"
          rows={3}
          value={request}
          placeholder="e.g. Add a Revenue column = Units × Price, and bold the header"
          onChange={(e) => setRequest(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && (e.metaKey || e.ctrlKey) && request.trim() && !busy && !problem) onPropose();
          }}
        />
        {settings.provider === "demo" && (
          <div className="chips">
            {DEMO_PROMPTS.map((p) => (
              <button key={p} className="chip" onClick={() => setRequest(p)}>{p}</button>
            ))}
          </div>
        )}
        <div className="row">
          <button className="primary" disabled={!request.trim() || !!busy || !!problem} onClick={onPropose}>
            {busy === "thinking" ? "Thinking…" : "Preview changes"}
          </button>
          <button disabled={!!busy || !!problem} onClick={onExplain} title="Explain the formula in the selected cell">
            {busy === "explaining" ? "Explaining…" : "Explain selected formula"}
          </button>
          {(busy === "thinking" || busy === "explaining") && (
            <button className="link" onClick={() => abortRef.current?.abort()}>Cancel</button>
          )}
        </div>
        {problem && <p className="warning-text">{problem}</p>}
      </section>

      {error && (
        <section className="card error" role="alert">
          <p>{error}</p>
          <IssueList issues={failedIssues} />
        </section>
      )}

      {explanation && (
        <section className="card">
          <h2>About {explanation.address}</h2>
          <code className="formula block">{explanation.formula}</code>
          <p className="explain">{explanation.text}</p>
        </section>
      )}

      {plan && (
        <section className="card preview">
          <h2>Preview</h2>
          <p>{plan.summary}</p>
          {plan.explanations.map((e, i) => (
            <p key={i} className="explain">{e.target && <strong>{e.target}: </strong>}{e.text}</p>
          ))}
          <IssueList issues={proposal.warnings} />
          {plan.ops.length > 0 && (
            <>
              <p className="muted small">
                {plan.changes.length} cell{plan.changes.length === 1 ? "" : "s"} will change
                {proposal.attempts > 1 && ` · fixed after ${proposal.attempts - 1} retry`}
              </p>
              <PreviewTable changes={plan.changes} multiSheet={multiSheet} />
            </>
          )}
          <div className="row">
            {plan.ops.length > 0 && (
              <button className="primary" disabled={!!busy} onClick={onApply}>
                {busy === "applying" ? "Applying…" : "Apply"}
              </button>
            )}
            <button disabled={!!busy} onClick={() => setProposal(null)}>
              {plan.ops.length > 0 ? "Discard" : "Close"}
            </button>
          </div>
        </section>
      )}

      {history.length > 0 && (
        <section className="card">
          <h2>Applied</h2>
          <ol className="history">
            {history.map((h, i) => (
              <li key={h.appliedAt.getTime() + i}>
                <span>{h.plan.summary}</span>
                <span className="muted small"> · {h.plan.changes.length} cells · {h.appliedAt.toLocaleTimeString()}</span>
                {i === 0 && (
                  <button disabled={!!busy} onClick={() => onUndo(false)}>
                    {busy === "undoing" ? "Undoing…" : "Undo"}
                  </button>
                )}
              </li>
            ))}
          </ol>
          {drift && (
            <div className="warning-box" role="alert">
              <p>
                You edited {drift.length} of these cells after applying (
                {drift.slice(0, 5).map((d) => d.address).join(", ")}{drift.length > 5 ? "…" : ""}).
                Undo will overwrite those edits.
              </p>
              <div className="row">
                <button className="danger" onClick={() => onUndo(true)}>Undo anyway</button>
                <button onClick={() => setDrift(null)}>Keep my edits</button>
              </div>
            </div>
          )}
          <p className="muted small">Undo restores the most recent change first. Use this button: Excel's own Ctrl+Z may not reverse add-in edits.</p>
        </section>
      )}
    </main>
  );
}

function providerLabel(s: Settings): string {
  if (s.provider === "anthropic") return `Claude · ${s.anthropicModel}`;
  if (s.provider === "gemini") return `Gemini · ${s.geminiModel}`;
  return "Demo mode";
}
