import { useState } from "react";
import type { ProviderId, Settings } from "../settings";

const PROVIDERS: { id: ProviderId; label: string }[] = [
  { id: "gemini", label: "Google Gemini (default)" },
  { id: "anthropic", label: "Anthropic Claude" },
  { id: "demo", label: "Demo (no key, canned answers)" },
];

export function SettingsPanel({ settings, onSave, onClose }: {
  settings: Settings;
  onSave: (s: Settings) => void;
  onClose: () => void;
}) {
  const [draft, setDraft] = useState(settings);
  const set = <K extends keyof Settings>(k: K, v: Settings[K]) => setDraft({ ...draft, [k]: v });

  return (
    <section className="card settings" aria-label="Settings">
      <h2>Settings</h2>
      <label>
        Provider
        <select value={draft.provider} onChange={(e) => set("provider", e.target.value as ProviderId)}>
          {PROVIDERS.map((p) => (
            <option key={p.id} value={p.id}>{p.label}</option>
          ))}
        </select>
      </label>

      {draft.provider === "anthropic" && (
        <>
          <label>
            Anthropic API key
            <input type="password" autoComplete="off" value={draft.anthropicKey} placeholder="sk-ant-…"
              onChange={(e) => set("anthropicKey", e.target.value)} />
          </label>
          <label>
            Model
            <input value={draft.anthropicModel} onChange={(e) => set("anthropicModel", e.target.value)} />
          </label>
        </>
      )}
      {draft.provider === "gemini" && (
        <>
          <label>
            Gemini API key
            <input type="password" autoComplete="off" value={draft.geminiKey}
              onChange={(e) => set("geminiKey", e.target.value)} />
          </label>
          <label>
            Model
            <input value={draft.geminiModel} onChange={(e) => set("geminiModel", e.target.value)} />
          </label>
        </>
      )}
      <p className="muted small">
        Keys are stored only in this browser's local storage for the add-in and sent only to the provider you pick.
      </p>
      <div className="row">
        <button className="primary" onClick={() => onSave(draft)}>Save</button>
        <button onClick={onClose}>Cancel</button>
      </div>
    </section>
  );
}
