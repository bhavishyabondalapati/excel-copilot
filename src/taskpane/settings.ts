// User settings (provider, model, API keys), kept in the browser's localStorage
// for this add-in's origin. Keys never go into source files or the build.

import { DEFAULT_ANTHROPIC_MODEL } from "../providers/anthropic";
import { DEFAULT_GEMINI_MODEL } from "../providers/gemini";

export type ProviderId = "anthropic" | "gemini" | "demo";

export interface Settings {
  provider: ProviderId;
  anthropicKey: string;
  anthropicModel: string;
  geminiKey: string;
  geminiModel: string;
}

export const DEFAULT_SETTINGS: Settings = {
  provider: "demo",
  anthropicKey: "",
  anthropicModel: DEFAULT_ANTHROPIC_MODEL,
  geminiKey: "",
  geminiModel: DEFAULT_GEMINI_MODEL,
};

const STORAGE_KEY = "excel-copilot.settings.v1";

export function loadSettings(storage: Pick<Storage, "getItem"> | undefined = safeStorage()): Settings {
  try {
    const raw = storage?.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const parsed = JSON.parse(raw) as Partial<Settings>;
    return { ...DEFAULT_SETTINGS, ...parsed };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(settings: Settings, storage: Pick<Storage, "setItem"> | undefined = safeStorage()): void {
  try {
    storage?.setItem(STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // Storage can be unavailable (private mode, blocked). Settings then last for this session only.
  }
}

function safeStorage(): Storage | undefined {
  try {
    return globalThis.localStorage;
  } catch {
    return undefined;
  }
}

/** Why the chosen provider can't be used yet, or null if it's ready. */
export function settingsProblem(s: Settings): string | null {
  if (s.provider === "anthropic" && !s.anthropicKey.trim()) return "Add an Anthropic API key in Settings.";
  if (s.provider === "gemini" && !s.geminiKey.trim()) return "Add a Gemini API key in Settings.";
  return null;
}
