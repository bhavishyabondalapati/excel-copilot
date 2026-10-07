import { DEFAULT_SETTINGS, loadSettings, saveSettings, settingsProblem } from "../src/taskpane/settings";

function memoryStorage() {
  const data = new Map<string, string>();
  return { getItem: (k: string) => data.get(k) ?? null, setItem: (k: string, v: string) => void data.set(k, v) };
}

describe("settings", () => {
  it("round-trips through storage and fills in defaults", () => {
    const s = memoryStorage();
    saveSettings({ ...DEFAULT_SETTINGS, provider: "gemini", geminiKey: "k" }, s);
    const loaded = loadSettings(s);
    expect(loaded.provider).toBe("gemini");
    expect(loaded.anthropicModel).toBe(DEFAULT_SETTINGS.anthropicModel);
  });

  it("falls back to defaults on corrupt or missing storage", () => {
    expect(loadSettings({ getItem: () => "{not json" })).toEqual(DEFAULT_SETTINGS);
    expect(loadSettings(undefined)).toEqual(DEFAULT_SETTINGS);
    expect(() => saveSettings(DEFAULT_SETTINGS, { setItem: () => { throw new Error("blocked"); } })).not.toThrow();
  });

  it("explains what is missing", () => {
    expect(DEFAULT_SETTINGS.provider).toBe("gemini");
    expect(settingsProblem(DEFAULT_SETTINGS)).toContain("Gemini API key"); // default needs a key
    expect(settingsProblem({ ...DEFAULT_SETTINGS, provider: "demo" })).toBeNull(); // demo needs nothing
    expect(settingsProblem({ ...DEFAULT_SETTINGS, provider: "anthropic" })).toContain("Anthropic API key");
    expect(settingsProblem({ ...DEFAULT_SETTINGS, provider: "gemini", geminiKey: "  " })).toContain("Gemini");
  });
});
