import { createDemoProvider } from "../demo/demoProvider";
import { AnthropicProvider } from "../providers/anthropic";
import { GeminiProvider } from "../providers/gemini";
import type { LLMProvider } from "../providers/types";
import type { Settings } from "./settings";

export function createProvider(s: Settings): LLMProvider {
  switch (s.provider) {
    case "anthropic":
      return new AnthropicProvider({ apiKey: s.anthropicKey.trim(), model: s.anthropicModel.trim() });
    case "gemini":
      return new GeminiProvider({ apiKey: s.geminiKey.trim(), model: s.geminiModel.trim() });
    case "demo":
      return createDemoProvider();
  }
}
