// Gemini via its REST API (generateContent), using plain fetch.
//
// Docs: https://ai.google.dev/api/generate-content
// JSON mode: generationConfig.responseMimeType = "application/json" plus
// generationConfig.responseJsonSchema = <our JSON Schema>.

import { ProviderError, type CompletionRequest, type LLMProvider } from "./types";

export const DEFAULT_GEMINI_MODEL = "gemini-3.8-flash";
const BASE_URL = "https://generativelanguage.googleapis.com/v1beta";

export interface GeminiProviderOptions {
  apiKey: string;
  model?: string;
  fetch?: typeof globalThis.fetch;
}

interface GeminiResponse {
  candidates?: { content?: { parts?: { text?: string }[] }; finishReason?: string }[];
  promptFeedback?: { blockReason?: string };
  error?: { message?: string; status?: string };
}

export class GeminiProvider implements LLMProvider {
  readonly name: string;
  private model: string;
  private apiKey: string;
  private fetchImpl: typeof globalThis.fetch;

  constructor(opts: GeminiProviderOptions) {
    this.model = opts.model || DEFAULT_GEMINI_MODEL;
    this.name = `gemini:${this.model}`;
    this.apiKey = opts.apiKey;
    this.fetchImpl = opts.fetch ?? globalThis.fetch.bind(globalThis);
  }

  async complete(req: CompletionRequest): Promise<string> {
    const body = {
      systemInstruction: { parts: [{ text: req.system }] },
      // Gemini calls the assistant role "model".
      contents: req.messages.map((m) => ({
        role: m.role === "assistant" ? "model" : "user",
        parts: [{ text: m.content }],
      })),
      generationConfig: {
        responseMimeType: "application/json",
        responseJsonSchema: req.jsonSchema,
        temperature: 0.2,
      },
    };

    let res: Response;
    try {
      res = await this.fetchImpl(`${BASE_URL}/models/${encodeURIComponent(this.model)}:generateContent`, {
        method: "POST",
        // Key goes in a header, not the URL, so it doesn't end up in logs.
        headers: { "content-type": "application/json", "x-goog-api-key": this.apiKey },
        body: JSON.stringify(body),
        ...(req.signal ? { signal: req.signal } : {}),
      });
    } catch (e) {
      throw new ProviderError(`Could not reach Gemini: ${(e as Error).message}`, "network");
    }

    let data: GeminiResponse;
    try {
      data = (await res.json()) as GeminiResponse;
    } catch {
      throw new ProviderError(`Gemini returned a non-JSON response (HTTP ${res.status})`, "bad_response", res.status);
    }

    if (!res.ok) {
      const msg = data.error?.message ?? `HTTP ${res.status}`;
      if (res.status === 401 || res.status === 403 || (res.status === 400 && /API key/i.test(msg))) {
        throw new ProviderError("Gemini rejected the API key. Check it in Settings.", "auth", res.status);
      }
      if (res.status === 429) throw new ProviderError("Gemini rate limit hit. Wait a moment and retry.", "rate_limit", 429);
      throw new ProviderError(`Gemini API error: ${msg}`, "other", res.status);
    }

    if (data.promptFeedback?.blockReason) {
      throw new ProviderError(`Gemini blocked the prompt (${data.promptFeedback.blockReason}).`, "refusal");
    }
    const candidate = data.candidates?.[0];
    const finish = candidate?.finishReason;
    if (finish === "MAX_TOKENS") throw new ProviderError("The reply was cut off. Try a smaller request.", "truncated");
    if (finish && finish !== "STOP") throw new ProviderError(`Gemini stopped early (${finish}).`, "refusal");
    const text = (candidate?.content?.parts ?? []).map((p) => p.text ?? "").join("");
    if (!text) throw new ProviderError("Gemini returned no text.", "bad_response");
    return text;
  }
}
