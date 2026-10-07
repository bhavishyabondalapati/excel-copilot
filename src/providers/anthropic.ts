// Claude via the official Anthropic TypeScript SDK.
//
// We ask for structured output (`output_config.format` with our JSON Schema), so
// the API constrains the reply to the schema. We still validate afterwards:
// the schema can't express rules like "values must match the range size".

import Anthropic from "@anthropic-ai/sdk";
import { ProviderError, type CompletionRequest, type LLMProvider } from "./types";

export const DEFAULT_ANTHROPIC_MODEL = "claude-opus-5-5";

export interface AnthropicProviderOptions {
  apiKey: string;
  model?: string;
  /** Injected in tests to avoid real network calls. */
  fetch?: typeof globalThis.fetch;
  maxRetries?: number;
}

export class AnthropicProvider implements LLMProvider {
  readonly name: string;
  private client: Anthropic;
  private model: string;

  constructor(opts: AnthropicProviderOptions) {
    this.model = opts.model || DEFAULT_ANTHROPIC_MODEL;
    this.name = `anthropic:${this.model}`;
    this.client = new Anthropic({
      apiKey: opts.apiKey,
      // The task pane is a web page, so the key lives in the user's browser.
      // Fine for a personal add-in; a shared deployment should proxy via a server.
      dangerouslyAllowBrowser: true,
      maxRetries: opts.maxRetries ?? 2,
      ...(opts.fetch ? { fetch: opts.fetch } : {}),
    });
  }

  async complete(req: CompletionRequest): Promise<string> {
    let response;
    try {
      response = await this.client.beta.messages.create(
        {
          model: this.model,
          max_tokens: 16000,
          system: req.system,
          messages: req.messages,
          output_config: {
            effort: "medium",
            format: { type: "json_schema", schema: req.jsonSchema },
          },
          // If a safety classifier declines, let the API retry on its recommended fallback model.
          betas: ["server-side-fallback-2026-07-01"],
          fallbacks: "default",
        },
        { signal: req.signal },
      );
    } catch (e) {
      throw toProviderError(e);
    }

    if (response.stop_reason === "refusal") {
      throw new ProviderError("Claude declined this request.", "refusal");
    }
    if (response.stop_reason === "max_tokens") {
      throw new ProviderError("The reply was cut off (max_tokens). Try a smaller request.", "truncated");
    }
    const text = response.content
      .filter((b) => b.type === "text")
      .map((b) => (b as { text: string }).text)
      .join("");
    if (!text) throw new ProviderError("Claude returned no text.", "bad_response");
    return text;
  }
}

function toProviderError(e: unknown): ProviderError {
  if (e instanceof Anthropic.AuthenticationError || e instanceof Anthropic.PermissionDeniedError) {
    return new ProviderError("Anthropic rejected the API key. Check it in Settings.", "auth", e.status);
  }
  if (e instanceof Anthropic.RateLimitError) {
    return new ProviderError("Anthropic rate limit hit. Wait a moment and retry.", "rate_limit", 429);
  }
  if (e instanceof Anthropic.APIConnectionError) {
    return new ProviderError(`Could not reach Anthropic: ${e.message}`, "network");
  }
  if (e instanceof Anthropic.APIError) {
    return new ProviderError(`Anthropic API error: ${e.message}`, "other", e.status);
  }
  return new ProviderError(String((e as Error)?.message ?? e), "other");
}
