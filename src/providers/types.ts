// The provider interface: the only thing the rest of the app knows about LLMs.
// Swapping Gemini for Anthropic (or for a fake in tests) means swapping the
// object that implements this — nothing else changes. This is the "strategy"
// pattern.

export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
}

export interface CompletionRequest {
  /** Instructions that frame the whole conversation. */
  system: string;
  /** Conversation so far; the last message is from the user. */
  messages: ChatMessage[];
  /** JSON Schema the reply must match (providers use their structured-output mode). */
  jsonSchema: Record<string, unknown>;
  signal?: AbortSignal;
}

export interface LLMProvider {
  /** Short label for the UI and logs, e.g. "anthropic:claude-opus-5-5". */
  readonly name: string;
  /** Returns the model's raw text reply (expected to be JSON). */
  complete(request: CompletionRequest): Promise<string>;
}

/** A provider failure with enough detail for the UI to say something useful. */
export class ProviderError extends Error {
  constructor(
    message: string,
    readonly kind: "auth" | "rate_limit" | "refusal" | "truncated" | "network" | "bad_response" | "other",
    readonly status?: number,
  ) {
    super(message);
    this.name = "ProviderError";
  }
}
