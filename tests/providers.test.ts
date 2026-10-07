import { AnthropicProvider } from "../src/providers/anthropic";
import { GeminiProvider } from "../src/providers/gemini";
import { ProviderError, type CompletionRequest } from "../src/providers/types";

const request: CompletionRequest = {
  system: "SYS",
  messages: [
    { role: "user", content: "hi" },
    { role: "assistant", content: "{}" },
    { role: "user", content: "again" },
  ],
  jsonSchema: { type: "object" },
};

/** A fetch stub that records calls and returns a canned JSON response. */
function stubFetch(status: number, body: unknown) {
  const calls: { url: string; init: RequestInit }[] = [];
  const fn = (async (input: RequestInfo | URL, init?: RequestInit) => {
    calls.push({ url: String(input instanceof Request ? input.url : input), init: init ?? {} });
    return new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });
  }) as typeof fetch;
  return { fn, calls };
}

const headerOf = (init: RequestInit, name: string) => new Headers(init.headers).get(name);

describe("GeminiProvider", () => {
  const ok = { candidates: [{ content: { parts: [{ text: '{"summary":"s",' }, { text: '"actions":[]}' }] }, finishReason: "STOP" }] };

  it("sends system, mapped roles and the JSON schema; joins text parts", async () => {
    const { fn, calls } = stubFetch(200, ok);
    const text = await new GeminiProvider({ apiKey: "k", model: "m1", fetch: fn }).complete(request);
    expect(text).toBe('{"summary":"s","actions":[]}');
    expect(calls[0]!.url).toBe("https://generativelanguage.googleapis.com/v1beta/models/m1:generateContent");
    expect(headerOf(calls[0]!.init, "x-goog-api-key")).toBe("k");
    const body = JSON.parse(String(calls[0]!.init.body));
    expect(body.systemInstruction.parts[0].text).toBe("SYS");
    expect(body.contents.map((c: { role: string }) => c.role)).toEqual(["user", "model", "user"]);
    expect(body.generationConfig).toMatchObject({ responseMimeType: "application/json", responseJsonSchema: { type: "object" } });
  });

  it.each([
    [400, { error: { message: "API key not valid" } }, "auth"],
    [429, { error: { message: "quota" } }, "rate_limit"],
    [500, { error: { message: "oops" } }, "other"],
    [200, { promptFeedback: { blockReason: "SAFETY" } }, "refusal"],
    [200, { candidates: [{ finishReason: "MAX_TOKENS", content: { parts: [{ text: "{" }] } }] }, "truncated"],
    [200, { candidates: [] }, "bad_response"],
  ])("maps HTTP %i to a %s-style error", async (status, body, kind) => {
    const { fn } = stubFetch(status, body);
    await expect(new GeminiProvider({ apiKey: "k", fetch: fn }).complete(request)).rejects.toMatchObject({ kind });
  });

  it("maps network failures", async () => {
    const fn = (async () => {
      throw new TypeError("Failed to fetch");
    }) as typeof fetch;
    const err = await new GeminiProvider({ apiKey: "k", fetch: fn }).complete(request).catch((e) => e);
    expect(err).toBeInstanceOf(ProviderError);
    expect(err.kind).toBe("network");
  });
});

describe("AnthropicProvider", () => {
  const message = (overrides: Record<string, unknown> = {}) => ({
    id: "msg_1",
    type: "message",
    role: "assistant",
    model: "claude-opus-5-5",
    content: [{ type: "text", text: '{"summary":"s","actions":[]}' }],
    stop_reason: "end_turn",
    stop_sequence: null,
    usage: { input_tokens: 1, output_tokens: 1 },
    ...overrides,
  });

  it("sends structured-output config and returns the text", async () => {
    const { fn, calls } = stubFetch(200, message());
    const p = new AnthropicProvider({ apiKey: "sk-test", fetch: fn, maxRetries: 0 });
    expect(p.name).toBe("anthropic:claude-opus-5-5");
    expect(await p.complete(request)).toBe('{"summary":"s","actions":[]}');

    const { url, init } = calls[0]!;
    expect(url).toContain("/v1/messages");
    expect(headerOf(init, "x-api-key")).toBe("sk-test");
    expect(headerOf(init, "anthropic-beta")).toContain("server-side-fallback-2026-07-01");
    const body = JSON.parse(String(init.body));
    expect(body).toMatchObject({
      model: "claude-opus-5-5",
      system: "SYS",
      fallbacks: "default",
      output_config: { format: { type: "json_schema", schema: { type: "object" } } },
    });
    expect(body.messages).toHaveLength(3);
  });

  it.each([
    ["refusal", "refusal"],
    ["max_tokens", "truncated"],
  ])("maps stop_reason %s", async (stop_reason, kind) => {
    const { fn } = stubFetch(200, message({ stop_reason }));
    const p = new AnthropicProvider({ apiKey: "k", fetch: fn, maxRetries: 0 });
    await expect(p.complete(request)).rejects.toMatchObject({ kind });
  });

  it.each([
    [401, "auth"],
    [429, "rate_limit"],
    [500, "other"],
  ])("maps HTTP %i", async (status, kind) => {
    const { fn } = stubFetch(status, { type: "error", error: { type: "x", message: "nope" } });
    const p = new AnthropicProvider({ apiKey: "k", fetch: fn, maxRetries: 0 });
    await expect(p.complete(request)).rejects.toMatchObject({ kind, status });
  });
});
