// A provider that never touches the network. Tests script its replies; it also
// records every request so tests can assert on the prompt we built.

import type { CompletionRequest, LLMProvider } from "./types";

type Reply = string | ((request: CompletionRequest) => string) | Error;

export class FakeProvider implements LLMProvider {
  readonly name = "fake";
  readonly requests: CompletionRequest[] = [];
  private queue: Reply[];

  /** Replies are used in order; the last one repeats once the queue is drained. */
  constructor(...replies: Reply[]) {
    if (replies.length === 0) throw new Error("FakeProvider needs at least one reply");
    this.queue = replies;
  }

  async complete(request: CompletionRequest): Promise<string> {
    this.requests.push(request);
    const reply = this.queue.length > 1 ? this.queue.shift()! : this.queue[0]!;
    if (reply instanceof Error) throw reply;
    if (typeof reply === "function") return reply(request);
    return reply;
  }
}

/** Convenience: a FakeProvider that returns these objects as JSON. */
export function fakeJson(...objects: unknown[]): FakeProvider {
  return new FakeProvider(...objects.map((o) => JSON.stringify(o)));
}
