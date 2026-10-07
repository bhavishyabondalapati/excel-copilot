// A keyword-driven stand-in for an LLM, tailored to the demo workbook.
// It exists so the full preview -> apply -> undo loop can be tried with no API
// key. It is NOT smart: it recognises a few phrases and returns canned actions.

import { FakeProvider } from "../providers/fake";
import type { CopilotResponse } from "../core/schema";
import { scanFormula } from "../core/formula";
import { formatRange } from "../core/a1";

export const DEMO_PROMPTS = [
  "Fill in the Revenue column",
  "Add a total row",
  "Make the header bold with a yellow fill",
  "Format prices as currency",
];

function respond(userText: string): CopilotResponse {
  // Only look at the request, not the workbook dump.
  const original = /<request>([\s\S]*)<\/request>/.exec(userText)?.[1] ?? userText;
  const request = original.toLowerCase();

  const explainMatch = /explain the formula in (\S+) in plain english[\s\S]*?formula: (=[^\n]*)/i.exec(original);
  if (explainMatch) {
    const formula = explainMatch[2]!.trim();
    const scan = scanFormula(formula);
    const fns = scan.functions.length ? `uses ${scan.functions.join(", ")}` : "uses no functions";
    const refs = scan.references.length ? ` and reads ${scan.references.map((r) => formatRange(r, true)).join(", ")}` : "";
    return {
      summary: "Explanation of the selected formula",
      actions: [
        {
          type: "explain",
          target: explainMatch[1]!,
          text: `(Demo mode) ${formula} ${fns}${refs}. Connect Anthropic or Gemini in Settings for a real plain-English explanation.`,
        },
      ],
    };
  }
  if (/revenue/.test(request)) {
    return {
      summary: "Revenue = Units × Unit price for each region.",
      actions: [{ type: "fill_range", range: "D2:D5", formulaTemplate: "=B{row}*C{row}" }],
    };
  }
  if (/total/.test(request)) {
    return {
      summary: "Adds a Total row under the data, summing Units and Revenue.",
      actions: [
        { type: "fill_range", range: "A6:D6", values: [["Total", "=SUM(B2:B5)", null, "=SUM(D2:D5)"]] },
        { type: "format", range: "A6:D6", style: { bold: true } },
      ],
    };
  }
  if (/header|bold|yellow/.test(request)) {
    return {
      summary: "Bold header row with a light yellow background.",
      actions: [{ type: "format", range: "A1:D1", style: { bold: true, fillColor: "#FFF2CC" } }],
    };
  }
  if (/currency|price|\$/.test(request)) {
    return {
      summary: "Shows prices and revenue as dollars.",
      actions: [{ type: "format", range: "C2:D7", style: { numberFormat: "$#,##0.00" } }],
    };
  }
  return {
    summary: "Demo mode didn't recognise that request.",
    actions: [
      {
        type: "explain",
        text: `Demo mode only understands a few phrases, e.g. "${DEMO_PROMPTS.join('", "')}". Choose Anthropic or Gemini in Settings for real requests.`,
      },
    ],
  };
}

export function createDemoProvider(): FakeProvider {
  const provider = new FakeProvider((req) => JSON.stringify(respond(req.messages.at(-1)?.content ?? "")));
  provider.name = "demo";
  return provider;
}
