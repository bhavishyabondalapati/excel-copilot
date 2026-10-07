// The contract between the LLM and the add-in.
//
// The LLM must answer with JSON shaped like `CopilotResponse`. We describe that
// shape once, with Zod, and derive two things from it:
//   1. a TypeScript type (compile-time safety for our own code), and
//   2. a JSON Schema (sent to the LLM so it knows what to produce).
// `.strict()` everywhere means unknown keys are rejected rather than ignored —
// if the model invents a field, we want to know.

import { z } from "zod";

/** A literal cell value. Strings starting with "=" are treated as formulas. */
export const CellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);
export type CellValue = z.infer<typeof CellValueSchema>;

const A1 = z
  .string()
  .describe('A1 address, optionally with sheet: "B2", "A1:C10", "Sheet2!A1:A5"');

export const WriteFormulaSchema = z
  .object({
    type: z.literal("write_formula"),
    cell: A1.describe('Single target cell, e.g. "D2" or "Sheet1!D2"'),
    formula: z.string().describe('Excel formula starting with "=", e.g. "=SUM(A2:A10)"'),
  })
  .strict();

export const FillRangeSchema = z
  .object({
    type: z.literal("fill_range"),
    range: A1.describe("Target block of cells"),
    values: z
      .array(z.array(CellValueSchema))
      .optional()
      .describe("2-D array (rows of columns) matching the range size exactly. Strings starting with = are formulas."),
    formulaTemplate: z
      .string()
      .optional()
      .describe(
        'One formula applied to every cell. "{row}" becomes the cell\'s row number and "{col}" its column letter, e.g. "=B{row}*C{row}". Use either values or formulaTemplate, never both.',
      ),
  })
  .strict();

export const HorizontalAlignmentSchema = z.enum(["General", "Left", "Center", "Right"]);

export const StyleSchema = z
  .object({
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    fontColor: z.string().optional().describe('Hex colour like "#1F4E79"'),
    fillColor: z.string().optional().describe('Hex background colour like "#FFF2CC"'),
    numberFormat: z.string().optional().describe('Excel number format, e.g. "0.00%", "$#,##0.00", "yyyy-mm-dd"'),
    horizontalAlignment: HorizontalAlignmentSchema.optional(),
  })
  .strict();
export type Style = z.infer<typeof StyleSchema>;

export const FormatSchema = z
  .object({
    type: z.literal("format"),
    range: A1,
    style: StyleSchema,
  })
  .strict();

export const ExplainSchema = z
  .object({
    type: z.literal("explain"),
    target: A1.optional().describe("Cell or range the explanation is about, if any"),
    text: z.string().describe("Plain-English explanation for the user"),
  })
  .strict();

export const ActionSchema = z.discriminatedUnion("type", [
  WriteFormulaSchema,
  FillRangeSchema,
  FormatSchema,
  ExplainSchema,
]);

export const CopilotResponseSchema = z
  .object({
    summary: z.string().describe("One or two sentences describing what the actions do"),
    actions: z.array(ActionSchema),
  })
  .strict();

export type WriteFormulaAction = z.infer<typeof WriteFormulaSchema>;
export type FillRangeAction = z.infer<typeof FillRangeSchema>;
export type FormatAction = z.infer<typeof FormatSchema>;
export type ExplainAction = z.infer<typeof ExplainSchema>;
export type Action = z.infer<typeof ActionSchema>;
export type ActionType = Action["type"];
export type CopilotResponse = z.infer<typeof CopilotResponseSchema>;

export const ACTION_TYPES: readonly ActionType[] = ["write_formula", "fill_range", "format", "explain"];

/**
 * JSON Schema for the LLM. Generated from the Zod schema so the two can never
 * drift apart. Providers' "structured output" modes accept a subset of JSON
 * Schema, so we normalise a couple of constructs:
 *   - `oneOf` -> `anyOf` (wider provider support; the discriminator keeps them exclusive)
 *   - `type: ["string", "null"]` -> `anyOf: [{type: "string"}, {type: "null"}]`
 *   - drop the `$schema` key (some APIs reject unknown top-level keys)
 */
export function responseJsonSchema(): Record<string, unknown> {
  const raw = z.toJSONSchema(CopilotResponseSchema, { target: "draft-7" }) as Record<string, unknown>;
  return normaliseJsonSchema(raw) as Record<string, unknown>;
}

function normaliseJsonSchema(node: unknown): unknown {
  if (Array.isArray(node)) return node.map(normaliseJsonSchema);
  if (node === null || typeof node !== "object") return node;
  const out: Record<string, unknown> = {};
  const typed = node as { type?: unknown };
  if (Array.isArray(typed.type)) {
    const { type, ...rest } = node as Record<string, unknown>;
    const variants = (type as string[]).map((t) => ({ type: t }));
    return normaliseJsonSchema({ ...rest, anyOf: variants });
  }
  for (const [key, value] of Object.entries(node)) {
    if (key === "$schema") continue;
    out[key === "oneOf" ? "anyOf" : key] = normaliseJsonSchema(value);
  }
  return out;
}
