import { CopilotResponseSchema, responseJsonSchema } from "../src/core/schema";

describe("CopilotResponseSchema", () => {
  it("accepts every action type", () => {
    const ok = CopilotResponseSchema.safeParse({
      summary: "demo",
      actions: [
        { type: "write_formula", cell: "D2", formula: "=B2*C2" },
        { type: "fill_range", range: "A1:B1", values: [["x", 1]] },
        { type: "fill_range", range: "E2:E5", formulaTemplate: "=D{row}*2" },
        { type: "format", range: "A1:E1", style: { bold: true, fillColor: "#FFF2CC" } },
        { type: "explain", text: "hello" },
      ],
    });
    expect(ok.success).toBe(true);
  });

  it("rejects unknown action types and extra keys (strict)", () => {
    expect(CopilotResponseSchema.safeParse({ summary: "", actions: [{ type: "delete_sheet" }] }).success).toBe(false);
    expect(
      CopilotResponseSchema.safeParse({
        summary: "",
        actions: [{ type: "write_formula", cell: "A1", formula: "=1", sneaky: true }],
      }).success,
    ).toBe(false);
  });
});

describe("responseJsonSchema", () => {
  const schema = JSON.stringify(responseJsonSchema());
  it("is provider friendly", () => {
    expect(schema).not.toContain("oneOf");
    expect(schema).not.toContain("$schema");
    expect(schema).not.toMatch(/"type":\[/); // no type arrays
    expect(schema).toContain('"additionalProperties":false');
  });
  it("mentions every action type", () => {
    for (const t of ["write_formula", "fill_range", "format", "explain"]) expect(schema).toContain(t);
  });
});
