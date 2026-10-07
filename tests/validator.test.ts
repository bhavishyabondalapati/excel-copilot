import { describeIssues, validateResponse } from "../src/core/validator";

const ctx = { sheetNames: ["Sheet1", "Data"], activeSheet: "Sheet1" };
const run = (actions: unknown[], limits?: Parameters<typeof validateResponse>[2]) =>
  validateResponse({ summary: "s", actions }, ctx, limits);
const errors = (r: ReturnType<typeof run>) => r.issues.filter((i) => i.severity === "error").map((i) => i.message);

describe("validateResponse", () => {
  it("passes a good response and resolves sheets", () => {
    const r = run([
      { type: "write_formula", cell: "D2", formula: "=B2*C2" },
      { type: "format", range: "data!A1:B1", style: { bold: true } },
    ]);
    expect(r.ok).toBe(true);
    expect(r.actions.map((a) => a.range?.sheet)).toEqual(["Sheet1", "Data"]);
  });

  it("reports schema errors with the action index", () => {
    const r = validateResponse({ summary: "s", actions: [{ type: "write_formula", cell: "A1" }] }, ctx);
    expect(r.ok).toBe(false);
    expect(r.issues[0]!.actionIndex).toBe(0);
    expect(r.issues[0]!.message).toContain("formula");
  });

  it("rejects non-JSON-object input", () => {
    expect(validateResponse("nope", ctx).ok).toBe(false);
  });

  it("rejects bad addresses and unknown sheets", () => {
    expect(errors(run([{ type: "write_formula", cell: "ZZZZ1", formula: "=1" }]))[0]).toContain("Invalid address");
    expect(errors(run([{ type: "write_formula", cell: "Nope!A1", formula: "=1" }]))[0]).toContain("does not exist");
  });

  it("rejects write_formula on a range", () => {
    expect(errors(run([{ type: "write_formula", cell: "A1:A3", formula: "=1" }]))[0]).toContain("fill_range");
  });

  it("checks fill_range shape", () => {
    expect(errors(run([{ type: "fill_range", range: "A1:B2", values: [[1, 2]] }]))[0]).toContain("1x2");
    expect(errors(run([{ type: "fill_range", range: "A1:B2" }]))[0]).toContain("exactly one");
    expect(
      errors(run([{ type: "fill_range", range: "A1:A2", values: [[1], [2]], formulaTemplate: "=1" }]))[0],
    ).toContain("exactly one");
    expect(errors(run([{ type: "fill_range", range: "A1:A2", formulaTemplate: "=B{r}" }]))[0]).toContain("{r}");
    expect(errors(run([{ type: "fill_range", range: "A1:A2", values: [["=SUM(B1"], [2]] }]))[0]).toContain(
      "Unbalanced",
    );
  });

  it("checks style values", () => {
    expect(errors(run([{ type: "format", range: "A1", style: { fontColor: "red" } }]))[0]).toContain("hex");
    expect(errors(run([{ type: "format", range: "A1", style: {} }]))[0]).toContain("empty");
  });

  it("blocks unsafe formulas", () => {
    expect(errors(run([{ type: "write_formula", cell: "A1", formula: '=WEBSERVICE("http://x")' }]))[0]).toContain(
      "blocked",
    );
  });

  it("warns about circular references and overlapping writes", () => {
    const r = run([
      { type: "write_formula", cell: "A1", formula: "=A1+1" },
      { type: "fill_range", range: "A1:A2", values: [[1], [2]] },
    ]);
    expect(r.ok).toBe(true);
    const warnings = r.issues.filter((i) => i.severity === "warning").map((i) => i.message);
    expect(warnings.join(" ")).toContain("circular");
    expect(warnings.join(" ")).toContain("Overwrites");
  });

  it("enforces size limits", () => {
    const limits = { maxCellsPerAction: 4, maxTotalCells: 5, maxActions: 2 };
    expect(errors(run([{ type: "format", range: "A1:C3", style: { bold: true } }], limits))[0]).toContain("limit is 4");
    const many = run(
      [
        { type: "format", range: "A1:A3", style: { bold: true } },
        { type: "format", range: "B1:B3", style: { bold: true } },
        { type: "explain", text: "x" },
      ],
      limits,
    );
    expect(errors(many).join(" ")).toContain("Too many actions");
    expect(errors(many).join(" ")).toContain("6 cells in total");
  });

  it("describes issues for humans", () => {
    const r = run([{ type: "format", range: "A1", style: {} }]);
    expect(describeIssues(r.issues)).toBe("- [error] action 1: format.style is empty");
  });
});
