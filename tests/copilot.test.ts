import { Copilot, extractJson, StalePreviewError, ValidationFailedError } from "../src/core/copilot";
import { MemoryWorkbook } from "../src/core/memoryWorkbook";
import { FakeProvider, fakeJson } from "../src/providers/fake";

const sales = () =>
  new MemoryWorkbook({
    sheets: { Sheet1: { A1: "Item", B1: "Qty", C1: "Price", A2: "Apple", B2: 3, C2: 0.5, A3: "Pear", B3: 2, C3: 0.75 } },
    selection: "B2",
  });

const totals = {
  summary: "Add a Total column",
  actions: [
    { type: "write_formula", cell: "D1", formula: '="Total"' },
    { type: "fill_range", range: "D2:D3", formulaTemplate: "=B{row}*C{row}" },
    { type: "format", range: "A1:D1", style: { bold: true } },
  ],
};

describe("extractJson", () => {
  it("handles plain, fenced and chatty replies", () => {
    expect(extractJson('{"a":1}')).toEqual({ a: 1 });
    expect(extractJson('```json\n{"a":1}\n```')).toEqual({ a: 1 });
    expect(extractJson('Sure! {"a":1} Hope that helps')).toEqual({ a: 1 });
    expect(extractJson("no json here")).toBeUndefined();
  });
});

describe("Copilot end to end (fake provider + memory workbook)", () => {
  it("proposes, applies and undoes", async () => {
    const wb = sales();
    const provider = fakeJson(totals);
    const copilot = new Copilot(provider, wb);

    const proposal = await copilot.propose("add a total column");
    expect(proposal.attempts).toBe(1);
    expect(proposal.plan.changes.map((c) => c.address)).toEqual(["A1", "B1", "C1", "D1", "D2", "D3"]);
    expect(wb.get("Sheet1", "D2").formula).toBeNull(); // preview doesn't write

    // The prompt includes the sheet contents and the request.
    const prompt = provider.requests[0]!.messages[0]!.content;
    expect(prompt).toContain("Apple");
    expect(prompt).toContain("add a total column");

    const applied = await copilot.apply(proposal);
    expect(wb.get("Sheet1", "D3").formula).toBe("=B3*C3");
    expect(await copilot.checkDrift(applied)).toEqual([]);

    await copilot.undo(applied);
    expect(wb.get("Sheet1", "D3").formula).toBeNull();
    expect(wb.get("Sheet1", "A1").format.bold).toBe(false);
  });

  it("repairs an invalid reply by sending the problems back", async () => {
    const bad = { summary: "x", actions: [{ type: "fill_range", range: "D2:D3", values: [[1]] }] };
    const provider = fakeJson(bad, totals);
    const proposal = await new Copilot(provider, sales()).propose("totals");
    expect(proposal.attempts).toBe(2);
    const repair = provider.requests[1]!.messages;
    expect(repair).toHaveLength(3);
    expect(repair[2]!.content).toContain("values is 1x1 but range D2:D3 is 2x1");
  });

  it("gives up after maxRepairs with the issues attached", async () => {
    const provider = new FakeProvider("not json at all");
    await expect(new Copilot(provider, sales()).propose("x", { maxRepairs: 2 })).rejects.toMatchObject({
      name: "ValidationFailedError",
      attempts: 3,
    });
    expect(provider.requests).toHaveLength(3);
  });

  it("refuses to apply a stale preview", async () => {
    const wb = sales();
    const copilot = new Copilot(fakeJson(totals), wb);
    const proposal = await copilot.propose("totals");
    wb.set("Sheet1", "D2", "typed by user");
    await expect(copilot.apply(proposal)).rejects.toBeInstanceOf(StalePreviewError);
    expect(wb.get("Sheet1", "D2").value).toBe("typed by user");
  });

  it("reports drift before undo", async () => {
    const wb = sales();
    const copilot = new Copilot(fakeJson(totals), wb);
    const applied = await copilot.apply(await copilot.propose("totals"));
    wb.set("Sheet1", "D2", 99);
    expect(await copilot.checkDrift(applied)).toEqual(["Sheet1!D2"]);
  });

  it("does not mistake Excel's normalisation of a formula for a user edit", async () => {
    // Real Excel upper-cases function names; simulate that.
    class NormalisingWorkbook extends MemoryWorkbook {
      override async applyOps(ops: Parameters<MemoryWorkbook["applyOps"]>[0]) {
        await super.applyOps(
          ops.map((op) =>
            op.kind === "setContents"
              ? { ...op, contents: op.contents.map((r) => r.map((c) => (typeof c === "string" ? c.toUpperCase() : c))) }
              : op,
          ),
        );
      }
    }
    const wb = new NormalisingWorkbook({ sheets: { Sheet1: { A1: 1, A2: 2 } } });
    const copilot = new Copilot(fakeJson({ summary: "s", actions: [{ type: "write_formula", cell: "A3", formula: "=sum(a1:a2)" }] }), wb);
    const applied = await copilot.apply(await copilot.propose("sum"));
    expect(wb.get("Sheet1", "A3").formula).toBe("=SUM(A1:A2)");
    expect(await copilot.checkDrift(applied)).toEqual([]);
  });

  it("surfaces provider errors unchanged", async () => {
    const copilot = new Copilot(new FakeProvider(new Error("boom")), sales());
    await expect(copilot.propose("x")).rejects.toThrow("boom");
  });

  it("never applies errors as a partial plan", async () => {
    const provider = fakeJson({
      summary: "x",
      actions: [
        { type: "write_formula", cell: "D2", formula: "=B2*C2" },
        { type: "write_formula", cell: "D3", formula: '=WEBSERVICE("http://x")' },
      ],
    });
    const err = await new Copilot(provider, sales()).propose("x", { maxRepairs: 0 }).catch((e) => e);
    expect(err).toBeInstanceOf(ValidationFailedError);
  });
});

describe("explainSelection", () => {
  it("explains the formula in the selected cell", async () => {
    const wb = sales();
    wb.set("Sheet1", "D2", "=B2*C2");
    wb.selection = "D2";
    const provider = fakeJson({
      summary: "Explanation",
      actions: [{ type: "explain", target: "D2", text: "Multiplies quantity by price." }],
    });
    const result = await new Copilot(provider, wb).explainSelection();
    expect(result).toEqual({ address: "Sheet1!D2", formula: "=B2*C2", text: "Multiplies quantity by price." });
    expect(provider.requests[0]!.messages[0]!.content).toContain("=B2*C2");
  });

  it("explains nothing when the cell has no formula", async () => {
    const copilot = new Copilot(fakeJson({ summary: "", actions: [] }), sales());
    await expect(copilot.explainSelection()).rejects.toThrow("doesn't contain a formula");
  });
});
