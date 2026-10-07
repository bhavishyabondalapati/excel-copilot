import { buildPlan, buildUndoOps, rangesToSnapshot, snapshotDifferences } from "../src/core/planner";
import { validateResponse } from "../src/core/validator";
import { MemoryWorkbook } from "../src/core/memoryWorkbook";
import { displayContent } from "../src/core/cells";

function setup() {
  const wb = new MemoryWorkbook({
    sheets: {
      Sheet1: { A1: "Item", B1: "Qty", C1: "Price", A2: "Apple", B2: 3, C2: 0.5, A3: "Pear", B3: 2, C3: 0.75 },
      Other: {},
    },
  });
  return wb;
}

async function plan(wb: MemoryWorkbook, actions: unknown[]) {
  const ctx = await wb.getContext();
  const v = validateResponse({ summary: "test", actions }, ctx);
  expect(v.issues.filter((i) => i.severity === "error")).toEqual([]);
  const snapshot = await wb.readCells(rangesToSnapshot(v.actions));
  return { plan: buildPlan("test", v.actions, snapshot), snapshot };
}

describe("buildPlan", () => {
  it("diffs a formula write", async () => {
    const wb = setup();
    const { plan: p } = await plan(wb, [{ type: "write_formula", cell: "D2", formula: "=B2*C2" }]);
    expect(p.changes).toHaveLength(1);
    const c = p.changes[0]!;
    expect([c.key, displayContent(c.before), displayContent(c.after)]).toEqual(["Sheet1!D2", "", "=B2*C2"]);
    expect(p.ops).toEqual([{ kind: "setContents", sheet: "Sheet1", address: "D2", contents: [["=B2*C2"]] }]);
  });

  it("expands formula templates per row", async () => {
    const wb = setup();
    const { plan: p } = await plan(wb, [{ type: "fill_range", range: "D2:D3", formulaTemplate: "=B{row}*C{row}" }]);
    expect(p.changes.map((c) => c.after.formula)).toEqual(["=B2*C2", "=B3*C3"]);
  });

  it("omits cells that do not actually change", async () => {
    const wb = setup();
    const { plan: p } = await plan(wb, [{ type: "fill_range", range: "A1:B1", values: [["Item", "Quantity"]] }]);
    expect(p.changes.map((c) => c.address)).toEqual(["B1"]);
  });

  it("combines content and format changes on the same cell", async () => {
    const wb = setup();
    const { plan: p } = await plan(wb, [
      { type: "write_formula", cell: "D1", formula: '="Total"' },
      { type: "format", range: "A1:D1", style: { bold: true } },
    ]);
    const d1 = p.changes.find((c) => c.address === "D1")!;
    expect(d1.contentChanged && d1.formatChanged).toBe(true);
    expect(d1.actionIndex).toBe(1);
    expect(p.changes).toHaveLength(4);
  });

  it("collects explanations without producing ops", async () => {
    const wb = setup();
    const { plan: p } = await plan(wb, [{ type: "explain", target: "B2", text: "Quantity of apples" }]);
    expect(p.ops).toEqual([]);
    expect(p.explanations).toEqual([{ target: "B2", text: "Quantity of apples" }]);
  });

  it("sorts changes by sheet, row, column", async () => {
    const wb = setup();
    const { plan: p } = await plan(wb, [
      { type: "write_formula", cell: "B9", formula: "=1" },
      { type: "write_formula", cell: "A9", formula: "=1" },
      { type: "write_formula", cell: "Other!A1", formula: "=1" },
    ]);
    expect(p.changes.map((c) => c.key)).toEqual(["Other!A1", "Sheet1!A9", "Sheet1!B9"]);
  });
});

describe("apply + undo round trip", () => {
  it("restores content and format exactly", async () => {
    const wb = setup();
    const before = await wb.readCells([{ sheet: "Sheet1", start: { row: 0, col: 0 }, end: { row: 4, col: 4 } }]);
    const { plan: p, snapshot } = await plan(wb, [
      { type: "fill_range", range: "D1:D3", values: [["Total"], ["=B2*C2"], ["=B3*C3"]] },
      { type: "format", range: "A1:D1", style: { bold: true, fillColor: "#FFF2CC" } },
      { type: "fill_range", range: "B2:B3", values: [[null], [10]] },
    ]);
    await wb.applyOps(p.ops);
    expect(wb.get("Sheet1", "D2").formula).toBe("=B2*C2");
    expect(wb.get("Sheet1", "A1").format.bold).toBe(true);
    expect(wb.get("Sheet1", "B2").value).toBe("");

    await wb.applyOps(buildUndoOps(p.touchedRanges, snapshot));
    const after = await wb.readCells([{ sheet: "Sheet1", start: { row: 0, col: 0 }, end: { row: 4, col: 4 } }]);
    expect(snapshotDifferences(before, after)).toEqual([]);
  });
});

describe("MemoryWorkbook.getContext", () => {
  it("reports sheets, used range and selection", async () => {
    const wb = setup();
    wb.selection = "B2:B3";
    const ctx = await wb.getContext();
    expect(ctx.sheetNames).toEqual(["Sheet1", "Other"]);
    expect(ctx.usedRange?.address).toBe("A1:C3");
    expect(ctx.selection?.cells.map((r) => r[0]!.value)).toEqual([3, 2]);
  });
});
