import {
  cellsInRange,
  colToLetters,
  formatRange,
  lettersToCol,
  parseCell,
  parseRange,
  rangeDims,
  rangesOverlap,
  splitSheet,
} from "../src/core/a1";

describe("column letters", () => {
  it.each([
    [0, "A"],
    [25, "Z"],
    [26, "AA"],
    [701, "ZZ"],
    [702, "AAA"],
    [16383, "XFD"],
  ])("%i <-> %s", (n, letters) => {
    expect(colToLetters(n)).toBe(letters);
    expect(lettersToCol(letters)).toBe(n);
  });
});

describe("parseCell", () => {
  it("parses plain and absolute refs", () => {
    expect(parseCell("B3")).toEqual({ row: 2, col: 1 });
    expect(parseCell("$b$3")).toEqual({ row: 2, col: 1 });
  });
  it.each(["", "3B", "A0", "XFE1", "A1048577", "A1:B2"])("rejects %j", (bad) => {
    expect(() => parseCell(bad)).toThrow();
  });
});

describe("parseRange", () => {
  it("parses ranges with and without sheets", () => {
    expect(parseRange("A1:B2")).toEqual({ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } });
    expect(parseRange("Sheet2!C5").sheet).toBe("Sheet2");
    expect(parseRange("'My ''Q1'' Data'!A1:A3").sheet).toBe("My 'Q1' Data");
  });
  it("normalises reversed corners", () => {
    expect(formatRange(parseRange("C3:A1"))).toBe("A1:C3");
  });
  it("rejects whole-column and malformed refs", () => {
    expect(() => parseRange("A:A")).toThrow();
    expect(() => parseRange("A1:")).toThrow();
    expect(() => parseRange("A1:B2:C3")).toThrow();
    expect(() => splitSheet("'Unclosed!A1")).toThrow();
  });
  it("formats with quoted sheet names when needed", () => {
    expect(formatRange(parseRange("'My Sheet'!A1:B2"), true)).toBe("'My Sheet'!A1:B2");
    expect(formatRange(parseRange("Data!A1"), true)).toBe("Data!A1");
  });
});

describe("range helpers", () => {
  it("dims, iteration and overlap", () => {
    const r = parseRange("B2:C4");
    expect(rangeDims(r)).toEqual({ rows: 3, cols: 2 });
    expect([...cellsInRange(r)]).toHaveLength(6);
    expect(rangesOverlap(r, parseRange("C4:D5"))).toBe(true);
    expect(rangesOverlap(r, parseRange("D1:D9"))).toBe(false);
    expect(rangesOverlap({ ...r, sheet: "A" }, { ...r, sheet: "B" })).toBe(false);
  });
});
