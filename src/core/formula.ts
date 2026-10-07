// Lightweight formula inspection. This is NOT a full Excel parser — it is just
// enough to catch common LLM mistakes (unbalanced brackets, missing "=") and
// unsafe constructs (network calls, DDE, links to other workbooks) before
// anything touches the user's sheet.

import { parseRange, type RangeRef } from "./a1";

export const MAX_FORMULA_LENGTH = 8192; // Excel's own limit

/** Functions we refuse to write: they reach outside the workbook. */
export const BLOCKED_FUNCTIONS = new Set([
  "WEBSERVICE", // makes HTTP requests — could leak sheet data
  "RTD", // talks to real-time data servers
  "CALL", // calls DLL functions (legacy XLM)
  "REGISTER.ID",
  "EXEC",
]);

/** Functions we allow but flag, so the user takes a second look. */
export const WARN_FUNCTIONS = new Map([
  ["HYPERLINK", "creates a clickable link; check the destination"],
  ["INDIRECT", "builds references from text, so the preview cannot show what it reads"],
  ["OFFSET", "is volatile and recalculates on every change"],
]);

export interface FormulaScan {
  /** Upper-cased function names, in order of appearance (duplicates removed). */
  functions: string[];
  /** Cell/range references found outside string literals (best-effort). */
  references: RangeRef[];
  /** Hard problems: the formula must not be written. */
  errors: string[];
  /** Soft problems: show to the user, but allow. */
  warnings: string[];
}

/**
 * Replace the contents of string literals with spaces so later regexes cannot
 * be fooled by text like ="WEBSERVICE(". Returns null if a string is unterminated.
 * Excel escapes a quote inside a string by doubling it: "say ""hi""".
 */
export function maskStrings(formula: string): string | null {
  let out = "";
  let inString = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i]!;
    if (inString) {
      if (ch === '"' && formula[i + 1] === '"') {
        out += "  ";
        i++;
      } else if (ch === '"') {
        inString = false;
        out += '"';
      } else {
        out += " ";
      }
    } else {
      if (ch === '"') inString = true;
      out += ch;
    }
  }
  return inString ? null : out;
}

const FUNCTION_RE = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
// Optional sheet prefix ('Quoted Name'! or Plain!) then A1 or A1:B2.
const REF_RE =
  /(?:('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?[0-9]+(?::\$?[A-Za-z]{1,3}\$?[0-9]+)?)(?![A-Za-z0-9_(!])/g;

export function scanFormula(formula: string): FormulaScan {
  const scan: FormulaScan = { functions: [], references: [], errors: [], warnings: [] };

  if (!formula.startsWith("=")) scan.errors.push('Formula must start with "="');
  if (formula.length > MAX_FORMULA_LENGTH) {
    scan.errors.push(`Formula is longer than Excel's limit of ${MAX_FORMULA_LENGTH} characters`);
  }
  if (formula.trim() === "=") scan.errors.push("Formula is empty");

  const masked = maskStrings(formula);
  if (masked === null) {
    scan.errors.push("Unterminated string literal (odd number of double quotes)");
    return scan;
  }

  // Parentheses must balance and never close before they open.
  let depth = 0;
  for (const ch of masked) {
    if (ch === "(") depth++;
    if (ch === ")") depth--;
    if (depth < 0) break;
  }
  if (depth !== 0) scan.errors.push("Unbalanced parentheses");

  if (masked.includes("|")) scan.errors.push("DDE-style references (\"|\") are not allowed");
  // Table refs like Table1[Col] are fine; [Book.xlsx]Sheet!A1 or [1]Sheet!A1 are external links.
  if (/\[[^\]]*\.(xls[xmb]?|csv)\]|\[\d+\]/i.test(masked)) {
    scan.errors.push("References to other workbooks are not allowed");
  }

  const seen = new Set<string>();
  for (const m of masked.matchAll(FUNCTION_RE)) {
    const name = m[1]!.toUpperCase().replace(/^_XLFN\./, "");
    if (seen.has(name)) continue;
    seen.add(name);
    scan.functions.push(name);
    if (BLOCKED_FUNCTIONS.has(name)) scan.errors.push(`Function ${name} is blocked for safety`);
    const warn = WARN_FUNCTIONS.get(name);
    if (warn) scan.warnings.push(`${name} ${warn}`);
  }

  for (const m of masked.matchAll(REF_RE)) {
    // A match preceded by a letter/digit is part of a longer name, not a reference.
    const before = m.index === 0 ? "" : masked[m.index! - 1]!;
    if (/[A-Za-z0-9_.]/.test(before)) continue;
    try {
      const sheetPrefix = m[1] ? `${m[1]}!` : "";
      scan.references.push(parseRange(sheetPrefix + m[2]!));
    } catch {
      // Something like "ABCD1" (4 letters) — not a reference, ignore.
    }
  }

  return scan;
}

/**
 * Expand a fill template for one cell.
 * "=B{row}*C{row}" at row index 4 -> "=B5*C5". "{col}" becomes the column letter.
 */
export function expandTemplate(template: string, rowNumber: number, colLetters: string): string {
  return template.replace(/\{row\}/g, String(rowNumber)).replace(/\{col\}/g, colLetters);
}

/**
 * Placeholders a template may use are {row} and {col}. Any other {word} is
 * probably a typo like {r}. Array constants such as {1,2,3} are left alone.
 */
export function unknownPlaceholders(template: string): string[] {
  const masked = maskStrings(template) ?? template;
  const found = [...masked.matchAll(/\{([A-Za-z_]+)\}/g)].map((m) => m[1]!);
  return found.filter((p) => p !== "row" && p !== "col");
}
