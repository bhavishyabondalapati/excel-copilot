// Small helpers for comparing and displaying cell states.

import type { CellValue, Style } from "./schema";
import type { CellContent, CellFormat, CellState } from "./types";

/** What you would type into the cell: the formula if there is one, else the value. */
export function contentOf(cell: CellState): CellContent {
  if (cell.formula !== null) return cell.formula;
  return cell.value === null ? "" : cell.value;
}

/** Human-readable content for the preview table. */
export function displayContent(cell: CellState): string {
  const c = contentOf(cell);
  return typeof c === "string" ? c : String(c);
}

/**
 * Are two cells' contents the same? Loose on purpose: Excel may hand back
 * the number 5 where we wrote the string "5".
 */
export function sameContent(a: CellState, b: CellState): boolean {
  return String(contentOf(a)) === String(contentOf(b));
}

const norm = (c: string | null) => (c === null ? null : c.toUpperCase());

export function sameFormat(a: CellFormat, b: CellFormat): boolean {
  return (
    a.bold === b.bold &&
    a.italic === b.italic &&
    norm(a.fontColor) === norm(b.fontColor) &&
    norm(a.fillColor) === norm(b.fillColor) &&
    a.numberFormat === b.numberFormat &&
    a.horizontalAlignment === b.horizontalAlignment
  );
}

/** Apply a (partial) Style from a `format` action on top of an existing format. */
export function applyStyle(format: CellFormat, style: Style): CellFormat {
  return {
    bold: style.bold ?? format.bold,
    italic: style.italic ?? format.italic,
    fontColor: style.fontColor ?? format.fontColor,
    fillColor: style.fillColor ?? format.fillColor,
    numberFormat: style.numberFormat ?? format.numberFormat,
    horizontalAlignment: style.horizontalAlignment ?? format.horizontalAlignment,
  };
}

/** The cell state after writing `content` (format unchanged). */
export function withContent(cell: CellState, content: CellValue): CellState {
  if (typeof content === "string" && content.startsWith("=")) {
    // We don't calculate formulas; Excel will. The value is unknown until then.
    return { ...cell, formula: content, value: null };
  }
  return { ...cell, formula: null, value: content === null ? "" : content };
}

/** Describe which format properties differ, e.g. ["bold", "fill #FFF2CC"]. */
export function formatDiff(before: CellFormat, after: CellFormat): string[] {
  const out: string[] = [];
  if (before.bold !== after.bold) out.push(after.bold ? "bold" : "not bold");
  if (before.italic !== after.italic) out.push(after.italic ? "italic" : "not italic");
  if (norm(before.fontColor) !== norm(after.fontColor)) out.push(`font ${after.fontColor}`);
  if (norm(before.fillColor) !== norm(after.fillColor)) out.push(`fill ${after.fillColor ?? "none"}`);
  if (before.numberFormat !== after.numberFormat) out.push(`format ${after.numberFormat}`);
  if (before.horizontalAlignment !== after.horizontalAlignment) out.push(`align ${after.horizontalAlignment}`);
  return out;
}
