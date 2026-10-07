// Plain data types shared by the core, the Excel adapter and the UI.
// Keeping them JSON-friendly (no classes, no Maps) makes them easy to log,
// test and store.

import type { CellValue, Style } from "./schema";

/** The handful of format properties we can preview and undo. */
export interface CellFormat {
  bold: boolean;
  italic: boolean;
  /** "#RRGGBB" */
  fontColor: string;
  /** "#RRGGBB", or null when the cell has no fill. */
  fillColor: string | null;
  numberFormat: string;
  horizontalAlignment: string;
}

export const DEFAULT_FORMAT: CellFormat = {
  bold: false,
  italic: false,
  fontColor: "#000000",
  fillColor: null,
  numberFormat: "General",
  horizontalAlignment: "General",
};

/** Everything we know about one cell at one moment. */
export interface CellState {
  /** The displayed/calculated value ("" for an empty cell). */
  value: CellValue;
  /** The formula text if the cell holds a formula, otherwise null. */
  formula: string | null;
  format: CellFormat;
}

export const EMPTY_CELL: CellState = { value: "", formula: null, format: DEFAULT_FORMAT };

/** Cells keyed by "Sheet1!B2". Only the cells we asked for are present. */
export type Snapshot = Record<string, CellState>;

/** A small window onto part of the sheet, used to give the LLM context. */
export interface GridSample {
  /** Address of the sampled block on `sheet`, e.g. "A1:D20". */
  address: string;
  sheet: string;
  /** Row-major grid of cell states (may be truncated vs. the real range). */
  cells: CellState[][];
  /** True when the real range was larger than the sample. */
  truncated: boolean;
}

/** What the add-in tells the LLM (and the validator) about the workbook. */
export interface WorkbookContext {
  sheetNames: string[];
  activeSheet: string;
  /** The user's current selection, if any. */
  selection: GridSample | null;
  /** The top of the active sheet's used range (headers + a few rows). */
  usedRange: GridSample | null;
}

/** Content for setContents: formulas are strings starting with "=", "" clears a cell. */
export type CellContent = string | number | boolean;

/** Low-level writes the Excel adapter knows how to perform. */
export type WriteOp =
  | { kind: "setContents"; sheet: string; address: string; contents: CellContent[][] }
  | { kind: "setStyle"; sheet: string; address: string; style: Style }
  | { kind: "setFormats"; sheet: string; address: string; formats: CellFormat[][] };
