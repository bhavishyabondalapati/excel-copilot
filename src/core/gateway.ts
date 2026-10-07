// The one seam between the pure core and a real spreadsheet.
//
// The core only ever talks to a `WorkbookGateway`. In Excel that is backed by
// Office.js (src/excel/excelGateway.ts); in tests and the browser demo it is the
// in-memory `MemoryWorkbook`. This is the "ports and adapters" pattern: the
// core defines the port, each environment provides an adapter.

import type { RangeRef } from "./a1";
import type { Snapshot, WorkbookContext, WriteOp } from "./types";

export interface WorkbookGateway {
  /** Sheet names, selection and a sample of the sheet, for the LLM prompt. */
  getContext(): Promise<WorkbookContext>;
  /** Read content + format for every cell in the ranges (sheet must be set). */
  readCells(ranges: RangeRef[]): Promise<Snapshot>;
  /** Perform the writes, in order, as one batch. */
  applyOps(ops: WriteOp[]): Promise<void>;
}

/** How much of the sheet we show the LLM. Keeps prompts small and cheap. */
export const SAMPLE_LIMITS = { maxRows: 30, maxCols: 15 };
