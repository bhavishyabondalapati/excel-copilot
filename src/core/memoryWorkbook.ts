// An in-memory stand-in for an Excel workbook. Used by tests and by the
// browser demo mode (when the task pane is opened outside Excel).
//
// It stores cells sparsely (only non-empty ones) and does not calculate
// formulas — a formula cell's value stays null, just like the planner assumes.

import {
  cellKey,
  cellsInRange,
  formatCell,
  formatRange,
  parseCell,
  parseRange,
  rangeDims,
  type CellRef,
  type RangeRef,
} from "./a1";
import { applyStyle, sameFormat, withContent } from "./cells";
import { SAMPLE_LIMITS, type WorkbookGateway } from "./gateway";
import type { CellValue } from "./schema";
import { DEFAULT_FORMAT, EMPTY_CELL, type CellState, type GridSample, type Snapshot, type WorkbookContext, type WriteOp } from "./types";

export interface MemoryWorkbookInit {
  /** sheet name -> { "A1": value-or-formula } */
  sheets: Record<string, Record<string, CellValue>>;
  activeSheet?: string;
  /** e.g. "B2:B5" on the active sheet */
  selection?: string;
}

export class MemoryWorkbook implements WorkbookGateway {
  private sheets = new Map<string, Map<string, CellState>>();
  activeSheet: string;
  selection: string | null;
  /** Every batch passed to applyOps, for assertions in tests. */
  readonly appliedBatches: WriteOp[][] = [];

  constructor(init: MemoryWorkbookInit) {
    for (const [name, cells] of Object.entries(init.sheets)) {
      const sheet = new Map<string, CellState>();
      for (const [addr, value] of Object.entries(cells)) {
        sheet.set(formatCell(parseCell(addr)), withContent(EMPTY_CELL, value));
      }
      this.sheets.set(name, sheet);
    }
    const first = Object.keys(init.sheets)[0];
    if (!first) throw new Error("MemoryWorkbook needs at least one sheet");
    this.activeSheet = init.activeSheet ?? first;
    this.selection = init.selection ?? null;
  }

  // ---- direct access (tests / demo grid) -----------------------------------

  sheetNames(): string[] {
    return [...this.sheets.keys()];
  }

  get(sheet: string, address: string): CellState {
    return this.sheet(sheet).get(formatCell(parseCell(address))) ?? EMPTY_CELL;
  }

  /** Simulate the user typing into a cell. */
  set(sheet: string, address: string, value: CellValue): void {
    this.put(sheet, parseCell(address), withContent(this.get(sheet, address), value));
  }

  /** Bounding box of non-empty cells on a sheet, or null if empty. */
  usedRange(sheet: string): RangeRef | null {
    const refs = [...this.sheet(sheet).keys()].map((a) => parseCell(a));
    if (refs.length === 0) return null;
    return {
      sheet,
      start: { row: Math.min(...refs.map((r) => r.row)), col: Math.min(...refs.map((r) => r.col)) },
      end: { row: Math.max(...refs.map((r) => r.row)), col: Math.max(...refs.map((r) => r.col)) },
    };
  }

  // ---- WorkbookGateway ------------------------------------------------------

  async getContext(): Promise<WorkbookContext> {
    const used = this.usedRange(this.activeSheet);
    return {
      sheetNames: this.sheetNames(),
      activeSheet: this.activeSheet,
      selection: this.selection ? this.sample({ ...parseRange(this.selection), sheet: this.activeSheet }) : null,
      usedRange: used ? this.sample(used) : null,
    };
  }

  async readCells(ranges: RangeRef[]): Promise<Snapshot> {
    const snap: Snapshot = {};
    for (const r of ranges) {
      for (const ref of cellsInRange(r)) {
        snap[cellKey(r.sheet!, ref)] = this.getRef(r.sheet!, ref);
      }
    }
    return snap;
  }

  async applyOps(ops: WriteOp[]): Promise<void> {
    this.appliedBatches.push(ops);
    for (const op of ops) {
      const r = parseRange(op.address);
      const { rows, cols } = rangeDims(r);
      for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
          const ref = { row: r.start.row + i, col: r.start.col + j };
          const cell = this.getRef(op.sheet, ref);
          if (op.kind === "setContents") {
            this.put(op.sheet, ref, withContent(cell, op.contents[i]![j]!));
          } else if (op.kind === "setStyle") {
            this.put(op.sheet, ref, { ...cell, format: applyStyle(cell.format, op.style) });
          } else {
            this.put(op.sheet, ref, { ...cell, format: op.formats[i]![j]! });
          }
        }
      }
    }
  }

  // ---- internals -----------------------------------------------------------

  private sheet(name: string): Map<string, CellState> {
    const s = this.sheets.get(name);
    if (!s) throw new Error(`No sheet named "${name}"`);
    return s;
  }

  private getRef(sheet: string, ref: CellRef): CellState {
    return this.sheet(sheet).get(formatCell(ref)) ?? EMPTY_CELL;
  }

  private put(sheet: string, ref: CellRef, cell: CellState): void {
    const isEmpty = cell.formula === null && (cell.value === "" || cell.value === null) && sameFormat(cell.format, DEFAULT_FORMAT);
    if (isEmpty) this.sheet(sheet).delete(formatCell(ref));
    else this.sheet(sheet).set(formatCell(ref), cell);
  }

  private sample(r: RangeRef): GridSample {
    const { rows, cols } = rangeDims(r);
    const shownRows = Math.min(rows, SAMPLE_LIMITS.maxRows);
    const shownCols = Math.min(cols, SAMPLE_LIMITS.maxCols);
    const cells = Array.from({ length: shownRows }, (_, i) =>
      Array.from({ length: shownCols }, (_, j) => this.getRef(r.sheet!, { row: r.start.row + i, col: r.start.col + j })),
    );
    const shown: RangeRef = {
      sheet: r.sheet!,
      start: r.start,
      end: { row: r.start.row + shownRows - 1, col: r.start.col + shownCols - 1 },
    };
    return {
      address: formatRange(shown),
      sheet: r.sheet!,
      cells,
      truncated: shownRows < rows || shownCols < cols,
    };
  }
}
