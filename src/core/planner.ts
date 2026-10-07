// The planner turns validated actions into:
//   - a *diff*: which cells change, from what, to what (for the preview), and
//   - *ops*: the low-level writes the Excel adapter performs, and
//   - *undo ops*: writes that put every touched cell back the way it was.
//
// It is a pure function of (actions, snapshot). It never talks to Excel; the
// caller reads the snapshot first and hands it in. That is what makes the
// preview and undo logic testable without Excel.

import { cellKey, cellsInRange, colToLetters, formatRange, parseCell, rangeDims, type RangeRef } from "./a1";
import { applyStyle, contentOf, sameContent, sameFormat, withContent } from "./cells";
import { expandTemplate } from "./formula";
import type { CellValue } from "./schema";
import { EMPTY_CELL, type CellContent, type CellState, type Snapshot, type WriteOp } from "./types";
import type { ValidatedAction } from "./validator";

export interface CellChange {
  key: string;
  sheet: string;
  /** Single-cell address, e.g. "B2". */
  address: string;
  row: number;
  col: number;
  before: CellState;
  after: CellState;
  contentChanged: boolean;
  formatChanged: boolean;
  /** Index of the last action that touched this cell. */
  actionIndex: number;
}

export interface Explanation {
  target?: string;
  text: string;
}

export interface Plan {
  summary: string;
  actions: ValidatedAction[];
  explanations: Explanation[];
  /** Only cells whose content or format actually differs. Sorted by sheet, row, col. */
  changes: CellChange[];
  /** Writes that perform the plan, in action order. */
  ops: WriteOp[];
  /** Every range the plan writes to: the area we must snapshot for undo. */
  touchedRanges: RangeRef[];
}

/** Which ranges must be read before planning (and again before applying). */
export function rangesToSnapshot(actions: ValidatedAction[]): RangeRef[] {
  return actions.filter((a) => a.action.type !== "explain" && a.range).map((a) => a.range!);
}

export function buildPlan(summary: string, actions: ValidatedAction[], snapshot: Snapshot): Plan {
  const touchedRanges = rangesToSnapshot(actions);
  const ops: WriteOp[] = [];
  const explanations: Explanation[] = [];

  // Simulate every action on a copy of the snapshot.
  const state = new Map<string, CellState>();
  const lastTouch = new Map<string, number>();
  const get = (key: string) => state.get(key) ?? snapshot[key] ?? EMPTY_CELL;
  const set = (key: string, cell: CellState, index: number) => {
    state.set(key, cell);
    lastTouch.set(key, index);
  };

  for (const { action, range, index } of actions) {
    if (action.type === "explain") {
      explanations.push(action.target ? { target: action.target, text: action.text } : { text: action.text });
      continue;
    }
    const r = range!;
    const sheet = r.sheet!;
    const address = formatRange(r);

    if (action.type === "write_formula") {
      const key = cellKey(sheet, r.start);
      set(key, withContent(get(key), action.formula), index);
      ops.push({ kind: "setContents", sheet, address, contents: [[action.formula]] });
    } else if (action.type === "fill_range") {
      const grid = fillGrid(r, action.values, action.formulaTemplate);
      for (const ref of cellsInRange(r)) {
        const key = cellKey(sheet, ref);
        const value = grid[ref.row - r.start.row]![ref.col - r.start.col]!;
        set(key, withContent(get(key), value), index);
      }
      ops.push({ kind: "setContents", sheet, address, contents: grid.map((row) => row.map(toContent)) });
    } else {
      for (const ref of cellsInRange(r)) {
        const key = cellKey(sheet, ref);
        const cell = get(key);
        set(key, { ...cell, format: applyStyle(cell.format, action.style) }, index);
      }
      ops.push({ kind: "setStyle", sheet, address, style: action.style });
    }
  }

  const changes: CellChange[] = [];
  for (const [key, after] of state) {
    const before = snapshot[key] ?? EMPTY_CELL;
    const contentChanged = !sameContent(before, after);
    const formatChanged = !sameFormat(before.format, after.format);
    if (!contentChanged && !formatChanged) continue;
    const [sheet, address] = splitKey(key);
    const { row, col } = parseCell(address);
    changes.push({
      key,
      sheet,
      address,
      row,
      col,
      before,
      after,
      contentChanged,
      formatChanged,
      actionIndex: lastTouch.get(key)!,
    });
  }
  changes.sort((a, b) => a.sheet.localeCompare(b.sheet) || a.row - b.row || a.col - b.col);

  return { summary, actions, explanations, changes, ops, touchedRanges };
}

/** The 2-D grid of values a fill_range writes (template expanded per cell). */
function fillGrid(r: RangeRef, values: CellValue[][] | undefined, template: string | undefined): CellValue[][] {
  if (values) return values;
  const { rows, cols } = rangeDims(r);
  return Array.from({ length: rows }, (_, i) =>
    Array.from({ length: cols }, (_, j) => expandTemplate(template!, r.start.row + i + 1, colToLetters(r.start.col + j))),
  );
}

const toContent = (v: CellValue): CellContent => (v === null ? "" : v);

/**
 * Ops that restore every touched range to its snapshot: formats first (so a
 * restored text format like "@" is in place before the content goes back),
 * then contents. We restore whole ranges, not just changed cells, which keeps
 * the op list short and also undoes side effects such as Excel auto-applying a
 * date format when a date was typed.
 */
export function buildUndoOps(touchedRanges: RangeRef[], snapshot: Snapshot): WriteOp[] {
  const formats: WriteOp[] = [];
  const contents: WriteOp[] = [];
  for (const r of touchedRanges) {
    const sheet = r.sheet!;
    const address = formatRange(r);
    const grid = gridOf(r, (ref) => snapshot[cellKey(sheet, ref)] ?? EMPTY_CELL);
    formats.push({ kind: "setFormats", sheet, address, formats: grid.map((row) => row.map((c) => c.format)) });
    contents.push({ kind: "setContents", sheet, address, contents: grid.map((row) => row.map(contentOf)) });
  }
  // Overlapping ranges are fine: every op writes data from the same snapshot.
  return [...formats, ...contents];
}

/** Cells the user has edited since the plan was applied (content or format). */
export function detectDrift(plan: Plan, current: Snapshot): CellChange[] {
  return plan.changes.filter((c) => {
    const now = current[c.key] ?? EMPTY_CELL;
    const contentMoved = c.contentChanged && !sameContent(now, c.after);
    const formatMoved = c.formatChanged && !sameFormat(now.format, c.after.format);
    return contentMoved || formatMoved;
  });
}

/** Cells whose content/format differs between two snapshots (used for stale-preview checks). */
export function snapshotDifferences(a: Snapshot, b: Snapshot): string[] {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  return [...keys].filter((k) => {
    const x = a[k] ?? EMPTY_CELL;
    const y = b[k] ?? EMPTY_CELL;
    return !sameContent(x, y) || !sameFormat(x.format, y.format);
  });
}

function gridOf<T>(r: RangeRef, f: (ref: { row: number; col: number }) => T): T[][] {
  const { rows, cols } = rangeDims(r);
  return Array.from({ length: rows }, (_, i) =>
    Array.from({ length: cols }, (_, j) => f({ row: r.start.row + i, col: r.start.col + j })),
  );
}

function splitKey(key: string): [string, string] {
  const bang = key.lastIndexOf("!");
  return [key.slice(0, bang), key.slice(bang + 1)];
}
