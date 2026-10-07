// The Office.js adapter: implements WorkbookGateway against real Excel.
//
// Office.js basics, in one paragraph: everything happens inside
// `Excel.run(async (context) => ...)`. Calls like `range.load("values")` and
// `range.values = ...` are *queued* on proxy objects; nothing reaches Excel
// until `await context.sync()`. So the pattern is always: queue loads -> sync ->
// read properties; queue writes -> sync. Fewer syncs = faster add-in.
//
// This file is deliberately thin. All decisions (what to write, what to
// restore) are made by the pure core; this only moves data in and out.

import { cellKey, formatRange, parseRange, rangeDims, splitSheet, type RangeRef } from "../core/a1";
import { SAMPLE_LIMITS, type WorkbookGateway } from "../core/gateway";
import type { CellValue } from "../core/schema";
import { DEFAULT_FORMAT, type CellFormat, type CellState, type GridSample, type Snapshot, type WorkbookContext, type WriteOp } from "../core/types";

/** Format properties we read/write via get/setCellProperties (ExcelApi 1.9). */
const FORMAT_LOAD: Excel.CellPropertiesLoadOptions = {
  format: {
    font: { bold: true, italic: true, color: true },
    fill: { color: true, pattern: true },
    horizontalAlignment: true,
  },
};

export class ExcelGateway implements WorkbookGateway {
  async getContext(): Promise<WorkbookContext> {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      const active = sheets.getActiveWorksheet();
      active.load("name");
      const selected = context.workbook.getSelectedRange();
      selected.load("address, rowCount, columnCount");
      const used = active.getUsedRangeOrNullObject(true);
      used.load("address, rowCount, columnCount");
      await context.sync();

      const selSample = sampleOf(selected);
      const usedSample = used.isNullObject ? null : sampleOf(used);
      await context.sync();

      return {
        sheetNames: sheets.items.map((s) => s.name),
        activeSheet: active.name,
        selection: selSample ? toGridSample(selSample) : null,
        usedRange: usedSample ? toGridSample(usedSample) : null,
      };
    });
  }

  async readCells(ranges: RangeRef[]): Promise<Snapshot> {
    if (ranges.length === 0) return {};
    return Excel.run(async (context) => {
      const loaded = ranges.map((r) => {
        const range = context.workbook.worksheets.getItem(r.sheet!).getRange(formatRange(r));
        range.load("values, formulas, numberFormat");
        const props = range.getCellProperties(FORMAT_LOAD);
        return { r, range, props };
      });
      await context.sync();

      const snap: Snapshot = {};
      for (const { r, range, props } of loaded) {
        const { rows, cols } = rangeDims(r);
        for (let i = 0; i < rows; i++) {
          for (let j = 0; j < cols; j++) {
            const ref = { row: r.start.row + i, col: r.start.col + j };
            snap[cellKey(r.sheet!, ref)] = toCellState(
              range.values[i]![j],
              range.formulas[i]![j],
              range.numberFormat[i]![j],
              props.value[i]![j]!,
            );
          }
        }
      }
      return snap;
    });
  }

  async applyOps(ops: WriteOp[]): Promise<void> {
    if (ops.length === 0) return;
    await Excel.run(async (context) => {
      for (const op of ops) {
        const range = context.workbook.worksheets.getItem(op.sheet).getRange(op.address);
        const { rows, cols } = rangeDims(parseRange(op.address));
        const fill2d = <T,>(v: T): T[][] => Array.from({ length: rows }, () => Array<T>(cols).fill(v));

        if (op.kind === "setContents") {
          // Assigning to `formulas` writes formulas *and* constants; "" clears.
          range.formulas = op.contents;
        } else if (op.kind === "setStyle") {
          const s = op.style;
          if (s.bold !== undefined) range.format.font.bold = s.bold;
          if (s.italic !== undefined) range.format.font.italic = s.italic;
          if (s.fontColor !== undefined) range.format.font.color = s.fontColor;
          if (s.fillColor !== undefined) range.format.fill.color = s.fillColor;
          if (s.horizontalAlignment !== undefined) range.format.horizontalAlignment = s.horizontalAlignment;
          if (s.numberFormat !== undefined) range.numberFormat = fill2d(s.numberFormat);
        } else {
          // Undo path: restore each cell's exact format.
          range.setCellProperties(
            op.formats.map((row) =>
              row.map((f) => ({
                format: {
                  font: { bold: f.bold, italic: f.italic, color: f.fontColor },
                  horizontalAlignment: f.horizontalAlignment as Excel.HorizontalAlignment,
                  ...(f.fillColor ? { fill: { color: f.fillColor } } : {}),
                },
              })),
            ),
          );
          // "No fill" can't be expressed as a colour, so clear those cells explicitly.
          op.formats.forEach((row, i) =>
            row.forEach((f, j) => {
              if (f.fillColor === null) range.getCell(i, j).format.fill.clear();
            }),
          );
          range.numberFormat = op.formats.map((row) => row.map((f) => f.numberFormat));
        }
      }
      // One sync: Excel receives the whole batch together.
      await context.sync();
    });
  }
}

/** Load values/formulas for the top-left corner of a range (bounded by SAMPLE_LIMITS). */
function sampleOf(range: Excel.Range) {
  const rows = Math.min(range.rowCount, SAMPLE_LIMITS.maxRows);
  const cols = Math.min(range.columnCount, SAMPLE_LIMITS.maxCols);
  const corner = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
  corner.load("address, values, formulas");
  return { corner, truncated: rows < range.rowCount || cols < range.columnCount };
}

function toGridSample({ corner, truncated }: ReturnType<typeof sampleOf>): GridSample {
  const { sheet, address } = splitSheet(corner.address);
  return {
    sheet: sheet ?? "",
    address,
    truncated,
    cells: corner.values.map((row, i) =>
      row.map((value, j) => toCellState(value, corner.formulas[i]![j], "General", null)),
    ),
  };
}

function toCellState(
  value: unknown,
  formula: unknown,
  numberFormat: unknown,
  props: Excel.CellProperties | null,
): CellState {
  const isFormula = typeof formula === "string" && formula.startsWith("=");
  return {
    value: normaliseValue(value),
    formula: isFormula ? formula : null,
    format: props ? toCellFormat(props, String(numberFormat ?? "General")) : DEFAULT_FORMAT,
  };
}

function normaliseValue(v: unknown): CellValue {
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  return v === null || v === undefined ? "" : String(v);
}

function toCellFormat(p: Excel.CellProperties, numberFormat: string): CellFormat {
  const fill = p.format?.fill;
  const noFill = !fill || fill.pattern === "None" || !fill.color;
  return {
    bold: p.format?.font?.bold ?? false,
    italic: p.format?.font?.italic ?? false,
    fontColor: p.format?.font?.color ?? "#000000",
    fillColor: noFill ? null : fill!.color!,
    numberFormat,
    horizontalAlignment: p.format?.horizontalAlignment ?? "General",
  };
}

/** Call `onChange` whenever the user selects different cells. Returns an unsubscribe function. */
export function watchSelection(onChange: () => void): () => void {
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onChange);
  return () =>
    Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: onChange });
}
