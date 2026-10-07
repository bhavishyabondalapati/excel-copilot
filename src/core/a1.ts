// A1-notation helpers: turn "Sheet1!B2:C5" into numbers and back.
// Everything here is pure: no Excel, no DOM, so it is easy to unit-test.

/** Excel's hard limits (Excel 2007+). */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384; // column "XFD"

/** Zero-based cell position. Row 0 / col 0 is "A1". */
export interface CellRef {
  row: number;
  col: number;
}

/** A rectangular block of cells, optionally on a named sheet. */
export interface RangeRef {
  sheet?: string;
  start: CellRef;
  end: CellRef;
}

/** 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function colToLetters(col: number): string {
  if (!Number.isInteger(col) || col < 0) throw new Error(`Invalid column index: ${col}`);
  let n = col + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** "A" -> 0, "AA" -> 26. Case-insensitive. */
export function lettersToCol(letters: string): number {
  if (!/^[A-Za-z]+$/.test(letters)) throw new Error(`Invalid column letters: ${letters}`);
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;

/** "B3" or "$B$3" -> { row: 2, col: 1 }. Throws on anything else. */
export function parseCell(text: string): CellRef {
  const m = CELL_RE.exec(text.trim());
  if (!m) throw new Error(`Not a cell address: "${text}"`);
  const col = lettersToCol(m[1]!);
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= MAX_ROWS) throw new Error(`Row out of bounds in "${text}"`);
  if (col >= MAX_COLS) throw new Error(`Column out of bounds in "${text}"`);
  return { row, col };
}

/** { row: 2, col: 1 } -> "B3". */
export function formatCell(ref: CellRef): string {
  return `${colToLetters(ref.col)}${ref.row + 1}`;
}

/**
 * Split "Sheet1!A1:B2" / "'My Sheet'!A1" into sheet + address.
 * Quoted sheet names may contain "!" and escaped quotes ('').
 */
export function splitSheet(text: string): { sheet?: string; address: string } {
  const t = text.trim();
  if (t.startsWith("'")) {
    // Find the closing quote that is followed by "!".
    for (let i = 1; i < t.length; i++) {
      if (t[i] === "'" && t[i + 1] === "'") {
        i++; // escaped quote
        continue;
      }
      if (t[i] === "'") {
        if (t[i + 1] !== "!") throw new Error(`Expected "!" after sheet name in "${text}"`);
        return { sheet: t.slice(1, i).replace(/''/g, "'"), address: t.slice(i + 2) };
      }
    }
    throw new Error(`Unterminated sheet name in "${text}"`);
  }
  const bang = t.lastIndexOf("!");
  if (bang === -1) return { address: t };
  const sheet = t.slice(0, bang);
  if (!sheet) throw new Error(`Empty sheet name in "${text}"`);
  return { sheet, address: t.slice(bang + 1) };
}

/** Parse "A1", "A1:B5", "Sheet1!A1:B5". Corners are normalised so start <= end. */
export function parseRange(text: string): RangeRef {
  const { sheet, address } = splitSheet(text);
  const parts = address.split(":");
  if (parts.length > 2 || parts.some((p) => p === "")) {
    throw new Error(`Not a range address: "${text}"`);
  }
  const a = parseCell(parts[0]!);
  const b = parts.length === 2 ? parseCell(parts[1]!) : a;
  const range: RangeRef = {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
  if (sheet !== undefined) range.sheet = sheet;
  return range;
}

/** Quote a sheet name only when Excel would require it. */
export function quoteSheet(sheet: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
}

/** RangeRef -> "A1:B2" (or "A1" for a single cell). Sheet is omitted unless includeSheet. */
export function formatRange(range: RangeRef, includeSheet = false): string {
  const a = formatCell(range.start);
  const b = formatCell(range.end);
  const addr = a === b ? a : `${a}:${b}`;
  return includeSheet && range.sheet ? `${quoteSheet(range.sheet)}!${addr}` : addr;
}

export function rangeDims(range: RangeRef): { rows: number; cols: number } {
  return {
    rows: range.end.row - range.start.row + 1,
    cols: range.end.col - range.start.col + 1,
  };
}

export function cellCount(range: RangeRef): number {
  const { rows, cols } = rangeDims(range);
  return rows * cols;
}

export function isSingleCell(range: RangeRef): boolean {
  return cellCount(range) === 1;
}

/** Visit every cell in row-major order (left-to-right, then down). */
export function* cellsInRange(range: RangeRef): Generator<CellRef> {
  for (let row = range.start.row; row <= range.end.row; row++) {
    for (let col = range.start.col; col <= range.end.col; col++) {
      yield { row, col };
    }
  }
}

/** Do two ranges (on the same sheet) share at least one cell? */
export function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
  if ((a.sheet ?? "") !== (b.sheet ?? "")) return false;
  return !(
    a.end.row < b.start.row ||
    b.end.row < a.start.row ||
    a.end.col < b.start.col ||
    b.end.col < a.start.col
  );
}

/** Stable key used to index cells in snapshots: "Sheet1!B2". */
export function cellKey(sheet: string, ref: CellRef): string {
  return `${sheet}!${formatCell(ref)}`;
}
