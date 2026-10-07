import { colToLetters, formatCell } from "../../core/a1";
import { displayContent } from "../../core/cells";
import type { MemoryWorkbook } from "../../core/memoryWorkbook";

/** A tiny read-only spreadsheet view, shown only in browser demo mode. */
export function DemoGrid({ workbook, pending, onSelect }: {
  workbook: MemoryWorkbook;
  /** Keys ("Sheet!A1") of cells the current preview would change. */
  pending: Set<string>;
  onSelect: (address: string) => void;
}) {
  const sheet = workbook.activeSheet;
  const used = workbook.usedRange(sheet);
  const rows = Math.max(8, (used?.end.row ?? 0) + 2);
  const cols = Math.max(5, (used?.end.col ?? 0) + 2);
  const selected = workbook.selection;

  return (
    <div className="demo-grid table-wrap" aria-label="Demo sheet">
      <div className="muted small">Demo workbook · sheet <strong>{sheet}</strong> · click a cell to select it</div>
      <table>
        <thead>
          <tr>
            <th />
            {Array.from({ length: cols }, (_, c) => <th key={c}>{colToLetters(c)}</th>)}
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: rows }, (_, r) => (
            <tr key={r}>
              <th>{r + 1}</th>
              {Array.from({ length: cols }, (_, c) => {
                const addr = formatCell({ row: r, col: c });
                const cell = workbook.get(sheet, addr);
                const cls = [
                  pending.has(`${sheet}!${addr}`) ? "pending" : "",
                  selected === addr ? "selected" : "",
                  cell.format.bold ? "bold" : "",
                  cell.format.italic ? "italic" : "",
                ].join(" ");
                return (
                  <td key={c} className={cls} title={cell.formula ?? undefined}
                    style={{ background: cell.format.fillColor ?? undefined, color: cell.format.fontColor !== "#000000" ? cell.format.fontColor : undefined }}
                    onClick={() => onSelect(addr)}>
                    {displayContent(cell)}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
