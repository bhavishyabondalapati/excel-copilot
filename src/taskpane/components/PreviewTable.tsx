import { displayContent, formatDiff } from "../../core/cells";
import type { CellChange } from "../../core/planner";

const MAX_ROWS = 200;

/** The "diff": one row per cell that will change, before -> after. */
export function PreviewTable({ changes, multiSheet }: { changes: CellChange[]; multiSheet: boolean }) {
  if (changes.length === 0) return <p className="muted">No cells will change.</p>;
  const shown = changes.slice(0, MAX_ROWS);
  return (
    <div className="table-wrap">
      <table className="diff" aria-label="Cells that will change">
        <thead>
          <tr>
            <th>Cell</th>
            <th>Before</th>
            <th>After</th>
          </tr>
        </thead>
        <tbody>
          {shown.map((c) => (
            <tr key={c.key}>
              <td className="addr">{multiSheet ? `${c.sheet}!${c.address}` : c.address}</td>
              <td className="before">
                {c.contentChanged ? <Content text={displayContent(c.before)} /> : <span className="muted">same</span>}
              </td>
              <td className="after">
                {c.contentChanged && <Content text={displayContent(c.after)} />}
                {c.formatChanged && (
                  <div className="fmt">{formatDiff(c.before.format, c.after.format).join(", ")}</div>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {changes.length > MAX_ROWS && (
        <p className="muted">…and {changes.length - MAX_ROWS} more cells.</p>
      )}
    </div>
  );
}

function Content({ text }: { text: string }) {
  if (text === "") return <span className="muted">(empty)</span>;
  return <code className={text.startsWith("=") ? "formula" : undefined}>{text}</code>;
}
