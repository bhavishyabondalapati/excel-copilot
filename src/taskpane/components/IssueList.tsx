import type { ValidationIssue } from "../../core/validator";

export function IssueList({ issues }: { issues: ValidationIssue[] }) {
  if (issues.length === 0) return null;
  return (
    <ul className="issues">
      {issues.map((i, n) => (
        <li key={n} className={i.severity}>
          <strong>{i.severity === "error" ? "Error" : "Warning"}</strong>
          {i.actionIndex !== null && ` (action ${i.actionIndex + 1})`}: {i.message}
        </li>
      ))}
    </ul>
  );
}
