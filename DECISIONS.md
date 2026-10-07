# Design decisions

A running log of choices made while building this project, and why.
Newest entries go at the bottom of each section.

## Tooling

- **Node via the existing nvm install (v22), no conda env.** This is a
  TypeScript project; `node_modules/` already isolates dependencies per project,
  so a conda environment would add nothing. Nothing is installed globally.
- **Vite instead of webpack.** Microsoft's Yeoman template uses webpack, but Vite
  needs far less config, starts instantly and shares its config with Vitest.
- **Vitest for tests.** Same transform pipeline as Vite, Jest-compatible API.
- **TypeScript pinned to ~5.9** rather than the brand-new 7.x native compiler, to
  avoid tooling incompatibilities while learning.
- **HTTPS certs from `office-addin-dev-certs`.** Office requires add-ins to be
  served over HTTPS. `vite.config.ts` reads the certs if present and falls back to
  HTTP so the in-browser demo still works without them.

## Schema

- **Zod is the single source of truth.** The TypeScript types *and* the JSON
  Schema sent to the LLM are derived from the same Zod definitions, so they cannot
  drift apart.
- **`.strict()` objects.** Unknown keys are errors, not silently dropped: if the
  model invents a field, we want the repair loop to tell it.
- **`write_formula` is single-cell only.** Writing one formula into a multi-cell
  range is ambiguous (are references relative?). Multi-cell writes go through
  `fill_range`, which is explicit.
- **`fill_range` has `values` *or* `formulaTemplate`.** Templates use `{row}` and
  `{col}` placeholders (`"=B{row}*C{row}"`). This is simpler and more predictable
  than implementing Excel's relative-reference shifting, and the preview can show
  the exact formula every cell will get.
- **JSON Schema normalisation.** Zod emits `oneOf` and `type: [..]` arrays; both
  are rewritten to `anyOf`, which more structured-output implementations accept.

## Validation

- **Two layers: shape (Zod) then meaning (custom checks).** Shape errors stop
  early; meaning checks collect every problem so the LLM can fix them all at once.
- **Validator never throws.** It returns issues with an action index so the UI
  can point at the bad action and the orchestrator can ask the LLM to repair it.
- **Errors vs warnings.** Errors block applying (bad address, wrong array size,
  unsafe formula). Warnings are shown but allowed (circular reference,
  overlapping writes, INDIRECT/OFFSET/HYPERLINK).
- **Blocked functions:** WEBSERVICE, RTD, CALL, REGISTER.ID, EXEC, plus DDE
  (`|`) and links to other workbooks. These reach outside the workbook and could
  leak data if a prompt injection in cell content tricked the model.
- **Formula scanning is a lightweight lexer, not a full parser.** It masks
  string literals, then uses regexes. Good enough for safety checks and
  self-reference warnings; documented as best-effort.
- **Default limits:** 2,000 cells per action, 10,000 total, 25 actions.
- **Whole-column refs like `A:A` are rejected** (1M+ cells, can't preview).
