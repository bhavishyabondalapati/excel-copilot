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
- **Blocked functions:** anything that can send workbook data outside the
  workbook or run external code, so a prompt injection in cell text can't use the
  model to leak data. Reviewed against Microsoft's function docs in October 2026:
  - network / cloud: WEBSERVICE, IMAGE (any URL), IMPORTTEXT and IMPORTCSV (URL or
    local file), COPILOT (sends cell context to an AI service that can search the
    web), TRANSLATE and DETECTLANGUAGE (Microsoft Translator), STOCKHISTORY
    (internet data), PY (Python in Excel runs in Microsoft's cloud);
  - external code / servers: RTD, CALL, REGISTER.ID, EXEC;
  - plus DDE (`|`) and links to other workbooks.
  IMAGE was missed in the first version and added after review. Names are
  matched case-insensitively, with or without Excel's `_xlfn.` / `_xlws.` storage
  prefixes. Each blocked function carries a reason that is shown to the user and
  sent back to the model in the repair loop. The same list is put into the system
  prompt so the model avoids these functions in the first place.
- **Reviewed but not blocked:** HYPERLINK (warned: leaks only if the user clicks),
  CUBE* functions (they query connections already in the workbook and can't target
  an arbitrary URL), FILTERXML and ENCODEURL (no network on their own).
  **Known gap:** custom functions from other installed add-ins (e.g.
  `CONTOSO.FETCH`) can do anything; they can't be told apart from built-in dotted
  names like `NORM.DIST` without an allow-list of every built-in function.
- **Formula scanning is a lightweight lexer, not a full parser.** It masks
  string literals, then uses regexes. Good enough for safety checks and
  self-reference warnings; documented as best-effort.
- **Default limits:** 2,000 cells per action, 10,000 total, 25 actions.
- **Whole-column refs like `A:A` are rejected** (1M+ cells, can't preview).

## Planner, preview and undo

- **Planner is a pure function of (actions, snapshot).** The caller reads the
  cells first; the planner simulates every action on a copy and diffs the result.
  No Excel needed to test preview or undo.
- **Formulas are not evaluated in the preview.** The "after" column shows the
  formula text; Excel calculates the value once applied. Writing a formula
  evaluator would be a big project of its own.
- **Unchanged cells are left out of the diff** (writing "Item" over "Item").
- **Undo restores whole touched ranges, content *and* format, every time.** Even
  a content-only action can change formatting (Excel auto-formats typed dates),
  so restoring both is more correct and still cheap. Formats go first so a text
  number format ("@") is back in place before contents are rewritten.
- **Drift detection before undo.** If the user edited a cell after applying, undo
  would silently overwrite that edit, so the UI warns first. Drift is measured
  against a re-read taken *right after applying*, not against what the plan meant
  to write: Excel normalises input (`=sum(a1)` → `=SUM(A1)`, number formats), and
  comparing with the plan's intent would cry wolf on every formula. (The first
  version compared with the plan; caught in self-review, fixed with a regression
  test that simulates Excel's upper-casing.)
- **`WorkbookGateway` interface ("ports and adapters").** The core only knows this
  interface. `MemoryWorkbook` implements it for tests and the browser demo; the
  Office.js adapter implements it for real Excel.

## LLM providers

- **`LLMProvider` interface with one method, `complete()`.** Returns raw text;
  parsing and validation live in the core so every provider is treated the same.
- **Anthropic via the official `@anthropic-ai/sdk`** with
  `dangerouslyAllowBrowser: true` (the task pane is a browser page). Default model
  `claude-opus-5-5`, `output_config.effort: "medium"`, structured output via
  `output_config.format: {type: "json_schema", ...}`. Server-side refusal
  fallbacks (`fallbacks: "default"`, beta `server-side-fallback-2026-07-01`) are on
  so a safety-classifier decline retries on Anthropic's recommended model.
  `stop_reason` is checked (`refusal`, `max_tokens`) before reading content.
- **Gemini via plain `fetch` to `generateContent`** rather than Google's SDK: one
  endpoint, easy to stub in tests, one fewer dependency. JSON mode uses
  `responseMimeType: "application/json"` + `responseJsonSchema`. Default model
  `gemini-3.8-flash` (from the Gemini models page, October 2026). The key is sent
  in the `x-goog-api-key` header, never in the URL.
- **Neither real provider was called live while building** (no API keys were
  available to the agent). Both are tested against stubbed `fetch` responses
  shaped like the documented API responses.
- **`FakeProvider`** replays scripted replies and records requests. Used by the
  tests and selectable in the UI so the whole flow works without a key.
- **Errors are normalised to `ProviderError` with a `kind`** (auth, rate_limit,
  refusal, truncated, network, bad_response, other) so the UI can show a useful
  message whatever the provider.
- **API keys are typed into the Settings panel and kept in the browser's
  `localStorage`**, never in source files or the bundle. Fine for a personal
  add-in; a shared deployment should put a small server in front of the LLM APIs.

## Orchestrator

- **One automatic repair attempt.** If the reply fails validation, the model sees
  its own reply plus the list of problems and gets one more try. Costs at most one
  extra call and fixes most size/format slips.
- **All-or-nothing.** If any action has an error the whole reply is rejected; we
  never apply "the valid half" of a plan.
- **Stale-preview check.** `apply()` re-reads the cells; if anything changed since
  the preview, it refuses and asks for a fresh preview. That same read is the undo
  snapshot, so undo always restores what was there immediately before writing.
- **Sheet data is marked as data in the prompt** (`<workbook>` vs `<request>` tags
  plus a rule to ignore instructions inside cells) to reduce prompt injection
  from spreadsheet contents. The validator's blocked-function list is the second
  line of defence.
- **"Explain this formula" reuses the same pipeline** with a fixed request and
  only keeps the `explain` actions, so it can never edit the sheet.

## Excel adapter

- **ExcelApi 1.9 minimum** (declared in the manifest) for
  `getCellProperties`/`setCellProperties`, which read and restore per-cell
  formats in one call. Excel for Mac (Microsoft 365) supports it.
- **Writes go through `range.formulas`**, which accepts formulas and constants
  alike; `""` clears a cell. All ops in a plan are queued and sent with a single
  `context.sync()`.
- **"No fill" is restored with `format.fill.clear()`** per cell, because it can't
  be expressed as a colour in `setCellProperties`.
- **Context sent to the LLM is capped at 30 rows × 15 columns** of the used range
  (plus the selection), with a "truncated" note, to keep prompts small.
- **The adapter has no unit tests.** It needs a running Excel. It is kept thin on
  purpose (it only moves data) so the tested core holds the logic. Verify in Excel
  after sideloading.

## UI

- **Plain React + CSS, no Fluent UI.** Fewer dependencies and easier to read while
  learning; colours follow Excel's green and support dark mode.
- **Browser demo mode.** When `Office.onReady` reports no Excel host, the app uses
  `MemoryWorkbook` with sample data, a mini grid, and a keyword-based demo
  provider. This makes the UI testable in a normal browser and in jsdom.
- **Undo is last-in-first-out only.** Undoing an older change while a newer one
  overlaps it would restore inconsistent data, so only the newest entry has an
  Undo button.
- **Undo history is in memory.** It is lost when the pane closes. Persisting it
  (e.g. in document settings) is a possible extension.
- **Manifest version is 1.0.0.0**: the validator rejects versions below 1.0.

- **Gemini is the default provider** (changed at the user's request). Claude and
  Demo stay selectable in Settings. Until a Gemini key is saved, the pane shows
  "Add a Gemini API key" and disables Preview, so trying the browser demo now
  means picking *Demo* in Settings first. The UI tests pick Demo explicitly.

## Project conventions

- **No `requirements.txt` / `.env.example`.** This is a Node project:
  `package.json` + `package-lock.json` play the role of `requirements.txt`. No
  `.env` is used because API keys are entered in the Settings panel at runtime,
  never written to files. `.gitignore` still excludes `.env` files defensively.
- **`npm run check`** (typecheck + tests) exists because a piped test command once
  hid a failing exit code from a commit; the next commit fixed it.

## Verification done while building

- `npm run check`: typecheck clean, 101 tests passing.
- `vite build` succeeds.
- `office-addin-manifest validate manifest.xml`: "The manifest is valid."
- Demo mode clicked through in a browser: preview → apply → explain → undo ×2
  restored the sheet; missing-key warning shows for Anthropic.
- **Not done:** loading the add-in inside real Excel, and live calls to the Claude
  or Gemini APIs (no Excel automation or API keys available to the agent).
