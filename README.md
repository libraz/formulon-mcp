# formulon-mcp

MCP server for [Formulon](https://github.com/libraz/formulon). It uses the
published npm package `@libraz/formulon@0.10.0` and exposes Excel-compatible
formula and `.xlsx` / `.xlsb` workbook operations over stdio.

This is designed for agent use: open a workbook once, inspect it, mutate cells,
recalculate, read ranges, save, and close the in-memory session.

## Install

Requires Node.js 22+. No clone needed — `npx` fetches and runs the server on
demand. The CLI binary is `formulon-mcp`.

### Claude Code

```sh
claude mcp add --scope user formulon -- npx -y @libraz/formulon-mcp
```

Verify with `claude mcp list` — `formulon` should report `✓ Connected`.

### Codex CLI

Add to `~/.codex/config.toml`:

```toml
[mcp_servers.formulon]
command = "npx"
args = ["-y", "@libraz/formulon-mcp"]
```

### Claude Desktop

Add to `claude_desktop_config.json`
(`~/Library/Application Support/Claude/` on macOS,
`%APPDATA%\Claude\` on Windows):

```json
{
  "mcpServers": {
    "formulon": {
      "command": "npx",
      "args": ["-y", "@libraz/formulon-mcp"]
    }
  }
}
```

### Other MCP clients

Any stdio-capable MCP client works. Point it at `npx -y @libraz/formulon-mcp`,
or run `formulon-mcp` directly after `npm install -g @libraz/formulon-mcp`.

### Interactive setup (optional)

If you'd rather not edit config files by hand for Codex CLI or Claude Desktop,
run the bundled installer:

```sh
npx -y @libraz/formulon-mcp init
```

Pick one or more targets (comma-separated, e.g. `1,3,4`):

- **Claude Code — user** (`~/.claude.json`)
- **Claude Code — project** (`./.mcp.json`)
- **Codex CLI** (`~/.codex/config.toml`)
- **Claude Desktop** (`claude_desktop_config.json` at the platform path above)

Re-running `init` safely replaces the existing `formulon` entry without
touching other servers. Restart your MCP client to pick up the change.

To remove the entry later:

```sh
npx -y @libraz/formulon-mcp uninstall
```

It drops only the `formulon` server; other entries are kept.

### From source

For development or to pin a fork, clone and build instead of using npm:

```sh
git clone https://github.com/libraz/formulon-mcp.git
cd formulon-mcp
yarn install
yarn run build
```

Then register the absolute path to `dist/index.js`, e.g.:

```sh
claude mcp add --scope user formulon node /absolute/path/to/formulon-mcp/dist/index.js
```

Or install the latest `main` directly without a local clone:

```sh
npx -y github:libraz/formulon-mcp
```

## Development

- Node.js 22 via mise
- Yarn 4 with `nodeLinker: node-modules`
- Biome 2 for format/lint
- TypeScript 7
- Vitest for tests (`yarn test`, `yarn test:watch`, `yarn test:coverage`)

```sh
yarn install
yarn run check
yarn run build
yarn run test
```

Run the server directly for local debugging:

```sh
node ./dist/index.js
```

## Tools

- `formulon_version`: returns the loaded Formulon engine version and the MCP
  server version.
- `formulon_eval_formula`: evaluates one Excel formula. With `sessionId`, it
  evaluates read-only against an open workbook, resolving references, defined
  names, and `ROW()`/`COLUMN()` anchored at the given cell.
- `formulon_open_workbook`: creates a workbook session from an `.xlsx` / `.xlsb`
  path, or creates a new default workbook. Anything the reader could not decode
  is reported as `loadLosses` on the session.
- `formulon_list_sessions`: lists open workbook sessions.
- `formulon_close_workbook`: releases a session.
- `formulon_inspect_session`: returns sheets, defined names, tables, and
  optionally sparse cell entries for an open session.
- `formulon_set_cells`: applies mutations to a session. Cells can be addressed
  with A1 refs like `Sheet1!B2` or zero-based `sheet`/`row`/`col`. Writes are
  bounds-checked against Excel's grid, and formula cells that evaluate to an
  error are reported back in `errorCells`.
- `formulon_set_range`: writes a 2D block of values from an anchor cell; each
  element's JSON type picks the cell type, `{"f":"=…"}` writes a formula, and
  `null` skips a cell. Much more compact than `set_cells` for tables.
- `formulon_sheet_operation`: adds, removes, renames, or moves sheets.
- `formulon_set_defined_name`: adds, replaces, or removes defined names, either
  workbook-scoped or local to one sheet. Print settings such as
  `_xlnm.Print_Area` must be sheet-scoped, since Excel ignores a
  workbook-scoped one without reporting an error.
- `formulon_edit_structure`: inserts or deletes rows and columns.
- `formulon_set_sheet_view`: sets zoom, frozen panes, or sheet-tab hidden state.
- `formulon_recalc_session`: recalculates an open session.
- `formulon_find_cells`: searches cell values (text, numbers, booleans) and/or
  formula text in a session.
- `formulon_replace_cells`: replaces matching text cell values and/or formula
  text in a session.
- `formulon_inspect_layout`: returns stable per-sheet layout data, including
  used ranges, merges, row/column overrides, protection, cells, calculated
  values, formulas, and optional style details.
- `formulon_detect_regions`: detects table-like regions, label-value pairs, and
  total-like fields with rule-based confidence and evidence.
- `formulon_analyze_workbook`: classifies workbook shape such as invoice, list,
  report, schedule, or form using deterministic features and evidence.
- `formulon_get_cell`: reads one cell from a session or directly from a path,
  including its formula text (empty for constants). Date/currency/percent cells
  carry a decoded `formatted` string alongside the raw value.
- `formulon_get_range`: reads an A1 rectangular range from a session as a sparse
  cell list — blanks omitted, clipped to the sheet's used range, and capped at
  `maxCells`. Set `includeFormulas` to annotate computed cells. Formatted
  numeric cells (dates, currency, percent) carry a decoded `formatted` string.
- `formulon_dimension_operation`: lists column-width / row-height overrides, or
  sets width/height, hidden, or outline level. Columns act on an inclusive
  `[first, last]` span; rows act on a single row index.
- `formulon_save_session`: writes a session out. The container follows the
  output extension — `.xlsb` writes XLSB, anything else XLSX — and whatever the
  writer had to drop or downgrade is reported as `losses`.
- `formulon_session_metadata`: reads function names or external links.
- `formulon_merge_operation`: lists, adds, removes, or clears merged ranges.
- `formulon_comment_operation`: lists, gets, sets, or removes cell comments.
- `formulon_hyperlink_operation`: lists, adds, removes, or clears hyperlinks. An
  added link covers one cell, or the rectangle through `lastRow`/`lastCol`; pass
  `location` (with an empty `target`) for an in-workbook destination.
- `formulon_validation_operation`: lists, adds, removes, or clears data
  validations.
- `formulon_conditional_format_operation`: lists, adds, removes, clears, or
  evaluates conditional formats.
- `formulon_trace`: reads precedents, dependents, or spill info. Results carry
  A1 references with sheet names, not raw indices.
- `formulon_function_lookup`: lists functions and resolves function metadata or
  localized names.
- `formulon_workbook_call`: allowlisted low-level access to the Formulon
  `Workbook` API for advanced features, including PivotTables, PivotCaches,
  worksheet tables and their AutoFilter, styles and differential formats,
  merges, comments, hyperlinks, validations, conditional formatting, sheet
  display flags such as gridlines and the page-layout view, dependency graph
  queries, function metadata, spill info, phonetic guides, print pagination,
  and the workbook clock pin.
- `formulon_inspect_workbook`: one-shot workbook summary from path.
- `formulon_update_workbook`: one-shot load/create, mutate, recalc, save.

Unless A1 notation is used, sheet, row, and column indexes are zero-based to
match the Formulon API. Operation tools (`merge`, `comment`, `hyperlink`,
`validation`, `conditional_format`, `trace`, `dimension`) accept a sheet name in
their `sheet` argument as well as a zero-based index.

Cell values are returned as `{kind, value}` envelopes. Formula cells are
recalculated when a session is opened, so the first read returns the computed
value rather than a blank. Error cells include an `errorName` Excel literal
(`#DIV/0!`, `#REF!`, `#NAME?`, `#SPILL!`, …) beside the numeric `errorCode`; the
literal comes from the engine, so it covers every error kind the engine can
produce. Cells whose
number format is a date, currency, or percent carry a `numberFormat` code, a
`formatKind`, and — for dates and percentages — a decoded `formatted` string
(for example an ISO date), while `value` keeps the raw Excel serial / number.
The full cell style is available through `inspect_layout` with `includeStyles`.

## Agent Workflow

Open a new workbook:

```json
{
  "path": "input.xlsx",
  "sessionId": "work"
}
```

Set cells:

```json
{
  "sessionId": "work",
  "mutations": [
    { "type": "number", "a1": "Sheet1!A1", "value": 41 },
    { "type": "formula", "a1": "Sheet1!B1", "formula": "=A1+1" }
  ],
  "recalc": true
}
```

Read a range:

```json
{
  "sessionId": "work",
  "range": "Sheet1!A1:B1"
}
```

Search and replace:

```json
{
  "sessionId": "work",
  "query": "budget",
  "target": "both",
  "matchCase": false
}
```

```json
{
  "sessionId": "work",
  "query": "budget",
  "replacement": "forecast",
  "target": "texts",
  "recalc": true
}
```

Save:

```json
{
  "sessionId": "work",
  "outputPath": "output.xlsx"
}
```

Low-level API access:

```json
{
  "sessionId": "work",
  "method": "addMerge",
  "args": [0, { "firstRow": 0, "firstCol": 0, "lastRow": 0, "lastCol": 2 }]
}
```

The low-level tool only dispatches methods explicitly allowlisted in
`src/sessions.ts`. It does not evaluate arbitrary code.

Two low-level calls are worth knowing about:

- `setPinnedNow` (`[year, month, day, hour, minute, second]`) pins the workbook
  clock, so `NOW()`, `TODAY()` and the pivot relative-period filters all agree
  on one instant instead of each reading the host clock. The pin is model state:
  it is not saved, and `clearPinnedNow` returns to the host clock.
- `pivotCacheSetWorksheetSource` must be called on a PivotCache built through the
  API before saving. A cache with no declared worksheet source produces a file
  Excel offers to repair, so the writer rejects it.

## License

Apache-2.0. See [LICENSE](./LICENSE).
