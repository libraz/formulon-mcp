# formulon-mcp

MCP server for [Formulon](https://github.com/libraz/formulon). It uses the
published npm package `@libraz/formulon@0.11.1` and exposes Excel-compatible
formula and `.xlsx` / `.xlsb` workbook operations over stdio.

This is designed for agent use: open a workbook once, inspect it, mutate cells,
recalculate, read ranges, save, and close the in-memory session. Documents can
be authored as well as edited — fonts, fills, ruled borders, number formats and
the page setup a printed invoice or receipt needs are all writable from a blank
workbook.

Authoring works at two levels. `formulon_build_document` takes a document as a
stack of blocks — title, fields, table, summary — and resolves every position,
rule and cross-reference itself, so nothing has to compute which row the total
lands on. The primitives (`set_cells`, `style_range`, `print_settings`) then
refine the result, using the A1 map the build call hands back.

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
- `formulon_set_sheet_view`: sets zoom, frozen panes, or sheet-tab visibility.
  `visibility` reaches all three states — Excel leaves a `veryHidden` sheet out
  of its "Unhide" dialog, which `hidden` cannot express.
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
- `formulon_build_document`: writes a whole document from a vertical stack of
  blocks — `title`, `text`, `fields`, `table`, `summary`, `spacer` — resolving
  positions, ruling, number formats, column widths, merges and the print area
  from the layout. Blocks reference each other by name rather than by address:
  a table column registers as `{table.<header>}` over its body range, and a
  field or summary item registers under its label, so `=SUM({table.Amount})`
  and `={Subtotal}+{Tax}` bind to the right cells once the layout is known —
  a `SUM` cannot end up one row short. `sameRow` puts a block beside the
  previous one instead of below it. The response maps every block and name to
  A1, which is what makes the result refinable with the tools below. It carries
  no document semantics: labels, tax rules and totals are the caller's own
  formulas.
- `formulon_style_range`: applies fonts, fills, borders, number formats, and
  alignment across an A1 range using names and `#RRGGBB` colors instead of
  OOXML ordinals. `border.all` rules every cell, `border.outline` boxes the
  range, and both together give a gridded table with a heavier frame. Blank
  cells are materialized so an empty ruled box renders.
- `formulon_default_font`: reads or redeclares the workbook default font — the
  one every cell that was never styled resolves to. A new workbook is seeded
  with Excel's Calibri 11, so a Japanese document declares its own default once
  here instead of styling every cell.
- `formulon_print_settings`: reads or sets page setup, margins, print options,
  header/footer, print area, print titles, and manual page breaks. A read also
  reports the resulting `pageCount`, so a layout can be checked without saving
  the file.
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
  validations. An omitted boolean field on a rule defaults to false, so a rule
  that should accept empty cells has to spell `allowBlank: true`.
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
  queries, function metadata, spill info, phonetic guides — whole-cell or span
  by span, with the kana form and alignment they render in — print pagination,
  the raw print-settings XML fragments, and the workbook clock pin.
- `formulon_inspect_workbook`: one-shot workbook summary from path.
- `formulon_update_workbook`: one-shot load/create, mutate, recalc, save.

Unless A1 notation is used, sheet, row, and column indexes are zero-based to
match the Formulon API. Operation tools (`merge`, `comment`, `hyperlink`,
`validation`, `conditional_format`, `trace`, `dimension`, `print_settings`)
accept a sheet name in their `sheet` argument as well as a zero-based index.

`style_range` treats each property as a delta: a cell keeps whatever the call
does not state, so ruling a table does not undo the number format already on
its amount column. Style a document in passes — the grid over the whole table
first, then the header row, then a total cell. Naming a typeface also cuts the
font's theme link, because a theme-linked font is re-resolved from the theme
rather than keeping the name it was given; `default_font` is the workbook-wide
counterpart, reaching the cells no style names at all.

`build_document` is the same idea one level up: reach for it first when writing
a document from scratch, and use the primitives on the ranges it names when the
result needs adjusting. It is not a substitute for them — it only knows how to
stack blocks, and stays out of what any particular document means.

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

Build a whole document in one call — `build_document`:

```json
{
  "sessionId": "work",
  "print": "a4-portrait-fit",
  "blocks": [
    { "type": "title", "text": "Invoice" },
    { "type": "spacer" },
    { "type": "text", "name": "to", "text": "Sample Co.", "span": 2 },
    {
      "type": "fields",
      "align": "right",
      "sameRow": true,
      "items": [
        { "label": "No.", "value": "INV-0001" },
        { "label": "Date", "value": "2026-08-22", "format": "date" }
      ]
    },
    { "type": "spacer" },
    {
      "type": "table",
      "columns": [
        { "header": "Item", "key": "name", "width": 30 },
        { "header": "Qty", "key": "qty", "format": "number", "align": "right" },
        { "header": "Unit", "key": "unit", "format": "number", "align": "right" },
        { "header": "Amount", "formula": "={qty}*{unit}", "format": "number", "align": "right" }
      ],
      "rows": [
        { "name": "Design", "qty": 3, "unit": 120000 },
        { "name": "Build", "qty": 5, "unit": 98000 }
      ]
    },
    { "type": "spacer" },
    {
      "type": "summary",
      "items": [
        { "label": "Subtotal", "formula": "=SUM({table.Amount})", "format": "number" },
        { "label": "Tax", "formula": "=ROUND({Subtotal}*0.1,0)", "format": "number" },
        { "label": "Total", "formula": "={Subtotal}+{Tax}", "format": "number", "emphasis": true }
      ]
    }
  ]
}
```

The response reports where everything landed, which is what the following calls
address:

```json
{
  "range": "B2:E13",
  "width": 4,
  "pageCount": 1,
  "names": {
    "title": "B2:E2",
    "to": "B4:C4",
    "Date": "E5",
    "table.header": "B7:E7",
    "table.body": "B8:E9",
    "table.Amount": "E8:E9",
    "Subtotal": "E11",
    "Tax": "E12",
    "Total": "E13"
  }
}
```

A `{name}` that matches nothing is an error rather than a formula that quietly
points at the wrong range. Braces holding a comma or semicolon are an Excel
array constant (`{1,2;3,4}`) and pass through untouched.

Rule a table and box it — `style_range`:

```json
{
  "sessionId": "work",
  "range": "Sheet1!B2:E12",
  "style": {
    "border": { "all": "thin", "outline": { "style": "medium", "color": "#1F4E79" } }
  }
}
```

Then give the header row its own band, without disturbing the ruling:

```json
{
  "sessionId": "work",
  "range": "Sheet1!B2:E2",
  "style": {
    "font": { "bold": true, "color": "#FFFFFF" },
    "fill": { "color": "#1F4E79" },
    "align": { "horizontal": "center" }
  }
}
```

Make it printable — `print_settings`:

```json
{
  "sessionId": "work",
  "pageSetup": { "orientation": "portrait", "paperSize": 9, "fitToPage": true, "fitToWidth": 1 },
  "margins": { "left": 0.6, "right": 0.6 },
  "printArea": "B2:E40",
  "printTitles": { "repeatRows": "2:2" },
  "headerFooter": { "oddFooter": "&C&P / &N" }
}
```

Header and footer text uses Excel's own codes — `&L`/`&C`/`&R` pick the section,
`&P` is the page number, `&N` the page count, and `&&` a literal ampersand. The
engine escapes them for the file, so no caller assembles XML.

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
