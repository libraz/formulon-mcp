# formulon-mcp

MCP server for [Formulon](https://github.com/libraz/formulon). It uses the
published npm package `@libraz/formulon@0.9.0` and exposes Excel-compatible
formula and `.xlsx` workbook operations over stdio.

This is designed for agent use: open a workbook once, inspect it, mutate cells,
recalculate, read ranges, save, and close the in-memory session.

## Toolchain

- Node.js 22 via Volta
- Yarn 4 with `nodeLinker: node-modules`
- Biome 2 for format/lint
- TypeScript 6

```sh
yarn install
yarn run check
yarn run build
```

## Run

```sh
yarn run build
node ./dist/index.js
```

MCP client config:

```json
{
  "mcpServers": {
    "formulon": {
      "command": "node",
      "args": ["/absolute/path/to/formulon-mcp/dist/index.js"]
    }
  }
}
```

## Tools

- `formulon_version`: returns the loaded Formulon engine version.
- `formulon_eval_formula`: evaluates one Excel formula.
- `formulon_open_workbook`: creates a workbook session from an `.xlsx` path, or
  creates a new default workbook.
- `formulon_list_sessions`: lists open workbook sessions.
- `formulon_close_workbook`: releases a session.
- `formulon_inspect_session`: returns sheets, defined names, tables, and
  optionally sparse cell entries for an open session.
- `formulon_set_cells`: applies mutations to a session. Cells can be addressed
  with A1 refs like `Sheet1!B2` or zero-based `sheet`/`row`/`col`.
- `formulon_sheet_operation`: adds, removes, renames, or moves sheets.
- `formulon_set_defined_name`: adds, replaces, or removes workbook-scoped
  defined names.
- `formulon_edit_structure`: inserts or deletes rows and columns.
- `formulon_set_sheet_view`: sets zoom, frozen panes, or sheet-tab hidden state.
- `formulon_recalc_session`: recalculates an open session.
- `formulon_find_cells`: searches text cell values and/or formula text in a
  session.
- `formulon_replace_cells`: replaces matching text cell values and/or formula
  text in a session.
- `formulon_inspect_layout`: returns stable per-sheet layout data, including
  used ranges, merges, row/column overrides, protection, cells, calculated
  values, formulas, and optional style details.
- `formulon_detect_regions`: detects table-like regions, label-value pairs, and
  total-like fields with rule-based confidence and evidence.
- `formulon_analyze_workbook`: classifies workbook shape such as invoice, list,
  report, schedule, or form using deterministic features and evidence.
- `formulon_get_cell`: reads one cell from a session or directly from a path.
- `formulon_get_range`: reads an A1 rectangular range from a session.
- `formulon_save_session`: writes a session to `.xlsx`.
- `formulon_session_metadata`: reads function names or external links.
- `formulon_merge_operation`: lists, adds, removes, or clears merged ranges.
- `formulon_comment_operation`: gets, sets, or removes cell comments.
- `formulon_hyperlink_operation`: lists, adds, removes, or clears hyperlinks.
- `formulon_validation_operation`: lists, adds, removes, or clears data
  validations.
- `formulon_conditional_format_operation`: lists, adds, removes, clears, or
  evaluates conditional formats.
- `formulon_trace`: reads precedents, dependents, or spill info.
- `formulon_function_lookup`: lists functions and resolves function metadata or
  localized names.
- `formulon_workbook_call`: allowlisted low-level access to the Formulon
  `Workbook` API for advanced features, including PivotTables, PivotCaches,
  styles, merges, comments, hyperlinks, validations, conditional formatting,
  dependency graph queries, function metadata, and spill info.
- `formulon_inspect_workbook`: one-shot workbook summary from path.
- `formulon_update_workbook`: one-shot load/create, mutate, recalc, save.

Unless A1 notation is used, sheet, row, and column indexes are zero-based to
match the Formulon API.

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

## License

Apache-2.0. See [LICENSE](./LICENSE).
