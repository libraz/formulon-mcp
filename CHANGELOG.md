# Changelog

All notable changes to formulon-mcp are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### BREAKING

- The server runs on `@libraz/formulon` 0.19. Tool results that pass an engine
  record through (cell XF records, sheet views, column and row overrides,
  table properties, external links, page layout records, validation and
  conditional-format lists) carry the 0.19 field names; see the engine's
  [upgrade guide](https://github.com/libraz/formulon/blob/main/docs/releases/upgrading-0.19.md).
- `formulon_workbook_call` addresses the object a method belongs to. `target`
  names a handle kind of the installed engine (`Workbook`, `Sheet`, `Table`,
  `Pivot`, `PivotCache`, `Chart` and the other drawing kinds, the `ooxml`
  layers) and the key each accessor on the way takes, or `Formulon` for the
  engine's module functions (`compareText`, `getCollationClasses`,
  `getLocaleFacts`, ...); an omitted `target` is the workbook. The callable
  methods are read from the engine's declarations instead of a fixed list, and
  the result carries `target`.

  ```jsonc
  // before
  { "sessionId": "s", "method": "getMerges", "args": [0] }
  // after
  { "sessionId": "s", "target": { "kind": "Sheet", "sheet": 0 }, "method": "listMerges", "args": [] }
  ```

- `formulon_workbook_api` lists the methods of every handle kind, each with
  its `target` (`kind` and `keys`), so an entry feeds `formulon_workbook_call`
  directly. `describe` takes `Kind.method` (`Sheet.addValidation`) or a bare
  name that only one kind has.
- `formulon_validation_operation` takes `validation` as the engine's
  `DataValidationInput`, with `ranges` and its enum fields given by member
  name. The engine checks the record, and its refusal is the tool error.

  ```jsonc
  // before
  { "ranges": [{ "firstRow": 0, "firstCol": 2, "lastRow": 9, "lastCol": 2 }],
    "type": 1, "op": 0, "formula1": "1", "formula2": "10" }
  // after
  { "ranges": "C1:C10", "type": "Whole", "operator": "Between", "formula1": "1", "formula2": "10" }
  ```

- `formulon_conditional_format_operation` takes `rule` as the engine's
  `ConditionalFormatInput`, with `ranges` and its enum fields (rule type,
  operator, time period, threshold types, icon sets, data-bar axis and
  direction) given by member name. `evaluate` without `todaySerial` evaluates
  time-period rules against the pinned clock, else today; before, it
  skipped them.

  ```jsonc
  // before
  { "sqref": [{ "firstRow": 0, "firstCol": 0, "lastRow": 9, "lastCol": 0 }],
    "type": 0, "formula1": "=A1>0" }
  // after
  { "ranges": "A1:A10", "type": "Expression", "formula1": "=A1>0" }
  ```

- Tools that report `{ session, method, result }` (merge, comment, hyperlink,
  validation, conditional format, function lookup, session-mode formula
  evaluation) name the engine call they made in `method`, for example
  `Sheet.listMerges` or `Formulon.listFunctionNames`, and `result` is that
  call's return value.
- An engine failure is reported as `<StatusName>: <message> (<context>)`,
  followed by a JSON content item with `code`, `message`, `context` and, for a
  formula syntax error, `parseError` with its positions.

- `formulon_table_operation` takes `columns` as one record per column,
  `{ name?, totalsRowFunction?, totalsRowLabel?, totalsRowFormula?,
  calculatedFormula? }`, and `totals` is removed. Totals functions use the
  engine's names (`countNumbers`; `null` for none). The `table` record in a
  result is the engine's (`headerRow` and `totalsRow` are booleans). Turning
  on `totalsRow` adds a row below the data instead of turning the last data
  row into the totals row. Blank and duplicate headers get the names Excel
  gives them (`Column1`, `qty2`) instead of failing, and a table name that
  reads as a cell address or a table overlapping another is refused.

  ```jsonc
  // before
  { "columns": ["Item", "Qty"], "totals": [{ "column": "Qty", "function": "sum" }] }
  // after
  { "columns": [{ "name": "Item" }, { "name": "Qty", "totalsRowFunction": "sum" }] }
  ```

- `formulon_create_pivot` and `formulon_trace` report a PivotTable's spill as
  `null` or `{ blocked, cause, anchor, rows, cols, range }` (it was
  `{ engaged, ... }`) and its drawn region as `layout.rect` (it was `top`,
  `left`, `rows`, `cols`). `formulon_create_pivot` returns `spill`. The
  report's cells read through `formulon_get_cell` and `formulon_get_range`,
  and a report that would overlap another is refused with
  `PreconditionFailed`.
- `formulon_hyperlink_operation` follows the engine's range-keyed hyperlinks.
  `list` returns an array of links, each with its `range` and a `target` that
  is `null` for a link to a location in the workbook, and URLs come back
  normalised (`https://example.com/`). `add` on cells that already hold a link
  replaces it, and writes the link text and the Hyperlink style into the cell.
  `remove` of a link that does not exist fails with `NotFound`, and `add`
  without a target or a location is refused.
- Validation and conditional-format listings carry the engine's records:
  formulas start with `=`, an unset text is `null`, `showInCellDropdown`
  replaces `showDropDown`, and `allowBlank`, `showInputMessage` and
  `showErrorMessage` are `true` unless set otherwise. The conditional-format
  `evaluate` result nests the data bar in `bar` and names the format
  `dxfIndex`.
- Calls Excel refuses now fail instead of succeeding: an overlapping merge
  (`AlreadyExists`), a defined name Excel does not accept (`InvalidArgument`),
  and a formula with a syntax error in `formulon_eval_formula` and in the
  localize and canonicalize lookups (`ParserUnexpectedToken`, where it gave
  `#NAME?`). Reading a note that does not exist returns `result: null`
  instead of failing, and a note without an author reports `null`.
- `formulon_inspect_layout` and `formulon_inspect_session` count and list only
  cells with content in `cellCount`, `usedRange` and `cells`; a cell that
  carries only a format is listed in `formattedCells` (with `includeStyles`)
  and counted in `formattedCellCount`. A sheet's tab visibility is
  `visibility` (`visible`, `hidden` or `veryHidden`), next to `view`, which no
  longer carries it.
- `formulon_print_settings` raw records write booleans as `1` and `0`, locale
  facts give `dateOrder` as the engine's number, and the first picture placed
  on a sheet is named `Picture 1`.

### Added

- `formulon_function_lookup` translates a formula's references to and from
  R1C1 (`r1c1`, with `baseRow` and `baseCol` naming the cell the references
  are relative to).

- `formulon_sparkline_operation` takes and lists `rightToLeft`, which plots a
  group's points from right to left.
- `formulon_recalc_session` returns `circularReferences`, every cell
  (`Sheet!A1`) that sits in an unresolved circular reference.
- `formulon_get_cell` and `formulon_get_range` mark a formula cell
  `stale: true` when manual calculation left its value behind a change to its
  inputs. Saving the session computes such formulas first, which clears the
  mark, unless the workbook's `calcOnSave` is off.
- `formulon_theme_operation` applies a `.thmx` theme file to a session's
  workbook or saves the workbook's theme as one.
- `formulon_image_operation` places, reads and removes a PNG or JPEG picture
  held in a cell (`insertInCell`, `getInCell`, `removeInCell`), as Excel's
  Place in Cell does; formulas that read such a cell see `#VALUE!`.

### Changed

- Errors the server itself raises (an index out of range, a missing or invalid
  argument, an unknown session or sheet) now use the same
  `<StatusName>: <message> (<context>)` text and second JSON content item
  (`{code, message, context}`) as engine errors. Removing a merge that does not
  exist used to read `merge index out of range: 9 (1 merges)`; it now reads
  `IndexOutOfRange: merge index out of range (index=9 count=1)`.

- `LEN`, `MID`, `FIND`, `SEARCH` and `REPLACE` count a character outside the
  Basic Multilingual Plane (an emoji, a rare CJK ideograph) by the workbook's
  Compatibility Version, as the engine now does: a new workbook and a file
  stamped with version 2 count it as one position, and a file without the
  stamp or stamped with version 1 counts it as two. Before, the default
  profile always counted two.
- `formulon_save_session` reports that `.xlsb` output leaves out a
  PivotTable built in the session and the custom views of a workbook opened
  from `.xlsx`; tables, print settings, notes, AutoFilter and the other
  features the tools author are kept. An `.xlsb` save of a data bar whose
  minimum or maximum is a formula fails with `IoFeatureNotWritable` and
  writes nothing.
