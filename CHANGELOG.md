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

### Added

- `formulon_sparkline_operation` takes and lists `rightToLeft`, which plots a
  group's points from right to left.
- `formulon_recalc_session` returns `circularReferences`, every cell
  (`Sheet!A1`) that sits in an unresolved circular reference.
- `formulon_get_cell` and `formulon_get_range` mark a formula cell
  `stale: true` when manual calculation left its value behind a change to its
  inputs.
- `formulon_theme_operation` applies a `.thmx` theme file to a session's
  workbook or saves the workbook's theme as one.
- `formulon_image_operation` places, reads and removes a PNG or JPEG picture
  held in a cell (`insertInCell`, `getInCell`, `removeInCell`), as Excel's
  Place in Cell does; formulas that read such a cell see `#VALUE!`.

### Changed

- `LEN`, `MID`, `FIND`, `SEARCH` and `REPLACE` count a character outside the
  Basic Multilingual Plane (an emoji, a rare CJK ideograph) by the workbook's
  Compatibility Version, as the engine now does: a new workbook and a file
  stamped with version 2 count it as one position, and a file without the
  stamp or stamped with version 1 counts it as two. Before, the default
  profile always counted two.
- `formulon_save_session` reports that `.xlsb` output leaves out a
  PivotTable built in the session and a note on the third or later sheet;
  tables, print settings, notes on the first two sheets, AutoFilter and the
  other features the tools author are kept.
- `formulon_inspect_workbook` and `formulon_get_cell` with a `path` and
  `recalc: false` read the values the file cached, as Excel shows them in
  Manual calculation; the workbook is not recalculated on open.
