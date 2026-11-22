# Changelog

All notable changes to formulon-mcp are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.12.0] - 2026-11-23

### BREAKING

- The server runs on `@libraz/formulon` 0.19. Tool results that pass an engine
  record through (cell XF records, sheet views, column and row overrides,
  table properties, external links, page layout records, validation and
  conditional-format lists) carry the 0.19 field names; see the engine's
  [upgrade guide](https://github.com/libraz/formulon/blob/main/docs/releases/upgrading-0.19.md).
- `formulon_workbook_call` addresses the object a method belongs to. `target`
  names a handle kind of the installed engine (`Workbook`, `Sheet`, `Table`,
  `Pivot`, `PivotCache`, the drawing kinds, the `ooxml` layers) and the key each
  accessor on the way takes, or `Formulon` for the engine's module functions;
  an omitted `target` is the workbook. The callable methods are read from the
  engine's declarations, and the result carries `target`. `formulon_workbook_api`
  lists every kind's methods with their `target`.

  ```jsonc
  // before
  { "sessionId": "s", "method": "getMerges", "args": [0] }
  // after
  { "sessionId": "s", "target": { "kind": "Sheet", "sheet": 0 }, "method": "listMerges", "args": [] }
  ```

- `formulon_validation_operation` and `formulon_conditional_format_operation`
  take the engine's `DataValidationInput` and `ConditionalFormatInput`: `ranges`
  is an A1 string and every enum field is a member name. Listings use the same
  names, so a listed record can be added back as it is. Records carry the
  engine's fields: formulas start with `=`, an unset text is `null`, and
  `showInCellDropdown` replaces `showDropDown`. Conditional-format `evaluate`
  nests the data bar in `bar`, names the format `dxfIndex`, and evaluates
  time-period rules against the pinned clock, else today, where it skipped them.

- Tools that report `{ session, method, result }` (merge, comment, hyperlink,
  validation, conditional format, function lookup, session-mode formula
  evaluation) name the engine call in `method`, for example `Sheet.listMerges`,
  and `result` is that call's return value.
- Errors, the server's own included, read `<StatusName>: <message> (<context>)`
  followed by a JSON content item with `code`, `message`, `context` and, for a
  formula syntax error, `parseError`. A merge index out of range now reads
  `IndexOutOfRange: merge index out of range (index=9 count=1)`.
- `formulon_table_operation` takes `columns` as one record per column
  (`{ name?, totalsRowFunction?, totalsRowLabel?, totalsRowFormula?,
  calculatedFormula? }`) and `totals` is removed. Totals functions use the
  engine's names (`countNumbers`; `null` for none), and turning on `totalsRow`
  adds a row below the data. Blank and duplicate headers get Excel's names
  (`Column1`, `qty2`); an empty `columns[].name`, a table name that reads as a
  cell address and a table overlapping another are refused.
- `formulon_create_pivot` and `formulon_trace` report a PivotTable's spill as
  `null` or `{ blocked, cause, anchor, rows, cols, range }` (it was
  `{ engaged, ... }`) and its drawn region as `layout.rect`. A report that
  would overlap another is refused with `PreconditionFailed`.
- `formulon_hyperlink_operation` follows range-keyed hyperlinks: `list` returns
  links with their `range`, a `target` that is `null` inside the workbook, and
  normalised URLs; `add` over a cell that holds a link replaces it and writes
  the link text and Hyperlink style into the cell; `remove` of a missing link
  fails with `NotFound`.
- `formulon_inspect_layout` and `formulon_inspect_session` count and list only
  cells with content in `cellCount`, `usedRange` and `cells`; a cell that
  carries only a format is in `formattedCells` and `formattedCellCount`. A
  sheet's tab visibility is `visibility`, no longer part of `view`.
- Clearing a setting takes `null`: `numberFormat`, `printArea`,
  `printTitles.repeatRows` and `repeatCols`, header and footer sections, the raw
  `*Xml` print fragments, a table `style` and `styleName`, a sparkline
  `dateAxis`, a defined name's `formula` and `comment`, a hyperlink `target`.
  An empty string no longer removes a raw print fragment or a table style; the
  engine refuses it. Print settings read back `null`, not `""`, for no print
  area, no repeated rows or columns and an absent raw fragment.
- Calls Excel refuses now fail: an overlapping merge (`AlreadyExists`), a
  defined name Excel does not accept (`InvalidArgument`), a formula with a
  syntax error in `formulon_eval_formula` and the localize and canonicalize
  lookups (`ParserUnexpectedToken`, where it gave `#NAME?`). Formula entry
  refuses what Excel refuses, such as a wrong argument count or a value where
  a function takes only a reference. Reading a missing note returns `null`.
- `LEN`, `MID`, `FIND`, `SEARCH` and `REPLACE` count a character outside the
  Basic Multilingual Plane by the workbook's Compatibility Version: one
  position in a new workbook or a file stamped version 2, two in a file without
  the stamp or stamped version 1. Before, the default profile counted two.
- A picture held in a cell is a value of its own kind: `TYPE` gives 128 and
  `ISERROR` FALSE where the cell read `#VALUE!`. Defined-name formulas read back
  with Excel's quoting (`=Sheet2!$A$1:$C$3`, not `='Sheet2'!$A$1:$C$3`).
- A save to `.xlsb` of something the format cannot hold, such as a protected
  range whose security descriptor it cannot keep, fails with
  `IoFeatureNotWritable` and writes nothing instead of dropping it.

### Added

- `formulon_theme_operation` applies a `.thmx` theme file to a session's
  workbook or saves the workbook's theme as one.
- `formulon_image_operation` places, reads and removes a PNG or JPEG picture
  held in a cell (`insertInCell`, `getInCell`, `removeInCell`).
- `formulon_function_lookup` translates a formula's references to and from
  R1C1 (`r1c1`, with `baseRow` and `baseCol` naming the base cell).
- `formulon_recalc_session` returns `circularReferences`, every cell
  (`Sheet!A1`) in an unresolved circular reference.
- `formulon_get_cell` and `formulon_get_range` mark a formula cell
  `stale: true` when manual calculation left its value behind a change to its
  inputs; saving computes such formulas first unless `calcOnSave` is off.
- `formulon_sparkline_operation` takes and lists `rightToLeft`.
- `.xlsb` saves write PivotTables, calculated fields included, and the custom
  views of a workbook opened from `.xlsx`. A PivotTable with alternative text,
  or whose cache is saved without its data, is left out and counted in
  `losses.deferredFeatures`.

### Changed

- An omitted picture `name` or `description` takes the engine's default
  instead of an empty text.

### Fixed

- Opening a file with a profile (`profile`, or `FORMULON_MCP_PROFILE`) loads it
  under that profile, so a manual-calculation workbook keeps the values it
  cached and computes only what the file left stale; before, every formula was
  recomputed. An unknown profile is reported as `unknown excel profile` for a
  file as for a new workbook.

[Unreleased]: https://github.com/libraz/formulon-mcp/compare/v0.12.0...HEAD
[0.12.0]: https://github.com/libraz/formulon-mcp/compare/v0.11.0...v0.12.0
