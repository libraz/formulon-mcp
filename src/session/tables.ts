import {
  type Table,
  type TableColumnPatch,
  type TableInput,
  type TableStyleOptionsPatch,
  TableTotalsRowFunction,
} from "@libraz/formulon";
import { normalizeRange, parseRangeRef } from "../a1.js";
import { findSheetIndex, okStatus, rectA1, resultToJson, sheetAt } from "../formulon.js";
import { assertInGrid, rangeA1, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

export type TableOperation = "list" | "create" | "update" | "remove";

const TOTALS_FUNCTIONS = {
  none: TableTotalsRowFunction.None,
  sum: TableTotalsRowFunction.Sum,
  min: TableTotalsRowFunction.Min,
  max: TableTotalsRowFunction.Max,
  average: TableTotalsRowFunction.Average,
  count: TableTotalsRowFunction.Count,
  countNums: TableTotalsRowFunction.CountNumbers,
  stdDev: TableTotalsRowFunction.StdDev,
  var: TableTotalsRowFunction.Var,
  custom: TableTotalsRowFunction.Custom,
} as const;

export type TableTotalsFunctionName = keyof typeof TOTALS_FUNCTIONS;

/** One column's totals-row cell: a function, a text label, or a custom formula. */
export type TableColumnTotals = {
  column: string | number;
  function?: TableTotalsFunctionName;
  label?: string;
  formula?: string;
};

export type TableStyleToggles = {
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
};

export type TableOperationOptions = TableStyleToggles & {
  sheet?: number | string;
  range?: string;
  ref?: string;
  name?: string;
  columns?: string[];
  style?: string;
  styleName?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  totals?: TableColumnTotals[];
  index?: number;
};

const STYLE_TOGGLES = [
  "showFirstColumn",
  "showLastColumn",
  "showRowStripes",
  "showColumnStripes",
] as const;

function styleToggles(options: TableStyleToggles): TableStyleToggles {
  const toggles: TableStyleToggles = {};
  for (const key of STYLE_TOGGLES) {
    if (options[key] !== undefined) {
      toggles[key] = options[key];
    }
  }
  return toggles;
}

function columnIndex(columns: string[], column: string | number): number {
  if (typeof column === "number") {
    if (!Number.isInteger(column) || column < 0 || column >= columns.length) {
      throw new Error(`table column index must be in 0..${columns.length - 1}: ${column}`);
    }
    return column;
  }
  const key = column.toLocaleLowerCase();
  const index = columns.findIndex((name) => name.toLocaleLowerCase() === key);
  if (index < 0) {
    throw new Error(`table has no column named ${JSON.stringify(column)}`);
  }
  return index;
}

/** The element-wise `Table.update` columns patch for `totals`; a later entry for the same column wins key by key. */
function totalsPatch(table: Table, totals: TableColumnTotals[]): TableColumnPatch[] {
  const columns = table.getProperties().columns.map((column) => column.name);
  const patch: TableColumnPatch[] = [];
  for (const { column, function: fn, label, formula } of totals) {
    const index = columnIndex(columns, column);
    while (patch.length <= index) {
      patch.push({});
    }
    const trimmed = formula?.trim();
    patch[index] = {
      ...patch[index],
      ...(fn !== undefined ? { totalsRowFunction: TOTALS_FUNCTIONS[fn] } : {}),
      ...(label !== undefined ? { totalsRowLabel: label } : {}),
      ...(trimmed !== undefined ? { totalsRowFormula: trimmed } : {}),
    };
  }
  return patch;
}

type TableRange = {
  sheetIndex: number;
  sheetName: string;
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
  ref: string;
};

function requireIndex(index: number | undefined, operation: string): number {
  if (index === undefined) {
    throw new Error(`index is required for table ${operation}`);
  }
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`index must be a nonnegative integer: ${index}`);
  }
  return index;
}

function requiredText(value: string | undefined, name: string): string {
  const trimmed = value?.trim() ?? "";
  if (!trimmed) {
    throw new Error(`${name} is required`);
  }
  return trimmed;
}

function tableRef(options: TableOperationOptions): string {
  const range = options.range?.trim();
  const ref = options.ref?.trim();
  if (range && ref) {
    throw new Error("provide only one of range or ref");
  }
  return requiredText(range ?? ref, "range");
}

function resolveTableRange(
  session: ReturnType<typeof getSession>,
  reference: string,
  fallbackSheet: number | string | undefined,
  expectedSheet?: number,
): TableRange {
  const parsed = normalizeRange(parseRangeRef(reference));
  assertInGrid(parsed.start.row, parsed.start.col);
  assertInGrid(parsed.end.row, parsed.end.col);
  if (parsed.end.row < parsed.start.row || parsed.end.col < parsed.start.col) {
    throw new Error(`table range must have a positive width and height: ${reference}`);
  }
  const sheetIndex = findSheetIndex(session.workbook, parsed.sheetName ?? fallbackSheet);
  if (expectedSheet !== undefined && sheetIndex !== expectedSheet) {
    throw new Error(
      `table range sheet ${sheetName(session, sheetIndex)} does not match the existing table sheet ${sheetName(session, expectedSheet)}`,
    );
  }
  const firstRow = parsed.start.row;
  const firstCol = parsed.start.col;
  const lastRow = parsed.end.row;
  const lastCol = parsed.end.col;
  return {
    sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    firstRow,
    firstCol,
    lastRow,
    lastCol,
    ref: rangeA1(firstRow, firstCol, lastRow, lastCol),
  };
}

function validateColumns(columns: string[], width: number): string[] {
  if (columns.length !== width) {
    throw new Error(`table columns count ${columns.length} does not match range width ${width}`);
  }
  const seen = new Set<string>();
  return columns.map((column, index) => {
    const value = column.trim();
    if (!value) {
      throw new Error(`table column ${index + 1} must be a non-empty text name`);
    }
    const key = value.toLocaleLowerCase();
    if (seen.has(key)) {
      throw new Error(`table column names must be unique; duplicate ${JSON.stringify(value)}`);
    }
    seen.add(key);
    return value;
  });
}

function deriveColumns(session: ReturnType<typeof getSession>, range: TableRange): string[] {
  const sheet = sheetAt(session.workbook, range.sheetIndex);
  const columns: string[] = [];
  for (let col = range.firstCol; col <= range.lastCol; col += 1) {
    const value = sheet.getValue(range.firstRow, col);
    if (typeof value !== "string" || !value.trim()) {
      throw new Error(
        `table header ${range.sheetName}!${rangeA1(range.firstRow, col, range.firstRow, col)} must be non-empty text`,
      );
    }
    columns.push(value);
  }
  return validateColumns(columns, range.lastCol - range.firstCol + 1);
}

type IndexedTable = { index: number; table: Table };

/** Every table in workbook order: sheets in order, each sheet's tables in order. */
function workbookTables(session: ReturnType<typeof getSession>): IndexedTable[] {
  const tables: IndexedTable[] = [];
  for (const sheet of session.workbook.listSheets()) {
    for (const table of sheet.listTables()) {
      tables.push({ index: tables.length, table });
    }
  }
  return tables;
}

function tableAtIndex(session: ReturnType<typeof getSession>, index: number): IndexedTable {
  const found = workbookTables(session)[index];
  if (found === undefined) {
    throw new Error(`table index out of bounds: ${index}`);
  }
  return found;
}

function readTable(session: ReturnType<typeof getSession>, { index, table }: IndexedTable) {
  const properties = table.getProperties();
  const sheetIndex = properties.range.sheet;
  return {
    index,
    name: properties.name,
    displayName: properties.name,
    ref: rectA1(properties.range.rect),
    sheet: sheetIndex,
    sheetName: sheetName(session, sheetIndex),
    // The engine's record, with the style toggles the table entry carried alongside it.
    raw: resultToJson({ ...properties, ...table.getStyleOptions() }),
  };
}

function listTables(session: ReturnType<typeof getSession>, sheet?: number | string) {
  const selectedSheet = sheet === undefined ? undefined : findSheetIndex(session.workbook, sheet);
  const tables = workbookTables(session)
    .map((entry) => readTable(session, entry))
    .filter((table) => selectedSheet === undefined || table.sheet === selectedSheet);
  touch(session, false);
  return { session: publicInfo(session), count: tables.length, tables };
}

function createTable(session: ReturnType<typeof getSession>, options: TableOperationOptions) {
  const reference = tableRef(options);
  const range = resolveTableRange(session, reference, options.sheet);
  const width = range.lastCol - range.firstCol + 1;
  const headerRow = options.headerRow ?? true;
  const columns =
    options.columns === undefined
      ? headerRow
        ? deriveColumns(session, range)
        : (() => {
            throw new Error("columns are required when headerRow is false");
          })()
      : validateColumns(options.columns, width);
  const name = requiredText(options.name, "name");
  const styleName = options.styleName ?? options.style;
  const input: TableInput = {
    range: range.ref,
    name,
    columns,
    headerRow,
    totalsRow: options.totalsRow ?? false,
    ...(styleName === undefined ? {} : { styleName }),
    ...styleToggles(options),
  };
  // A native call can fail after changing part of the workbook. Record the
  // session as dirty before entering the engine and keep that state on error.
  touch(session, true);
  const created = sheetAt(session.workbook, range.sheetIndex).addTable(input);
  if (options.totals !== undefined) {
    created.update({ columns: totalsPatch(created, options.totals) });
  }
  const entry = workbookTables(session).find(({ table }) => table === created);
  if (entry === undefined) {
    throw new Error(`created table ${name} is not listed`);
  }
  return {
    session: publicInfo(session),
    table: readTable(session, entry),
    status: okStatus(),
  };
}

function updateTable(session: ReturnType<typeof getSession>, options: TableOperationOptions) {
  const index = requireIndex(options.index, "update");
  const entry = tableAtIndex(session, index);
  const current = readTable(session, entry);
  const hasRange = options.range !== undefined || options.ref !== undefined;
  const hasStyle = options.style !== undefined || options.styleName !== undefined;
  const toggles = styleToggles(options);
  if (
    !(hasRange || hasStyle) &&
    options.name === undefined &&
    Object.keys(toggles).length === 0 &&
    options.headerRow === undefined &&
    options.totalsRow === undefined &&
    options.totals === undefined
  ) {
    throw new Error(
      "table update requires at least one of range, name, style, a style toggle, headerRow, totalsRow, or totals",
    );
  }
  const patch: {
    range?: string;
    headerRow?: boolean;
    totalsRow?: boolean;
    columns?: TableColumnPatch[];
  } = {};
  const name = options.name === undefined ? undefined : requiredText(options.name, "name");
  if (hasRange) {
    const range = resolveTableRange(session, tableRef(options), current.sheet);
    const oldRange = parseRangeRef(current.ref);
    const oldWidth = Math.abs(oldRange.end.col - oldRange.start.col) + 1;
    const width = range.lastCol - range.firstCol + 1;
    if (width !== oldWidth) {
      throw new Error(`updated table range must keep ${oldWidth} columns; received ${width}`);
    }
    patch.range = range.ref;
  }
  if (options.headerRow !== undefined) {
    patch.headerRow = options.headerRow;
  }
  if (options.totalsRow !== undefined) {
    patch.totalsRow = options.totalsRow;
  }
  if (options.totals !== undefined) {
    patch.columns = totalsPatch(entry.table, options.totals);
  }
  const styleOptions: TableStyleOptionsPatch = {
    ...toggles,
    ...(hasStyle ? { styleName: options.styleName ?? options.style ?? "" } : {}),
  };
  touch(session, true);
  const { table } = entry;
  if (Object.keys(patch).length > 0) {
    table.update(patch);
  }
  if (Object.keys(styleOptions).length > 0) {
    table.updateStyleOptions(styleOptions);
  }
  if (name !== undefined) {
    table.rename(name);
  }
  return {
    session: publicInfo(session),
    table: readTable(session, entry),
    status: okStatus(),
  };
}

function removeTable(session: ReturnType<typeof getSession>, options: TableOperationOptions) {
  const index = requireIndex(options.index, "remove");
  const entry = tableAtIndex(session, index);
  const table = readTable(session, entry);
  touch(session, true);
  entry.table.sheet.removeTable(entry.table);
  return { session: publicInfo(session), removed: table, status: okStatus() };
}

/** Lists and authors native worksheet tables, preserving the engine's global index. */
export function tableOperation(
  id: string,
  operation: TableOperation,
  options: TableOperationOptions = {},
) {
  const session = getSession(id);
  if (operation === "list") {
    return listTables(session, options.sheet);
  }
  if (operation === "create") {
    return createTable(session, options);
  }
  if (operation === "update") {
    return updateTable(session, options);
  }
  return removeTable(session, options);
}

export const listSessionTables = (id: string, sheet?: number | string) =>
  tableOperation(id, "list", { sheet });
