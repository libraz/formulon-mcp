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

const lowerFirst = (text: string) => text.charAt(0).toLowerCase() + text.slice(1);

// The tool's totals-function names, derived from the engine enum (`CountNumbers` is `countNumbers`); `None` is null.
const TOTALS_FUNCTION_ORDINALS: ReadonlyMap<string, TableTotalsRowFunction> = new Map(
  Object.entries(TableTotalsRowFunction)
    .filter(([key]) => key !== "None")
    .map(([key, ordinal]) => [lowerFirst(key), ordinal]),
);
const TOTALS_FUNCTION_NAMES: ReadonlyMap<TableTotalsRowFunction, string> = new Map(
  [...TOTALS_FUNCTION_ORDINALS].map(([name, ordinal]) => [ordinal, name]),
);

/** Every totals-function name the table tool accepts and reports. */
export const TABLE_TOTALS_FUNCTION_NAMES = [...TOTALS_FUNCTION_ORDINALS.keys()] as [
  string,
  ...string[],
];

/** One element of the tool's `columns`: element i is column i. `name` names it on create; `null` clears a setting. */
export type TableColumnSettings = {
  name?: string;
  totalsRowFunction?: string | null;
  totalsRowLabel?: string | null;
  totalsRowFormula?: string | null;
  calculatedFormula?: string | null;
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
  columns?: TableColumnSettings[];
  style?: string;
  styleName?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
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

function totalsFunctionOrdinal(name: string): TableTotalsRowFunction {
  const ordinal = TOTALS_FUNCTION_ORDINALS.get(name);
  if (ordinal === undefined) {
    throw new Error(
      `unknown totalsRowFunction ${JSON.stringify(name)}; expected one of ${TABLE_TOTALS_FUNCTION_NAMES.join(", ")} or null`,
    );
  }
  return ordinal;
}

/** The element-wise `Table.update` columns patch; an element without settings is `{}`. */
function columnsPatch(columns: TableColumnSettings[]): TableColumnPatch[] {
  return columns.map(
    ({ totalsRowFunction, totalsRowLabel, totalsRowFormula, calculatedFormula }) => ({
      ...(totalsRowFunction !== undefined
        ? {
            totalsRowFunction:
              totalsRowFunction === null ? null : totalsFunctionOrdinal(totalsRowFunction),
          }
        : {}),
      ...(totalsRowLabel !== undefined ? { totalsRowLabel } : {}),
      ...(totalsRowFormula !== undefined ? { totalsRowFormula } : {}),
      ...(calculatedFormula !== undefined ? { calculatedFormula } : {}),
    }),
  );
}

const hasSettings = (patch: TableColumnPatch[]) =>
  patch.some((element) => Object.keys(element).length > 0);

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
    // The engine's record, with the style toggles the table entry carried alongside it and
    // each column's totals function under the tool's name (null for none).
    raw: resultToJson({
      ...properties,
      columns: properties.columns.map((column) => ({
        ...column,
        totalsRowFunction: TOTALS_FUNCTION_NAMES.get(column.totalsRowFunction) ?? null,
      })),
      ...table.getStyleOptions(),
    }),
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
  const settings = options.columns ?? [];
  if (settings.length > width) {
    throw new Error(`table columns count ${settings.length} exceeds range width ${width}`);
  }
  // The engine names the columns; given names must cover every column.
  const names = settings.flatMap(({ name }) => (name === undefined ? [] : [name]));
  if (names.length > 0 && names.length !== width) {
    throw new Error(
      "give every column a name in columns, one per range column, or none to let the engine name them",
    );
  }
  const patch = columnsPatch(settings);
  const name = requiredText(options.name, "name");
  const styleName = options.styleName ?? options.style;
  const input: TableInput = {
    range: range.ref,
    name,
    ...(names.length > 0 ? { columns: names } : {}),
    ...(options.headerRow === undefined ? {} : { headerRow: options.headerRow }),
    ...(options.totalsRow === undefined ? {} : { totalsRow: options.totalsRow }),
    ...(styleName === undefined ? {} : { styleName }),
    ...styleToggles(options),
  };
  // A native call can fail after changing part of the workbook. Record the
  // session as dirty before entering the engine and keep that state on error.
  touch(session, true);
  const created = sheetAt(session.workbook, range.sheetIndex).addTable(input);
  if (hasSettings(patch)) {
    created.update({ columns: patch });
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
    options.columns === undefined
  ) {
    throw new Error(
      "table update requires at least one of range, name, style, a style toggle, headerRow, totalsRow, or columns",
    );
  }
  const named = options.columns?.findIndex((column) => column.name !== undefined) ?? -1;
  if (named >= 0) {
    throw new Error(
      `columns[${named}].name names a column on create only; rename a column with Table.renameColumn through formulon_workbook_call`,
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
  if (options.columns !== undefined) {
    patch.columns = columnsPatch(options.columns);
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
