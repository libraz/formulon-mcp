import type { TableInput, TableUpdateInput } from "@libraz/formulon";
import { normalizeRange, parseRangeRef } from "../a1.js";
import { assertStatus, findSheetIndex, resultToJson, statusToJson } from "../formulon.js";
import { assertInGrid, rangeA1, sheetName } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

export type TableOperation = "list" | "create" | "update" | "remove";

export type TableOperationOptions = {
  sheet?: number | string;
  range?: string;
  ref?: string;
  name?: string;
  columns?: string[];
  style?: string;
  styleName?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  index?: number;
};

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
  const columns: string[] = [];
  for (let col = range.firstCol; col <= range.lastCol; col += 1) {
    const result = session.workbook.getValue(range.sheetIndex, range.firstRow, col);
    assertStatus(
      result.status,
      `read table header ${sheetName(session, range.sheetIndex)}!${rangeA1(range.firstRow, col, range.firstRow, col)}`,
    );
    const value = result.value;
    if (value.kind !== 3 || !value.text.trim()) {
      throw new Error(
        `table header ${sheetName(session, range.sheetIndex)}!${rangeA1(range.firstRow, col, range.firstRow, col)} must be non-empty text`,
      );
    }
    columns.push(value.text);
  }
  return validateColumns(columns, range.lastCol - range.firstCol + 1);
}

function readTable(session: ReturnType<typeof getSession>, index: number) {
  const entry = session.workbook.tableAt(index);
  assertStatus(entry.status, `read table ${index}`);
  const sheetIndex = entry.sheetIndex;
  return {
    index,
    name: entry.name ?? entry.displayName ?? "",
    displayName: entry.displayName ?? entry.name ?? "",
    ref: entry.ref ?? "",
    sheet: sheetIndex ?? null,
    sheetName: sheetIndex === undefined ? null : sheetName(session, sheetIndex),
    raw: resultToJson(entry),
  };
}

function listTables(session: ReturnType<typeof getSession>, sheet?: number | string) {
  const selectedSheet = sheet === undefined ? undefined : findSheetIndex(session.workbook, sheet);
  const countResult = session.workbook.tableCount();
  assertStatus(countResult.status, "read table count");
  const tables = [];
  for (let index = 0; index < countResult.value; index += 1) {
    const table = readTable(session, index);
    if (selectedSheet === undefined || table.sheet === selectedSheet) {
      tables.push(table);
    }
  }
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
    sheetIndex: range.sheetIndex,
    ref: range.ref,
    name,
    columns,
    headerRow,
    totalsRow: options.totalsRow ?? false,
    ...(styleName === undefined ? {} : { styleName }),
  };
  // A native call can fail after changing part of the workbook. Record the
  // session as dirty before entering the engine and keep that state on error.
  touch(session, true);
  const result = session.workbook.createTable(input);
  assertStatus(result.status, "create table");
  const table = readTable(session, result.index);
  return { session: publicInfo(session), table, status: statusToJson(result.status) };
}

function updateTable(session: ReturnType<typeof getSession>, options: TableOperationOptions) {
  const index = requireIndex(options.index, "update");
  const current = readTable(session, index);
  const hasRange = options.range !== undefined || options.ref !== undefined;
  const hasStyle = options.style !== undefined || options.styleName !== undefined;
  if (
    !(hasRange || hasStyle) &&
    options.headerRow === undefined &&
    options.totalsRow === undefined
  ) {
    throw new Error("table update requires at least one of range, style, headerRow, or totalsRow");
  }
  const input: TableUpdateInput = {};
  if (hasRange) {
    if (current.sheet === null) {
      throw new Error(`table ${index} has no worksheet index`);
    }
    const range = resolveTableRange(session, tableRef(options), current.sheet);
    const oldRange = parseRangeRef(current.ref);
    const oldWidth = Math.abs(oldRange.end.col - oldRange.start.col) + 1;
    const width = range.lastCol - range.firstCol + 1;
    if (width !== oldWidth) {
      throw new Error(`updated table range must keep ${oldWidth} columns; received ${width}`);
    }
    input.ref = range.ref;
  }
  if (hasStyle) {
    input.styleName = options.styleName ?? options.style ?? "";
  }
  if (options.headerRow !== undefined) {
    input.headerRow = options.headerRow;
  }
  if (options.totalsRow !== undefined) {
    input.totalsRow = options.totalsRow;
  }
  touch(session, true);
  const status = session.workbook.updateTable(index, input);
  assertStatus(status, `update table ${index}`);
  const table = readTable(session, index);
  return { session: publicInfo(session), table, status: statusToJson(status) };
}

function removeTable(session: ReturnType<typeof getSession>, options: TableOperationOptions) {
  const index = requireIndex(options.index, "remove");
  const table = readTable(session, index);
  touch(session, true);
  const status = session.workbook.removeTable(index);
  assertStatus(status, `remove table ${index}`);
  return { session: publicInfo(session), removed: table, status: statusToJson(status) };
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
