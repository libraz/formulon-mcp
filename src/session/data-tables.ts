import type { DataTable, DataTableInput } from "@libraz/formulon";
import { parseCellRef, parseRangeRef } from "../a1.js";
import { assertStatus, findSheetIndex, statusToJson } from "../formulon.js";
import { assertInGrid, rangeA1, refA1 } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

type Session = ReturnType<typeof getSession>;

function sheetCell(ref: string, label: string) {
  const cell = parseCellRef(ref);
  if (cell.sheetName !== undefined) {
    throw new Error(`${label} must be a cell on the selected sheet, without a sheet name`);
  }
  assertInGrid(cell.row, cell.col);
  return { row: cell.row, col: cell.col };
}

function describeInput(input: DataTableInput | null) {
  if (!input) {
    return undefined;
  }
  return { cell: refA1(input.row, input.col), ...(input.deleted ? { deleted: true } : {}) };
}

function describeTable(table: DataTable) {
  const rowInput = describeInput(table.rowInput);
  const colInput = describeInput(table.colInput);
  return {
    range: rangeA1(table.firstRow, table.firstCol, table.lastRow, table.lastCol),
    twoDimensional: table.twoDimensional,
    ...(rowInput ? { rowInput } : {}),
    ...(colInput ? { colInput } : {}),
  };
}

function tableAt(session: Session, sheetIndex: number, cell: { row: number; col: number }) {
  const result = session.workbook.getDataTable(sheetIndex, cell.row, cell.col);
  assertStatus(result.status, "read data table");
  return result.dataTable ? describeTable(result.dataTable) : null;
}

/** Reads the What-If data table whose result range covers `cell`, or null. */
export function getSessionDataTable(id: string, sheet: number | string | undefined, cell: string) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const dataTable = tableAt(session, sheetIndex, sheetCell(cell, "cell"));
  touch(session, false);
  return { session: publicInfo(session), sheet: sheetIndex, dataTable };
}

/**
 * Creates a What-If data table over a result range and recalculates so its
 * cells hold results. The row input takes the values in the row above the
 * range, the column input those in the column left of it.
 */
export function addSessionDataTable(
  id: string,
  sheet: number | string | undefined,
  input: { range: string; rowInput?: string; colInput?: string },
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const range = parseRangeRef(input.range);
  if (range.sheetName !== undefined) {
    throw new Error("range must be on the selected sheet, without a sheet name");
  }
  assertInGrid(range.end.row, range.end.col);
  if (input.rowInput === undefined && input.colInput === undefined) {
    throw new Error("a data table needs a rowInput, a colInput, or both");
  }
  const rowInput = input.rowInput === undefined ? null : sheetCell(input.rowInput, "rowInput");
  const colInput = input.colInput === undefined ? null : sheetCell(input.colInput, "colInput");
  touch(session, true);
  assertStatus(
    session.workbook.addDataTable(sheetIndex, {
      firstRow: range.start.row,
      firstCol: range.start.col,
      lastRow: range.end.row,
      lastCol: range.end.col,
      rowInput,
      colInput,
    }),
    "add data table",
  );
  assertStatus(session.workbook.recalc(), "recalc workbook");
  return {
    session: publicInfo(session),
    sheet: sheetIndex,
    dataTable: tableAt(session, sheetIndex, range.start),
  };
}

/** Removes the data table covering `cell` and blanks its result cells. */
export function removeSessionDataTable(
  id: string,
  sheet: number | string | undefined,
  cell: string,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  const target = sheetCell(cell, "cell");
  const removed = tableAt(session, sheetIndex, target);
  if (!removed) {
    throw new Error(`no data table covers ${cell}`);
  }
  touch(session, true);
  assertStatus(
    session.workbook.removeDataTable(sheetIndex, target.row, target.col),
    "remove data table",
  );
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, removed };
}

/** Recomputes every data table, including under the automatic-except-tables calc mode. */
export function recalcSessionDataTables(id: string) {
  const session = getSession(id);
  touch(session, true);
  const status = session.workbook.recalcDataTables();
  assertStatus(status, "recalc data tables");
  return { session: publicInfo(session), status: statusToJson(status) };
}
