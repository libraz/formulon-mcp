import type { DataTable, DataTableInputCell } from "@libraz/formulon";
import { parseCellRef, parseRangeRef } from "../a1.js";
import { findSheetIndex, okStatus, StatusCode, sheetAt, statusError } from "../formulon.js";
import { assertInGrid, rangeA1, refA1 } from "./shared.js";
import { getSession, publicInfo, touch } from "./state.js";

type Session = ReturnType<typeof getSession>;

function sheetCell(ref: string, label: string) {
  const cell = parseCellRef(ref);
  if (cell.sheetName !== undefined) {
    throw statusError(
      StatusCode.InvalidArgument,
      `${label} must be a cell on the selected sheet, without a sheet name`,
    );
  }
  assertInGrid(cell.row, cell.col);
  return { row: cell.row, col: cell.col };
}

function describeInput(input: DataTableInputCell | null) {
  if (!input) {
    return undefined;
  }
  return { cell: refA1(input.row, input.col), ...(input.deleted ? { deleted: true } : {}) };
}

function describeTable(table: DataTable) {
  const rowInput = describeInput(table.rowInput);
  const colInput = describeInput(table.columnInput);
  const { firstRow, firstCol, lastRow, lastCol } = table.range;
  return {
    range: rangeA1(firstRow, firstCol, lastRow, lastCol),
    twoDimensional: table.twoDimensional,
    ...(rowInput ? { rowInput } : {}),
    ...(colInput ? { colInput } : {}),
  };
}

function tableAt(session: Session, sheetIndex: number, cell: { row: number; col: number }) {
  const dataTable = sheetAt(session.workbook, sheetIndex).getDataTable(cell.row, cell.col);
  return dataTable ? describeTable(dataTable) : null;
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
    throw statusError(
      StatusCode.InvalidArgument,
      "range must be on the selected sheet, without a sheet name",
    );
  }
  assertInGrid(range.end.row, range.end.col);
  if (input.rowInput === undefined && input.colInput === undefined) {
    throw statusError(
      StatusCode.InvalidArgument,
      "a data table needs a rowInput, a colInput, or both",
    );
  }
  const rowInput = input.rowInput === undefined ? null : sheetCell(input.rowInput, "rowInput");
  const colInput = input.colInput === undefined ? null : sheetCell(input.colInput, "colInput");
  touch(session, true);
  sheetAt(session.workbook, sheetIndex).addDataTable({
    range: {
      firstRow: range.start.row,
      firstCol: range.start.col,
      lastRow: range.end.row,
      lastCol: range.end.col,
    },
    rowInput,
    columnInput: colInput,
  });
  session.workbook.recalc();
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
    throw statusError(StatusCode.NotFound, `no data table covers ${cell}`);
  }
  touch(session, true);
  sheetAt(session.workbook, sheetIndex).removeDataTable(target.row, target.col);
  touch(session, true);
  return { session: publicInfo(session), sheet: sheetIndex, removed };
}

/** Recomputes every data table, including under the automatic-except-tables calc mode. */
export function recalcSessionDataTables(id: string) {
  const session = getSession(id);
  touch(session, true);
  session.workbook.recalcDataTables();
  return { session: publicInfo(session), status: okStatus() };
}
