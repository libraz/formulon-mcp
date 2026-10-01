import { cellToA1, normalizeRange, parseCellRef, parseRangeRef } from "../a1.js";
import {
  applyMutation,
  assertStatus,
  type CellMutation,
  errorName,
  findSheetIndex,
  statusToJson,
  valueToJson,
} from "../formulon.js";
import {
  assertInGrid,
  decodeNumberCell,
  formulaAt,
  formulaMap,
  jsonCellValue,
  MAX_RANGE_CELLS,
  type PrimitiveCellValue,
  sheetName,
  sheetUsedBounds,
  type WorkbookOwner,
} from "./shared.js";
import { getSession, publicInfo, touch, type WorkbookSession } from "./state.js";

type A1MutationBase = {
  sheet?: number | string;
  a1?: string;
  row?: number;
  col?: number;
};

export type FlexibleCellMutation =
  | (A1MutationBase & { type: "number"; value: number })
  | (A1MutationBase & { type: "bool"; value: boolean })
  | (A1MutationBase & { type: "text"; value: string })
  | (A1MutationBase & { type: "blank" })
  | (A1MutationBase & { type: "formula"; formula: string });
function resolveCell(session: WorkbookSession, mutation: FlexibleCellMutation) {
  if (mutation.a1) {
    const parsed = parseCellRef(mutation.a1);
    assertInGrid(parsed.row, parsed.col);
    return {
      sheet: findSheetIndex(session.workbook, parsed.sheetName ?? mutation.sheet),
      row: parsed.row,
      col: parsed.col,
    };
  }
  if (mutation.row === undefined || mutation.col === undefined) {
    throw new Error("mutation requires either a1 or row/col");
  }
  assertInGrid(mutation.row, mutation.col);
  return {
    sheet: findSheetIndex(session.workbook, mutation.sheet),
    row: mutation.row,
    col: mutation.col,
  };
}

/** Applies flexible cell mutations using either A1 references or zero-based coordinates. */
export function applySessionMutations(
  id: string,
  mutations: FlexibleCellMutation[],
  recalc: boolean,
) {
  const session = getSession(id);
  for (const [idx, mutation] of mutations.entries()) {
    if (mutation.type === "number" && !Number.isFinite(mutation.value)) {
      throw new Error(`mutation ${idx}: number value must be finite`);
    }
  }
  // Resolve (and bounds-check) every address up front so a bad ref or sheet
  // name aborts before any cell is written, keeping addressing errors atomic.
  const resolved = mutations.map((mutation, idx) => {
    try {
      return resolveCell(session, mutation);
    } catch (error) {
      throw new Error(`mutation ${idx}: ${error instanceof Error ? error.message : String(error)}`);
    }
  });
  const applied = [];
  let appliedCount = 0;
  let attemptedCount = 0;
  try {
    for (const [idx, mutation] of mutations.entries()) {
      const address = resolved[idx];
      const concrete = { ...mutation, ...address } as CellMutation;
      attemptedCount += 1;
      // A mutating engine call may fail after changing part of the workbook;
      // mark dirty before entering it so thrown and failed statuses retain the
      // conservative state.
      touch(session, true);
      const status = applyMutation(session.workbook, concrete);
      assertStatus(status, `apply mutation ${idx}`);
      appliedCount += 1;
      applied.push({
        index: idx,
        a1: cellToA1(address.row, address.col),
        sheet: address.sheet,
        status: statusToJson(status),
      });
    }
  } finally {
    // Even on a mid-batch failure, recalc what landed and mark the session
    // dirty so the partial write is not silently reported as unchanged.
    if (attemptedCount > 0) {
      touch(session, true);
      if (recalc) {
        if (appliedCount > 0) {
          assertStatus(session.workbook.recalc(), "recalc workbook");
        }
      }
    }
  }
  // Surface formula cells that evaluated to an error, so a mistyped function
  // (#NAME?, #REF!, ...) is not reported as an all-green success.
  const errorCells = [];
  for (const [idx, mutation] of mutations.entries()) {
    if (mutation.type !== "formula") {
      continue;
    }
    const address = resolved[idx];
    const result = session.workbook.getValue(address.sheet, address.row, address.col);
    assertStatus(result.status, `read formula result ${idx}`);
    const value = jsonCellValue(result.value);
    if (value.kind === "error") {
      errorCells.push({
        index: idx,
        ref: `${sheetName(session, address.sheet)}!${cellToA1(address.row, address.col)}`,
        errorName: errorName(value.errorCode ?? -1),
      });
    }
  }
  return { session: publicInfo(session), applied, errorCells };
}

/** One cell of a `set_range` block: a scalar literal, a formula, a blank, or a skip. */
export type RangeWriteCell = number | boolean | string | null | { f: string } | { blank: true };

/**
 * Reads one cell from a workbook without requiring it to be registered as a
 * session. The caller owns the workbook handle and is responsible for
 * releasing it; this keeps one-shot path reads identical to session reads.
 */
export function readWorkbookCell(
  owner: WorkbookOwner,
  sheet: number | string | undefined,
  row: number,
  col: number,
  recalc = false,
) {
  const sheetIndex = findSheetIndex(owner.workbook, sheet);
  assertInGrid(row, col);
  if (recalc) {
    assertStatus(owner.workbook.recalc(), "recalc workbook");
  }
  const result = owner.workbook.getValue(sheetIndex, row, col);
  assertStatus(result.status, "get cell");
  const value = valueToJson(result.value) as PrimitiveCellValue;
  return {
    address: { sheet: sheetIndex, row, col, a1: cellToA1(row, col) },
    status: statusToJson(result.status),
    value,
    formula: formulaAt(owner, sheetIndex, row, col),
    ...decodeNumberCell(owner, sheetIndex, row, col, value),
  };
}

/**
 * Writes a 2D block of values starting at an anchor cell. Each element's JSON
 * type selects the cell type (number/bool/text); `{f: "..."}` writes a formula;
 * `null` leaves the cell untouched; `{blank: true}` explicitly clears it.
 * Returns a compact summary plus any error cells produced by written formulas.
 */
export function setSessionRange(
  id: string,
  start: string,
  values: RangeWriteCell[][],
  fallbackSheet: number | string | undefined,
  recalc: boolean,
) {
  const session = getSession(id);
  const anchor = parseCellRef(start);
  const sheetIndex = findSheetIndex(session.workbook, anchor.sheetName ?? fallbackSheet);
  assertInGrid(anchor.row, anchor.col);
  const maxColumns = values.reduce((max, row) => Math.max(max, row.length), 0);
  if (values.length > 0 && maxColumns > 0) {
    assertInGrid(anchor.row + values.length - 1, anchor.col + maxColumns - 1);
  }
  const mutations: FlexibleCellMutation[] = [];
  for (const [r, rowValues] of values.entries()) {
    for (const [c, cell] of rowValues.entries()) {
      if (cell === null) {
        continue;
      }
      const row = anchor.row + r;
      const col = anchor.col + c;
      assertInGrid(row, col);
      const base = { sheet: sheetIndex, row, col };
      if (typeof cell === "number") {
        mutations.push({ ...base, type: "number", value: cell });
      } else if (typeof cell === "boolean") {
        mutations.push({ ...base, type: "bool", value: cell });
      } else if (typeof cell === "string") {
        mutations.push({ ...base, type: "text", value: cell });
      } else if (typeof cell === "object" && "f" in cell && typeof cell.f === "string") {
        mutations.push({ ...base, type: "formula", formula: cell.f });
      } else if (typeof cell === "object" && "blank" in cell && cell.blank === true) {
        mutations.push({ ...base, type: "blank" });
      } else {
        throw new Error(
          `invalid cell at row ${r}, col ${c}: expected number|boolean|string|{f}|{blank:true}|null`,
        );
      }
    }
  }
  if (mutations.length === 0) {
    return {
      session: publicInfo(session),
      range: { sheet: sheetIndex, start: cellToA1(anchor.row, anchor.col) },
      cellsWritten: 0,
      errorCells: [],
    };
  }
  const result = applySessionMutations(id, mutations, recalc);
  const rows = values.length;
  const cols = values.reduce((max, row) => Math.max(max, row.length), 0);
  return {
    session: result.session,
    range: {
      sheet: sheetIndex,
      start: cellToA1(anchor.row, anchor.col),
      end: cellToA1(anchor.row + rows - 1, anchor.col + cols - 1),
    },
    cellsWritten: result.applied.length,
    errorCells: result.errorCells,
  };
}

/** Reads one cell from an open session using zero-based coordinates. */
export function getSessionCell(
  id: string,
  sheet: number | string | undefined,
  row: number,
  col: number,
  recalc = false,
) {
  const session = getSession(id);
  const sheetIndex = findSheetIndex(session.workbook, sheet);
  assertInGrid(row, col);
  if (recalc) {
    touch(session, true);
  }
  const cell = readWorkbookCell(session, sheetIndex, row, col, recalc);
  touch(session, false);
  return {
    session: publicInfo(session),
    ...cell,
  };
}

/**
 * Reads one cell from an open session using an A1 reference. When the ref has
 * no sheet prefix, `fallbackSheet` (the tool's `sheet` argument) is used.
 */
export function getSessionCellByA1(
  id: string,
  ref: string,
  fallbackSheet?: number | string,
  recalc = false,
) {
  const parsed = parseCellRef(ref);
  return getSessionCell(id, parsed.sheetName ?? fallbackSheet, parsed.row, parsed.col, recalc);
}

export type RangeReadOptions = {
  sheet?: number | string;
  maxCells: number;
  recalc: boolean;
  includeFormulas: boolean;
};

/**
 * Reads a rectangular A1 range from an open session as a sparse cell list.
 *
 * The requested rectangle is clipped to the sheet's used range, blank cells are
 * omitted, and materialization is capped at `maxCells` (with a `truncated`
 * flag) so an over-wide request cannot blow up the response or hang the server.
 */
export function getSessionRange(id: string, ref: string, options: RangeReadOptions) {
  const session = getSession(id);
  const parsed = parseRangeRef(ref);
  const range = normalizeRange(parsed);
  assertInGrid(range.start.row, range.start.col);
  assertInGrid(range.end.row, range.end.col);
  const sheetIndex = findSheetIndex(session.workbook, range.sheetName ?? options.sheet);
  if (options.recalc) {
    touch(session, true);
    assertStatus(session.workbook.recalc(), "recalc workbook");
  }

  // Clip the requested rectangle to the sheet's used extent so a generous
  // "A1:Z1000" request only scans cells that can actually hold data.
  const used = sheetUsedBounds(session, sheetIndex);
  const startRow = range.start.row;
  const startCol = range.start.col;
  const endRow = used ? Math.min(range.end.row, used.maxRow) : range.start.row - 1;
  const endCol = used ? Math.min(range.end.col, used.maxCol) : range.start.col - 1;

  const rectCells = (endRow - startRow + 1) * (endCol - startCol + 1);
  if (rectCells > MAX_RANGE_CELLS) {
    throw new Error(
      `range ${ref} spans ${rectCells} cells (limit ${MAX_RANGE_CELLS}); read a smaller range or use inspect_layout for the used range`,
    );
  }

  const formulas = options.includeFormulas ? formulaMap(session, sheetIndex) : undefined;
  const cells = [];
  let truncated = false;
  for (let row = startRow; row <= endRow && !truncated; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      const result = session.workbook.getValue(sheetIndex, row, col);
      assertStatus(result.status, `get cell ${cellToA1(row, col)}`);
      const value = valueToJson(result.value) as PrimitiveCellValue;
      const formula = formulas?.get(`${row},${col}`) ?? "";
      // Skip cells that are both blank and formula-free to keep output sparse.
      if (value.kind === "blank" && !formula) {
        continue;
      }
      if (cells.length >= options.maxCells) {
        truncated = true;
        break;
      }
      cells.push({
        a1: cellToA1(row, col),
        row,
        col,
        value,
        ...decodeNumberCell(session, sheetIndex, row, col, value),
        ...(options.includeFormulas ? { formula } : {}),
      });
    }
  }
  touch(session, false);
  return {
    session: publicInfo(session),
    range: {
      sheet: sheetIndex,
      sheetName: sheetName(session, sheetIndex),
      ref,
      start: cellToA1(range.start.row, range.start.col),
      end: cellToA1(range.end.row, range.end.col),
    },
    cellCount: cells.length,
    truncated,
    cells,
  };
}
