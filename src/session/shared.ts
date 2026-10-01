import { cellToA1 } from "../a1.js";
import {
  assertStatus,
  findSheetIndex,
  readCell,
  requireCell,
  requireNumberResult,
  resultToJson,
  type Status,
  valueToJson,
} from "../formulon.js";
import { describeNumberFormat, type NumberFormatInfo } from "../numfmt.js";
import type { WorkbookSession } from "./state.js";

/** Excel's maximum zero-based row index (1,048,576 rows). */
export const MAX_ROW = 1_048_575;
/** Excel's maximum zero-based column index (16,384 columns, XFD). */
export const MAX_COL = 16_383;
/** Hard ceiling on cells materialized by a single range read. */
export const MAX_RANGE_CELLS = 100_000;
export type PrimitiveCellValue = {
  kind: string;
  value?: number | boolean | string;
  errorCode?: number;
};
/** The workbook-only view shared by registered sessions and one-shot reads. */
export type WorkbookOwner = Pick<WorkbookSession, "workbook">;

export function sheetName(session: WorkbookOwner, sheet: number): string {
  const name = session.workbook.sheetName(sheet);
  assertStatus(name.status, `read sheet ${sheet} name`);
  return name.value;
}

export function refA1(row: number, col: number): string {
  return cellToA1(row, col);
}

export function rangeA1(
  firstRow: number,
  firstCol: number,
  lastRow: number,
  lastCol: number,
): string {
  return `${refA1(firstRow, firstCol)}:${refA1(lastRow, lastCol)}`;
}

/** Rejects a cell address that falls outside Excel's grid limits. */
export function assertInGrid(row: number, col: number): void {
  if (!(Number.isInteger(row) && Number.isInteger(col)) || row < 0 || col < 0) {
    throw new Error(`cell coordinates must be finite nonnegative integers: row ${row}, col ${col}`);
  }
  if (row > MAX_ROW || col > MAX_COL) {
    throw new Error(
      `cell ${cellToA1(row, col)} exceeds Excel sheet limits (max row ${MAX_ROW + 1}, max col XFD)`,
    );
  }
}

/**
 * Builds a `"row,col" -> formula` map for a sheet by scanning its non-empty
 * cells once, so a range read can annotate which cells are computed without a
 * per-cell scan.
 */
export function formulaMap(session: WorkbookOwner, sheet: number): Map<string, string> {
  const map = new Map<string, string>();
  const count = requireNumberResult(
    session.workbook.cellCount(sheet),
    `read sheet ${sheet} cell count`,
  );
  for (let idx = 0; idx < count; idx += 1) {
    const cell = readCell(session.workbook.cellAt(sheet, idx));
    if (cell?.formula) {
      map.set(`${cell.row},${cell.col}`, cell.formula);
    }
  }
  return map;
}

/** Reads the formula text at one cell, or an empty string for a constant. */
export function formulaAt(session: WorkbookOwner, sheet: number, row: number, col: number): string {
  const count = requireNumberResult(
    session.workbook.cellCount(sheet),
    `read sheet ${sheet} cell count`,
  );
  for (let idx = 0; idx < count; idx += 1) {
    const cell = readCell(session.workbook.cellAt(sheet, idx));
    if (cell && cell.row === row && cell.col === col) {
      return cell.formula;
    }
  }
  return "";
}

/**
 * Resolves the number-format code applied to a cell by chaining
 * `getCellXfIndex -> getCellXf -> getNumFmt`. Returns null when the cell has no
 * resolvable format (empty cell, default style, or an engine read error).
 */
export function cellNumberFormat(
  session: WorkbookOwner,
  sheet: number,
  row: number,
  col: number,
): { numFmtId: number; formatCode: string } | null {
  const xfIndexResult = safeWorkbookCall(session, "getCellXfIndex", [sheet, row, col]) as {
    xfIndex?: number;
  } | null;
  if (typeof xfIndexResult?.xfIndex !== "number") {
    return null;
  }
  const xf = safeWorkbookCall(session, "getCellXf", [xfIndexResult.xfIndex]) as {
    numFmtId?: number;
  } | null;
  if (typeof xf?.numFmtId !== "number") {
    return null;
  }
  const numFmt = safeWorkbookCall(session, "getNumFmt", [xf.numFmtId]) as {
    formatCode?: string;
  } | null;
  return { numFmtId: xf.numFmtId, formatCode: numFmt?.formatCode ?? "" };
}

/**
 * Decodes a numeric cell's applied format into a readable annotation (ISO date,
 * currency/percent label). Returns undefined for non-numeric cells and plain
 * unformatted numbers so the output stays sparse.
 */
export function decodeNumberCell(
  session: WorkbookOwner,
  sheet: number,
  row: number,
  col: number,
  value: PrimitiveCellValue,
): NumberFormatInfo | undefined {
  if (value.kind !== "number" || typeof value.value !== "number") {
    return undefined;
  }
  const format = cellNumberFormat(session, sheet, row, col);
  if (!format) {
    return undefined;
  }
  return describeNumberFormat(format.numFmtId, format.formatCode, value.value);
}

export function safeWorkbookCall(session: WorkbookOwner, method: string, args: unknown[]): unknown {
  const callable = (session.workbook as unknown as Record<string, unknown>)[method];
  if (typeof callable !== "function") {
    return null;
  }
  try {
    const raw = (callable as (...methodArgs: unknown[]) => unknown).apply(session.workbook, args);
    const status = resultStatus(raw);
    if (status) {
      assertStatus(status, `${method}`);
    }
    // High-level inspection payloads use ordinary arrays. The low-level
    // workbook-call surface retains the ListResult envelope via resultToJson.
    const value = isListResultLike(raw) ? Array.from(raw) : raw;
    return resultToJson(value);
  } catch {
    return null;
  }
}

export function resultStatus(value: unknown): Status | undefined {
  if (typeof value !== "object" || value === null) {
    return undefined;
  }
  const record = value as { ok?: unknown; status?: unknown };
  if (typeof record.ok === "boolean" && typeof record.status === "number") {
    return value as Status;
  }
  if (
    typeof record.status === "object" &&
    record.status !== null &&
    typeof (record.status as { ok?: unknown }).ok === "boolean" &&
    typeof (record.status as { status?: unknown }).status === "number"
  ) {
    return record.status as Status;
  }
  return undefined;
}

export function isListResultLike(value: unknown): value is readonly unknown[] & { status: Status } {
  return Array.isArray(value) && resultStatus(value) !== undefined && "status" in value;
}

export function jsonCellValue(value: unknown): PrimitiveCellValue {
  return valueToJson(value as Parameters<typeof valueToJson>[0]) as PrimitiveCellValue;
}

export function primitiveText(value: PrimitiveCellValue): string {
  if (value.value === undefined) {
    return "";
  }
  return String(value.value);
}

export function isNonBlank(value: PrimitiveCellValue, formula: string): boolean {
  return formula.length > 0 || value.kind !== "blank";
}

export function cellKind(value: PrimitiveCellValue, formula: string): string {
  if (formula) {
    return "formula";
  }
  if (value.kind === "number") {
    return "number";
  }
  if (value.kind === "bool") {
    return "boolean";
  }
  if (value.kind === "text") {
    const text = primitiveText(value);
    if (/^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}/.test(text)) {
      return "date";
    }
    if (/^[¥$€]?\s*-?\d[\d,]*(?:\.\d+)?$/.test(text)) {
      return "money";
    }
    return "text";
  }
  return value.kind;
}

export function isNumericLikeKind(kind: string): boolean {
  return kind === "number" || kind === "money" || kind === "formula";
}

export function stableCells(
  session: WorkbookSession,
  sheet: number,
  maxCells: number,
  includeStyles: boolean,
) {
  const count = requireNumberResult(
    session.workbook.cellCount(sheet),
    `read sheet ${sheet} cell count`,
  );
  const limit = Math.max(0, Math.min(maxCells, count));
  const cells = [];
  for (let index = 0; index < limit; index += 1) {
    const cell = requireCell(
      session.workbook.cellAt(sheet, index),
      `read sheet ${sheet} cell ${index}`,
    );
    const value = jsonCellValue(cell.value);
    const formula = cell.formula;
    const xfIndexResult = includeStyles
      ? (safeWorkbookCall(session, "getCellXfIndex", [sheet, cell.row, cell.col]) as {
          xfIndex?: number;
        } | null)
      : null;
    const xfIndex = typeof xfIndexResult?.xfIndex === "number" ? xfIndexResult.xfIndex : null;
    cells.push({
      row: cell.row,
      col: cell.col,
      a1: refA1(cell.row, cell.col),
      value,
      formula,
      kind: cellKind(value, formula),
      style: includeStyles
        ? {
            xfIndex,
            xf: xfIndex === null ? null : safeWorkbookCall(session, "getCellXf", [xfIndex]),
          }
        : undefined,
    });
  }
  cells.sort((left, right) => left.row - right.row || left.col - right.col);
  return { cells, cellCount: count, truncated: limit < count };
}

export function usedRange(cells: { row: number; col: number }[]): string | null {
  if (cells.length === 0) {
    return null;
  }
  const rows = cells.map((cell) => cell.row);
  const cols = cells.map((cell) => cell.col);
  return rangeA1(Math.min(...rows), Math.min(...cols), Math.max(...rows), Math.max(...cols));
}

export function searchSheets(
  session: WorkbookSession,
  sheet: number | string | undefined,
): number[] {
  if (sheet !== undefined) {
    return [findSheetIndex(session.workbook, sheet)];
  }
  const count = requireNumberResult(session.workbook.sheetCount(), "read sheet count");
  return Array.from({ length: count }, (_, index) => index);
}

export function textValue(value: unknown): string | undefined {
  if (
    typeof value === "object" &&
    value !== null &&
    "kind" in value &&
    (value as { kind: unknown }).kind === 3 &&
    "text" in value &&
    typeof (value as { text: unknown }).text === "string"
  ) {
    return (value as { text: string }).text;
  }
  return undefined;
}

/**
 * Renders a constant (non-formula) cell as searchable text so `find_cells` can
 * match numbers and booleans by value, not just text cells. Blank/other kinds
 * return undefined and are skipped.
 */
export function cellDisplayText(value: unknown): string | undefined {
  if (typeof value !== "object" || value === null || !("kind" in value)) {
    return undefined;
  }
  const envelope = value as { kind: number; number?: number; boolean?: number; text?: string };
  switch (envelope.kind) {
    case 1:
      return typeof envelope.number === "number" ? String(envelope.number) : undefined;
    case 2:
      return envelope.boolean ? "TRUE" : "FALSE";
    case 3:
      return typeof envelope.text === "string" ? envelope.text : undefined;
    default:
      return undefined;
  }
}
/** Returns the max non-empty (row, col) extent of a sheet, or null when empty. */
export function sheetUsedBounds(
  session: WorkbookOwner,
  sheet: number,
): { maxRow: number; maxCol: number } | null {
  const count = requireNumberResult(
    session.workbook.cellCount(sheet),
    `read sheet ${sheet} cell count`,
  );
  if (count === 0) {
    return null;
  }
  let maxRow = 0;
  let maxCol = 0;
  for (let idx = 0; idx < count; idx += 1) {
    const cell = readCell(session.workbook.cellAt(sheet, idx));
    if (!cell) {
      continue;
    }
    if (cell.row > maxRow) {
      maxRow = cell.row;
    }
    if (cell.col > maxCol) {
      maxCol = cell.col;
    }
  }
  return { maxRow, maxCol };
}
export function requiredString(value: string | undefined, name: string): string {
  if (value === undefined || value.length === 0) {
    throw new Error(`${name} is required`);
  }
  return value;
}

export function requiredNumber(value: number | undefined, name: string): number {
  if (value === undefined) {
    throw new Error(`${name} is required`);
  }
  if (!Number.isFinite(value)) {
    throw new Error(`${name} must be finite`);
  }
  return value;
}

export function requiredBoolean(value: boolean | undefined): boolean {
  if (value === undefined) {
    throw new Error("hidden is required");
  }
  return value;
}
