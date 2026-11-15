import {
  type CellRect,
  type RangeCell,
  type Sheet,
  SpillCause,
  type Value,
  ValueKind,
} from "@libraz/formulon";
import { cellToA1, MAX_COL, MAX_ROW } from "../a1.js";
import {
  DisplayStatus,
  ExcelError,
  findSheetIndex,
  isDate1904,
  resultToJson,
  StatusCode,
  statusError,
  valueToJson,
} from "../formulon.js";
import { classifyNumberFormat, describeNumberFormat, type NumberFormatInfo } from "../numfmt.js";
import type { WorkbookSession } from "./state.js";

export { MAX_COL, MAX_ROW };
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
  return session.workbook.sheet(sheet).name;
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

/** `PascalCase` engine enum key to the camelCase name the tools use. */
export const lowerFirst = (text: string) => text.charAt(0).toLowerCase() + text.slice(1);

const SPILL_CAUSE_NAMES: ReadonlyMap<SpillCause, string> = new Map(
  Object.entries(SpillCause).map(([key, ordinal]) => [ordinal, lowerFirst(key)]),
);

/**
 * The spill or PivotTable report region `Sheet.getSpill` reads at a cell, with A1 references;
 * `null` when the cell is in none.
 */
export function spillRecord(session: WorkbookOwner, sheet: number, row: number, col: number) {
  const info = session.workbook.sheet(sheet).getSpill(row, col);
  if (info === null) {
    return null;
  }
  const { firstRow, firstCol, lastRow, lastCol } = info.range;
  const { row: anchorRow, col: anchorCol } = info.anchor;
  const name = sheetName(session, sheet);
  return {
    blocked: info.blocked,
    cause: info.cause === null ? null : (SPILL_CAUSE_NAMES.get(info.cause) ?? null),
    anchor: {
      sheet,
      sheetName: name,
      row: anchorRow,
      col: anchorCol,
      ref: `${name}!${refA1(anchorRow, anchorCol)}`,
    },
    rows: lastRow - firstRow + 1,
    cols: lastCol - firstCol + 1,
    range: `${name}!${rangeA1(firstRow, firstCol, lastRow, lastCol)}`,
  };
}

/** Rejects a cell address that falls outside Excel's grid limits. */
export function assertInGrid(row: number, col: number): void {
  if (!(Number.isInteger(row) && Number.isInteger(col)) || row < 0 || col < 0) {
    throw statusError(
      StatusCode.InvalidArgument,
      `cell coordinates must be finite nonnegative integers: row ${row}, col ${col}`,
    );
  }
  if (row > MAX_ROW || col > MAX_COL) {
    throw statusError(
      StatusCode.InvalidArgument,
      `cell ${cellToA1(row, col)} exceeds Excel sheet limits (max row ${MAX_ROW + 1}, max col XFD)`,
    );
  }
}

/** One populated cell of a rectangle read, with its formula text ("" for a constant). */
export type RectCell = { row: number; col: number; formula: string; value: Value };

/** The tagged `Value` of a range cell read, which carries its value natively with its kind. */
export function taggedValue(sheet: Sheet, cell: RangeCell): Value {
  const value = cell.value;
  if (value === null) {
    return { kind: ValueKind.Blank };
  }
  if (typeof value === "number") {
    return { kind: ValueKind.Number, number: value };
  }
  if (typeof value === "boolean") {
    return { kind: ValueKind.Bool, boolean: value };
  }
  if (typeof value === "string") {
    return { kind: ValueKind.Text, text: value };
  }
  if (value instanceof ExcelError) {
    return { kind: ValueKind.Error, errorCode: value.code };
  }
  return sheet.getTypedValue(cell.row, cell.col);
}

/**
 * Pages the engine's populated cells inside a rectangle in row-major order.
 * `pageSize` caps each engine call; stopping the iteration early stops the paging.
 */
export function* rectCells(
  session: WorkbookOwner,
  sheet: number,
  rect: CellRect,
  pageSize?: number,
): Generator<RectCell> {
  const handle = session.workbook.sheet(sheet);
  let cursor: number | null = null;
  do {
    const page = handle.getCellsInRange(rect, { cursor, limit: pageSize ?? null });
    for (const cell of page.cells) {
      yield {
        row: cell.row,
        col: cell.col,
        formula: cell.formula ?? "",
        value: taggedValue(handle, cell),
      };
    }
    cursor = page.nextCursor;
  } while (cursor !== null);
}

/** Reads the formula text at one cell (leading `=` included), or "" for a constant. */
export function formulaAt(session: WorkbookOwner, sheet: number, row: number, col: number): string {
  return session.workbook.sheet(sheet).getFormula(row, col) ?? "";
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
  return safeRead(() => {
    const xfIndex = session.workbook.sheet(sheet).ooxml.getCellXfIndex(row, col);
    if (xfIndex === null) {
      return null;
    }
    const numFmtId = session.workbook.ooxml.getCellXf(xfIndex).numFmtId;
    return { numFmtId, formatCode: session.workbook.ooxml.getNumberFormat(numFmtId) ?? "" };
  });
}

/** Format annotation of a numeric cell plus the engine's display text. */
export type CellFormatAnnotation = Partial<NumberFormatInfo> & {
  text?: string;
  textStatus?: "overflow" | "invalidFormat";
};

/**
 * Decodes a numeric cell's applied format into a readable annotation (ISO date,
 * currency/percent label) and the text Excel displays for it. Returns undefined
 * for non-numeric cells and General-format numbers so the output stays sparse.
 */
export function decodeNumberCell(
  session: WorkbookOwner,
  sheet: number,
  row: number,
  col: number,
  value: PrimitiveCellValue,
): CellFormatAnnotation | undefined {
  if (value.kind !== "number" || typeof value.value !== "number") {
    return undefined;
  }
  const format = cellNumberFormat(session, sheet, row, col);
  if (!format || classifyNumberFormat(format.numFmtId, format.formatCode) === "general") {
    return undefined;
  }
  const display = safeRead(() => session.workbook.sheet(sheet).getDisplayText(row, col));
  const annotation: CellFormatAnnotation = {
    ...describeNumberFormat(
      format.numFmtId,
      format.formatCode,
      value.value,
      isDate1904(session.workbook),
    ),
  };
  if (display !== null) {
    annotation.text = display.text;
    if (display.displayStatus === DisplayStatus.Overflow) {
      annotation.textStatus = "overflow";
    } else if (display.displayStatus === DisplayStatus.InvalidFormat) {
      annotation.textStatus = "invalidFormat";
    }
  }
  return annotation;
}

/**
 * `{ stale: true }` for a formula cell whose value predates a change to its inputs, which only
 * manual or automatic-except-tables calculation leaves behind; empty otherwise, so the output
 * stays sparse.
 */
export function staleMark(
  session: WorkbookOwner,
  sheet: number,
  row: number,
  col: number,
  formula: string,
): { stale?: true } {
  return formula !== "" && session.workbook.sheet(sheet).isCellStale(row, col)
    ? { stale: true }
    : {};
}

/**
 * Runs an optional engine read for an inspection payload and converts its result to JSON;
 * a read the engine refuses yields `null` instead of failing the whole tool call.
 */
export function safeRead<T>(read: () => T): T | null {
  try {
    return read();
  } catch {
    return null;
  }
}

/** `safeRead` of a read whose result is reported as-is, converted to JSON. */
export function safeReadJson(read: () => unknown): unknown {
  return safeRead(() => resultToJson(read()));
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
  const handle = session.workbook.sheet(sheet);
  const all = handle.listCells();
  const limit = Math.max(0, Math.min(maxCells, all.length));
  // One read of every cell with a non-default XF; a stored cell absent from it has XF 0, as
  // `getCellXfIndex` reads it.
  const ownXfCells = includeStyles ? (safeRead(() => handle.ooxml.listCellXfIndices()) ?? []) : [];
  const ownXf = new Map(ownXfCells.map((entry) => [refA1(entry.row, entry.col), entry.xfIndex]));
  const xfJson = new Map<number, unknown>();
  const styleOf = (a1: string) => {
    const xfIndex = ownXf.get(a1) ?? 0;
    if (!xfJson.has(xfIndex)) {
      xfJson.set(
        xfIndex,
        safeReadJson(() => session.workbook.ooxml.getCellXf(xfIndex)),
      );
    }
    return { xfIndex, xf: xfJson.get(xfIndex) };
  };
  const cells = [];
  const stored = new Set<string>();
  for (const cell of all) {
    stored.add(refA1(cell.row, cell.col));
  }
  for (const cell of all.slice(0, limit)) {
    const value = jsonCellValue(taggedValue(handle, cell));
    const formula = cell.formula ?? "";
    const a1 = refA1(cell.row, cell.col);
    cells.push({
      row: cell.row,
      col: cell.col,
      a1,
      value,
      formula,
      kind: cellKind(value, formula),
      style: includeStyles ? styleOf(a1) : undefined,
    });
  }
  cells.sort((left, right) => left.row - right.row || left.col - right.col);
  // Cells formatted but holding no content, which `listCells` does not list.
  const formattedOnly = ownXfCells.filter((entry) => !stored.has(refA1(entry.row, entry.col)));
  const formattedLimit = Math.max(0, Math.min(maxCells, formattedOnly.length));
  const formattedCells = includeStyles
    ? formattedOnly.slice(0, formattedLimit).map(({ row, col }) => {
        const a1 = refA1(row, col);
        return { row, col, a1, style: styleOf(a1) };
      })
    : undefined;
  return {
    cells,
    cellCount: all.length,
    truncated: limit < all.length,
    formattedCells,
    formattedCellCount: includeStyles ? formattedOnly.length : undefined,
    formattedTruncated: includeStyles ? formattedLimit < formattedOnly.length : undefined,
  };
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
  return Array.from({ length: session.workbook.sheetCount }, (_, index) => index);
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
export function requiredString(value: string | undefined, name: string): string {
  if (value === undefined || value.length === 0) {
    throw statusError(StatusCode.InvalidArgument, `${name} is required`);
  }
  return value;
}

export function requiredNumber(value: number | undefined, name: string): number {
  if (value === undefined) {
    throw statusError(StatusCode.InvalidArgument, `${name} is required`);
  }
  if (!Number.isFinite(value)) {
    throw statusError(StatusCode.InvalidArgument, `${name} must be finite`);
  }
  return value;
}

export function requiredBoolean(value: boolean | undefined, name: string): boolean {
  if (value === undefined) {
    throw statusError(StatusCode.InvalidArgument, `${name} is required`);
  }
  return value;
}
