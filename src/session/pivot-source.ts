import {
  type ErrorValue,
  ExcelError,
  type NativeValue,
  PivotAggregation,
  PivotAxis,
  PivotReportLayout,
  ValueKind,
} from "@libraz/formulon";
import { normalizeRange, parseCellRef, parseRangeRef } from "../a1.js";
import { findSheetIndex, StatusCode, sheetAt, statusError } from "../formulon.js";
import { assertInGrid, rangeA1, rectCells, sheetName } from "./shared.js";
import type { WorkbookSession } from "./state.js";

export type PivotAggregationName =
  | "sum"
  | "count"
  | "average"
  | "max"
  | "min"
  | "product"
  | "countNumbers"
  | "stddev"
  | "stddevp"
  | "var"
  | "varp";

export type PivotValueOption = {
  field: string;
  aggregation?: PivotAggregationName;
  name?: string;
  numberFormat?: string;
  format?: string;
};

export type PivotGrandTotals = {
  rows?: boolean;
  columns?: boolean;
};

export type CreatePivotOptions = {
  sourceRange?: string;
  source?: string;
  sourceSheet?: number | string;
  target?: string;
  targetAnchor?: string;
  targetSheet?: number | string;
  name?: string;
  rows?: string[];
  columns?: string[];
  pages?: string[];
  values?: PivotValueOption[];
  layout?: "compact" | "tabular" | "outline";
  grandTotals?: boolean | PivotGrandTotals;
  sourceLimit?: number;
};

export type PivotRange = {
  sheet: number;
  sheetName: string;
  firstRow: number;
  firstCol: number;
  lastRow: number;
  lastCol: number;
  ref: string;
};

export type SourceValue = null | number | boolean | string | ErrorValue;

export type SourceData = {
  range: PivotRange;
  headers: string[];
  records: SourceValue[][];
};

export const AGGREGATIONS: Record<PivotAggregationName, PivotAggregation> = {
  sum: PivotAggregation.Sum,
  count: PivotAggregation.Count,
  average: PivotAggregation.Average,
  max: PivotAggregation.Max,
  min: PivotAggregation.Min,
  product: PivotAggregation.Product,
  countNumbers: PivotAggregation.CountNumbers,
  stddev: PivotAggregation.StdDev,
  stddevp: PivotAggregation.StdDevP,
  var: PivotAggregation.Var,
  varp: PivotAggregation.VarP,
};

const AGGREGATION_LABELS: Record<PivotAggregationName, string> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
  product: "Product",
  countNumbers: "Count Numbers",
  stddev: "StdDev",
  stddevp: "StdDevP",
  var: "Var",
  varp: "VarP",
};

export function requiredText(value: string | undefined, name: string): string {
  const trimmed = value?.trim() ?? "";
  if (!trimmed) {
    throw statusError(StatusCode.InvalidArgument, `${name} is required`);
  }
  return trimmed;
}

export function sourceReference(options: CreatePivotOptions): string {
  const sourceRange = options.sourceRange?.trim();
  const source = options.source?.trim();
  if (sourceRange && source) {
    throw statusError(StatusCode.InvalidArgument, "provide only one of sourceRange or source");
  }
  return requiredText(sourceRange ?? source, "sourceRange");
}

export function targetReference(options: CreatePivotOptions): string {
  const target = options.target?.trim();
  const targetAnchor = options.targetAnchor?.trim();
  if (target && targetAnchor) {
    throw statusError(StatusCode.InvalidArgument, "provide only one of target or targetAnchor");
  }
  return requiredText(target ?? targetAnchor, "target");
}

export function resolveRange(
  session: WorkbookSession,
  reference: string,
  fallbackSheet: number | string | undefined,
): PivotRange {
  const parsed = normalizeRange(parseRangeRef(reference));
  assertInGrid(parsed.start.row, parsed.start.col);
  assertInGrid(parsed.end.row, parsed.end.col);
  const sheet = findSheetIndex(session.workbook, parsed.sheetName ?? fallbackSheet);
  const firstRow = parsed.start.row;
  const firstCol = parsed.start.col;
  const lastRow = parsed.end.row;
  const lastCol = parsed.end.col;
  return {
    sheet,
    sheetName: sheetName(session, sheet),
    firstRow,
    firstCol,
    lastRow,
    lastCol,
    ref: rangeA1(firstRow, firstCol, lastRow, lastCol),
  };
}

export function resolveTarget(
  session: WorkbookSession,
  reference: string,
  fallbackSheet: number | string | undefined,
) {
  const parsed = parseCellRef(reference);
  assertInGrid(parsed.row, parsed.col);
  const sheet = findSheetIndex(session.workbook, parsed.sheetName ?? fallbackSheet);
  return { sheet, sheetName: sheetName(session, sheet), row: parsed.row, col: parsed.col };
}

export function validateRangeSize(range: PivotRange, sourceLimit: number): void {
  const rows = range.lastRow - range.firstRow + 1;
  const cols = range.lastCol - range.firstCol + 1;
  if (rows < 2) {
    throw statusError(
      StatusCode.InvalidArgument,
      "pivot source must include a header row and at least one data row",
    );
  }
  if (cols < 1) {
    throw statusError(StatusCode.InvalidArgument, "pivot source must include at least one column");
  }
  const cells = rows * cols;
  if (cells > sourceLimit) {
    throw statusError(
      StatusCode.InvalidArgument,
      `pivot source has ${cells} cells; sourceLimit is ${sourceLimit}`,
    );
  }
}

export function cellRef(range: PivotRange, row: number, col: number): string {
  return `${range.sheetName}!${rangeA1(row, col, row, col)}`;
}

export function sourceValue(value: NativeValue, ref: string): SourceValue {
  if (value === null || typeof value === "boolean" || typeof value === "string") {
    return value;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw statusError(
        StatusCode.InvalidArgument,
        `pivot source cell ${ref} is not a finite number`,
      );
    }
    return value;
  }
  if (value instanceof ExcelError) {
    return { kind: ValueKind.Error, errorCode: value.code };
  }
  throw statusError(
    StatusCode.InvalidArgument,
    `pivot source cell ${ref} has an unsupported value; resolve arrays, references, and lambdas before creating the pivot`,
  );
}

export function formulaInRange(session: WorkbookSession, range: PivotRange): boolean {
  for (const cell of rectCells(session, range.sheet, range)) {
    if (cell.formula) {
      return true;
    }
  }
  return false;
}

export function readSource(
  session: WorkbookSession,
  range: PivotRange,
  sourceLimit: number,
): SourceData {
  validateRangeSize(range, sourceLimit);
  const headers = sourceHeaders(session, range);
  const sheet = sheetAt(session.workbook, range.sheet);
  const records: SourceValue[][] = [];
  for (let row = range.firstRow + 1; row <= range.lastRow; row += 1) {
    const record: SourceValue[] = [];
    for (let col = range.firstCol; col <= range.lastCol; col += 1) {
      record.push(sourceValue(sheet.getValue(row, col), cellRef(range, row, col)));
    }
    records.push(record);
  }
  if (records.length === 0) {
    throw statusError(
      StatusCode.InvalidArgument,
      "pivot source must include at least one data row",
    );
  }
  return { range, headers, records };
}

export function normalizeFieldList(fields: string[] | undefined, axis: string): string[] {
  const seen = new Set<string>();
  return (fields ?? []).map((field) => {
    const name = requiredText(field, `${axis} field`);
    const key = name.toLocaleLowerCase();
    if (seen.has(key)) {
      throw statusError(
        StatusCode.InvalidArgument,
        `${axis} fields must not repeat ${JSON.stringify(name)}`,
      );
    }
    seen.add(key);
    return name;
  });
}

export function resolveFieldIndex(headers: string[], name: string): number {
  const lowered = name.toLocaleLowerCase();
  const index = headers.findIndex((header) => header.toLocaleLowerCase() === lowered);
  if (index < 0) {
    throw statusError(
      StatusCode.InvalidArgument,
      `pivot field ${JSON.stringify(name)} is not present in source headers: ${headers.join(", ")}`,
    );
  }
  return index;
}

export function assignAxes(
  headers: string[],
  rows: string[],
  columns: string[],
  pages: string[],
  values: PivotValueOption[],
) {
  const assignments = new Map<number, PivotAxis>();
  const assign = (fields: string[], axis: PivotAxis, label: string) => {
    for (const field of fields) {
      const index = resolveFieldIndex(headers, field);
      const previous = assignments.get(index);
      if (previous !== undefined) {
        throw statusError(
          StatusCode.InvalidArgument,
          `pivot field ${JSON.stringify(headers[index])} is assigned to more than one axis (${label} and ${axisName(previous)}); assign each source header exactly once`,
        );
      }
      assignments.set(index, axis);
    }
  };
  assign(rows, PivotAxis.Row, "rows");
  assign(columns, PivotAxis.Column, "columns");
  assign(pages, PivotAxis.Page, "pages");

  for (const value of values) {
    const index = resolveFieldIndex(headers, requiredText(value.field, "value field"));
    const previous = assignments.get(index);
    if (previous !== undefined && previous !== PivotAxis.Value) {
      throw statusError(
        StatusCode.InvalidArgument,
        `pivot field ${JSON.stringify(headers[index])} is assigned to ${axisName(previous)} and values; assign each source header exactly once`,
      );
    }
    assignments.set(index, PivotAxis.Value);
  }

  const unused = headers.filter((_, index) => !assignments.has(index));
  if (unused.length > 0) {
    throw statusError(
      StatusCode.InvalidArgument,
      `pivot source headers are unused: ${unused.join(", ")}; assign every header to rows, columns, pages, or values`,
    );
  }
  return { assignments };
}

export function axisName(axis: PivotAxis): string {
  switch (axis) {
    case PivotAxis.Row:
      return "rows";
    case PivotAxis.Column:
      return "columns";
    case PivotAxis.Page:
      return "pages";
    default:
      return "values";
  }
}

export function normalizedValues(values: PivotValueOption[] | undefined) {
  if (!values || values.length === 0) {
    throw statusError(
      StatusCode.InvalidArgument,
      "values must contain at least one pivot value field",
    );
  }
  const names = new Set<string>();
  return values.map((value, index) => {
    const field = requiredText(value.field, `values[${index}].field`);
    const aggregation = value.aggregation ?? "sum";
    if (!(aggregation in AGGREGATIONS)) {
      throw statusError(
        StatusCode.InvalidArgument,
        `values[${index}].aggregation is unsupported: ${aggregation}`,
      );
    }
    const format = value.numberFormat ?? value.format;
    if (format !== undefined && !format.trim()) {
      throw statusError(
        StatusCode.InvalidArgument,
        `values[${index}].numberFormat must be a non-empty format code`,
      );
    }
    const name = (value.name?.trim() || `${AGGREGATION_LABELS[aggregation]} of ${field}`).trim();
    const key = name.toLocaleLowerCase();
    if (names.has(key)) {
      throw statusError(
        StatusCode.InvalidArgument,
        `pivot value names must be unique; duplicate ${JSON.stringify(name)}`,
      );
    }
    names.add(key);
    return { field, aggregation, format, name };
  });
}

export function layoutKind(layout: CreatePivotOptions["layout"]): PivotReportLayout {
  switch (layout ?? "compact") {
    case "compact":
      return PivotReportLayout.Compact;
    case "tabular":
      return PivotReportLayout.Tabular;
    case "outline":
      return PivotReportLayout.Outline;
    default:
      throw statusError(StatusCode.InvalidArgument, `unsupported pivot layout: ${layout}`);
  }
}

export function grandTotalsOptions(grandTotals: CreatePivotOptions["grandTotals"]) {
  if (grandTotals === undefined) {
    return undefined;
  }
  if (typeof grandTotals === "boolean") {
    return { rows: grandTotals, columns: grandTotals };
  }
  return { rows: grandTotals.rows ?? true, columns: grandTotals.columns ?? true };
}

export function sourceHeadersForValidation(session: WorkbookSession, range: PivotRange): string[] {
  return sourceHeaders(session, range);
}

function sourceHeaders(session: WorkbookSession, range: PivotRange): string[] {
  const headers: string[] = [];
  const seen = new Set<string>();
  const sheet = sheetAt(session.workbook, range.sheet);
  for (let col = range.firstCol; col <= range.lastCol; col += 1) {
    const value = sheet.getValue(range.firstRow, col);
    const ref = cellRef(range, range.firstRow, col);
    if (typeof value !== "string" || !value.trim()) {
      throw statusError(StatusCode.InvalidArgument, `pivot header ${ref} must be non-empty text`);
    }
    const header = value.trim();
    const key = header.toLocaleLowerCase();
    if (seen.has(key)) {
      throw statusError(
        StatusCode.InvalidArgument,
        `pivot headers must be unique; duplicate ${JSON.stringify(header)} at ${ref}`,
      );
    }
    seen.add(key);
    headers.push(header);
  }
  return headers;
}
