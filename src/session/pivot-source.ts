import { PivotAggregation, PivotAxis, PivotReportLayout, type Value } from "@libraz/formulon";
import { normalizeRange, parseCellRef, parseRangeRef } from "../a1.js";
import { assertStatus, findSheetIndex, type Workbook } from "../formulon.js";
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

export type SourceValue =
  | { kind: 0 }
  | { kind: 1; number: number }
  | { kind: 2; boolean: boolean }
  | { kind: 3; text: string }
  | { kind: 4; errorCode: number };

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
    throw new Error(`${name} is required`);
  }
  return trimmed;
}

export function sourceReference(options: CreatePivotOptions): string {
  const sourceRange = options.sourceRange?.trim();
  const source = options.source?.trim();
  if (sourceRange && source) {
    throw new Error("provide only one of sourceRange or source");
  }
  return requiredText(sourceRange ?? source, "sourceRange");
}

export function targetReference(options: CreatePivotOptions): string {
  const target = options.target?.trim();
  const targetAnchor = options.targetAnchor?.trim();
  if (target && targetAnchor) {
    throw new Error("provide only one of target or targetAnchor");
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
    throw new Error("pivot source must include a header row and at least one data row");
  }
  if (cols < 1) {
    throw new Error("pivot source must include at least one column");
  }
  const cells = rows * cols;
  if (cells > sourceLimit) {
    throw new Error(`pivot source has ${cells} cells; sourceLimit is ${sourceLimit}`);
  }
}

export function cellRef(range: PivotRange, row: number, col: number): string {
  return `${range.sheetName}!${rangeA1(row, col, row, col)}`;
}

export function sourceValue(value: Value, ref: string): SourceValue {
  switch (value.kind) {
    case 0:
      return { kind: 0 };
    case 1:
      if (!Number.isFinite(value.number)) {
        throw new Error(`pivot source cell ${ref} is not a finite number`);
      }
      return { kind: 1, number: value.number };
    case 2:
      return { kind: 2, boolean: Boolean(value.boolean) };
    case 3:
      return { kind: 3, text: value.text };
    case 4:
      return { kind: 4, errorCode: value.errorCode };
    default:
      throw new Error(
        `pivot source cell ${ref} has unsupported value kind ${value.kind}; resolve arrays, references, and lambdas before creating the pivot`,
      );
  }
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
  const records: SourceValue[][] = [];
  for (let row = range.firstRow + 1; row <= range.lastRow; row += 1) {
    const record: SourceValue[] = [];
    for (let col = range.firstCol; col <= range.lastCol; col += 1) {
      const result = session.workbook.getValue(range.sheet, row, col);
      const ref = cellRef(range, row, col);
      assertStatus(result.status, `read pivot source ${ref}`);
      record.push(sourceValue(result.value, ref));
    }
    records.push(record);
  }
  if (records.length === 0) {
    throw new Error("pivot source must include at least one data row");
  }
  return { range, headers, records };
}

export function normalizeFieldList(fields: string[] | undefined, axis: string): string[] {
  const seen = new Set<string>();
  return (fields ?? []).map((field) => {
    const name = requiredText(field, `${axis} field`);
    const key = name.toLocaleLowerCase();
    if (seen.has(key)) {
      throw new Error(`${axis} fields must not repeat ${JSON.stringify(name)}`);
    }
    seen.add(key);
    return name;
  });
}

export function resolveFieldIndex(headers: string[], name: string): number {
  const lowered = name.toLocaleLowerCase();
  const index = headers.findIndex((header) => header.toLocaleLowerCase() === lowered);
  if (index < 0) {
    throw new Error(
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
        throw new Error(
          `pivot field ${JSON.stringify(headers[index])} is assigned to more than one axis (${label} and ${axisName(previous)}); assign each source header exactly once`,
        );
      }
      assignments.set(index, axis);
    }
  };
  assign(rows, PivotAxis.Row, "rows");
  assign(columns, PivotAxis.Col, "columns");
  assign(pages, PivotAxis.Page, "pages");

  for (const value of values) {
    const index = resolveFieldIndex(headers, requiredText(value.field, "value field"));
    const previous = assignments.get(index);
    if (previous !== undefined && previous !== PivotAxis.Value) {
      throw new Error(
        `pivot field ${JSON.stringify(headers[index])} is assigned to ${axisName(previous)} and values; assign each source header exactly once`,
      );
    }
    assignments.set(index, PivotAxis.Value);
  }

  const unused = headers.filter((_, index) => !assignments.has(index));
  if (unused.length > 0) {
    throw new Error(
      `pivot source headers are unused: ${unused.join(", ")}; assign every header to rows, columns, pages, or values`,
    );
  }
  return { assignments };
}

export function axisName(axis: PivotAxis): string {
  switch (axis) {
    case PivotAxis.Row:
      return "rows";
    case PivotAxis.Col:
      return "columns";
    case PivotAxis.Page:
      return "pages";
    default:
      return "values";
  }
}

export function normalizedValues(values: PivotValueOption[] | undefined) {
  if (!values || values.length === 0) {
    throw new Error("values must contain at least one pivot value field");
  }
  const names = new Set<string>();
  return values.map((value, index) => {
    const field = requiredText(value.field, `values[${index}].field`);
    const aggregation = value.aggregation ?? "sum";
    if (!(aggregation in AGGREGATIONS)) {
      throw new Error(`values[${index}].aggregation is unsupported: ${aggregation}`);
    }
    const format = value.numberFormat ?? value.format;
    if (format !== undefined && !format.trim()) {
      throw new Error(`values[${index}].numberFormat must be a non-empty format code`);
    }
    const name = (value.name?.trim() || `${AGGREGATION_LABELS[aggregation]} of ${field}`).trim();
    const key = name.toLocaleLowerCase();
    if (names.has(key)) {
      throw new Error(`pivot value names must be unique; duplicate ${JSON.stringify(name)}`);
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
      throw new Error(`unsupported pivot layout: ${layout}`);
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
  for (let col = range.firstCol; col <= range.lastCol; col += 1) {
    const result = session.workbook.getValue(range.sheet, range.firstRow, col);
    const ref = cellRef(range, range.firstRow, col);
    assertStatus(result.status, `read pivot header ${ref}`);
    if (result.value.kind !== 3 || !result.value.text.trim()) {
      throw new Error(`pivot header ${ref} must be non-empty text`);
    }
    const header = result.value.text.trim();
    const key = header.toLocaleLowerCase();
    if (seen.has(key)) {
      throw new Error(
        `pivot headers must be unique; duplicate ${JSON.stringify(header)} at ${ref}`,
      );
    }
    seen.add(key);
    headers.push(header);
  }
  return headers;
}

export function registerPivotNumberFormat(workbook: Workbook, format: string): number {
  const result = workbook.addNumFmt(format);
  assertStatus(result.status, `register pivot number format ${format}`);
  return result.numFmtId;
}
