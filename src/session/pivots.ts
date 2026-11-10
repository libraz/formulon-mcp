import { PivotAxis } from "@libraz/formulon";
import { nativeToJson, okStatus, resultToJson } from "../formulon.js";
import {
  AGGREGATIONS,
  assignAxes,
  type CreatePivotOptions,
  formulaInRange,
  grandTotalsOptions,
  layoutKind,
  normalizedValues,
  normalizeFieldList,
  readSource,
  resolveFieldIndex,
  resolveRange,
  resolveTarget,
  sourceHeadersForValidation,
  sourceReference,
  targetReference,
  validateRangeSize,
} from "./pivot-source.js";
import { getSession, publicInfo, touch, type WorkbookSession } from "./state.js";

export type {
  CreatePivotOptions,
  PivotAggregationName,
  PivotGrandTotals,
  PivotValueOption,
} from "./pivot-source.js";

function cleanupPivot(
  session: WorkbookSession,
  targetSheet: number,
  pivotIndex: number | undefined,
  cacheId: number | undefined,
): void {
  if (pivotIndex !== undefined) {
    try {
      session.workbook.sheet(targetSheet).removePivot(pivotIndex);
    } catch {
      // Keep the original engine error; cleanup is best effort.
    }
  }
  if (cacheId !== undefined) {
    try {
      session.workbook.removePivotCache(cacheId);
    } catch {
      // A cache still referenced by a partially-created pivot is retained.
    }
  }
}

function validatePivotRequest(session: WorkbookSession, options: CreatePivotOptions) {
  const sourceRange = resolveRange(session, sourceReference(options), options.sourceSheet);
  const sourceLimit = options.sourceLimit ?? 10_000;
  if (!Number.isInteger(sourceLimit) || sourceLimit < 2 || sourceLimit > 10_000) {
    throw new Error("sourceLimit must be an integer from 2 through 10000");
  }
  validateRangeSize(sourceRange, sourceLimit);
  const rows = normalizeFieldList(options.rows, "row");
  const columns = normalizeFieldList(options.columns, "column");
  const pages = normalizeFieldList(options.pages, "page");
  const values = normalizedValues(options.values);
  const name = requiredText(options.name, "name");
  const layout = layoutKind(options.layout);
  const totals = grandTotalsOptions(options.grandTotals);
  // This preflight intentionally reads only headers. Unknown, unused, and
  // cross-axis assignments therefore fail before recalc/cache mutation.
  assignAxes(sourceHeadersForValidation(session, sourceRange), rows, columns, pages, values);
  const target = resolveTarget(
    session,
    targetReference(options),
    options.targetSheet ?? sourceRange.sheet,
  );
  return { sourceRange, sourceLimit, rows, columns, pages, values, name, layout, totals, target };
}

function requiredText(value: string | undefined, name: string): string {
  const trimmed = value?.trim() ?? "";
  if (!trimmed) {
    throw new Error(`${name} is required`);
  }
  return trimmed;
}

/** Creates a native PivotTable and its worksheet-backed cache in one operation. */
export function createSessionPivot(id: string, options: CreatePivotOptions) {
  const session = getSession(id);
  const request = validatePivotRequest(session, options);

  if (formulaInRange(session, request.sourceRange)) {
    touch(session, true);
    session.workbook.recalc();
  }
  const source = readSource(session, request.sourceRange, request.sourceLimit);
  // Formula headers can change after recalc, so repeat all header/axis checks
  // against the source values that will actually enter the cache.
  const resolved = assignAxes(
    source.headers,
    request.rows,
    request.columns,
    request.pages,
    request.values,
  );

  let cacheId: number | undefined;
  let pivotIndex: number | undefined;
  try {
    touch(session, true);
    const cache = session.workbook.addPivotCache();
    cacheId = cache.cacheId;
    const { firstRow, firstCol, lastRow, lastCol } = source.range;
    cache.setWorksheetSource({
      range: { sheet: source.range.sheet, rect: { firstRow, firstCol, lastRow, lastCol } },
    });

    for (const [index, header] of source.headers.entries()) {
      const added = cache.addField(header);
      if (added !== index) {
        throw new Error(`pivot cache field order mismatch: expected ${index}, received ${added}`);
      }
    }
    for (const [recordIndex, record] of source.records.entries()) {
      const added = cache.addRecord();
      if (added !== recordIndex) {
        throw new Error(
          `pivot cache record order mismatch: expected ${recordIndex}, received ${added}`,
        );
      }
      for (const [fieldIndex, value] of record.entries()) {
        cache.setRecordValue(recordIndex, fieldIndex, value);
      }
    }

    const pivot = session.workbook.sheet(request.target.sheet).addPivot({
      name: request.name,
      cacheId,
      anchor: { row: request.target.row, col: request.target.col },
    });
    pivotIndex = pivot.index;

    // A value field is placed by addDataField below; only the row, column and page axes are set here.
    for (const fieldIndex of source.headers.keys()) {
      const axis = resolved.assignments.get(fieldIndex);
      if (axis !== undefined && axis !== PivotAxis.Value) {
        pivot.setFieldAxis(fieldIndex, axis);
      }
    }

    pivot.setRowFieldOrder(request.rows.map((field) => resolveFieldIndex(source.headers, field)));
    pivot.setColumnFieldOrder(
      request.columns.map((field) => resolveFieldIndex(source.headers, field)),
    );

    for (const value of request.values) {
      pivot.addDataField({
        name: value.name,
        fieldIndex: resolveFieldIndex(source.headers, value.field),
        aggregation: AGGREGATIONS[value.aggregation],
        ...(value.format === undefined ? {} : { numberFormat: value.format }),
      });
    }

    pivot.setReportLayout(request.layout);
    if (request.totals) {
      pivot.updateGrandTotals(request.totals);
    }
    const layout = pivot.getLayout();
    return {
      session: publicInfo(session),
      source: {
        sheet: source.range.sheet,
        sheetName: source.range.sheetName,
        ref: source.range.ref,
        headers: source.headers,
        rows: source.records.length,
      },
      target: request.target,
      cacheId,
      pivotIndex,
      layout: resultToJson({
        rect: layout.range,
        cells: layout.cells.map((cell) => ({ ...cell, value: nativeToJson(cell.value) })),
      }),
      status: okStatus(),
    };
  } catch (error) {
    // Only remove objects created by this operation. Any failure remains dirty
    // and unrelated tables, pivots, and workbook edits are left untouched.
    cleanupPivot(session, request.target.sheet, pivotIndex, cacheId);
    touch(session, true);
    throw error;
  }
}
