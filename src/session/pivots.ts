import type { PivotAxis } from "@libraz/formulon";
import { assertStatus, resultToJson, statusToJson, type Workbook } from "../formulon.js";
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
  registerPivotNumberFormat,
  resolveFieldIndex,
  resolveRange,
  resolveTarget,
  type SourceValue,
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

function writeCacheValue(
  workbook: Workbook,
  cacheId: number,
  record: number,
  field: number,
  value: SourceValue,
): void {
  const status =
    value.kind === 0
      ? workbook.pivotCacheRecordSetBlank(cacheId, record, field)
      : value.kind === 1
        ? workbook.pivotCacheRecordSetNumber(cacheId, record, field, value.number)
        : value.kind === 2
          ? workbook.pivotCacheRecordSetBool(cacheId, record, field, value.boolean)
          : value.kind === 3
            ? workbook.pivotCacheRecordSetText(cacheId, record, field, value.text)
            : workbook.pivotCacheRecordSetError(cacheId, record, field, value.errorCode);
  assertStatus(status, `write pivot cache record ${record}, field ${field}`);
}

function cleanupPivot(
  session: WorkbookSession,
  targetSheet: number,
  pivotIndex: number | undefined,
  cacheId: number | undefined,
): void {
  if (pivotIndex !== undefined) {
    try {
      session.workbook.pivotRemove(targetSheet, pivotIndex);
    } catch {
      // Keep the original engine error; cleanup is best effort.
    }
  }
  if (cacheId !== undefined) {
    try {
      session.workbook.pivotCacheRemove(cacheId);
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
    assertStatus(session.workbook.recalc(), "recalc pivot source");
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
    const cache = session.workbook.pivotCacheCreate(0);
    assertStatus(cache.status, "create pivot cache");
    cacheId = cache.index;
    assertStatus(
      session.workbook.pivotCacheSetWorksheetSource(cacheId, {
        present: true,
        ref: source.range.ref,
        sheet: source.range.sheetName,
      }),
      "set pivot cache worksheet source",
    );

    for (const [index, header] of source.headers.entries()) {
      const added = session.workbook.pivotCacheFieldAdd(cacheId, header);
      assertStatus(added.status, `add pivot cache field ${header}`);
      if (added.index !== index) {
        throw new Error(
          `pivot cache field order mismatch: expected ${index}, received ${added.index}`,
        );
      }
    }
    for (const [recordIndex, record] of source.records.entries()) {
      const added = session.workbook.pivotCacheRecordAdd(cacheId);
      assertStatus(added.status, `add pivot cache record ${recordIndex}`);
      if (added.index !== recordIndex) {
        throw new Error(
          `pivot cache record order mismatch: expected ${recordIndex}, received ${added.index}`,
        );
      }
      for (const [fieldIndex, value] of record.entries()) {
        writeCacheValue(session.workbook, cacheId, recordIndex, fieldIndex, value);
      }
    }

    const created = session.workbook.pivotCreate(
      request.target.sheet,
      request.name,
      cacheId,
      request.target.row,
      request.target.col,
    );
    assertStatus(created.status, "create pivot table");
    pivotIndex = created.index;

    const pivotFieldIndexes: number[] = [];
    for (const [fieldIndex, header] of source.headers.entries()) {
      const added = session.workbook.pivotFieldAdd(request.target.sheet, pivotIndex, {
        sourceName: header,
        axis: resolved.assignments.get(fieldIndex) as PivotAxis,
      });
      assertStatus(added.status, `add pivot field ${header}`);
      if (added.index !== fieldIndex) {
        throw new Error(
          `pivot field order mismatch: expected ${fieldIndex}, received ${added.index}`,
        );
      }
      pivotFieldIndexes.push(added.index);
    }

    const rowOrder = request.rows.map(
      (field) => pivotFieldIndexes[resolveFieldIndex(source.headers, field)],
    );
    const colOrder = request.columns.map(
      (field) => pivotFieldIndexes[resolveFieldIndex(source.headers, field)],
    );
    assertStatus(
      session.workbook.pivotSetRowFieldOrder(request.target.sheet, pivotIndex, rowOrder),
      "set pivot row field order",
    );
    assertStatus(
      session.workbook.pivotSetColFieldOrder(request.target.sheet, pivotIndex, colOrder),
      "set pivot column field order",
    );

    for (const value of request.values) {
      const format =
        value.format === undefined
          ? undefined
          : String(registerPivotNumberFormat(session.workbook, value.format));
      const added = session.workbook.pivotDataFieldAdd(request.target.sheet, pivotIndex, {
        name: value.name,
        fieldIndex: pivotFieldIndexes[resolveFieldIndex(source.headers, value.field)],
        aggregation: AGGREGATIONS[value.aggregation],
        ...(format === undefined ? {} : { numberFormat: format }),
      });
      assertStatus(added.status, `add pivot data field ${value.name}`);
    }

    assertStatus(
      session.workbook.pivotSetLayout(request.target.sheet, pivotIndex, request.layout),
      "set pivot layout",
    );
    if (request.totals) {
      assertStatus(
        session.workbook.pivotSetGrandTotals(
          request.target.sheet,
          pivotIndex,
          request.totals.rows,
          request.totals.columns,
        ),
        "set pivot grand totals",
      );
    }
    const layout = session.workbook.pivotLayout(request.target.sheet, pivotIndex);
    assertStatus(layout.status, "read pivot layout");
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
      layout: resultToJson(layout),
      status: statusToJson(layout.status),
    };
  } catch (error) {
    // Only remove objects created by this operation. Any failure remains dirty
    // and unrelated tables, pivots, and workbook edits are left untouched.
    cleanupPivot(session, request.target.sheet, pivotIndex, cacheId);
    touch(session, true);
    throw error;
  }
}
