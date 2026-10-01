import assert from "node:assert/strict";
import { test } from "vitest";
import { callWorkbookMethod, closeSession, openSession } from "../../dist/sessions.js";

test("supports pivot cache and pivot table creation through workbook calls", async () => {
  await openSession(undefined, "pivot");
  try {
    const cache = callWorkbookMethod("pivot", "pivotCacheCreate", [0]);
    assert.equal(cache.result.status.ok, true);
    const cacheId = cache.result.index;
    assert.equal(callWorkbookMethod("pivot", "pivotCacheCount", []).result.value, 1);

    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheFieldAdd", [cacheId, "Region"]).result.index,
      0,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheFieldAdd", [cacheId, "Amount"]).result.index,
      1,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotCacheFieldCount", [cacheId]).result.value, 2);
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheFieldName", [cacheId, 0]).result.value,
      "Region",
    );

    assert.equal(callWorkbookMethod("pivot", "pivotCacheRecordAdd", [cacheId]).result.index, 0);
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheRecordSetText", [cacheId, 0, 0, "East"]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotCacheRecordSetNumber", [cacheId, 0, 1, 10]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotCacheRecordCount", [cacheId]).result.value, 1);

    assert.equal(
      callWorkbookMethod("pivot", "pivotCreate", [0, "Pivot1", cacheId, 4, 0]).result.index,
      0,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotCount", [0]).result.value, 1);
    assert.equal(
      callWorkbookMethod("pivot", "pivotFieldAdd", [0, 0, { sourceName: "Region", axis: 0 }]).result
        .index,
      0,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotFieldAdd", [0, 0, { sourceName: "Amount", axis: 2 }]).result
        .index,
      1,
    );
    assert.equal(
      callWorkbookMethod("pivot", "pivotDataFieldAdd", [
        0,
        0,
        { name: "Sum of Amount", fieldIndex: 1, aggregation: 0 },
      ]).result.index,
      0,
    );

    const layout = callWorkbookMethod("pivot", "pivotLayout", [0, 0]);
    assert.equal(layout.result.status.ok, true);
    assert.equal(
      layout.result.cells.some((cell) => cell.value.kind === "number" && cell.value.value === 10),
      true,
    );

    assert.equal(
      callWorkbookMethod("pivot", "pivotFilterAdd", [
        0,
        0,
        { axis: 0, fieldName: "Region", type: 3, valueKind: 2, valueText: "Ea" },
      ]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("pivot", "pivotFilterCount", [0, 0]).result.value, 1);
    assert.equal(callWorkbookMethod("pivot", "pivotFilterClear", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotDataFieldClear", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotFieldClear", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotRemove", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotCount", [0]).result.value, 0);
    assert.equal(callWorkbookMethod("pivot", "pivotCacheRemove", [cacheId]).result.ok, true);
    assert.equal(callWorkbookMethod("pivot", "pivotCacheCount", []).result.value, 0);
  } finally {
    closeSession("pivot");
  }
});

test("groups pivot dates in seven-day intervals and formats aggregated values", async () => {
  await openSession(undefined, "pivot-days");
  try {
    const cache = callWorkbookMethod("pivot-days", "pivotCacheCreate", [0]);
    const cacheId = cache.result.index;
    callWorkbookMethod("pivot-days", "pivotCacheFieldAdd", [cacheId, "Date"]);
    callWorkbookMethod("pivot-days", "pivotCacheFieldAdd", [cacheId, "Amount"]);
    for (const [index, serial] of [45_000, 45_003, 45_009].entries()) {
      callWorkbookMethod("pivot-days", "pivotCacheRecordAdd", [cacheId]);
      callWorkbookMethod("pivot-days", "pivotCacheRecordSetNumber", [cacheId, index, 0, serial]);
      callWorkbookMethod("pivot-days", "pivotCacheRecordSetNumber", [
        cacheId,
        index,
        1,
        (index + 1) * 10,
      ]);
    }
    callWorkbookMethod("pivot-days", "pivotCreate", [0, "ByDays", cacheId, 4, 0]);
    callWorkbookMethod("pivot-days", "pivotFieldAdd", [0, 0, { sourceName: "Date", axis: 0 }]);
    callWorkbookMethod("pivot-days", "pivotFieldAdd", [0, 0, { sourceName: "Amount", axis: 2 }]);
    callWorkbookMethod("pivot-days", "pivotSetRowFieldOrder", [0, 0, [0]]);
    assert.equal(
      callWorkbookMethod("pivot-days", "pivotFieldSetDateGroup", [0, 0, 0, 4, 0, -1, -1, 7, -1, -1])
        .result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("pivot-days", "pivotFieldSetNumberFormat", [0, 0, 1, "14"]).result.ok,
      true,
    );
    assert.throws(
      () => callWorkbookMethod("pivot-days", "pivotFieldSetNumberFormat", [0, 0, 1, "0.00"]),
      /failed/,
    );
    const numFmtId = callWorkbookMethod("pivot-days", "addNumFmt", ["0.00"]).result.numFmtId;
    callWorkbookMethod("pivot-days", "pivotDataFieldAdd", [
      0,
      0,
      { name: "Sum", fieldIndex: 1, aggregation: 0, numberFormat: String(numFmtId) },
    ]);
    const layout = callWorkbookMethod("pivot-days", "pivotLayout", [0, 0]).result;
    assert.equal(layout.status.ok, true);
    assert.deepEqual(
      layout.cells.filter((cell) => cell.kind === 1).map((cell) => cell.value.value),
      ["2023/3/15 - 2023/3/21", "2023/3/22 - 2023/3/25"],
    );
    assert.deepEqual(
      layout.cells
        .filter((cell) => cell.value.kind === "number")
        .map((cell) => [cell.value.value, cell.numberFormat]),
      [
        [30, String(numFmtId)],
        [30, String(numFmtId)],
        [60, String(numFmtId)],
      ],
    );
  } finally {
    closeSession("pivot-days");
  }
});
