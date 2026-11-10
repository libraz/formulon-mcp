import assert from "node:assert/strict";
import { test } from "vitest";
import { callWorkbookMethod, closeSession, openSession } from "../../dist/sessions.js";

const call = (id, target, method, args = []) => callWorkbookMethod(id, target, method, args).result;
const cacheOf = (cacheId) => ({ kind: "PivotCache", pivotCache: cacheId });
const pivotOf = { kind: "Pivot", sheet: 0, pivot: 0 };
const sheetOf = { kind: "Sheet", sheet: 0 };

test("supports pivot cache and pivot table creation through workbook calls", async () => {
  await openSession(undefined, "pivot");
  try {
    const cacheId = 7;
    const cache = call("pivot", undefined, "addPivotCache", [cacheId]);
    assert.equal(cache.handle, "PivotCache");
    assert.equal(call("pivot", undefined, "listPivotCaches").length, 1);

    assert.equal(call("pivot", cacheOf(cacheId), "addField", ["Region"]), 0);
    assert.equal(call("pivot", cacheOf(cacheId), "addField", ["Amount"]), 1);
    const fields = call("pivot", cacheOf(cacheId), "listFields");
    assert.equal(fields.length, 2);
    assert.equal(fields[0].name, "Region");

    assert.equal(call("pivot", cacheOf(cacheId), "addRecord"), 0);
    call("pivot", cacheOf(cacheId), "setRecordValue", [0, 0, "East"]);
    call("pivot", cacheOf(cacheId), "setRecordValue", [0, 1, 10]);
    assert.equal(call("pivot", cacheOf(cacheId), "getRecordCount"), 1);

    const created = call("pivot", sheetOf, "addPivot", [{ name: "Pivot1", cacheId, anchor: "E1" }]);
    assert.equal(created.handle, "Pivot");
    assert.equal(call("pivot", sheetOf, "listPivots").length, 1);
    call("pivot", pivotOf, "setFieldAxis", [0, 0]);
    call("pivot", pivotOf, "setFieldAxis", [1, 2]);
    assert.equal(
      call("pivot", pivotOf, "addDataField", [
        { name: "Sum of Amount", fieldIndex: 1, aggregation: 0 },
      ]),
      0,
    );

    const layout = call("pivot", pivotOf, "getLayout");
    assert.equal(
      layout.cells.some((cell) => cell.value === 10),
      true,
    );

    // PivotFilterKind.LabelContains is 3; valueKind/valueText became a single `value`.
    assert.equal(call("pivot", pivotOf, "addFilter", [{ fieldIndex: 0, kind: 3, value: "Ea" }]), 0);
    assert.equal(call("pivot", pivotOf, "listFilters").length, 1);
    call("pivot", pivotOf, "clearFilters");
    assert.equal(call("pivot", pivotOf, "listFilters").length, 0);
    call("pivot", pivotOf, "clearDataFields");
    assert.equal(call("pivot", pivotOf, "listDataFields").length, 0);
    call("pivot", pivotOf, "clearFields");
    call("pivot", sheetOf, "removePivot", [0]);
    assert.equal(call("pivot", sheetOf, "listPivots").length, 0);
    call("pivot", undefined, "removePivotCache", [cacheId]);
    assert.equal(call("pivot", undefined, "listPivotCaches").length, 0);
  } finally {
    closeSession("pivot");
  }
});

test("groups pivot dates in seven-day intervals and formats aggregated values", async () => {
  await openSession(undefined, "pivot-days");
  try {
    const cacheId = 3;
    call("pivot-days", undefined, "addPivotCache", [cacheId]);
    call("pivot-days", cacheOf(cacheId), "addField", ["Date"]);
    call("pivot-days", cacheOf(cacheId), "addField", ["Amount"]);
    for (const [index, serial] of [45_000, 45_003, 45_009].entries()) {
      call("pivot-days", cacheOf(cacheId), "addRecord");
      call("pivot-days", cacheOf(cacheId), "setRecordValue", [index, 0, serial]);
      call("pivot-days", cacheOf(cacheId), "setRecordValue", [index, 1, (index + 1) * 10]);
    }
    call("pivot-days", sheetOf, "addPivot", [{ name: "ByDays", cacheId, anchor: "A1" }]);
    call("pivot-days", pivotOf, "setFieldAxis", [0, 0]);
    call("pivot-days", pivotOf, "setFieldAxis", [1, 2]);
    call("pivot-days", pivotOf, "setRowFieldOrder", [[0]]);
    // PivotDatePeriod.Days is 8 (it was 4 before the bit values were renumbered).
    call("pivot-days", pivotOf, "setFieldDateGroup", [0, { periods: 8, intervalDays: 7 }]);
    call("pivot-days", pivotOf, "updateField", [1, { numberFormat: "14" }]);
    call("pivot-days", pivotOf, "addDataField", [
      { name: "Sum", fieldIndex: 1, aggregation: 0, numberFormat: "0.00" },
    ]);
    const layout = call("pivot-days", pivotOf, "getLayout");
    assert.deepEqual(
      layout.cells.filter((cell) => cell.kind === 1).map((cell) => cell.value),
      ["3/15/23 - 3/21/23", "3/22/23 - 3/25/23"],
    );
    assert.deepEqual(
      layout.cells
        .filter((cell) => typeof cell.value === "number")
        .map((cell) => [cell.value, cell.numberFormat]),
      [
        [30, "0.00"],
        [30, "0.00"],
        [60, "0.00"],
      ],
    );
  } finally {
    closeSession("pivot-days");
  }
});
