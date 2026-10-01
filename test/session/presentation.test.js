import assert from "node:assert/strict";
import { test } from "vitest";
import {
  applySessionMutations,
  callWorkbookMethod,
  closeSession,
  getSessionCell,
  getSessionCellByA1,
  getSessionRange,
  openSession,
  sessionDimension,
} from "../../dist/sessions.js";

/** Applies a custom number format to a cell via the low-level style APIs. */
function applyNumberFormat(id, sheet, row, col, formatCode) {
  const numFmtId = callWorkbookMethod(id, "addNumFmt", [formatCode]).result.numFmtId;
  const xfIndex = callWorkbookMethod(id, "addXf", [
    {
      fontIndex: 0,
      fillIndex: 0,
      borderIndex: 0,
      numFmtId,
      horizontalAlign: 0,
      verticalAlign: 0,
      wrapText: false,
    },
  ]).result.index;
  callWorkbookMethod(id, "setCellXfIndex", [sheet, row, col, xfIndex]);
}
test("supports calc mode, profile, partial recalc, and sheet protection", async () => {
  await openSession(undefined, "settings");
  try {
    assert.equal(callWorkbookMethod("settings", "calcMode", []).result.value, 0);
    assert.equal(callWorkbookMethod("settings", "setCalcMode", [1]).result.ok, true);
    assert.equal(callWorkbookMethod("settings", "calcMode", []).result.value, 1);

    assert.equal(
      callWorkbookMethod("settings", "excelProfileId", []).result.value,
      "win-365-ja_JP",
    );
    assert.equal(
      callWorkbookMethod("settings", "setExcelProfileId", ["mac-365-ja_JP"]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("settings", "excelProfileId", []).result.value,
      "mac-365-ja_JP",
    );

    applySessionMutations(
      "settings",
      [
        { type: "number", a1: "Sheet1!A1", value: 20 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1+22" },
      ],
      false,
    );
    const partial = callWorkbookMethod("settings", "partialRecalc", [
      { sheet: 0, firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 1 },
    ]);
    assert.equal(partial.result.status.ok, true);
    assert.equal(partial.result.recomputed >= 1, true);
    assert.deepEqual(getSessionCellByA1("settings", "Sheet1!B1").value, {
      kind: "number",
      value: 42,
    });

    const protection = callWorkbookMethod("settings", "getSheetProtection", [0]).result.protection;
    const nextProtection = { ...protection, enabled: 1, sheet: 1, selectLockedCells: 1 };
    assert.equal(
      callWorkbookMethod("settings", "setSheetProtection", [0, nextProtection]).result.ok,
      true,
    );
    const protectedSheet = callWorkbookMethod("settings", "getSheetProtection", [0]);
    assert.equal(protectedSheet.result.status.ok, true);
    assert.equal(protectedSheet.result.protection.enabled, 1);
  } finally {
    closeSession("settings");
  }
});

test("decodes date, currency, and percent number formats on reads", async () => {
  await openSession(undefined, "numfmt");
  try {
    applySessionMutations(
      "numfmt",
      [
        { type: "number", row: 0, col: 0, value: 45000 },
        { type: "number", row: 1, col: 0, value: 1500 },
        { type: "number", row: 2, col: 0, value: 0.153 },
        { type: "number", row: 3, col: 0, value: 42 },
      ],
      false,
    );
    applyNumberFormat("numfmt", 0, 0, 0, "yyyy-mm-dd");
    applyNumberFormat("numfmt", 0, 1, 0, '"¥"#,##0');
    applyNumberFormat("numfmt", 0, 2, 0, "0.0%");

    const date = getSessionCell("numfmt", 0, 0, 0);
    assert.equal(date.formatKind, "date");
    assert.equal(date.formatted, "2023-03-15");
    assert.equal(date.value.value, 45000);

    const currency = getSessionCell("numfmt", 0, 1, 0);
    assert.equal(currency.formatKind, "currency");
    assert.equal(currency.formatted, undefined);

    const percent = getSessionCell("numfmt", 0, 2, 0);
    assert.equal(percent.formatKind, "percent");
    assert.equal(percent.formatted, "15.3%");

    // Unformatted number carries no format annotation.
    const plain = getSessionCell("numfmt", 0, 3, 0);
    assert.equal(plain.formatKind, undefined);
    assert.equal(plain.numberFormat, undefined);

    const range = getSessionRange("numfmt", "A1:A4", {
      maxCells: 100,
      recalc: false,
      includeFormulas: false,
    });
    const a1 = range.cells.find((cell) => cell.a1 === "A1");
    assert.equal(a1.formatted, "2023-03-15");
    const a4 = range.cells.find((cell) => cell.a1 === "A4");
    assert.equal(a4.formatKind, undefined);
  } finally {
    closeSession("numfmt");
  }
});

test("dimension operation lists and sets column width and row height", async () => {
  await openSession(undefined, "dims");
  try {
    const setWidth = sessionDimension("dims", "column", "size", {
      sheet: 0,
      first: 0,
      last: 2,
      size: 120,
    });
    assert.equal(setWidth.status.ok, true);
    const setHeight = sessionDimension("dims", "row", "size", { sheet: 0, row: 0, size: 30 });
    assert.equal(setHeight.status.ok, true);
    const cols = sessionDimension("dims", "column", "list", { sheet: 0 });
    assert.equal(cols.value.status.ok, true);
    assert.throws(
      () => sessionDimension("dims", "column", "size", { sheet: 0, first: 0 }),
      /size is required/,
    );
  } finally {
    closeSession("dims");
  }
});
