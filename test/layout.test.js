import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { GeometryMode } from "../dist/formulon.js";
import { mmToPoints, pointsToMm, pointsToPixels } from "../dist/preview/geometry.js";
import {
  columnCharsToPoints,
  columnPointsToChars,
  resolveColumnMetrics,
} from "../dist/session/geometry.js";
import { applySessionLayout } from "../dist/session/layout.js";
import {
  applySessionMutations,
  callWorkbookMethod,
  closeSession,
  getSession,
  inspectSession,
  openSession,
  saveSession,
} from "../dist/sessions.js";

test("column geometry follows the engine width model", async () => {
  await openSession(undefined, "layout-geometry");
  try {
    const { workbook } = getSession("layout-geometry");
    const metrics = resolveColumnMetrics(workbook, 0);
    const model = workbook.sheet(0).getWidthModel(GeometryMode.Display);
    assert.equal(metrics.approximate, !model.calibrated);
    assert.equal(metrics.pointsPerChar, model.pointsPerChar);
    const expected = workbook.sheet(0).convertColumnCharsToPt(30, GeometryMode.Display);
    assert.equal(columnCharsToPoints(workbook, 0, 30), expected);
    assert.ok(Math.abs(columnPointsToChars(workbook, 0, expected) - 30) < 1e-9);
    assert.equal(columnPointsToChars(workbook, 0, 0), 0);
    // Stored widths include the cell padding, so even 1pt is a positive width.
    const narrow = columnPointsToChars(workbook, 0, 1);
    assert.equal(narrow, workbook.sheet(0).convertColumnPtToChars(1, GeometryMode.Display));
    assert.ok(narrow > 0);
    assert.equal(pointsToMm(mmToPoints(25.4)), 25.4);
  } finally {
    closeSession("layout-geometry");
  }
});

test("layout batches preserve operation order and report physical geometry", async () => {
  await openSession(undefined, "layout-order");
  try {
    applySessionMutations(
      "layout-order",
      [
        { type: "text", a1: "A1", value: "Header" },
        { type: "number", a1: "B1", value: 42 },
      ],
      false,
    );
    const result = applySessionLayout("layout-order", 0, [
      { type: "column", first: 0, size: 30 },
      { type: "row", first: 0, size: 40, unit: "px" },
      { type: "style", range: "A1:B1", style: { border: { all: "thin" } } },
      { type: "style", range: "A1:A1", style: { fill: { color: "#FF0000" } } },
      { type: "style", range: "A1:A1", style: { font: { bold: true } } },
      { type: "merge", operation: "add", range: "A2:B2" },
    ]);
    assert.equal(result.operations.length, 6);
    const { workbook } = getSession("layout-order");
    assert.equal(
      result.operations[0].displayGeometry.pixels,
      pointsToPixels(workbook.sheet(0).convertColumnCharsToPt(30, GeometryMode.Display)),
    );
    assert.equal(result.operations[1].displayGeometry.pixels, 40);
    assert.equal(result.operations[5].range, "A2:B2");

    const tables = (method, ...args) =>
      callWorkbookMethod("layout-order", { kind: "WorkbookOoxml" }, method, args).result;
    const xfIndex = callWorkbookMethod(
      "layout-order",
      { kind: "SheetOoxml", sheet: 0 },
      "getCellXfIndex",
      [0, 0],
    ).result;
    const xf = tables("getCellXf", xfIndex);
    const font = tables("getFont", xf.fontIndex);
    const fill = tables("getFill", xf.fillIndex);
    const border = tables("getBorder", xf.borderIndex);
    assert.equal(font.bold, true);
    assert.equal(fill.foreground.argb, 0xffff0000);
    assert.equal(border.left.color.argb, 0xff000000);
    assert.equal(
      callWorkbookMethod("layout-order", { kind: "Sheet", sheet: 0 }, "listMerges", []).result
        .length,
      1,
    );
  } finally {
    closeSession("layout-order");
  }
});

test("layout accepts millimetres and stores a narrow column with its padding", async () => {
  await openSession(undefined, "layout-units");
  try {
    const result = applySessionLayout("layout-units", 0, [
      { type: "column", first: 0, size: 25.4, unit: "mm" },
      { type: "row", first: 0, size: 10, unit: "mm" },
      { type: "column", first: 1, size: 1, unit: "pt" },
    ]);
    assert.ok(Math.abs(result.operations[0].displayGeometry.mm - 25.4) < 1e-10);
    assert.ok(Math.abs(result.operations[1].displayGeometry.mm - 10) < 1e-10);
    assert.ok(result.operations[2].storedSize > 0);
    // Display snaps to whole pixels: 1pt shows as one 0.75pt pixel.
    assert.equal(result.operations[2].displayGeometry.points, 0.75);
  } finally {
    closeSession("layout-units");
  }
});

test("layout preflight leaves a clean session when a later operation is invalid", async () => {
  await openSession(undefined, "layout-preflight");
  try {
    assert.throws(
      () =>
        applySessionLayout("layout-preflight", 0, [
          { type: "column", first: 0, size: 30 },
          { type: "style", range: "A1048577", style: { font: { bold: true } } },
        ]),
      /exceeds Excel sheet limits/,
    );
    assert.equal(inspectSession("layout-preflight", false, 10).session.dirty, false);
    assert.deepEqual(
      callWorkbookMethod("layout-preflight", { kind: "Sheet", sheet: 0 }, "listColumnOverrides", [])
        .result,
      [],
    );
    for (const operation of [
      { type: "column", first: 0, size: 10, unit: "bogus" },
      { type: "row", first: 0, size: 10, unit: "bogus" },
      { type: "merge", operation: "bogus", range: "A1:B1" },
    ]) {
      assert.throws(() => applySessionLayout("layout-preflight", 0, [operation]), /unit|operation/);
      assert.equal(inspectSession("layout-preflight", false, 10).session.dirty, false);
    }
  } finally {
    closeSession("layout-preflight");
  }
});

test("layout mutations survive an XLSX save and reload", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-layout-test-"));
  const outputPath = path.join(dir, "layout.xlsx");
  let originalOpen = false;
  try {
    await openSession(undefined, "layout-reload");
    originalOpen = true;
    applySessionMutations("layout-reload", [{ type: "text", a1: "A1", value: "saved" }], false);
    applySessionLayout("layout-reload", 0, [
      { type: "column", first: 0, size: 30 },
      { type: "row", first: 0, size: 40, unit: "px" },
      { type: "merge", operation: "add", range: "A2:B2" },
      { type: "style", range: "A1:B2", style: { font: { bold: true } } },
      {
        type: "print",
        settings: { printArea: "A1:B2", printTitles: { repeatRows: " 1:2 ", repeatCols: " A:A " } },
      },
    ]);
    await saveSession("layout-reload", outputPath);
    closeSession("layout-reload");
    originalOpen = false;

    await openSession(outputPath, "layout-reloaded");
    try {
      const sheet = (method, ...args) =>
        callWorkbookMethod("layout-reloaded", { kind: "Sheet", sheet: 0 }, method, args).result;
      const columns = sheet("listColumnOverrides");
      const rows = sheet("listRowOverrides");
      const merges = sheet("listMerges");
      const printArea = sheet("getPrintArea");
      assert.equal(columns[0].widthChars, 30);
      assert.equal(rows[0].heightPt, 30);
      assert.equal(merges[0].firstRow, 1);
      assert.deepEqual(printArea, [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }]);
      assert.deepEqual(sheet("getTypedValue", 0, 0), { kind: "text", value: "saved" });
    } finally {
      closeSession("layout-reloaded");
    }
  } finally {
    if (originalOpen) {
      closeSession("layout-reload");
    }
    await rm(dir, { recursive: true, force: true });
  }
});
