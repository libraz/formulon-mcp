import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  columnCharsToPoints,
  getColumnMetrics,
  mmToPoints,
  pointsToMm,
  pointsToPixels,
} from "../dist/preview/geometry.js";
import { applySessionLayout } from "../dist/session/layout.js";
import {
  applySessionMutations,
  callWorkbookMethod,
  closeSession,
  inspectSession,
  openSession,
  saveSession,
} from "../dist/sessions.js";

test("column geometry follows the sibling engine calibration", () => {
  const metrics = getColumnMetrics("Calibri", 11);
  assert.equal(metrics.approximate, false);
  assert.equal(columnCharsToPoints(30, metrics), 161.25);
  assert.equal(pointsToPixels(columnCharsToPoints(30, metrics)), 215);
  assert.equal(pointsToMm(mmToPoints(25.4)), 25.4);
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
    assert.equal(result.operations[0].displayGeometry.pixels, 215);
    assert.equal(result.operations[1].displayGeometry.pixels, 40);
    assert.equal(result.operations[5].range, "A2:B2");

    const xfIndex = callWorkbookMethod("layout-order", "getCellXfIndex", [0, 0, 0]).result;
    const xf = callWorkbookMethod("layout-order", "getCellXf", [xfIndex.xfIndex]).result;
    const font = callWorkbookMethod("layout-order", "getFont", [xf.fontIndex]).result;
    const fill = callWorkbookMethod("layout-order", "getFill", [xf.fillIndex]).result;
    const border = callWorkbookMethod("layout-order", "getBorder", [xf.borderIndex]).result;
    assert.equal(font.bold, true);
    assert.equal(fill.fgArgb, 0xffff0000);
    assert.equal(border.left.colorArgb, 0xff000000);
    assert.equal(callWorkbookMethod("layout-order", "getMerges", [0]).result.items.length, 1);
  } finally {
    closeSession("layout-order");
  }
});

test("layout accepts millimetres and rejects an unrepresentable narrow column", async () => {
  await openSession(undefined, "layout-units");
  try {
    const result = applySessionLayout("layout-units", 0, [
      { type: "column", first: 0, size: 25.4, unit: "mm" },
      { type: "row", first: 0, size: 10, unit: "mm" },
    ]);
    assert.ok(Math.abs(result.operations[0].displayGeometry.mm - 25.4) < 1e-10);
    assert.ok(Math.abs(result.operations[1].displayGeometry.mm - 10) < 1e-10);
    assert.throws(
      () =>
        applySessionLayout("layout-units", 0, [{ type: "column", first: 1, size: 1, unit: "pt" }]),
      /must exceed.*padding|exactly zero/,
    );
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
      callWorkbookMethod("layout-preflight", "getSheetColumns", [0]).result.columns,
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
      const columns = callWorkbookMethod("layout-reloaded", "getSheetColumns", [0]).result.columns;
      const rows = callWorkbookMethod("layout-reloaded", "getSheetRowOverrides", [0]).result.rows;
      const merges = callWorkbookMethod("layout-reloaded", "getMerges", [0]).result.items;
      const printArea = callWorkbookMethod("layout-reloaded", "getSheetPrintArea", [0]).result;
      assert.equal(columns[0].width, 30);
      assert.equal(rows[0].height, 30);
      assert.equal(merges[0].firstRow, 1);
      assert.equal(printArea.ranges, "A1:B2");
      const savedValue = callWorkbookMethod("layout-reloaded", "getValue", [0, 0, 0]).result;
      assert.equal(savedValue.status.ok, true);
      assert.deepEqual(savedValue.value, {
        kind: "text",
        value: "saved",
      });
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
