import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { loadWorkbook } from "../../dist/formulon.js";
import { callWorkbookMethod, closeSession, openSession, saveSession } from "../../dist/sessions.js";

test("supports layout, comments, hyperlinks, validations, and conditional formats", async () => {
  await openSession(undefined, "objects");
  try {
    assert.equal(callWorkbookMethod("objects", "setColumnWidth", [0, 0, 1, 22]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "setColumnHidden", [0, 2, 2, true]).result.ok, true);
    const columns = callWorkbookMethod("objects", "getSheetColumns", [0]);
    assert.equal(columns.result.status.ok, true);
    assert.equal(
      columns.result.columns.some((column) => column.width === 22),
      true,
    );

    assert.equal(callWorkbookMethod("objects", "setRowHeight", [0, 0, 30]).result.ok, true);
    const rows = callWorkbookMethod("objects", "getSheetRowOverrides", [0]);
    assert.equal(rows.result.status.ok, true);
    assert.equal(
      rows.result.rows.some((row) => row.height === 30),
      true,
    );

    assert.equal(
      callWorkbookMethod("objects", "setComment", [0, 0, 0, "tester", "hello"]).result.ok,
      true,
    );
    assert.deepEqual(callWorkbookMethod("objects", "getComment", [0, 0, 0]).result, {
      author: "tester",
      text: "hello",
    });
    assert.deepEqual(callWorkbookMethod("objects", "getCommentResult", [0, 0, 0]).result, {
      status: { ok: true, status: 0, message: "", context: "" },
      comment: { author: "tester", text: "hello" },
    });

    assert.equal(
      callWorkbookMethod("objects", "addHyperlink", [
        0,
        0,
        1,
        "https://example.com",
        "Example",
        "Example tooltip",
        "",
      ]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects", "addHyperlinkRange", [
        0,
        2,
        0,
        2,
        3,
        "",
        "Jump",
        "In-workbook",
        "Sheet1!A1",
      ]).result.ok,
      true,
    );
    const hyperlinks = callWorkbookMethod("objects", "getHyperlinks", [0]).result.items;
    assert.equal(hyperlinks.length, 2);
    assert.deepEqual(
      hyperlinks.map((link) => [link.row, link.col, link.lastRow, link.lastCol, link.location]),
      [
        [0, 1, 0, 1, ""],
        [2, 0, 2, 3, "Sheet1!A1"],
      ],
    );

    assert.equal(
      callWorkbookMethod("objects", "addValidation", [
        0,
        {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 1,
          op: 0,
          formula1: "1",
          formula2: "10",
        },
      ]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("objects", "getValidations", [0]).result.items.length, 1);

    assert.equal(
      callWorkbookMethod("objects", "addConditionalFormat", [
        0,
        {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 0,
          formula1: "=A1>0",
        },
      ]).result.status.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects", "getConditionalFormats", [0]).result.items.length,
      1,
    );

    assert.equal(callWorkbookMethod("objects", "removeHyperlink", [0, 0, 1]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getHyperlinks", [0]).result.items.length, 1);
    assert.equal(callWorkbookMethod("objects", "clearHyperlinks", [0]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getHyperlinks", [0]).result.items.length, 0);
    assert.equal(callWorkbookMethod("objects", "removeValidationAt", [0, 0]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getValidations", [0]).result.items.length, 0);
    assert.equal(
      callWorkbookMethod("objects", "removeConditionalFormatAt", [0, 0]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects", "getConditionalFormats", [0]).result.items.length,
      0,
    );
    assert.equal(callWorkbookMethod("objects", "clearMerges", [0]).result.ok, true);
    assert.equal(callWorkbookMethod("objects", "getMerges", [0]).result.items.length, 0);
  } finally {
    closeSession("objects");
  }
});

test("round-trips comments, validations, conditional formats, and protection through xlsx/xlsb", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-objects-"));
  const file = path.join(dir, "objects.xlsx");
  const xlsb = path.join(dir, "objects.xlsb");
  await openSession(undefined, "objects-roundtrip");
  try {
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "setComment", [0, 0, 0, "tester", "hello"]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "addHyperlink", [
        0,
        0,
        1,
        "https://example.com",
        "Example",
        "Example tooltip",
        "",
      ]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "addValidation", [
        0,
        {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 1,
          op: 0,
          formula1: "1",
          formula2: "10",
        },
      ]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "addConditionalFormat", [
        0,
        {
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
          type: 0,
          formula1: "=A1>0",
        },
      ]).result.status.ok,
      true,
    );
    const protection = callWorkbookMethod("objects-roundtrip", "getSheetProtection", [0]).result;
    assert.equal(protection.status.ok, true);
    assert.equal(
      callWorkbookMethod("objects-roundtrip", "setSheetProtection", [
        0,
        { ...protection.protection, enabled: 1, sheet: 1 },
      ]).result.ok,
      true,
    );
    await saveSession("objects-roundtrip", file);
    await saveSession("objects-roundtrip", xlsb);
  } finally {
    closeSession("objects-roundtrip");
  }

  for (const output of [file, xlsb]) {
    const wb = await loadWorkbook(output);
    try {
      if (output === file) {
        assert.deepEqual(wb.getComment(0, 0, 0), { author: "tester", text: "hello" });
        assert.equal(wb.getHyperlinks(0).length, 1);
      }
      assert.equal(wb.getValidations(0).length, 1);
      assert.equal(wb.getConditionalFormats(0).length, 1);
      assert.equal(wb.getSheetProtection(0).protection.enabled, 1);
    } finally {
      wb.delete();
    }
  }
  await rm(dir, { recursive: true, force: true });
});

test("supports style and number format APIs", async () => {
  await openSession(undefined, "styles");
  try {
    const font = callWorkbookMethod("styles", "addFont", [
      {
        name: "Arial",
        size: 11,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff112233,
      },
    ]);
    assert.equal(font.result.status.ok, true);

    const fill = callWorkbookMethod("styles", "addFill", [
      { pattern: 1, fgArgb: 0xffddeeff, bgArgb: 0xffffffff },
    ]);
    assert.equal(fill.result.status.ok, true);

    const side = { style: 1, colorArgb: 0xff000000 };
    const border = callWorkbookMethod("styles", "addBorder", [
      {
        left: side,
        right: side,
        top: side,
        bottom: side,
        diagonal: { style: 0, colorArgb: 0 },
        diagonalUp: false,
        diagonalDown: false,
      },
    ]);
    assert.equal(border.result.status.ok, true);

    const numFmt = callWorkbookMethod("styles", "addNumFmt", ["0.00"]);
    assert.equal(numFmt.result.status.ok, true);

    const xf = callWorkbookMethod("styles", "addXf", [
      {
        fontIndex: font.result.index,
        fillIndex: fill.result.index,
        borderIndex: border.result.index,
        numFmtId: numFmt.result.numFmtId,
        horizontalAlign: 2,
        verticalAlign: 1,
        wrapText: true,
      },
    ]);
    assert.equal(xf.result.status.ok, true);
    assert.equal(
      callWorkbookMethod("styles", "setCellXfIndex", [0, 0, 0, xf.result.index]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("styles", "getCellXfIndex", [0, 0, 0]).result.xfIndex,
      xf.result.index,
    );
  } finally {
    closeSession("styles");
  }
});

test("round-trips style, layout, and protection metadata through xlsx", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-style-"));
  const file = path.join(dir, "style.xlsx");
  await openSession(undefined, "style-roundtrip");
  try {
    const numFmt = callWorkbookMethod("style-roundtrip", "addNumFmt", ["0.00"]);
    const xf = callWorkbookMethod("style-roundtrip", "addXf", [
      {
        fontIndex: 0,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: numFmt.result.numFmtId,
        horizontalAlign: 2,
        verticalAlign: 1,
        wrapText: true,
      },
    ]);
    assert.equal(
      callWorkbookMethod("style-roundtrip", "setCellXfIndex", [0, 0, 0, xf.result.index]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("style-roundtrip", "setColumnWidth", [0, 0, 0, 18]).result.ok,
      true,
    );
    assert.equal(callWorkbookMethod("style-roundtrip", "setRowHeight", [0, 0, 24]).result.ok, true);
    const protection = callWorkbookMethod("style-roundtrip", "getSheetProtection", [0]).result
      .protection;
    assert.equal(
      callWorkbookMethod("style-roundtrip", "setSheetProtection", [
        0,
        { ...protection, enabled: 1, sheet: 1 },
      ]).result.ok,
      true,
    );
    await saveSession("style-roundtrip", file);
  } finally {
    closeSession("style-roundtrip");
  }

  const wb = await loadWorkbook(file);
  try {
    assert.equal(wb.getCellXfIndex(0, 0, 0).status.ok, true);
    const columns = wb.getSheetColumns(0);
    assert.equal(columns.status.ok, true);
    assert.equal(columns.columns[0].width, 18);
    const rows = wb.getSheetRowOverrides(0);
    assert.equal(rows.status.ok, true);
    assert.equal(rows.rows[0].height, 24);
    assert.equal(wb.getSheetProtection(0).protection.enabled, 1);
  } finally {
    wb.delete();
    await rm(dir, { recursive: true, force: true });
  }
});
