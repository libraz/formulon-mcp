import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { loadWorkbook } from "../../dist/formulon.js";
import { callWorkbookMethod, closeSession, openSession, saveSession } from "../../dist/sessions.js";

const sheet0 = { kind: "Sheet", sheet: 0 };
const ooxml = { kind: "WorkbookOoxml" };
const sheetOoxml0 = { kind: "SheetOoxml", sheet: 0 };
const call = (id, target, method, args = []) => callWorkbookMethod(id, target, method, args).result;
const column0 = { firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 };

test("supports layout, comments, hyperlinks, validations, and conditional formats", async () => {
  await openSession(undefined, "objects");
  try {
    call("objects", sheet0, "setColumnWidth", [0, 1, 22]);
    call("objects", sheet0, "setColumnHidden", [2, 2, true]);
    const columns = call("objects", sheet0, "listColumnOverrides");
    assert.equal(
      columns.some((column) => column.widthChars === 22),
      true,
    );

    call("objects", sheet0, "setRowHeight", [0, 0, 30]);
    const rows = call("objects", sheet0, "listRowOverrides");
    assert.equal(
      rows.some((row) => row.heightPt === 30),
      true,
    );

    call("objects", sheet0, "setComment", [0, 0, { author: "tester", text: "hello" }]);
    assert.deepEqual(call("objects", sheet0, "getComment", [0, 0]), {
      author: "tester",
      text: "hello",
    });

    call("objects", sheet0, "setHyperlink", [
      {
        target: "https://example.com",
        display: "Example",
        tooltip: "Example tooltip",
        range: "B1",
      },
    ]);
    call("objects", sheet0, "setHyperlink", [
      {
        target: null,
        location: "Sheet1!A1",
        display: "Jump",
        tooltip: "In-workbook",
        range: { firstRow: 2, firstCol: 0, lastRow: 2, lastCol: 3 },
      },
    ]);
    const hyperlinks = call("objects", sheet0, "listHyperlinks");
    assert.equal(hyperlinks.length, 2);
    assert.deepEqual(
      hyperlinks.map((link) => [
        link.range.firstRow,
        link.range.firstCol,
        link.range.lastRow,
        link.range.lastCol,
        link.location,
      ]),
      [
        [0, 1, 0, 1, null],
        [2, 0, 2, 3, "Sheet1!A1"],
      ],
    );

    call("objects", sheet0, "addValidation", [
      { ranges: [column0], type: 1, operator: 0, formula1: "1", formula2: "10" },
    ]);
    assert.equal(call("objects", sheet0, "listValidations").length, 1);

    call("objects", sheet0, "addCfRule", [{ ranges: [column0], type: 0, formula1: "=A1>0" }]);
    assert.equal(call("objects", sheet0, "listCfRules").length, 1);

    call("objects", sheet0, "removeHyperlink", ["B1"]);
    assert.equal(call("objects", sheet0, "listHyperlinks").length, 1);
    call("objects", sheet0, "clearHyperlinks");
    assert.equal(call("objects", sheet0, "listHyperlinks").length, 0);
    call("objects", sheet0, "removeValidation", [0]);
    assert.equal(call("objects", sheet0, "listValidations").length, 0);
    call("objects", sheet0, "removeCfRule", [0]);
    assert.equal(call("objects", sheet0, "listCfRules").length, 0);
    call("objects", sheet0, "clearMerges");
    assert.equal(call("objects", sheet0, "listMerges").length, 0);
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
    const id = "objects-roundtrip";
    call(id, sheet0, "setComment", [0, 0, { author: "tester", text: "hello" }]);
    call(id, sheet0, "setHyperlink", [
      {
        target: "https://example.com",
        display: "Example",
        tooltip: "Example tooltip",
        range: "B1",
      },
    ]);
    call(id, sheet0, "addValidation", [
      { ranges: [column0], type: 1, operator: 0, formula1: "1", formula2: "10" },
    ]);
    call(id, sheet0, "addCfRule", [{ ranges: [column0], type: 0, formula1: "=A1>0" }]);
    assert.equal(call(id, sheet0, "getSheetProtection"), null);
    call(id, sheet0, "setSheetProtection", [{}]);
    await saveSession(id, file);
    await saveSession(id, xlsb);
  } finally {
    closeSession("objects-roundtrip");
  }

  for (const output of [file, xlsb]) {
    const wb = await loadWorkbook(output);
    try {
      const ws = wb.sheet(0);
      if (output === file) {
        assert.deepEqual(ws.getComment(0, 0), { author: "tester", text: "hello" });
        assert.equal(ws.listHyperlinks().length, 1);
      }
      assert.equal(ws.listValidations().length, 1);
      assert.equal(ws.listCfRules().length, 1);
      assert.notEqual(ws.getSheetProtection(), null);
    } finally {
      wb.dispose();
    }
  }
  await rm(dir, { recursive: true, force: true });
});

test("supports style and number format APIs", async () => {
  await openSession(undefined, "styles");
  try {
    const font = call("styles", ooxml, "addFont", [
      {
        name: "Arial",
        sizePt: 11,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        color: "#112233",
      },
    ]);
    assert.equal(call("styles", ooxml, "getFont", [font]).name, "Arial");

    const fill = call("styles", ooxml, "addFill", [
      { pattern: 1, foreground: "#DDEEFF", background: "#FFFFFF" },
    ]);
    assert.equal(call("styles", ooxml, "getFill", [fill]).pattern, 1);

    const side = { style: 1, color: "#000000" };
    const border = call("styles", ooxml, "addBorder", [
      {
        left: side,
        right: side,
        top: side,
        bottom: side,
        diagonal: { style: 0 },
        diagonalUp: false,
        diagonalDown: false,
      },
    ]);
    assert.equal(call("styles", ooxml, "getBorder", [border]).left.style, 1);

    const numFmtId = call("styles", ooxml, "addNumberFormat", ["0.00"]);
    assert.equal(call("styles", ooxml, "getNumberFormat", [numFmtId]), "0.00");

    const xf = call("styles", ooxml, "addCellXf", [
      {
        fontIndex: font,
        fillIndex: fill,
        borderIndex: border,
        numFmtId,
        horizontal: 2,
        vertical: 1,
        wrapText: true,
      },
    ]);
    const record = call("styles", ooxml, "getCellXf", [xf]);
    assert.equal(record.horizontal, 2);
    assert.equal(record.wrapText, true);
    call("styles", sheetOoxml0, "setCellXfIndex", [0, 0, xf]);
    assert.equal(call("styles", sheetOoxml0, "getCellXfIndex", [0, 0]), xf);
  } finally {
    closeSession("styles");
  }
});

test("round-trips style, layout, and protection metadata through xlsx", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-style-"));
  const file = path.join(dir, "style.xlsx");
  await openSession(undefined, "style-roundtrip");
  try {
    const id = "style-roundtrip";
    const numFmtId = call(id, ooxml, "addNumberFormat", ["0.00"]);
    const xf = call(id, ooxml, "addCellXf", [
      {
        fontIndex: 0,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId,
        horizontal: 2,
        vertical: 1,
        wrapText: true,
      },
    ]);
    call(id, sheetOoxml0, "setCellXfIndex", [0, 0, xf]);
    call(id, sheet0, "setColumnWidth", [0, 0, 18]);
    call(id, sheet0, "setRowHeight", [0, 0, 24]);
    call(id, sheet0, "setSheetProtection", [{}]);
    await saveSession(id, file);
  } finally {
    closeSession("style-roundtrip");
  }

  const wb = await loadWorkbook(file);
  try {
    const ws = wb.sheet(0);
    assert.notEqual(ws.ooxml.getCellXfIndex(0, 0), null);
    assert.equal(ws.listColumnOverrides()[0].widthChars, 18);
    assert.equal(ws.listRowOverrides()[0].heightPt, 24);
    assert.notEqual(ws.getSheetProtection(), null);
  } finally {
    wb.dispose();
    await rm(dir, { recursive: true, force: true });
  }
});
