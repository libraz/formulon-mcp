import assert from "node:assert/strict";
import { deflateSync } from "node:zlib";
import { test } from "vitest";
import { GeometryMode } from "../dist/formulon.js";
import { previewSessionRange } from "../dist/session/preview.js";
import {
  callWorkbookMethod,
  closeSession,
  getSession,
  openSession,
  setSessionRange,
  styleSessionRange,
} from "../dist/sessions.js";
import { textPayload, withClient } from "./helpers/mcp.js";

function pngDimensions(png) {
  assert.equal(png.subarray(0, 8).toString("hex"), "89504e470d0a1a0a");
  return { width: png.readUInt32BE(16), height: png.readUInt32BE(20) };
}

function makeInvoice(id) {
  setSessionRange(
    id,
    "A1",
    [
      ["請求書😀<&", "発行日", "金額", "備考"],
      ["ACME", 45292, 1234.5, "税込"],
      ["商品", 1, 1234.5, "合計"],
      ["長い日本語タイトル😀<&", null, null, null],
    ],
    undefined,
    false,
  );
  styleSessionRange(
    id,
    "A1:D4",
    { border: { all: "thin" }, font: { name: "Calibri", size: 11 } },
    undefined,
    "existing",
  );
  styleSessionRange(
    id,
    "A1:D1",
    { fill: { color: "#DCE6F1" }, font: { bold: true }, align: { horizontal: "center" } },
    undefined,
    "existing",
  );
  styleSessionRange(id, "A4", { align: { wrapText: true } }, undefined, "existing");
  styleSessionRange(id, "B2:B2", { numberFormat: "yyyy-mm-dd" }, undefined, "existing");
  styleSessionRange(id, "C2:C3", { numberFormat: '"¥"#,##0.00' }, undefined, "existing");

  callWorkbookMethod(id, "setColumnWidth", [0, 0, 0, 4]);
  callWorkbookMethod(id, "setColumnWidth", [0, 1, 1, 8]);
  callWorkbookMethod(id, "setColumnWidth", [0, 2, 2, 22]);
  callWorkbookMethod(id, "setColumnWidth", [0, 3, 3, 5]);
  callWorkbookMethod(id, "setRowHeight", [0, 0, 30]);
  callWorkbookMethod(id, "addMerge", [0, { firstRow: 2, firstCol: 2, lastRow: 2, lastCol: 3 }]);
  callWorkbookMethod(id, "setSheetPrintArea", [0, "A1:D4"]);
  callWorkbookMethod(id, "addSheetRowBreak", [0, 2, true]);
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, crc]);
}

/** A solid-colour RGBA PNG of the given pixel size. */
function solidPng(width, height, [red, green, blue]) {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header.set([8, 6, 0, 0, 0], 8);
  const row = Buffer.concat([
    Buffer.from([0]),
    Buffer.from(Array.from({ length: width }, () => [red, green, blue, 255]).flat()),
  ]);
  const raw = Buffer.concat(Array.from({ length: height }, () => row));
  return Buffer.concat([
    Buffer.from("89504e470d0a1a0a", "hex"),
    pngChunk("IHDR", header),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

function hex(argb) {
  return `#${(argb & 0xffffff).toString(16).padStart(6, "0")}`;
}

test("renders a styled invoice range with merges, formatting, and pagination metadata", async () => {
  await openSession(undefined, "preview-invoice");
  try {
    makeInvoice("preview-invoice");
    const result = previewSessionRange("preview-invoice", {
      range: "A1:D4",
      showGridLines: true,
      showPageBreaks: true,
    });
    const dimensions = pngDimensions(result.png);
    assert.deepEqual(dimensions, {
      width: result.metadata.scaledWidth,
      height: result.metadata.scaledHeight,
    });
    assert.equal(result.metadata.range, "Sheet1!A1:D4");
    assert.equal(result.metadata.logicalUnit, "pt");
    assert.equal(result.metadata.cellGeometryUnit, "pt");
    assert.equal(result.metadata.rowBreaks.includes(2), true);
    assert.equal(result.metadata.pageCount >= 1, true);
    assert.match(result.svg, /¥1,234\.50/);
    assert.match(result.svg, /2024-01-01/);
    assert.equal(
      result.metadata.warnings.some((warning) => warning.includes("date/time format")),
      false,
    );
    assert.equal(result.metadata.pages.count, result.metadata.pageCount);
    assert.match(result.metadata.pages.ranges[0], /^A1:/);
    assert.match(result.svg, /&amp;/);
    assert.match(result.svg, /&lt;/);
    assert.match(result.svg, /😀/);
    assert.equal(result.svg.includes("\uFFFD"), false);
    assert.equal((result.svg.match(/<tspan /g) ?? []).length >= 2, true);
    assert.match(result.svg, /stroke="#000000"/);
    assert.match(result.svg, /clipPath/);
    // C3:D3 is an off-origin merge over unequal columns. The internal C|D
    // grid segment is suppressed while the merged right perimeter remains.
    assert.equal(
      /<line x1="253(?:\.000)?" y1="59\.429" x2="253(?:\.000)?" y2="[^"]+" stroke="#D9E0E7"/.test(
        result.svg,
      ),
      false,
    );
    assert.match(result.svg, /x1="293"[^>]*x2="293"/);
    for (const scale of [0.25, 1.5]) {
      const scaled = previewSessionRange("preview-invoice", {
        range: "A1:D4",
        scale,
      });
      assert.deepEqual(pngDimensions(scaled.png), {
        width: scaled.metadata.scaledWidth,
        height: scaled.metadata.scaledHeight,
      });
    }
  } finally {
    closeSession("preview-invoice");
  }
});

test("rejects invalid preview requests before an optional recalc dirties the session", async () => {
  await openSession(undefined, "preview-preflight");
  try {
    assert.throws(
      () => previewSessionRange("preview-preflight", { range: "A1048577", recalc: true }),
      /grid|row|range/i,
    );
    const session = (await import("../dist/sessions.js")).getSession("preview-preflight");
    assert.equal(session.dirty, false);
    assert.throws(
      () => previewSessionRange("preview-preflight", { range: "A1:Z1000", recalc: true }),
      /10,000|spans/,
    );
    assert.equal(session.dirty, false);
    assert.throws(
      () =>
        previewSessionRange("preview-preflight", {
          range: "A1",
          outputPath: "preview.txt",
          recalc: true,
        }),
      /outputPath/,
    );
    assert.equal(session.dirty, false);
  } finally {
    closeSession("preview-preflight");
  }
});

test("MCP preview returns image content before compact pagination metadata", async () => {
  await withClient(async (client) => {
    const sessionId = "preview-mcp";
    const opened = textPayload(
      await client.callTool({
        name: "formulon_open_workbook",
        arguments: { sessionId },
      }),
    );
    assert.equal(opened.session.id, sessionId);
    try {
      textPayload(
        await client.callTool({
          name: "formulon_set_range",
          arguments: {
            sessionId,
            start: "A1",
            values: [
              ["MCP preview", 10],
              ["second", 20],
            ],
            recalc: false,
          },
        }),
      );
      const preview = await client.callTool({
        name: "formulon_preview_range",
        arguments: { sessionId, range: "A1:B2", showPageBreaks: true },
      });
      assert.equal(preview.isError, undefined);
      assert.equal(preview.content[0].type, "image");
      assert.equal(preview.content[0].mimeType, "image/png");
      assert.match(preview.content[0].data, /^[A-Za-z0-9+/]+=*$/);
      assert.equal(preview.content[1].type, "text");
      const metadata = JSON.parse(preview.content[1].text);
      assert.equal(metadata.range, "Sheet1!A1:B2");
      assert.equal(Array.isArray(metadata.rowBreaks), true);
      assert.equal(typeof metadata.pageCount, "number");
    } finally {
      await client.callTool({ name: "formulon_close_workbook", arguments: { sessionId } });
    }
  });
});

test("a theme-coloured fill renders with the colour the engine resolves", async () => {
  await openSession(undefined, "preview-theme");
  try {
    const { workbook } = getSession("preview-theme");
    setSessionRange("preview-theme", "A1", [["x"]], undefined, false);
    const none = { kind: 0, rgb: 0, theme: 0, tint: 0, indexed: 0 };
    const fill = workbook.addFill({
      pattern: 1,
      fgArgb: 0,
      bgArgb: 0,
      fg: { ...none, kind: 2, theme: 4, tint: 0.4 },
      bg: none,
    });
    const base = workbook.getCellXf(0);
    const xf = workbook.addXf({ ...base, fillIndex: fill.index, applyFill: true });
    workbook.setCellXfIndex(0, 0, 0, xf.index);
    const expected = hex(workbook.getEffectiveStyle(0, 0, 0).fillForeground.argb);
    const result = previewSessionRange("preview-theme", { range: "A1:A1" });
    assert.ok(result.svg.includes(`fill="${expected}"`));
    assert.notEqual(expected, "#000000");
    assert.equal(
      result.metadata.warnings.some((entry) => entry.includes("ARGB fallback")),
      false,
    );
  } finally {
    closeSession("preview-theme");
  }
});

test("cell text is the engine display text for date, percent, and fraction formats", async () => {
  await openSession(undefined, "preview-display");
  try {
    const { workbook } = getSession("preview-display");
    setSessionRange("preview-display", "A1", [[45292], [0.256], [0.75]], undefined, false);
    styleSessionRange(
      "preview-display",
      "A1",
      { numberFormat: "yyyy-mm-dd" },
      undefined,
      "existing",
    );
    styleSessionRange("preview-display", "A2", { numberFormat: "0.0%" }, undefined, "existing");
    styleSessionRange("preview-display", "A3", { numberFormat: "# ?/?" }, undefined, "existing");
    const result = previewSessionRange("preview-display", { range: "A1:A3" });
    for (const row of [0, 1, 2]) {
      const text = workbook.getDisplayText(0, row, 0).text;
      assert.ok(text.length > 0);
      assert.equal(result.metadata.cells[row].text, text);
      assert.ok(result.svg.includes(`>${text}</tspan>`));
    }
    assert.equal(result.metadata.cells[1].text, "25.6%");
  } finally {
    closeSession("preview-display");
  }
});

test("sheet default row height drives preview geometry", async () => {
  await openSession(undefined, "preview-defaults");
  try {
    const { workbook } = getSession("preview-defaults");
    setSessionRange("preview-defaults", "A1", [["a"], ["b"]], undefined, false);
    const before = previewSessionRange("preview-defaults", { range: "A1:A2" });
    const status = workbook.setSheetFormatDefaults(0, {
      defaultRowHeight: 40,
      hasDefaultRowHeight: true,
    });
    assert.equal(status.ok, true);
    const result = previewSessionRange("preview-defaults", { range: "A1:A2" });
    const expected = workbook.getRowHeightPt(0, 0).value + workbook.getRowHeightPt(0, 1).value;
    assert.equal(result.metadata.logicalHeight, expected);
    assert.ok(result.metadata.logicalHeight > before.metadata.logicalHeight);
    const width = workbook.getColumnWidthPt(0, 0, GeometryMode.Display).value;
    assert.equal(result.metadata.logicalWidth, width);
  } finally {
    closeSession("preview-defaults");
  }
});

test("merges are read per range and a cutting range is rejected", async () => {
  await openSession(undefined, "preview-merge");
  try {
    setSessionRange("preview-merge", "A1", [["wide", null, "c"]], undefined, false);
    callWorkbookMethod("preview-merge", "addMerge", [
      0,
      { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 },
    ]);
    const result = previewSessionRange("preview-merge", { range: "A1:C2" });
    const covered = result.metadata.cells.filter((cell) => cell.text === "wide");
    assert.equal(covered.length, 1);
    const { workbook } = getSession("preview-merge");
    const first = workbook.getColumnWidthPt(0, 0, GeometryMode.Display).value;
    const second = workbook.getColumnWidthPt(0, 1, GeometryMode.Display).value;
    assert.equal(covered[0].width, first + second);
    assert.throws(() => previewSessionRange("preview-merge", { range: "B1:C2" }), /cuts merged/);
    const away = previewSessionRange("preview-merge", { range: "C1:C2" });
    assert.equal(away.metadata.cells.length, 1);
  } finally {
    closeSession("preview-merge");
  }
});

test("pictures render as embedded images and other objects as placeholders", async () => {
  await openSession(undefined, "preview-picture");
  try {
    const { workbook } = getSession("preview-picture");
    setSessionRange(
      "preview-picture",
      "A1",
      [
        ["a", "b"],
        ["c", "d"],
      ],
      undefined,
      false,
    );
    const inserted = workbook.insertImage(0, solidPng(4, 4, [255, 0, 0]), {
      row: 1,
      col: 1,
      rowOffEmu: 12700 * 2,
      colOffEmu: 12700 * 3,
      widthEmu: 12700 * 20,
      heightEmu: 12700 * 10,
    });
    assert.equal(inserted.status.ok, true);
    const cell = workbook.getCellRectPt(
      0,
      { firstRow: 1, firstCol: 1, lastRow: 1, lastCol: 1 },
      GeometryMode.Display,
    );
    const range = workbook.getCellRectPt(
      0,
      { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2 },
      GeometryMode.Display,
    );
    const result = previewSessionRange("preview-picture", { range: "A1:C3" });
    const match =
      /<image x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"[^>]*href="data:image\/png;base64,/.exec(
        result.svg,
      );
    assert.ok(match);
    const pt = (value) => (value * 96) / 72;
    assert.ok(Math.abs(Number(match[1]) - pt(cell.x + 3)) < 0.01);
    assert.ok(Math.abs(Number(match[2]) - pt(cell.y + 2)) < 0.01);
    assert.ok(Math.abs(Number(match[3]) - pt(20)) < 0.01);
    assert.ok(Math.abs(Number(match[4]) - pt(10)) < 0.01);
    assert.equal(result.metadata.objects.length, 1);
    assert.equal(result.metadata.objects[0].rendered, true);
    assert.ok(range.width > 0);

    // The raster must actually carry the picture's pixels.
    const { Resvg } = await import("@resvg/resvg-js");
    const rendered = new Resvg(result.svg).render();
    const x = Math.floor(pt(cell.x + 3 + 10));
    const y = Math.floor(pt(cell.y + 2 + 5));
    const offset = (y * rendered.width + x) * 4;
    assert.deepEqual(Array.from(rendered.pixels.subarray(offset, offset + 4)), [255, 0, 0, 255]);

    // A range that excludes the picture lists no object.
    const away = previewSessionRange("preview-picture", { range: "A1:A1" });
    assert.equal(away.metadata.objects.length, 0);
    assert.equal(away.svg.includes("<image"), false);
  } finally {
    closeSession("preview-picture");
  }
});
