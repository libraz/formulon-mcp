import assert from "node:assert/strict";
import { test } from "vitest";
import { formatPreviewNumber } from "../dist/preview/format.js";
import { previewSessionRange } from "../dist/session/preview.js";
import {
  callWorkbookMethod,
  closeSession,
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

test("formats bounded numeric previews and warns for unsupported scientific codes", () => {
  assert.equal(formatPreviewNumber(1234.5, 4, "#,##0.00").text, "1,234.50");
  assert.equal(formatPreviewNumber(1234.5, 164, '"¥"#,##0.00').text, "¥1,234.50");
  assert.equal(formatPreviewNumber(0.1234, 10, "0.00%").text, "12.34%");
  const scientific = formatPreviewNumber(1234.5, 0, "0.00E+00");
  assert.equal(scientific.text, "1234.5");
  assert.match(scientific.warning ?? "", /approximated/);
});

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
    assert.equal(
      result.metadata.warnings.some((warning) => warning.includes("date/time format yyyy-mm-dd")),
      true,
    );
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
