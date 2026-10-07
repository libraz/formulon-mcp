import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { deflateSync } from "node:zlib";
import { test } from "vitest";
import { callTool, errorPayload, textPayload, withClient } from "./helpers/mcp.js";

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k += 1) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  return c >>> 0;
});

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, crc]);
}

/** Builds a solid red RGB PNG of the given pixel size. */
function solidPng(width, height) {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header.set([8, 2, 0, 0, 0], 8);
  const row = Buffer.concat([Buffer.from([0]), Buffer.alloc(width * 3, Buffer.from([255, 0, 0]))]);
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", header),
    chunk("IDAT", deflateSync(Buffer.concat(Array.from({ length: height }, () => row)))),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

test("inserts, lists, reads, and removes pictures through a save round trip", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-images-"));
  try {
    const logo = path.join(dir, "logo.png");
    const png = solidPng(40, 20);
    await writeFile(logo, png);
    const saved = path.join(dir, "images.xlsx");

    await withClient(async (client) => {
      const sessionId = "images";
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
      textPayload(
        await callTool(client, "formulon_set_cells", {
          sessionId,
          mutations: [{ type: "formula", sheet: 0, row: 0, col: 0, formula: "=1+1" }],
        }),
      );

      const first = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor: "B2",
          offsetX: { size: 4, unit: "px" },
          width: { size: 60, unit: "pt" },
          description: "Company logo",
        }),
      );
      assert.deepEqual(first.source, { path: logo, format: "png", pxWidth: 40, pxHeight: 20 });
      assert.equal(first.placed.from, "B2");
      assert.equal(first.placed.offset.x.px, 4);
      // Height follows the source aspect ratio when only the width is given.
      assert.deepEqual(first.placed.width, { pt: 60, px: 80 });
      assert.deepEqual(first.placed.height, { pt: 30, px: 40 });
      assert.equal(first.placed.description, "Company logo");

      const second = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: logo,
          anchor: "E5",
        }),
      );
      assert.deepEqual(second.placed.width, { pt: 30, px: 40 });

      const got = await callTool(client, "formulon_image_operation", {
        sessionId,
        operation: "get",
        objectId: first.objectId,
        outputPath: path.join(dir, "copy.png"),
      });
      assert.equal(got.content[0].type, "image");
      assert.equal(got.content[0].mimeType, "image/png");
      assert.deepEqual(Buffer.from(got.content[0].data, "base64"), png);
      assert.deepEqual(await readFile(path.join(dir, "copy.png")), png);

      textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "remove",
          objectId: second.objectId,
        }),
      );
      textPayload(
        await callTool(client, "formulon_save_session", { sessionId, outputPath: saved }),
      );
      textPayload(await callTool(client, "formulon_close_workbook", { sessionId }));

      textPayload(
        await callTool(client, "formulon_open_workbook", { sessionId: "reopened", path: saved }),
      );
      const listed = textPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId: "reopened",
          operation: "list",
        }),
      );
      assert.equal(listed.count, 1);
      assert.equal(listed.objects[0].kind, "picture");
      assert.equal(listed.objects[0].from, "B2");
      assert.equal(listed.objects[0].description, "Company logo");
      assert.deepEqual(listed.objects[0].width, { pt: 60, px: 80 });
      const cell = textPayload(
        await callTool(client, "formulon_get_cell", { sessionId: "reopened", a1: "A1" }),
      );
      assert.equal(cell.formula, "=1+1");
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("rejects image operations missing their required arguments", async () => {
  await withClient(async (client) => {
    const sessionId = "image-errors";
    textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
    assert.match(
      errorPayload(
        await callTool(client, "formulon_image_operation", { sessionId, operation: "insert" }),
      ),
      /path and anchor are required/,
    );
    assert.match(
      errorPayload(
        await callTool(client, "formulon_image_operation", { sessionId, operation: "remove" }),
      ),
      /objectId is required/,
    );
    assert.match(
      errorPayload(
        await callTool(client, "formulon_image_operation", {
          sessionId,
          operation: "insert",
          path: "logo.png",
          anchor: "Sheet1!B2",
        }),
      ),
      /without a sheet name/,
    );
  });
});
