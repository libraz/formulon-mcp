import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { inflateRawSync } from "node:zlib";
import { test } from "vitest";
import { callTool, errorPayload, textPayload, withClient } from "./helpers/mcp.js";

/** An .xlsx holding one custom workbook view and one custom sheet view. */
const CUSTOM_VIEWS_FIXTURE = fileURLToPath(
  new URL("./fixtures/custom-views.xlsx", import.meta.url),
);

/** The bytes of the part `name` in the zip package `zip`, read through its central directory. */
const zipPart = (zip, name) => {
  let end = zip.length - 22;
  while (zip.readUInt32LE(end) !== 0x06054b50) {
    end -= 1;
  }
  let at = zip.readUInt32LE(end + 16);
  for (let i = 0; i < zip.readUInt16LE(end + 10); i += 1) {
    const method = zip.readUInt16LE(at + 10);
    const size = zip.readUInt32LE(at + 20);
    const nameLength = zip.readUInt16LE(at + 28);
    const local = zip.readUInt32LE(at + 42);
    if (zip.toString("utf8", at + 46, at + 46 + nameLength) === name) {
      const data = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28);
      const bytes = zip.subarray(data, data + size);
      return method === 8 ? inflateRawSync(bytes) : bytes;
    }
    at += 46 + nameLength + zip.readUInt16LE(at + 30) + zip.readUInt16LE(at + 32);
  }
  throw new Error(`${name} is not in the package`);
};

test("an .xlsb save keeps the custom views of an .xlsx workbook", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-xlsb-losses-"));
  try {
    await withClient(async (client) => {
      textPayload(
        await callTool(client, "formulon_open_workbook", {
          sessionId: "views",
          path: CUSTOM_VIEWS_FIXTURE,
        }),
      );
      const xlsbPath = path.join(dir, "views.xlsb");
      const xlsb = textPayload(
        await callTool(client, "formulon_save_session", {
          sessionId: "views",
          outputPath: xlsbPath,
        }),
      );
      assert.equal(xlsb.losses, undefined);
      assert.equal(xlsb.session.dirty, false);
      // BrtUserBookView (record 397) carries the custom workbook view.
      const book = zipPart(await readFile(xlsbPath), "xl/workbook.bin");
      assert.ok(book.includes(Buffer.from([0x8d, 0x03])));

      textPayload(
        await callTool(client, "formulon_open_workbook", { sessionId: "back", path: xlsbPath }),
      );
      const xlsxPath = path.join(dir, "views.xlsx");
      const xlsx = textPayload(
        await callTool(client, "formulon_save_session", {
          sessionId: "back",
          outputPath: xlsxPath,
        }),
      );
      assert.equal(xlsx.losses, undefined);
      const workbookXml = zipPart(await readFile(xlsxPath), "xl/workbook.xml").toString("utf8");
      assert.match(workbookXml, /<customWorkbookView /);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("an .xlsb save of a protected range with a descriptor the format cannot hold fails and writes nothing", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-xlsb-refused-"));
  try {
    await withClient(async (client) => {
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId: "guard" }));
      textPayload(
        await callTool(client, "formulon_workbook_call", {
          sessionId: "guard",
          target: { kind: "Sheet", sheet: 0 },
          method: "addProtectedRange",
          args: [{ name: "Inputs", ranges: "A1:A3", securityDescriptor: "not an sddl" }],
        }),
      );
      const xlsbPath = path.join(dir, "guard.xlsb");
      const refused = errorPayload(
        await callTool(client, "formulon_save_session", {
          sessionId: "guard",
          outputPath: xlsbPath,
        }),
      );
      assert.match(refused, /^IoFeatureNotWritable:/);
      assert.equal(existsSync(xlsbPath), false);

      const xlsx = textPayload(
        await callTool(client, "formulon_save_session", {
          sessionId: "guard",
          outputPath: path.join(dir, "guard.xlsx"),
        }),
      );
      assert.equal(xlsx.losses, undefined);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
