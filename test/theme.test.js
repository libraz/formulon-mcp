import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { callTool, errorPayload, textPayload, withClient } from "./helpers/mcp.js";

const ACCENT1 = 4;

test("saves a workbook's theme as a .thmx file and applies it to another workbook", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-theme-"));
  try {
    const thmx = path.join(dir, "brand.thmx");
    const book = path.join(dir, "book.xlsx");
    await withClient(async (client) => {
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId: "brand" }));
      const before = textPayload(
        await callTool(client, "formulon_workbook_call", {
          sessionId: "brand",
          method: "getThemeColors",
          args: [],
        }),
      ).result.colors;
      const colors = before.map((_, index) =>
        index === ACCENT1
          ? "#C0392B"
          : `#${(before[index] & 0xffffff).toString(16).padStart(6, "0")}`,
      );
      textPayload(
        await callTool(client, "formulon_workbook_call", {
          sessionId: "brand",
          method: "setThemeColors",
          args: [{ colors }],
        }),
      );
      const saved = textPayload(
        await callTool(client, "formulon_theme_operation", {
          sessionId: "brand",
          operation: "save",
          path: thmx,
        }),
      );
      assert.equal(saved.path, thmx);
      assert.equal(saved.colors.colors[ACCENT1] & 0xffffff, 0xc0392b);
      // A .thmx is a zip package.
      assert.deepEqual([...(await readFile(thmx)).subarray(0, 2)], [0x50, 0x4b]);

      textPayload(await callTool(client, "formulon_open_workbook", { sessionId: "plain" }));
      textPayload(
        await callTool(client, "formulon_save_session", { sessionId: "plain", outputPath: book }),
      );
      const applied = textPayload(
        await callTool(client, "formulon_theme_operation", {
          sessionId: "plain",
          operation: "apply",
          path: thmx,
        }),
      );
      assert.equal(applied.session.dirty, true);
      assert.deepEqual(applied.colors.colors, saved.colors.colors);

      // A workbook package is not a theme: the apply fails and changes nothing.
      const refused = errorPayload(
        await callTool(client, "formulon_theme_operation", {
          sessionId: "plain",
          operation: "apply",
          path: book,
        }),
      );
      assert.match(refused, /^IoRelationshipBroken:/);
      const after = textPayload(
        await callTool(client, "formulon_workbook_call", {
          sessionId: "plain",
          method: "getThemeColors",
          args: [],
        }),
      ).result;
      assert.deepEqual(after.colors, saved.colors.colors);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
