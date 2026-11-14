import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "vitest";
import { callTool, textPayload, withClient } from "./helpers/mcp.js";

/** An .xlsx holding one custom workbook view and one custom sheet view. */
const CUSTOM_VIEWS_FIXTURE = fileURLToPath(
  new URL("./fixtures/custom-views.xlsx", import.meta.url),
);

test("an .xlsb save that leaves out custom views counts them and keeps the session dirty", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-xlsb-losses-"));
  try {
    await withClient(async (client) => {
      textPayload(
        await callTool(client, "formulon_open_workbook", {
          sessionId: "views",
          path: CUSTOM_VIEWS_FIXTURE,
        }),
      );
      const xlsb = textPayload(
        await callTool(client, "formulon_save_session", {
          sessionId: "views",
          outputPath: path.join(dir, "views.xlsb"),
        }),
      );
      assert.equal(xlsb.losses?.deferredFeatures, 2);
      assert.equal(xlsb.session.dirty, true);

      const xlsx = textPayload(
        await callTool(client, "formulon_save_session", {
          sessionId: "views",
          outputPath: path.join(dir, "views.xlsx"),
        }),
      );
      assert.equal(xlsx.losses, undefined);
      assert.equal(xlsx.session.dirty, false);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
