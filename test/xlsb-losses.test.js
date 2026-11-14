import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { test } from "vitest";
import { callTool, errorPayload, textPayload, withClient } from "./helpers/mcp.js";

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

test("an .xlsb save of a data bar with a formula threshold fails and writes nothing", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-xlsb-refused-"));
  try {
    await withClient(async (client) => {
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId: "bar" }));
      textPayload(
        await callTool(client, "formulon_set_range", {
          sessionId: "bar",
          start: "A1",
          values: [[1], [2], [3]],
        }),
      );
      textPayload(
        await callTool(client, "formulon_conditional_format_operation", {
          sessionId: "bar",
          operation: "add",
          rule: {
            type: "DataBar",
            ranges: "A1:A3",
            dataBar: {
              min: { type: "Formula", value: "=A1" },
              max: { type: "Max" },
              fill: "#0064C8",
            },
          },
        }),
      );
      const xlsbPath = path.join(dir, "bar.xlsb");
      const refused = errorPayload(
        await callTool(client, "formulon_save_session", { sessionId: "bar", outputPath: xlsbPath }),
      );
      assert.match(refused, /^IoFeatureNotWritable:/);
      assert.equal(existsSync(xlsbPath), false);

      const xlsx = textPayload(
        await callTool(client, "formulon_save_session", {
          sessionId: "bar",
          outputPath: path.join(dir, "bar.xlsx"),
        }),
      );
      assert.equal(xlsx.losses, undefined);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
