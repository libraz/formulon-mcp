import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const require = createRequire(import.meta.url);
const FORMULON_VERSION = require("@libraz/formulon/package.json").version;

function textPayload(result) {
  assert.equal(result.content[0].type, "text");
  return JSON.parse(result.content[0].text);
}

function errorPayload(result) {
  assert.equal(result.isError, true);
  assert.equal(result.content[0].type, "text");
  return result.content[0].text;
}

async function withClient(fn) {
  const client = new Client({ name: "formulon-mcp-test", version: "0.1.0" });
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: ["./dist/index.js"],
    cwd: process.cwd(),
    stderr: "pipe",
  });
  await client.connect(transport);
  try {
    return await fn(client);
  } finally {
    await client.close();
  }
}

test("MCP stdio lists and calls core tools", async () => {
  await withClient(async (client) => {
    const tools = await client.listTools();
    const names = tools.tools.map((tool) => tool.name);
    assert.equal(names.includes("formulon_eval_formula"), true);
    assert.equal(names.includes("formulon_workbook_call"), true);
    assert.equal(names.includes("formulon_merge_operation"), true);
    assert.equal(names.includes("formulon_function_lookup"), true);
    assert.equal(names.includes("formulon_find_cells"), true);
    assert.equal(names.includes("formulon_replace_cells"), true);
    assert.equal(names.includes("formulon_inspect_layout"), true);
    assert.equal(names.includes("formulon_detect_regions"), true);
    assert.equal(names.includes("formulon_analyze_workbook"), true);
    assert.equal(
      tools.tools.every((tool) => tool.inputSchema && typeof tool.inputSchema === "object"),
      true,
    );

    const evalResult = textPayload(
      await client.callTool({
        name: "formulon_eval_formula",
        arguments: { formula: "=SUM(1,2,3)" },
      }),
    );
    assert.deepEqual(evalResult.value, { kind: "number", value: 6 });
  });
});

test("MCP stdio edits, reads, saves, and closes a workbook session", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-mcp-test-"));
  const outputPath = path.join(dir, "mcp.xlsx");
  try {
    await withClient(async (client) => {
      const opened = textPayload(
        await client.callTool({
          name: "formulon_open_workbook",
          arguments: { sessionId: "mcp-session" },
        }),
      );
      assert.equal(opened.session.id, "mcp-session");

      const setCells = textPayload(
        await client.callTool({
          name: "formulon_set_cells",
          arguments: {
            sessionId: "mcp-session",
            mutations: [
              { type: "number", a1: "Sheet1!A1", value: 7 },
              { type: "formula", a1: "Sheet1!B1", formula: "=A1*6" },
            ],
            recalc: true,
          },
        }),
      );
      assert.equal(setCells.applied.length, 2);

      const range = textPayload(
        await client.callTool({
          name: "formulon_get_range",
          arguments: { sessionId: "mcp-session", range: "Sheet1!A1:B1" },
        }),
      );
      assert.deepEqual(
        range.rows[0].map((entry) => entry.value),
        [
          { kind: "number", value: 7 },
          { kind: "number", value: 42 },
        ],
      );

      await client.callTool({
        name: "formulon_set_cells",
        arguments: {
          sessionId: "mcp-session",
          mutations: [{ type: "text", a1: "Sheet1!C1", value: "Draft budget" }],
          recalc: true,
        },
      });
      const found = textPayload(
        await client.callTool({
          name: "formulon_find_cells",
          arguments: { sessionId: "mcp-session", query: "budget" },
        }),
      );
      assert.deepEqual(
        found.results.map((result) => result.ref),
        ["Sheet1!C1"],
      );

      const replaced = textPayload(
        await client.callTool({
          name: "formulon_replace_cells",
          arguments: {
            sessionId: "mcp-session",
            query: "budget",
            replacement: "forecast",
            target: "texts",
          },
        }),
      );
      assert.equal(replaced.count, 1);
      const replacedCell = textPayload(
        await client.callTool({
          name: "formulon_get_cell",
          arguments: { sessionId: "mcp-session", a1: "Sheet1!C1" },
        }),
      );
      assert.deepEqual(replacedCell.value, { kind: "text", value: "Draft forecast" });

      const sheetCall = textPayload(
        await client.callTool({
          name: "formulon_workbook_call",
          arguments: { sessionId: "mcp-session", method: "sheetCount", args: [] },
        }),
      );
      assert.equal(sheetCall.result, 1);

      const saved = textPayload(
        await client.callTool({
          name: "formulon_save_session",
          arguments: { sessionId: "mcp-session", outputPath },
        }),
      );
      assert.equal(saved.bytes > 0, true);

      const closed = textPayload(
        await client.callTool({
          name: "formulon_close_workbook",
          arguments: { sessionId: "mcp-session" },
        }),
      );
      assert.equal(closed.session.id, "mcp-session");
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("MCP stdio exposes advanced dedicated workbook tools", async () => {
  await withClient(async (client) => {
    await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId: "advanced-session" },
    });

    try {
      const mergeAdd = textPayload(
        await client.callTool({
          name: "formulon_merge_operation",
          arguments: {
            sessionId: "advanced-session",
            operation: "add",
            range: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 },
          },
        }),
      );
      assert.equal(mergeAdd.result.ok, true);

      const merges = textPayload(
        await client.callTool({
          name: "formulon_merge_operation",
          arguments: { sessionId: "advanced-session", operation: "list" },
        }),
      );
      assert.deepEqual(merges.result, [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 }]);

      const commentSet = textPayload(
        await client.callTool({
          name: "formulon_comment_operation",
          arguments: {
            sessionId: "advanced-session",
            operation: "set",
            row: 1,
            col: 0,
            author: "tester",
            text: "note",
          },
        }),
      );
      assert.equal(commentSet.result.ok, true);

      const comment = textPayload(
        await client.callTool({
          name: "formulon_comment_operation",
          arguments: { sessionId: "advanced-session", operation: "get", row: 1, col: 0 },
        }),
      );
      assert.deepEqual(comment.result, { author: "tester", text: "note" });

      const hyperlinkAdd = textPayload(
        await client.callTool({
          name: "formulon_hyperlink_operation",
          arguments: {
            sessionId: "advanced-session",
            operation: "add",
            row: 2,
            col: 0,
            target: "https://example.com",
            display: "Example",
          },
        }),
      );
      assert.equal(hyperlinkAdd.result.ok, true);

      const hyperlinks = textPayload(
        await client.callTool({
          name: "formulon_hyperlink_operation",
          arguments: { sessionId: "advanced-session", operation: "list" },
        }),
      );
      assert.equal(hyperlinks.result.length, 1);

      const validationAdd = textPayload(
        await client.callTool({
          name: "formulon_validation_operation",
          arguments: {
            sessionId: "advanced-session",
            operation: "add",
            validation: {
              ranges: [{ firstRow: 0, firstCol: 2, lastRow: 9, lastCol: 2 }],
              type: 1,
              op: 0,
              formula1: "1",
              formula2: "10",
            },
          },
        }),
      );
      assert.equal(validationAdd.result.ok, true);

      const validations = textPayload(
        await client.callTool({
          name: "formulon_validation_operation",
          arguments: { sessionId: "advanced-session", operation: "list" },
        }),
      );
      assert.equal(validations.result.length, 1);

      const cfAdd = textPayload(
        await client.callTool({
          name: "formulon_conditional_format_operation",
          arguments: {
            sessionId: "advanced-session",
            operation: "add",
            rule: {
              sqref: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
              type: 0,
              formula1: "=A1>0",
            },
          },
        }),
      );
      assert.equal(cfAdd.result.ok, true);

      const cfs = textPayload(
        await client.callTool({
          name: "formulon_conditional_format_operation",
          arguments: { sessionId: "advanced-session", operation: "list" },
        }),
      );
      assert.equal(cfs.result.length, 1);

      await client.callTool({
        name: "formulon_set_cells",
        arguments: {
          sessionId: "advanced-session",
          mutations: [
            { type: "number", a1: "Sheet1!A1", value: 2 },
            { type: "formula", a1: "Sheet1!B1", formula: "=A1+1" },
          ],
          recalc: true,
        },
      });

      const precedents = textPayload(
        await client.callTool({
          name: "formulon_trace",
          arguments: {
            sessionId: "advanced-session",
            operation: "precedents",
            row: 0,
            col: 1,
          },
        }),
      );
      assert.deepEqual(precedents.result, [{ sheet: 0, row: 0, col: 0 }]);

      const metadata = textPayload(
        await client.callTool({
          name: "formulon_function_lookup",
          arguments: { sessionId: "advanced-session", operation: "metadata", name: "SUM" },
        }),
      );
      assert.equal(metadata.result.ok, true);
      assert.equal(metadata.result.name, "SUM");

      const mergeClear = textPayload(
        await client.callTool({
          name: "formulon_merge_operation",
          arguments: { sessionId: "advanced-session", operation: "clear" },
        }),
      );
      assert.equal(mergeClear.result.ok, true);

      const hyperlinkClear = textPayload(
        await client.callTool({
          name: "formulon_hyperlink_operation",
          arguments: { sessionId: "advanced-session", operation: "clear" },
        }),
      );
      assert.equal(hyperlinkClear.result.ok, true);

      const validationClear = textPayload(
        await client.callTool({
          name: "formulon_validation_operation",
          arguments: { sessionId: "advanced-session", operation: "clear" },
        }),
      );
      assert.equal(validationClear.result.ok, true);

      const cfClear = textPayload(
        await client.callTool({
          name: "formulon_conditional_format_operation",
          arguments: { sessionId: "advanced-session", operation: "clear" },
        }),
      );
      assert.equal(cfClear.result.ok, true);
    } finally {
      await client.callTool({
        name: "formulon_close_workbook",
        arguments: { sessionId: "advanced-session" },
      });
    }
  });
});

test("MCP stdio supports one-shot path tools", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-one-shot-"));
  const outputPath = path.join(dir, "one-shot.xlsx");
  try {
    await withClient(async (client) => {
      const updated = textPayload(
        await client.callTool({
          name: "formulon_update_workbook",
          arguments: {
            outputPath,
            mutations: [
              { type: "number", sheet: 0, row: 0, col: 0, value: 40 },
              { type: "formula", sheet: 0, row: 0, col: 1, formula: "=A1+2" },
            ],
          },
        }),
      );
      assert.equal(updated.bytes > 0, true);

      const inspected = textPayload(
        await client.callTool({
          name: "formulon_inspect_workbook",
          arguments: { path: outputPath, recalc: true, includeCells: true },
        }),
      );
      assert.equal(inspected.sheets[0].cellCount >= 2, true);

      const cell = textPayload(
        await client.callTool({
          name: "formulon_get_cell",
          arguments: { path: outputPath, sheet: 0, row: 0, col: 1 },
        }),
      );
      assert.deepEqual(cell.value, { kind: "number", value: 42 });
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("MCP stdio reports tool errors without crashing the server", async () => {
  await withClient(async (client) => {
    const missing = await client.callTool({
      name: "formulon_get_range",
      arguments: { sessionId: "missing", range: "A1:B1" },
    });
    assert.match(errorPayload(missing), /session not found/);

    await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId: "error-session" },
    });
    const rejected = await client.callTool({
      name: "formulon_workbook_call",
      arguments: { sessionId: "error-session", method: "constructor", args: [] },
    });
    assert.match(errorPayload(rejected), /not allowlisted/);

    await client.callTool({
      name: "formulon_close_workbook",
      arguments: { sessionId: "error-session" },
    });
    const afterClose = await client.callTool({
      name: "formulon_get_range",
      arguments: { sessionId: "error-session", range: "A1:B1" },
    });
    assert.match(errorPayload(afterClose), /session not found/);

    const version = textPayload(await client.callTool({ name: "formulon_version", arguments: {} }));
    assert.equal(version.version, FORMULON_VERSION);
  });
});
