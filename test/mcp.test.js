import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { test } from "vitest";

const require = createRequire(import.meta.url);
const FORMULON_VERSION = require("@libraz/formulon/package.json").version;
const SERVER_VERSION = require("../package.json").version;

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
    assert.equal(names.includes("formulon_dimension_operation"), true);
    assert.equal(names.includes("formulon_style_range"), true);
    assert.equal(names.includes("formulon_print_settings"), true);
    assert.equal(names.includes("formulon_build_document"), true);
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
        range.cells.map((entry) => entry.value),
        [
          { kind: "number", value: 7 },
          { kind: "number", value: 42 },
        ],
      );

      const sessionEval = textPayload(
        await client.callTool({
          name: "formulon_eval_formula",
          arguments: { sessionId: "mcp-session", formula: "=A1+B1" },
        }),
      );
      assert.deepEqual(sessionEval.result.value, { kind: "number", value: 49 });

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

test("MCP stdio builds a document from blocks in one call", async () => {
  await withClient(async (client) => {
    await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId: "blocks-session" },
    });

    try {
      const built = textPayload(
        await client.callTool({
          name: "formulon_build_document",
          arguments: {
            sessionId: "blocks-session",
            print: "a4-portrait-fit",
            theme: { accent: "#1F4E79" },
            blocks: [
              { type: "title", text: "Invoice" },
              { type: "spacer" },
              { type: "text", name: "to", text: "Sample Co.", span: 2 },
              {
                type: "fields",
                name: "meta",
                align: "right",
                sameRow: true,
                items: [
                  { label: "No.", value: "INV-1" },
                  { label: "Date", value: "2026-08-22", format: "date" },
                ],
              },
              { type: "spacer" },
              {
                type: "table",
                columns: [
                  { header: "Item", key: "name", width: 30 },
                  { header: "Qty", key: "qty", format: "number", align: "right" },
                  { header: "Unit", key: "unit", format: "number", align: "right" },
                  { header: "Amount", formula: "={qty}*{unit}", format: "number", align: "right" },
                ],
                rows: [
                  { name: "Design", qty: 3, unit: 120000 },
                  { name: "Build", qty: 5, unit: 98000 },
                ],
              },
              { type: "spacer" },
              {
                type: "summary",
                items: [
                  { label: "Subtotal", formula: "=SUM({table.Amount})", format: "number" },
                  { label: "Tax", formula: "=ROUND({Subtotal}*0.1,0)", format: "number" },
                  {
                    label: "Total",
                    formula: "={Subtotal}+{Tax}",
                    format: "number",
                    emphasis: true,
                  },
                ],
              },
            ],
          },
        }),
      );

      assert.equal(built.width, 4);
      assert.equal(built.pageCount, 1);
      assert.equal(built.names["table.Amount"], "E8:E9");
      // A sameRow pair sits side by side rather than stacking.
      assert.equal(built.names.to, "B4:C4");
      assert.equal(built.names.meta, "D4:E5");

      const total = textPayload(
        await client.callTool({
          name: "formulon_get_cell",
          arguments: { sessionId: "blocks-session", a1: built.names.Total },
        }),
      );
      assert.deepEqual(total.value, { kind: "number", value: 935000 });
      assert.equal(total.formula, "=E11+E12");

      // The returned name map is what makes the result refinable: styling a
      // block afterwards needs no knowledge of where it landed.
      const restyled = textPayload(
        await client.callTool({
          name: "formulon_style_range",
          arguments: {
            sessionId: "blocks-session",
            range: built.names.Total,
            style: { font: { size: 14 } },
          },
        }),
      );
      assert.equal(restyled.range.start, restyled.range.end);

      const unknown = errorPayload(
        await client.callTool({
          name: "formulon_build_document",
          arguments: {
            sessionId: "blocks-session",
            start: "H2",
            blocks: [{ type: "summary", items: [{ label: "Total", formula: "=SUM({nowhere})" }] }],
          },
        }),
      );
      assert.match(unknown, /unknown reference \{nowhere\}/);
    } finally {
      await client.callTool({
        name: "formulon_close_workbook",
        arguments: { sessionId: "blocks-session" },
      });
    }
  });
});

test("MCP stdio authors a styled, printable document", async () => {
  await withClient(async (client) => {
    await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId: "document-session" },
    });

    try {
      await client.callTool({
        name: "formulon_set_range",
        arguments: {
          sessionId: "document-session",
          start: "B2",
          values: [
            ["Item", "Qty", "Unit", "Amount"],
            ["Design", 3, 120000, { f: "=C3*D3" }],
            ["Build", 5, 98000, { f: "=C4*D4" }],
            ["Total", null, null, { f: "=SUM(E3:E4)" }],
          ],
        },
      });

      const ruled = textPayload(
        await client.callTool({
          name: "formulon_style_range",
          arguments: {
            sessionId: "document-session",
            range: "B2:E5",
            style: { border: { all: "thin", outline: { style: "medium", color: "#1F4E79" } } },
          },
        }),
      );
      // An outline splits the block into corners, edges, and interior.
      assert.equal(ruled.regions.length, 9);

      const header = textPayload(
        await client.callTool({
          name: "formulon_style_range",
          arguments: {
            sessionId: "document-session",
            range: "B2:E2",
            style: {
              font: { bold: true, color: "#FFFFFF" },
              fill: { color: "#1F4E79" },
              align: { horizontal: "center" },
            },
          },
        }),
      );
      assert.equal(header.range.start, "B2");
      assert.equal(header.range.end, "E2");

      const amounts = textPayload(
        await client.callTool({
          name: "formulon_style_range",
          arguments: {
            sessionId: "document-session",
            range: "E3:E5",
            style: { numberFormat: '"¥"#,##0' },
          },
        }),
      );
      assert.ok(amounts.regions.length >= 1);

      const printed = textPayload(
        await client.callTool({
          name: "formulon_print_settings",
          arguments: {
            sessionId: "document-session",
            pageSetup: { orientation: "portrait", paperSize: 9, fitToPage: true, fitToWidth: 1 },
            margins: { left: 0.6, right: 0.6 },
            printArea: "B2:E20",
            printTitles: { repeatRows: "2:2" },
            headerFooter: { oddFooter: "&C&P / &N" },
          },
        }),
      );
      assert.equal(printed.settings.pageSetup.orientation, "portrait");
      assert.equal(printed.settings.printArea, "B2:E20");
      assert.equal(printed.settings.printTitles.repeatRows, "2:2");
      assert.match(
        printed.settings.headerFooterXml,
        /<oddFooter>&amp;C&amp;P \/ &amp;N<\/oddFooter>/,
      );

      const read = textPayload(
        await client.callTool({
          name: "formulon_print_settings",
          arguments: { sessionId: "document-session" },
        }),
      );
      assert.equal(read.applied, undefined);
      assert.equal(read.settings.printArea, "B2:E20");
      assert.ok(read.settings.pageCount >= 1);

      const badStyle = errorPayload(
        await client.callTool({
          name: "formulon_style_range",
          arguments: {
            sessionId: "document-session",
            range: "B2",
            style: { border: { all: "hairline" } },
          },
        }),
      );
      assert.match(badStyle, /unknown border style: hairline/);
    } finally {
      await client.callTool({
        name: "formulon_close_workbook",
        arguments: { sessionId: "document-session" },
      });
    }
  });
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

      const commentList = textPayload(
        await client.callTool({
          name: "formulon_comment_operation",
          arguments: { sessionId: "advanced-session", operation: "list" },
        }),
      );
      assert.deepEqual(commentList.result, [{ row: 1, col: 0, author: "tester", text: "note" }]);

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

      const hyperlinkRangeAdd = textPayload(
        await client.callTool({
          name: "formulon_hyperlink_operation",
          arguments: {
            sessionId: "advanced-session",
            operation: "add",
            row: 3,
            col: 0,
            lastRow: 3,
            lastCol: 2,
            display: "Jump",
            location: "Sheet1!A1",
          },
        }),
      );
      assert.equal(hyperlinkRangeAdd.result.ok, true);

      const hyperlinks = textPayload(
        await client.callTool({
          name: "formulon_hyperlink_operation",
          arguments: { sessionId: "advanced-session", operation: "list" },
        }),
      );
      assert.equal(hyperlinks.result.length, 2);
      assert.deepEqual(
        hyperlinks.result.map((link) => [link.row, link.lastRow, link.lastCol, link.location]),
        [
          [2, 2, 0, ""],
          [3, 3, 2, "Sheet1!A1"],
        ],
      );

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
      assert.equal(cfAdd.result.status.ok, true);
      assert.equal(cfAdd.result.index, 0);

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
      assert.equal(precedents.operation, "precedents");
      assert.equal(precedents.count, 1);
      assert.deepEqual(precedents.cells, [
        { sheet: 0, sheetName: "Sheet1", row: 0, col: 0, ref: "Sheet1!A1" },
      ]);

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

test("MCP stdio supports dimensions and sheet-name references", async () => {
  await withClient(async (client) => {
    await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId: "dim-session" },
    });
    try {
      await client.callTool({
        name: "formulon_sheet_operation",
        arguments: { sessionId: "dim-session", operation: "rename", index: 0, newName: "Data" },
      });

      const width = textPayload(
        await client.callTool({
          name: "formulon_dimension_operation",
          arguments: {
            sessionId: "dim-session",
            sheet: "Data",
            axis: "column",
            operation: "size",
            first: 0,
            last: 1,
            size: 100,
          },
        }),
      );
      assert.equal(width.status.ok, true);
      assert.equal(width.sheet, 0);

      // An operation-tool must accept a sheet name, not just an index.
      const merges = textPayload(
        await client.callTool({
          name: "formulon_merge_operation",
          arguments: { sessionId: "dim-session", sheet: "Data", operation: "list" },
        }),
      );
      assert.equal(Array.isArray(merges.result), true);

      const badSheet = await client.callTool({
        name: "formulon_merge_operation",
        arguments: { sessionId: "dim-session", sheet: "Missing", operation: "list" },
      });
      assert.match(errorPayload(badSheet), /sheet not found/);
    } finally {
      await client.callTool({
        name: "formulon_close_workbook",
        arguments: { sessionId: "dim-session" },
      });
    }
  });
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
    // The server version is read from package.json at runtime and falls back to
    // "0.0.0" on any failure, so a broken path would otherwise ship silently.
    assert.equal(version.serverVersion, SERVER_VERSION);
  });
});
