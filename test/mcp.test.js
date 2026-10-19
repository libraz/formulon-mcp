import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { callTool, errorPayload, textPayload, withClient } from "./helpers/mcp.js";

const require = createRequire(import.meta.url);
const FORMULON_VERSION = require("@libraz/formulon/package.json").version;
const SERVER_VERSION = require("../package.json").version;

test("MCP stdio preserves accessor results and surfaces engine failures", async () => {
  await withClient(async (client) => {
    await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId: "status-results" },
    });
    try {
      const call = (method, args = []) =>
        client.callTool({
          name: "formulon_workbook_call",
          arguments: { sessionId: "status-results", method, args },
        });

      const count = textPayload(await call("cellCount", [0]));
      assert.equal(count.result.status.ok, true);
      assert.equal(count.result.value, 0);
      const empty = textPayload(await call("getMerges", [0]));
      assert.equal(empty.result.status.ok, true);
      assert.deepEqual(empty.result.items, []);
      const clock = textPayload(await call("pinnedNow"));
      assert.equal(clock.result.status.ok, true);
      assert.equal(clock.result.now, null);

      for (const [method, args] of [
        ["cellCount", [9]],
        ["getMerges", [9]],
        ["getValue", [0, 1_048_576, 0]],
        ["setNumber", [9, 0, 0, 1]],
      ]) {
        assert.match(errorPayload(await call(method, args)), /failed/i);
      }

      for (const operation of ["precedents", "dependents", "spillInfo"]) {
        const invalidTrace = await client.callTool({
          name: "formulon_trace",
          arguments: {
            sessionId: "status-results",
            operation,
            row: 1_048_576,
            col: 0,
          },
        });
        assert.match(errorPayload(invalidTrace), /failed|limits/i);
      }

      const names = textPayload(
        await client.callTool({
          name: "formulon_function_lookup",
          arguments: { sessionId: "status-results", operation: "names" },
        }),
      );
      assert.equal(names.result.status.ok, true);
      assert.equal(names.result.items.includes("USDOLLAR"), true);
      const localized = textPayload(
        await client.callTool({
          name: "formulon_function_lookup",
          arguments: { sessionId: "status-results", operation: "localize", name: "SUM" },
        }),
      );
      assert.equal(localized.result.status.ok, true);
      assert.equal(localized.result.value, "SUM");
      const lookup = async (args) =>
        textPayload(
          await client.callTool({
            name: "formulon_function_lookup",
            arguments: { sessionId: "status-results", ...args },
          }),
        ).result;
      assert.equal(
        (await lookup({ operation: "localize", name: "SUM", profile: "win-365-de_DE" })).value,
        "SUMME",
      );
      assert.equal(
        (
          await lookup({
            operation: "localizeFormula",
            formula: "=SUM(A1,1.5)",
            profile: "win-365-de_DE",
          })
        ).value,
        "=SUMME(A1;1,5)",
      );
      assert.equal(
        (
          await lookup({
            operation: "canonicalizeFormula",
            formula: "=SUMME(A1;1,5)",
            profile: "win-365-de_DE",
          })
        ).value,
        "=SUM(A1,1.5)",
      );
      // Without a profile the session's own profile is described.
      const facts = await lookup({ operation: "localeFacts" });
      assert.equal(facts.facts.listSeparator, ",");
      assert.equal(
        (await lookup({ operation: "localeFacts", profile: "win-365-de_DE" })).facts
          .decimalSeparator,
        ",",
      );
      assert.match(
        errorPayload(
          await client.callTool({
            name: "formulon_function_lookup",
            arguments: { sessionId: "status-results", operation: "localizeFormula" },
          }),
        ),
        /formula is required/,
      );

      const absent = errorPayload(
        await client.callTool({
          name: "formulon_comment_operation",
          arguments: { sessionId: "status-results", operation: "get", row: 0, col: 0 },
        }),
      );
      assert.match(absent, /no comment at cell/);

      const currency = textPayload(
        await client.callTool({
          name: "formulon_eval_formula",
          arguments: { formula: "=USDOLLAR(-1234.5)" },
        }),
      );
      assert.deepEqual(currency.value, { kind: "text", value: "($1,234.50)" });
    } finally {
      await client.callTool({
        name: "formulon_close_workbook",
        arguments: { sessionId: "status-results" },
      });
    }
  });
});

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
    assert.equal(names.includes("formulon_default_font"), true);
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
      assert.equal(sheetCall.result.status.ok, true);
      assert.equal(sheetCall.result.value, 1);

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
      const defaultFont = textPayload(
        await client.callTool({
          name: "formulon_default_font",
          arguments: { sessionId: "document-session", font: { name: "Meiryo", size: 10 } },
        }),
      );
      assert.equal(defaultFont.font.name, "Meiryo");
      assert.equal(defaultFont.font.size, 10);

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
      assert.match(badStyle, /Input validation error.*style\.border\.all/);
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
      assert.equal(merges.result.status.ok, true);
      assert.deepEqual(merges.result.items, [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 1 }]);

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
      assert.equal(comment.result.status.ok, true);
      assert.deepEqual(comment.result.comment, { author: "tester", text: "note" });

      const commentList = textPayload(
        await client.callTool({
          name: "formulon_comment_operation",
          arguments: { sessionId: "advanced-session", operation: "list" },
        }),
      );
      assert.equal(commentList.result.status.ok, true);
      assert.deepEqual(commentList.result.items, [
        { row: 1, col: 0, author: "tester", text: "note" },
      ]);

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
      assert.equal(hyperlinks.result.status.ok, true);
      assert.equal(hyperlinks.result.items.length, 2);
      assert.deepEqual(
        hyperlinks.result.items.map((link) => [
          link.row,
          link.lastRow,
          link.lastCol,
          link.location,
        ]),
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
      assert.equal(validations.result.status.ok, true);
      assert.equal(validations.result.items.length, 1);

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
      assert.equal(cfs.result.status.ok, true);
      assert.equal(cfs.result.items.length, 1);

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
              { type: "formula", sheet: 0, row: 0, col: 2, formula: '=ASC("ＡＢ")' },
            ],
            profile: "win-365-ja_JP",
          },
        }),
      );
      assert.equal(updated.bytes > 0, true);
      // The profile is a runtime setting, not stored in the file: the cached
      // result keeps the ja-JP fold, and a recalc needs the profile again.
      const ascAfter = async (args) => {
        const inspected = textPayload(
          await client.callTool({
            name: "formulon_inspect_workbook",
            arguments: { path: outputPath, includeCells: true, ...args },
          }),
        );
        return inspected.sheets[0].cells.find((cell) => cell.col === 2).value.value;
      };
      assert.equal(await ascAfter({}), "AB");
      assert.equal(await ascAfter({ recalc: true }), "ＡＢ");
      assert.equal(await ascAfter({ recalc: true, profile: "win-365-ja_JP" }), "AB");

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
      assert.equal(merges.result.status.ok, true);
      assert.equal(Array.isArray(merges.result.items), true);

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

// ---------------------------------------------------------------------------
// Excel profile resolution through the tools (tool x requested x environment x
// source; 21 rows covering every value pair, grouped by environment value
// because the environment belongs to the server process).
// ---------------------------------------------------------------------------

const PROFILE_ENV = "FORMULON_MCP_PROFILE";
const ASC_FORMULA = '=ASC("ＡＢ")';
const REQUESTED_PROFILES = {
  omitted: undefined,
  ja: "win-365-ja_JP",
  macde: "mac-365-de_DE",
  unknown: "win-365-xx_XX",
};
const ENV_PROFILES = {
  unset: undefined,
  valid: "mac-365-ja_JP",
  unknown: "win-365-xx_XX",
  empty: "",
  whitespace: "   ",
};
const NEW_SHEET_NAME = {
  "win-365-en_US": "Sheet1",
  "win-365-ja_JP": "Sheet1",
  "mac-365-ja_JP": "Sheet1",
  "mac-365-de_DE": "Tabelle1",
};

const MCP_PROFILE_CASES = [
  ["open", "macde", "empty", "loaded"],
  ["open", "omitted", "unknown", "new"],
  ["update", "ja", "unset", "loaded"],
  ["update", "unknown", "valid", "new"],
  ["inspect", "omitted", "whitespace", "loaded"],
  ["update", "omitted", "unknown", "loaded"],
  ["inspect", "unknown", "unknown", "loaded"],
  ["update", "macde", "whitespace", "new"],
  ["inspect", "macde", "valid", "loaded"],
  ["open", "omitted", "valid", "loaded"],
  ["inspect", "omitted", "empty", "loaded"],
  ["open", "ja", "whitespace", "loaded"],
  ["open", "unknown", "unset", "new"],
  ["inspect", "omitted", "unset", "loaded"],
  ["open", "ja", "empty", "new"],
  ["update", "unknown", "empty", "new"],
  ["open", "unknown", "whitespace", "new"],
  ["open", "ja", "valid", "new"],
  ["inspect", "macde", "unset", "loaded"],
  ["inspect", "ja", "unknown", "loaded"],
  ["open", "macde", "unknown", "loaded"],
  // Beyond the pairwise rows: the environment-only paths of each tool.
  ["inspect", "omitted", "valid", "loaded"],
  ["update", "omitted", "valid", "new"],
  ["update", "omitted", "valid", "loaded"],
  ["open", "omitted", "valid", "new"],
];

function resolveMcpProfile(requestedKey, envKey) {
  const requested = REQUESTED_PROFILES[requestedKey];
  const fromEnvironment = ENV_PROFILES[envKey]?.trim() || undefined;
  const chosen = requested ?? fromEnvironment;
  return {
    chosen,
    fromEnv: requested === undefined && fromEnvironment !== undefined,
    id: chosen ?? "win-365-en_US",
    known: chosen === undefined || !chosen.endsWith("xx_XX"),
  };
}

function assertProfileError(message, expected, source) {
  const head = source === "new" ? "unknown excel profile: " : "set excel profile ";
  const suffix = expected.fromEnv ? ` (from ${PROFILE_ENV})` : "";
  assert.ok(message.startsWith(`${head}${expected.chosen}${suffix}`), message);
  assert.equal(message.includes(PROFILE_ENV), expected.fromEnv, message);
}

async function runMcpProfileCase(client, file, [tool, requestedKey, envKey, source], index) {
  const expected = resolveMcpProfile(requestedKey, envKey);
  const requested = REQUESTED_PROFILES[requestedKey];
  const profile = requested === undefined ? {} : { profile: requested };
  const folds = expected.id.endsWith("ja_JP");
  const out = path.join(path.dirname(file), `out-${index}.xlsx`);
  const sessionId = `mcp-profile-${index}`;

  if (tool === "open") {
    const opened = await callTool(client, "formulon_open_workbook", {
      sessionId,
      ...(source === "loaded" ? { path: file } : {}),
      ...profile,
    });
    if (!expected.known) {
      assertProfileError(errorPayload(opened), expected, source);
      const { sessions } = textPayload(await callTool(client, "formulon_list_sessions"));
      assert.equal(
        sessions.some((session) => session.id === sessionId),
        false,
      );
      // The id is free again.
      textPayload(
        await callTool(client, "formulon_open_workbook", { sessionId, profile: "win-365-en_US" }),
      );
      textPayload(await callTool(client, "formulon_close_workbook", { sessionId }));
      return;
    }
    textPayload(opened);
    try {
      const call = async (method, args) =>
        textPayload(await callTool(client, "formulon_workbook_call", { sessionId, method, args }))
          .result;
      assert.equal((await call("excelProfileId", [])).value, expected.id);
      if (source === "new") {
        assert.equal((await call("sheetName", [0])).value, NEW_SHEET_NAME[expected.id] ?? "Sheet1");
        await call("setFormula", [0, 0, 0, ASC_FORMULA]);
        await call("recalc", []);
      }
      assert.equal((await call("getValue", [0, 0, 0])).value.value, folds ? "AB" : "ＡＢ");
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
    return;
  }

  if (tool === "inspect") {
    const inspected = await callTool(client, "formulon_inspect_workbook", {
      path: file,
      includeCells: true,
      recalc: true,
      ...profile,
    });
    if (!expected.known) {
      assertProfileError(errorPayload(inspected), expected, "loaded");
      return;
    }
    const summary = textPayload(inspected);
    assert.equal(summary.sheets[0].cells[0].value.value, folds ? "AB" : "ＡＢ");
    return;
  }

  const updated = await callTool(client, "formulon_update_workbook", {
    ...(source === "loaded" ? { inputPath: file } : {}),
    outputPath: out,
    mutations: [{ type: "formula", sheet: 0, row: 0, col: 1, formula: ASC_FORMULA }],
    ...profile,
  });
  if (!expected.known) {
    assertProfileError(errorPayload(updated), expected, source);
    assert.equal(existsSync(out), false, "a rejected profile writes no output");
    return;
  }
  const result = textPayload(updated);
  if (source === "new") {
    assert.equal(result.summary.sheets[0].name, NEW_SHEET_NAME[expected.id] ?? "Sheet1");
  }
  // Read the cached result back under an explicit profile, which wins over
  // any environment value without recalculating.
  const reread = textPayload(
    await callTool(client, "formulon_inspect_workbook", {
      path: out,
      includeCells: true,
      profile: "win-365-en_US",
    }),
  );
  const cell = reread.sheets[0].cells.find((item) => item.col === 1);
  assert.equal(cell.value.value, folds ? "AB" : "ＡＢ");
}

test("MCP tools resolve the Excel profile from the argument, the environment, then the default", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-profile-tools-"));
  const file = path.join(dir, "asc.xlsx");
  try {
    await withClient(async (client) => {
      textPayload(
        await callTool(client, "formulon_update_workbook", {
          outputPath: file,
          mutations: [{ type: "formula", sheet: 0, row: 0, col: 0, formula: ASC_FORMULA }],
        }),
      );
    });
    for (const [envKey, envValue] of Object.entries(ENV_PROFILES)) {
      await withClient(
        async (client) => {
          for (const [index, row] of MCP_PROFILE_CASES.entries()) {
            if (row[2] !== envKey) {
              continue;
            }
            try {
              await runMcpProfileCase(client, file, row, index);
            } catch (error) {
              error.message = `[${row.join(" / ")}] ${error.message}`;
              throw error;
            }
          }
        },
        envValue === undefined ? {} : { env: { [PROFILE_ENV]: envValue } },
      );
    }
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

const MALFORMED_PROFILES = [
  "ja_JP",
  "win-365-ja-JP",
  "WIN-365-ja_JP",
  "win-365-JA_JP",
  "win-365-ja_jp",
  "win-366-ja_JP",
  "linux-365-en_US",
  "win-365-ja_JPX",
  "win-365-j_JP",
  "win-365-jpn_JP",
  " win-365-ja_JP",
  "win-365-ja_JP\n",
  "",
];

test("MCP schema rejects malformed profile ids before any engine work", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-profile-schema-"));
  const file = path.join(dir, "schema.xlsx");
  try {
    await withClient(async (client) => {
      textPayload(
        await callTool(client, "formulon_update_workbook", {
          outputPath: file,
          mutations: [{ type: "number", sheet: 0, row: 0, col: 0, value: 1 }],
        }),
      );
      textPayload(await callTool(client, "formulon_open_workbook", { sessionId: "schema" }));
      const targets = [
        ["formulon_open_workbook", { sessionId: "schema-bad" }],
        ["formulon_inspect_workbook", { path: file }],
        [
          "formulon_update_workbook",
          {
            outputPath: path.join(dir, "never.xlsx"),
            mutations: [{ type: "number", sheet: 0, row: 0, col: 0, value: 1 }],
          },
        ],
        ["formulon_function_lookup", { sessionId: "schema", operation: "localeFacts" }],
      ];
      for (const [tool, base] of targets) {
        for (const profile of MALFORMED_PROFILES) {
          const message = errorPayload(await callTool(client, tool, { ...base, profile }));
          assert.match(
            message,
            /Input validation error/,
            `${tool} accepted ${JSON.stringify(profile)}`,
          );
          assert.match(message, /profile/);
        }
      }
      const { sessions } = textPayload(await callTool(client, "formulon_list_sessions"));
      assert.deepEqual(
        sessions.map((session) => session.id),
        ["schema"],
      );
      assert.equal(existsSync(path.join(dir, "never.xlsx")), false);
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

// ---------------------------------------------------------------------------
// formulon_function_lookup
// ---------------------------------------------------------------------------

const LOCAL_SUM = {
  "win-365-en_US": "SUM",
  "win-365-de_DE": "SUMME",
  "win-365-fr_FR": "SOMME",
};

/**
 * Pairwise matrix over (operation x name x formula x profile x session profile);
 * 23 rows covering every value pair.
 */
const LOOKUP_CASES = [
  ["canonicalize", "present", "absent", "valid", "en_US"],
  ["localize", "present", "present", "unknown", "de_DE"],
  ["names", "empty", "present", "valid", "de_DE"],
  ["localeFacts", "absent", "absent", "unknown", "en_US"],
  ["canonicalize", "empty", "absent", "omitted", "de_DE"],
  ["canonicalizeFormula", "absent", "present", "omitted", "de_DE"],
  ["localizeFormula", "empty", "present", "unknown", "en_US"],
  ["localize", "absent", "absent", "valid", "en_US"],
  ["metadata", "empty", "absent", "unknown", "de_DE"],
  ["localeFacts", "present", "present", "valid", "de_DE"],
  ["localizeFormula", "present", "absent", "omitted", "de_DE"],
  ["canonicalize", "absent", "present", "unknown", "de_DE"],
  ["metadata", "absent", "present", "omitted", "en_US"],
  ["localize", "empty", "present", "omitted", "de_DE"],
  ["localeFacts", "empty", "absent", "omitted", "de_DE"],
  ["canonicalizeFormula", "empty", "absent", "unknown", "de_DE"],
  ["metadata", "present", "absent", "valid", "en_US"],
  ["canonicalizeFormula", "present", "present", "valid", "en_US"],
  ["names", "present", "present", "unknown", "en_US"],
  ["names", "absent", "present", "omitted", "en_US"],
  ["localizeFormula", "empty", "present", "valid", "en_US"],
  ["names", "empty", "absent", "valid", "en_US"],
  ["localizeFormula", "absent", "absent", "valid", "en_US"],
  // Beyond the pairwise rows: both directions with a present name and formula.
  ["localize", "present", "present", "omitted", "de_DE"],
  ["localize", "present", "present", "valid", "de_DE"],
  ["canonicalize", "present", "present", "omitted", "de_DE"],
  ["canonicalize", "present", "present", "unknown", "en_US"],
  ["localizeFormula", "present", "present", "omitted", "de_DE"],
  ["localizeFormula", "present", "present", "unknown", "de_DE"],
  ["canonicalizeFormula", "present", "present", "omitted", "de_DE"],
  ["localeFacts", "present", "present", "omitted", "de_DE"],
  ["metadata", "present", "present", "omitted", "de_DE"],
];

/** `=SUM(A1,1.5)` as the given profile spells it. */
function localFormula(profile) {
  return profile === "win-365-en_US" ? "=SUM(A1,1.5)" : `=${LOCAL_SUM[profile] ?? "SUM"}(A1;1,5)`;
}

const UNKNOWN_PROFILE = "win-365-xx_XX";
const VALID_PROFILE = "win-365-fr_FR";

async function runLookupCase(client, sessionId, [operation, nameKey, formulaKey, profileKey]) {
  const args = { sessionId, operation };
  const sessionProfile = textPayload(
    await callTool(client, "formulon_workbook_call", {
      sessionId,
      method: "excelProfileId",
      args: [],
    }),
  ).result.value;
  const effective =
    profileKey === "valid"
      ? VALID_PROFILE
      : profileKey === "unknown"
        ? UNKNOWN_PROFILE
        : sessionProfile;
  if (profileKey !== "omitted") {
    args.profile = effective;
  }
  if (nameKey === "present") {
    // The function name in the direction the operation reads it.
    args.name = operation === "canonicalize" ? (LOCAL_SUM[effective] ?? "SUM") : "SUM";
  } else if (nameKey === "empty") {
    args.name = "";
  }
  if (formulaKey === "present") {
    args.formula = operation === "canonicalizeFormula" ? localFormula(effective) : "=SUM(A1,1.5)";
  }
  const result = await callTool(client, "formulon_function_lookup", args);

  const needsName = ["metadata", "localize", "canonicalize"].includes(operation);
  const needsFormula = operation === "localizeFormula" || operation === "canonicalizeFormula";
  const usesProfile = !["names", "metadata"].includes(operation);
  if (needsFormula && formulaKey === "absent") {
    assert.match(errorPayload(result), new RegExp(`formula is required for ${operation}`));
    return;
  }
  if (needsName && nameKey !== "present") {
    assert.match(errorPayload(result), /name is required/);
    return;
  }
  if (usesProfile && profileKey === "unknown") {
    assert.match(errorPayload(result), /unknown profile/);
    assert.match(errorPayload(result), new RegExp(UNKNOWN_PROFILE));
    return;
  }

  const { result: payload } = textPayload(result);
  switch (operation) {
    case "names":
      assert.equal(payload.items.includes("SUM"), true);
      break;
    case "metadata":
      assert.equal(payload.ok, true);
      assert.equal(payload.name, "SUM");
      assert.equal(payload.minArity, 1);
      break;
    case "localize":
      assert.equal(payload.value, LOCAL_SUM[effective]);
      break;
    case "canonicalize":
      assert.equal(payload.value, "SUM");
      break;
    case "localizeFormula":
      assert.equal(
        payload.value,
        effective === "win-365-en_US" ? "=SUM(A1,1.5)" : `=${LOCAL_SUM[effective]}(A1;1,5)`,
      );
      break;
    case "canonicalizeFormula":
      assert.equal(payload.value, "=SUM(A1,1.5)");
      break;
    case "localeFacts":
      assert.equal(payload.facts.decimalSeparator, effective === "win-365-en_US" ? "." : ",");
      assert.equal(payload.facts.listSeparator, effective === "win-365-en_US" ? "," : ";");
      break;
    default:
      assert.fail(`unhandled operation ${operation}`);
  }
}

test("function lookup follows the operation, argument and profile matrix", async () => {
  await withClient(async (client) => {
    for (const sessionKey of ["en_US", "de_DE"]) {
      const sessionId = `lookup-${sessionKey}`;
      textPayload(
        await callTool(client, "formulon_open_workbook", {
          sessionId,
          profile: `win-365-${sessionKey}`,
        }),
      );
      try {
        for (const row of LOOKUP_CASES.filter((item) => item[4] === sessionKey)) {
          try {
            await runLookupCase(client, sessionId, row);
          } catch (error) {
            error.message = `[${row.join(" / ")}] ${error.message}`;
            throw error;
          }
        }
      } finally {
        await callTool(client, "formulon_close_workbook", { sessionId });
      }
    }
  });
});

test("function lookup defaults to the session profile and follows a profile switch", async () => {
  await withClient(async (client) => {
    const sessionId = "lookup-switch";
    textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
    try {
      const lookup = async (args) =>
        textPayload(await callTool(client, "formulon_function_lookup", { sessionId, ...args }))
          .result;
      assert.equal((await lookup({ operation: "localize", name: "SUM" })).value, "SUM");

      textPayload(
        await callTool(client, "formulon_workbook_call", {
          sessionId,
          method: "setExcelProfileId",
          args: ["win-365-de_DE"],
        }),
      );
      assert.equal((await lookup({ operation: "localize", name: "SUM" })).value, "SUMME");
      assert.equal((await lookup({ operation: "canonicalize", name: "SUMME" })).value, "SUM");
      assert.equal(
        (await lookup({ operation: "localizeFormula", formula: "=SUM(1.5,2)" })).value,
        "=SUMME(1,5;2)",
      );
      assert.equal((await lookup({ operation: "localeFacts" })).facts.decimalSeparator, ",");
      // An explicit profile overrides the session's.
      assert.equal(
        (await lookup({ operation: "localize", name: "SUM", profile: "win-365-en_US" })).value,
        "SUM",
      );
      assert.equal(
        (await lookup({ operation: "localize", name: "SUM", profile: "win-365-ru_RU" })).value,
        "СУММ",
      );

      textPayload(
        await callTool(client, "formulon_workbook_call", {
          sessionId,
          method: "setExcelProfileId",
          args: ["win-365-ja_JP"],
        }),
      );
      assert.equal((await lookup({ operation: "localize", name: "SUM" })).value, "SUM");
      assert.equal((await lookup({ operation: "localeFacts" })).facts.decimalSeparator, ".");
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});

const ROUND_TRIP_PROFILES = [
  "win-365-de_DE",
  "win-365-fr_FR",
  "win-365-ja_JP",
  "win-365-en_US",
  "win-365-ru_RU",
  "mac-365-de_DE",
];

const ROUND_TRIP_FORMULAS = [
  "=SUM(A1:B2)",
  '=IF(A1>0,"a,b","c;d")',
  "=ROUND(3.14159,2)",
  "=AND(TRUE,FALSE)",
  "=IFERROR(1/0,#N/A)",
  "=SUM({1,2;3,4})",
  "=INDEX({1,2;3,4},2,1)",
  '={"a","b";"c","d"}',
  "=SUM(IF(A1>0,MAX(1.5,2.5),MIN(0.5,SUM(1,2))))",
  '="1,5"&"x"&"3.5"',
  "=A1+1.25*2-0.5",
  "=1000.25+0.5",
  '=TEXT(A1,"0.00")',
  "=VLOOKUP(1,A1:B2,2,FALSE)",
  "=#REF!",
  "=IF(ISERROR(A1),#DIV/0!,#VALUE!)",
  "=Sheet1!A1+'My Sheet'!B2",
  "=SUM(Table1[Col1],1.5)",
  "=LET(x,1.5,x*2)",
  "=IF(TRUE(),1,0)",
];

test("localizing then canonicalizing a formula returns the original", async () => {
  await withClient(async (client) => {
    const sessionId = "lookup-round-trip";
    textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
    try {
      const lookup = async (args) =>
        textPayload(await callTool(client, "formulon_function_lookup", { sessionId, ...args }))
          .result.value;
      for (const profile of ROUND_TRIP_PROFILES) {
        for (const formula of ROUND_TRIP_FORMULAS) {
          const local = await lookup({ operation: "localizeFormula", formula, profile });
          const back = await lookup({ operation: "canonicalizeFormula", formula: local, profile });
          assert.equal(back, formula, `${profile}: ${formula} -> ${local}`);
        }
      }
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});

/** Known spellings in each locale, checked against what Excel itself writes. */
const LOCALIZED_FORMULAS = [
  ["win-365-de_DE", "=IF(A1>0,SUM(1.5,2),0)", "=WENN(A1>0;SUMME(1,5;2);0)"],
  ["win-365-de_DE", "=AND(TRUE,FALSE)", "=UND(WAHR;FALSCH)"],
  ["win-365-de_DE", "=VLOOKUP(1,A1:B2,2,FALSE)", "=SVERWEIS(1;A1:B2;2;FALSCH)"],
  ["win-365-de_DE", "=IFERROR(1/0,#N/A)", "=WENNFEHLER(1/0;#NV)"],
  ["win-365-de_DE", "=#REF!", "=#BEZUG!"],
  ["win-365-de_DE", "=SUM({1,2;3,4})", "=SUMME({1.2;3.4})"],
  ["win-365-fr_FR", "=IF(A1>0,SUM(1.5,2),0)", "=SI(A1>0;SOMME(1,5;2);0)"],
  ["win-365-fr_FR", "=AND(TRUE,FALSE)", "=ET(VRAI;FAUX)"],
  ["win-365-fr_FR", "=VLOOKUP(1,A1:B2,2,FALSE)", "=RECHERCHEV(1;A1:B2;2;FAUX)"],
  ["win-365-ru_RU", "=IF(A1>0,SUM(1.5,2),0)", "=ЕСЛИ(A1>0;СУММ(1,5;2);0)"],
  ["win-365-ru_RU", "=AND(TRUE,FALSE)", "=И(ИСТИНА;ЛОЖЬ)"],
  ["win-365-ja_JP", "=IF(A1>0,SUM(1.5,2),0)", "=IF(A1>0,SUM(1.5,2),0)"],
  ["win-365-en_US", "=IF(A1>0,SUM(1.5,2),0)", "=IF(A1>0,SUM(1.5,2),0)"],
  // Text in string literals is never touched.
  ["win-365-de_DE", '="SUM(1.5,2)"&A1', '="SUM(1.5,2)"&A1'],
];

test("localizes formulas to the spelling each locale uses", async () => {
  await withClient(async (client) => {
    const sessionId = "lookup-spellings";
    textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
    try {
      for (const [profile, canonical, local] of LOCALIZED_FORMULAS) {
        const localized = textPayload(
          await callTool(client, "formulon_function_lookup", {
            sessionId,
            operation: "localizeFormula",
            formula: canonical,
            profile,
          }),
        ).result.value;
        assert.equal(localized, local, `${profile} ${canonical}`);
      }
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});

test("function lookup reports unknown functions and sessions", async () => {
  await withClient(async (client) => {
    const sessionId = "lookup-unknown";
    textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
    try {
      const call = (args) => callTool(client, "formulon_function_lookup", { sessionId, ...args });
      // Metadata for an unknown function is a negative answer, not a failure.
      assert.equal(
        textPayload(await call({ operation: "metadata", name: "NOPE" })).result.ok,
        false,
      );
      assert.match(
        errorPayload(await call({ operation: "localize", name: "NOPE", profile: "win-365-de_DE" })),
        /unknown function/,
      );
      assert.match(
        errorPayload(
          await call({ operation: "canonicalize", name: "NOPE", profile: "win-365-de_DE" }),
        ),
        /unknown function/,
      );
      assert.match(
        errorPayload(
          await callTool(client, "formulon_function_lookup", {
            sessionId: "lookup-missing",
            operation: "names",
          }),
        ),
        /session not found/,
      );
      assert.match(
        errorPayload(
          await callTool(client, "formulon_function_lookup", { sessionId, operation: "bogus" }),
        ),
        /Input validation error/,
      );
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});

test("MCP workbook_call serializes typed-array results as plain arrays and stays clean", async () => {
  await withClient(async (client) => {
    const sessionId = "call-typed-arrays";
    textPayload(await callTool(client, "formulon_open_workbook", { sessionId }));
    try {
      const call = async (method, args) =>
        textPayload(await callTool(client, "formulon_workbook_call", { sessionId, method, args }));
      const text = (value) => ({ kind: 3, text: value });

      const classes = await call("collationClasses", [
        [text("a"), text("b"), text("A"), { kind: 1, number: 1 }, text("a")],
        0,
        "mac-365-en_US",
      ]);
      assert.equal(Array.isArray(classes.result.classIds), true);
      // Class ids follow collation order, so the number sorts before the text.
      assert.deepEqual(classes.result.classIds, [1, 2, 1, 0, 1]);
      assert.equal(classes.result.classCount, 3);
      assert.equal(classes.session.dirty, false);

      const empty = await call("collationClasses", [[], 0, "mac-365-en_US"]);
      assert.deepEqual(empty.result.classIds, []);
      assert.equal(empty.result.classCount, 0);

      // Case-insensitive equality under a mac profile.
      const equal = await call("compareText", ["a", "A", 0, "mac-365-en_US"]);
      assert.equal(equal.result.order, 0);
      assert.equal(equal.session.dirty, false);
      assert.equal((await call("compareText", ["a", "b", 0, "mac-365-en_US"])).result.order, -1);
      assert.equal((await call("compareText", ["b", "a", 0, "mac-365-en_US"])).result.order, 1);

      for (const method of [
        "snapshotDrawing",
        "restoreDrawing",
        "snapshotRangeStyle",
        "restoreRangeStyle",
      ]) {
        const rejected = await callTool(client, "formulon_workbook_call", {
          sessionId,
          method,
          args: [],
        });
        assert.match(errorPayload(rejected), /not allowlisted/);
      }
      const { sessions } = textPayload(await callTool(client, "formulon_list_sessions"));
      assert.equal(sessions.find((session) => session.id === sessionId).dirty, false);
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});
