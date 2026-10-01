import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import { errorPayload, textPayload, withClient } from "./helpers/mcp.js";

test("AI tools publish validation and conditional-format shapes and reject invalid input before writing", async () => {
  await withClient(async (client) => {
    const tools = (await client.listTools()).tools;
    const conditional = tools.find((tool) => tool.name === "formulon_conditional_format_operation");
    const validation = tools.find((tool) => tool.name === "formulon_validation_operation");
    const rule = conditional.inputSchema.properties.rule;
    assert.equal(rule.properties.dataBar.properties.direction.maximum, 2);
    assert.equal(rule.properties.iconSet.properties.floor.properties.type.maximum, 7);
    assert.equal(
      validation.inputSchema.properties.validation.properties.allowBlank.type,
      "boolean",
    );
    const sessionId = "schema-preflight";
    textPayload(
      await client.callTool({ name: "formulon_open_workbook", arguments: { sessionId } }),
    );
    try {
      for (const [name, extra] of [
        [
          "formulon_conditional_format_operation",
          { rule: { type: 99, sqref: [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }] } },
        ],
        [
          "formulon_conditional_format_operation",
          {
            rule: {
              type: 3,
              sqref: [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }],
              dataBar: { min: { type: 3 }, max: { type: 4 }, fill: { r: 256, g: 0, b: 0 } },
            },
          },
        ],
        [
          "formulon_validation_operation",
          {
            validation: { type: 1, ranges: [{ firstRow: 2, firstCol: 0, lastRow: 1, lastCol: 1 }] },
          },
        ],
        ...[
          { type: 0, id: "not-a-guid", formula1: "TRUE" },
          {
            type: 3,
            dataBar: {
              min: { type: 3 },
              max: { type: 4 },
              fill: { r: 0, g: 100, b: 200 },
              minLengthPct: 90,
              maxLengthPct: 10,
            },
          },
          {
            type: 3,
            dataBar: {
              min: { type: 3 },
              max: { type: 4 },
              fill: { r: 0, g: 100, b: 200 },
              minLengthPct: 10.5,
            },
          },
          { type: 4, iconSet: { name: 0, thresholds: Array(4).fill({ type: 1, value: "50" }) } },
          { type: 4, iconSet: { name: 8, thresholds: Array(2).fill({ type: 1, value: "50" }) } },
        ].map((rule) => [
          "formulon_conditional_format_operation",
          { rule: { sqref: [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }], ...rule } },
        ]),
      ]) {
        assert.match(
          errorPayload(
            await client.callTool({
              name,
              arguments: { sessionId, operation: "add", ...extra },
            }),
          ),
          /validation|invalid/i,
        );
        const state = textPayload(
          await client.callTool({
            name: "formulon_inspect_session",
            arguments: { sessionId },
          }),
        );
        assert.equal(state.session.dirty, false);
      }
    } finally {
      textPayload(
        await client.callTool({ name: "formulon_close_workbook", arguments: { sessionId } }),
      );
    }
  });
});

test("path cell reads report invalid addresses as MCP errors", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-path-read-"));
  try {
    await withClient(async (client) => {
      const outputPath = path.join(dir, "book.xlsx");
      textPayload(
        await client.callTool({
          name: "formulon_update_workbook",
          arguments: {
            outputPath,
            mutations: [{ type: "number", sheet: 0, row: 0, col: 0, value: 7 }],
          },
        }),
      );
      const valid = textPayload(
        await client.callTool({
          name: "formulon_get_cell",
          arguments: { path: outputPath, sheet: 0, row: 0, col: 0 },
        }),
      );
      assert.deepEqual(valid.value, { kind: "number", value: 7 });
      for (const [address, expectedError] of [
        [{ sheet: 9, row: 0, col: 0 }, /sheet index out of bounds/],
        [{ sheet: 0, row: 1_048_576, col: 0 }, /exceeds Excel sheet limits/],
      ]) {
        assert.match(
          errorPayload(
            await client.callTool({
              name: "formulon_get_cell",
              arguments: { path: outputPath, ...address },
            }),
          ),
          expectedError,
        );
      }
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("AI tools preserve icon floors and data-bar directions through xlsx and xlsb", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-cf-workflow-"));
  try {
    await withClient(async (client) => {
      const call = async (name, args) =>
        textPayload(await client.callTool({ name, arguments: args }));
      const sessionId = "cf-workflow";
      await call("formulon_open_workbook", { sessionId });
      await call("formulon_set_range", {
        sessionId,
        start: "A1",
        values: [
          [5, 5],
          [60, 60],
          [90, 90],
        ],
      });
      await call("formulon_conditional_format_operation", {
        sessionId,
        operation: "add",
        rule: {
          type: 3,
          id: "{12345678-1234-1234-1234-123456789ABC}",
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
          dataBar: {
            min: { type: 0, value: "0" },
            max: { type: 0, value: "100" },
            fill: { r: 0, g: 100, b: 200, a: 255 },
            direction: 2,
          },
        },
      });
      await call("formulon_conditional_format_operation", {
        sessionId,
        operation: "add",
        rule: {
          type: 4,
          sqref: [{ firstRow: 0, firstCol: 1, lastRow: 2, lastCol: 1 }],
          iconSet: {
            name: 0,
            floor: { type: 0, value: "10", gte: true },
            thresholds: [
              { type: 0, value: "50" },
              { type: 0, value: "80" },
            ],
          },
        },
      });
      for (const format of ["xlsx", "xlsb"]) {
        const outputPath = path.join(dir, `rules.${format}`);
        await call("formulon_save_session", { sessionId, outputPath });
        const reloaded = `cf-${format}`;
        await call("formulon_open_workbook", { sessionId: reloaded, path: outputPath });
        try {
          const rules = await call("formulon_conditional_format_operation", {
            sessionId: reloaded,
            operation: "list",
          });
          assert.equal(rules.result.status.ok, true);
          assert.equal(rules.result.items.find((rule) => rule.type === 3).dataBar.direction, 2);
          assert.deepEqual(rules.result.items.find((rule) => rule.type === 4).iconSet.floor, {
            type: 0,
            value: "10",
            gte: true,
          });
          const evaluated = await call("formulon_conditional_format_operation", {
            sessionId: reloaded,
            operation: "evaluate",
            firstRow: 0,
            firstCol: 0,
            lastRow: 2,
            lastCol: 1,
          });
          assert.equal(evaluated.result.status.ok, true);
          const bar = evaluated.result.cells.find((cell) => cell.row === 1 && cell.col === 0);
          assert.equal(bar.matches.find((match) => match.kind === 2).barDirection, 2);
          assert.equal(
            evaluated.result.cells.some((cell) => cell.row === 0 && cell.col === 1),
            false,
          );
          const icon = evaluated.result.cells.find((cell) => cell.row === 2 && cell.col === 1);
          assert.equal(icon.matches.find((match) => match.kind === 3).iconIndex, 2);
        } finally {
          await call("formulon_close_workbook", { sessionId: reloaded });
        }
      }
      await call("formulon_close_workbook", { sessionId });
    });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
