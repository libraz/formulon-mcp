import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  callWorkbookMethod,
  closeSession,
  createSessionPivot,
  getSessionCellByA1,
  listSessions,
  openSession,
  saveSession,
  setSessionRange,
  tableOperation,
} from "../dist/sessions.js";
import { textPayload, withClient } from "./helpers/mcp.js";

const sales = [
  ["Region", "Product", "Amount"],
  ["East", "Pen", { f: "=20" }],
  ["East", "Book", 10],
  ["West", "Pen", 5],
];

function assertPivotTotals(layout) {
  const cells = layout.cells;
  const rowLabels = cells
    .filter((cell) => cell.kind === 1 && cell.value.kind === "text")
    .map((cell) => cell.value.value);
  assert.deepEqual([...new Set(rowLabels)], ["East", "West"]);

  const rowFor = (label) => {
    const matches = cells.filter(
      (cell) => cell.kind === 1 && cell.value.kind === "text" && cell.value.value === label,
    );
    assert.equal(matches.length, 1, `expected one row label for ${label}`);
    return matches[0].row;
  };
  const eastRow = rowFor("East");
  const westRow = rowFor("West");
  const eastNumbers = cells
    .filter((cell) => cell.row === eastRow && cell.value.kind === "number")
    .map((cell) => cell.value.value);
  const westNumbers = cells
    .filter((cell) => cell.row === westRow && cell.value.kind === "number")
    .map((cell) => cell.value.value);
  assert.equal(eastNumbers.includes(30), true);
  assert.equal(westNumbers.includes(5), true);

  const grandTotalNumbers = cells
    .filter((cell) => cell.kind === 6 && cell.value.kind === "number")
    .map((cell) => cell.value.value);
  assert.equal(grandTotalNumbers.includes(35), true);

  const formatted = cells
    .filter((cell) => cell.value.kind === "number" && [5, 30, 35].includes(cell.value.value))
    .map((cell) => cell.numberFormat);
  assert.equal(formatted.length >= 3, true);
  assert.equal(new Set(formatted).size, 1);
  assert.notEqual(formatted[0], "");
  return formatted[0];
}

test("authors a native worksheet Table with derived headers and global indexes", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-table-"));
  const file = path.join(dir, "table.xlsx");
  await openSession(undefined, "table-authoring");
  try {
    setSessionRange("table-authoring", "Sheet1!A1", sales, 0, true);
    const created = tableOperation("table-authoring", "create", {
      range: "Sheet1!A1:C4",
      name: "SalesTable",
      style: "TableStyleMedium2",
    });
    assert.equal(created.table.index, 0);
    assert.equal(created.table.ref, "A1:C4");
    assert.equal(created.table.sheetName, "Sheet1");

    const listed = tableOperation("table-authoring", "list");
    assert.equal(listed.tables.length, 1);
    assert.equal(listed.tables[0].index, 0);
    assert.equal(listed.tables[0].name, "SalesTable");

    setSessionRange("table-authoring", "Sheet1!A5", [["South", "Pen", 7]], 0, false);
    const updated = tableOperation("table-authoring", "update", {
      index: 0,
      ref: "A1:C5",
      totalsRow: true,
    });
    assert.equal(updated.table.ref, "A1:C5");
    const saved = await saveSession("table-authoring", file);
    assert.equal(saved.format, "xlsx");
    closeSession("table-authoring");

    await openSession(file, "table-reload");
    try {
      assert.equal(callWorkbookMethod("table-reload", "tableCount", []).result.value, 1);
      assert.equal(callWorkbookMethod("table-reload", "tableAt", [0]).result.ref, "A1:C5");
      assert.equal(tableOperation("table-reload", "remove", { index: 0 }).removed.index, 0);
      assert.equal(tableOperation("table-reload", "list").tables.length, 0);
    } finally {
      closeSession("table-reload");
    }
  } finally {
    if (listSessions().some((session) => session.id === "table-authoring")) {
      closeSession("table-authoring");
    }
    await rm(dir, { recursive: true, force: true });
  }
});

test("creates and round-trips a worksheet-backed PivotTable", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-table-pivot-"));
  const file = path.join(dir, "pivot.xlsx");
  let numberFormat;
  await openSession(undefined, "pivot-authoring");
  try {
    setSessionRange("pivot-authoring", "Sheet1!A1", sales, 0, true);
    assert.deepEqual(getSessionCellByA1("pivot-authoring", "Sheet1!C2").value, {
      kind: "number",
      value: 20,
    });
    const created = createSessionPivot("pivot-authoring", {
      sourceRange: "Sheet1!A1:C4",
      target: "Sheet1!E1",
      name: "SalesPivot",
      rows: ["Region"],
      columns: ["Product"],
      values: [{ field: "Amount", aggregation: "sum", numberFormat: "0.00" }],
      layout: "tabular",
      grandTotals: { rows: true, columns: true },
    });
    assert.equal(created.pivotIndex, 0);
    assert.equal(created.source.headers.join(","), "Region,Product,Amount");
    numberFormat = assertPivotTotals(created.layout);

    const source = callWorkbookMethod("pivot-authoring", "pivotCacheGetWorksheetSource", [
      created.cacheId,
    ]).result;
    assert.equal(source.present, true);
    assert.equal(source.sheet, "Sheet1");
    assert.equal(source.ref, "A1:C4");
    const saved = await saveSession("pivot-authoring", file);
    assert.equal(saved.format, "xlsx");
    assert.equal(saved.losses, undefined);
  } finally {
    closeSession("pivot-authoring");
  }

  await openSession(file, "pivot-reload");
  try {
    assert.equal(callWorkbookMethod("pivot-reload", "pivotCacheCount", []).result.value, 1);
    assert.equal(callWorkbookMethod("pivot-reload", "pivotCount", [0]).result.value, 1);
    const layout = callWorkbookMethod("pivot-reload", "pivotLayout", [0, 0]).result;
    assert.equal(layout.status.ok, true);
    assert.equal(assertPivotTotals(layout), numberFormat);
  } finally {
    closeSession("pivot-reload");
    await rm(dir, { recursive: true, force: true });
  }
});

test("rejects unknown and unused PivotTable fields before creating cache state", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-pivot-invalid-"));
  const file = path.join(dir, "fixture.xlsx");
  await openSession(undefined, "pivot-invalid");
  try {
    setSessionRange("pivot-invalid", "Sheet1!A1", sales, 0, true);
    await saveSession("pivot-invalid", file);
    assert.equal(listSessions().find((session) => session.id === "pivot-invalid")?.dirty, false);
    assert.throws(
      () =>
        createSessionPivot("pivot-invalid", {
          sourceRange: "A1:C4",
          target: "E1",
          name: "UnknownField",
          rows: ["Missing"],
          values: [{ field: "Amount" }],
        }),
      /not present in source headers/,
    );
    assert.equal(listSessions().find((session) => session.id === "pivot-invalid")?.dirty, false);
    assert.equal(callWorkbookMethod("pivot-invalid", "pivotCacheCount", []).result.value, 0);

    assert.throws(
      () =>
        createSessionPivot("pivot-invalid", {
          sourceRange: "A1:C4",
          target: "E1",
          name: "UnusedField",
          rows: ["Region"],
          values: [{ field: "Amount" }],
        }),
      /headers are unused: Product/,
    );
    assert.equal(listSessions().find((session) => session.id === "pivot-invalid")?.dirty, false);
    assert.equal(callWorkbookMethod("pivot-invalid", "pivotCount", [0]).result.value, 0);
  } finally {
    closeSession("pivot-invalid");
    await rm(dir, { recursive: true, force: true });
  }
});

test("registers and executes native Table/Pivot tools through MCP", async () => {
  await withClient(async (client) => {
    const tools = await client.listTools();
    const names = tools.tools.map((tool) => tool.name);
    assert.equal(names.includes("formulon_table_operation"), true);
    assert.equal(names.includes("formulon_create_pivot"), true);

    await client.callTool({
      name: "formulon_open_workbook",
      arguments: { sessionId: "mcp-table-pivot" },
    });
    try {
      const written = textPayload(
        await client.callTool({
          name: "formulon_set_range",
          arguments: {
            sessionId: "mcp-table-pivot",
            start: "Sheet1!A1",
            values: sales,
            recalc: true,
          },
        }),
      );
      assert.equal(written.cellsWritten, 12);

      const table = textPayload(
        await client.callTool({
          name: "formulon_table_operation",
          arguments: {
            sessionId: "mcp-table-pivot",
            operation: "create",
            range: "Sheet1!A1:C4",
            name: "McpSales",
          },
        }),
      );
      assert.equal(table.table.index, 0);
      assert.equal(table.table.ref, "A1:C4");

      const listed = textPayload(
        await client.callTool({
          name: "formulon_table_operation",
          arguments: { sessionId: "mcp-table-pivot", operation: "list" },
        }),
      );
      assert.deepEqual(
        listed.tables.map((entry) => entry.index),
        [0],
      );

      const pivot = textPayload(
        await client.callTool({
          name: "formulon_create_pivot",
          arguments: {
            sessionId: "mcp-table-pivot",
            sourceRange: "Sheet1!A1:C4",
            target: "Sheet1!E1",
            name: "McpSalesPivot",
            rows: ["Region"],
            columns: ["Product"],
            values: [{ field: "Amount", aggregation: "sum", numberFormat: "0.00" }],
            layout: "tabular",
            grandTotals: { rows: true, columns: true },
          },
        }),
      );
      assert.equal(pivot.pivotIndex, 0);
      assertPivotTotals(pivot.layout);
    } finally {
      await client.callTool({
        name: "formulon_close_workbook",
        arguments: { sessionId: "mcp-table-pivot" },
      });
    }
  });
});
