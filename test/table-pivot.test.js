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
  traceSession,
} from "../dist/sessions.js";
import { textPayload, withClient } from "./helpers/mcp.js";

const sheet0 = { kind: "Sheet", sheet: 0 };

const sales = [
  ["Region", "Product", "Amount"],
  ["East", "Pen", { f: "=20" }],
  ["East", "Book", 10],
  ["West", "Pen", 5],
];

// The create tool reports cell values as tagged `{ kind, value }`; a raw engine layout record
// (Pivot.getLayout) carries the plain value.
const taggedValue = (cell) => (cell.value.kind === "blank" ? null : cell.value.value);
const plainValue = (cell) => cell.value;

function assertPivotTotals(layout, cellValue = taggedValue) {
  const cells = layout.cells.map((cell) => ({ ...cell, value: cellValue(cell) }));
  const rowLabels = cells
    .filter((cell) => cell.kind === 1 && typeof cell.value === "string")
    .map((cell) => cell.value);
  assert.deepEqual([...new Set(rowLabels)], ["East", "West"]);

  const rowFor = (label) => {
    const matches = cells.filter((cell) => cell.kind === 1 && cell.value === label);
    assert.equal(matches.length, 1, `expected one row label for ${label}`);
    return matches[0].row;
  };
  const eastRow = rowFor("East");
  const westRow = rowFor("West");
  const eastNumbers = cells
    .filter((cell) => cell.row === eastRow && typeof cell.value === "number")
    .map((cell) => cell.value);
  const westNumbers = cells
    .filter((cell) => cell.row === westRow && typeof cell.value === "number")
    .map((cell) => cell.value);
  assert.equal(eastNumbers.includes(30), true);
  assert.equal(westNumbers.includes(5), true);

  const grandTotalNumbers = cells
    .filter((cell) => cell.kind === 6 && typeof cell.value === "number")
    .map((cell) => cell.value);
  assert.equal(grandTotalNumbers.includes(35), true);

  const formatted = cells
    .filter((cell) => typeof cell.value === "number" && [5, 30, 35].includes(cell.value))
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
      assert.equal(callWorkbookMethod("table-reload", sheet0, "listTables", []).result.length, 1);
      assert.deepEqual(
        callWorkbookMethod("table-reload", sheet0, "table", [0]).result.properties.range,
        { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 2 } },
      );
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
    assert.equal(created.spill.blocked, false);
    assert.equal(created.spill.anchor.ref, "Sheet1!E1");
    assert.equal(created.source.headers.join(","), "Region,Product,Amount");
    numberFormat = assertPivotTotals(created.layout);

    const source = callWorkbookMethod(
      "pivot-authoring",
      { kind: "PivotCache", pivotCache: created.cacheId },
      "getWorksheetSource",
      [],
    ).result;
    assert.deepEqual(source.range, {
      sheet: 0,
      rect: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 2 },
    });
    const saved = await saveSession("pivot-authoring", file);
    assert.equal(saved.format, "xlsx");
    assert.equal(saved.losses, undefined);
  } finally {
    closeSession("pivot-authoring");
  }

  await openSession(file, "pivot-reload");
  try {
    assert.equal(
      callWorkbookMethod("pivot-reload", undefined, "listPivotCaches", []).result.length,
      1,
    );
    assert.equal(callWorkbookMethod("pivot-reload", sheet0, "listPivots", []).result.length, 1);
    const layout = callWorkbookMethod(
      "pivot-reload",
      { kind: "Pivot", sheet: 0, pivot: 0 },
      "getLayout",
      [],
    ).result;
    assert.equal(assertPivotTotals(layout, plainValue), numberFormat);
  } finally {
    closeSession("pivot-reload");
    await rm(dir, { recursive: true, force: true });
  }
});

const cellAt = (ref, row, col) => ({ sheet: 0, sheetName: "Sheet1", row, col, ref });

test("reports a PivotTable report's region, blocked by occupied cells, a merge or the sheet edge", async () => {
  await openSession(undefined, "pivot-blocked");
  try {
    const regions = Array.from({ length: 25 }, (_, i) => [`R${String(i).padStart(2, "0")}`, i + 1]);
    setSessionRange("pivot-blocked", "Sheet1!A1", [["Region", "Amount"], ...regions], 0, true);

    // The empty report frame (3 columns by 18 rows) must be free; a refusal leaves no cache behind.
    setSessionRange("pivot-blocked", "Sheet1!E3", [["x"]], 0, true);
    assert.throws(
      () =>
        createSessionPivot("pivot-blocked", {
          sourceRange: "A1:B26",
          target: "D1",
          name: "Framed",
          rows: ["Region"],
          values: [{ field: "Amount" }],
        }),
      /would overwrite non-empty cell Sheet1!E3/,
    );
    assert.equal(
      callWorkbookMethod("pivot-blocked", undefined, "listPivotCaches", []).result.length,
      0,
    );
    setSessionRange("pivot-blocked", "Sheet1!E3", [[{ blank: true }]], 0, true);

    // The finished report needs D1:E27; D22 lies outside the frame but inside the report.
    setSessionRange("pivot-blocked", "Sheet1!D22", [["x"]], 0, true);
    const created = createSessionPivot("pivot-blocked", {
      sourceRange: "A1:B26",
      target: "D1",
      name: "Blocked",
      rows: ["Region"],
      values: [{ field: "Amount" }],
    });
    const blocked = {
      blocked: true,
      cause: "occupied",
      anchor: cellAt("Sheet1!D1", 0, 3),
      rows: 27,
      cols: 2,
      range: "Sheet1!D1:E27",
    };
    assert.deepEqual(created.spill, blocked);
    assert.deepEqual(created.layout.rect, { firstRow: 0, firstCol: 3, lastRow: 26, lastCol: 4 });
    assert.deepEqual(getSessionCellByA1("pivot-blocked", "Sheet1!D1").value, {
      kind: "error",
      errorCode: 8,
      errorName: "#SPILL!",
    });
    assert.deepEqual(traceSession("pivot-blocked", "spillInfo", 0, 0, 3, 1).spill, blocked);

    // Freeing the range draws the report on the next recalculation.
    setSessionRange("pivot-blocked", "Sheet1!D22", [[{ blank: true }]], 0, true);
    assert.deepEqual(traceSession("pivot-blocked", "spillInfo", 0, 0, 3, 1).spill, {
      ...blocked,
      blocked: false,
      cause: null,
    });
    assert.deepEqual(getSessionCellByA1("pivot-blocked", "Sheet1!E27").value, {
      kind: "number",
      value: 325,
    });

    // A merge inside the report's range reads as its own cause.
    callWorkbookMethod("pivot-blocked", sheet0, "addMerge", ["J20:K20"]);
    const merged = createSessionPivot("pivot-blocked", {
      sourceRange: "A1:B26",
      target: "J1",
      name: "Merged",
      rows: ["Region"],
      values: [{ field: "Amount" }],
    });
    assert.equal(merged.spill.blocked, true);
    assert.equal(merged.spill.cause, "merged");

    // A report running past the last column is blocked with its range clipped to the grid.
    const edge = createSessionPivot("pivot-blocked", {
      sourceRange: "A1:B6",
      target: "XFB1",
      name: "PastEdge",
      columns: ["Region"],
      values: [{ field: "Amount" }],
    });
    assert.equal(edge.spill.blocked, true);
    assert.equal(edge.spill.cause, "beyondEdge");
    assert.equal(edge.spill.range, "Sheet1!XFB1:XFD3");
    assert.deepEqual(traceSession("pivot-blocked", "spillInfo", 0, 0, 16381, 1).spill, edge.spill);
  } finally {
    closeSession("pivot-blocked");
  }
});

test("reads a spill's anchor at a phantom cell and a report's anchor below its page rows", async () => {
  await openSession(undefined, "spill-anchor");
  try {
    const rows = Array.from({ length: 5 }, (_, i) => [`R${i}`, `C${i % 2}`, i + 1]);
    setSessionRange("spill-anchor", "Sheet1!A1", [["Region", "Kind", "Amount"], ...rows], 0, true);
    const paged = createSessionPivot("spill-anchor", {
      sourceRange: "A1:C6",
      target: "F5",
      name: "Paged",
      rows: ["Region"],
      pages: ["Kind"],
      values: [{ field: "Amount" }],
    });
    assert.deepEqual(paged.spill, {
      blocked: false,
      cause: null,
      anchor: cellAt("Sheet1!F5", 4, 5),
      rows: 9,
      cols: 2,
      range: "Sheet1!F3:G11",
    });
    // A report's other cells, its page rows included, are ordinary cells.
    assert.equal(traceSession("spill-anchor", "spillInfo", 0, 2, 5, 1).spill, null);

    setSessionRange("spill-anchor", "Sheet1!K1", [[{ f: "=SEQUENCE(3)" }]], 0, true);
    const drawn = {
      blocked: false,
      cause: null,
      anchor: cellAt("Sheet1!K1", 0, 10),
      rows: 3,
      cols: 1,
      range: "Sheet1!K1:K3",
    };
    assert.deepEqual(traceSession("spill-anchor", "spillInfo", 0, 1, 10, 1).spill, drawn);
    setSessionRange("spill-anchor", "Sheet1!K3", [["x"]], 0, true);
    assert.deepEqual(traceSession("spill-anchor", "spillInfo", 0, 0, 10, 1).spill, {
      ...drawn,
      blocked: true,
      cause: "occupied",
    });
  } finally {
    closeSession("spill-anchor");
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
    assert.equal(
      callWorkbookMethod("pivot-invalid", undefined, "listPivotCaches", []).result.length,
      0,
    );

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
    assert.equal(callWorkbookMethod("pivot-invalid", sheet0, "listPivots", []).result.length, 0);
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
