import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  addSessionDataTable,
  addSessionSparklines,
  callWorkbookMethod,
  closeSession,
  getSession,
  getSessionCellByA1,
  getSessionDataTable,
  insertSessionShape,
  inspectSession,
  listSessionShapes,
  listSessionSparklines,
  openSession,
  recalcSessionDataTables,
  removeSessionDataTable,
  removeSessionShape,
  removeSessionSparklines,
  saveSession,
  sessionDimension,
  setSessionDefinedName,
  setSessionRange,
  setSessionSheetView,
  styleSessionRange,
  tableOperation,
  updateSessionShape,
  updateSessionSparklines,
} from "../dist/sessions.js";
import { callTool, textPayload, withClient } from "./helpers/mcp.js";

async function withSession(id, fn) {
  await openSession(undefined, id);
  try {
    return await fn();
  } finally {
    closeSession(id);
  }
}

async function withTempDir(fn) {
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-sheet-objects-"));
  try {
    return await fn(dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

// Pairwise over placement x geometry x fill x line x rotation x text, with
// text omitted for line geometry (a connector holds none).
const SHAPE_CASES = [
  ["oneCell", "omit", "none", "none", "negative", "value"],
  ["oneCell", "line", "omit", "omit", "omit", "omit"],
  ["twoCell", "roundRect", "hex", "none", "positive", "omit"],
  ["twoCell", "omit", "none", "hex", "omit", "omit"],
  ["twoCell", "roundRect", "hex", "omit", "omit", "value"],
  ["twoCell", "line", "omit", "none", "negative", "omit"],
  ["oneCell", "omit", "hex", "hex", "positive", "value"],
  ["oneCell", "roundRect", "omit", "hex", "negative", "omit"],
  ["twoCell", "omit", "omit", "omit", "positive", "value"],
  ["oneCell", "line", "hex", "hex", "negative", "omit"],
  ["twoCell", "roundRect", "none", "omit", "negative", "omit"],
  ["twoCell", "omit", "hex", "none", "omit", "value"],
  ["oneCell", "roundRect", "none", "omit", "positive", "value"],
  ["twoCell", "line", "none", "none", "positive", "omit"],
];

test.each(SHAPE_CASES)(
  "inserts a %s shape (geometry %s, fill %s, line %s, rotation %s, text %s)",
  async (placement, geometry, fill, line, rotation, text) => {
    await withSession("shape-insert", () => {
      const input = {
        anchor: "B2",
        ...(placement === "twoCell"
          ? { to: "E6" }
          : { width: { size: 120 }, height: { size: 40 } }),
        ...(geometry === "omit" ? {} : { geometry }),
        ...(fill === "omit" ? {} : { fill: fill === "hex" ? "#336699" : "none" }),
        ...(line === "omit" ? {} : { line: line === "hex" ? "#AA0000" : "none" }),
        ...(rotation === "omit" ? {} : { rotation: rotation === "positive" ? 45 : -90 }),
        ...(text === "omit" ? {} : { text: "Note\nSecond line" }),
      };
      const inserted = insertSessionShape("shape-insert", 0, input);
      const shape = inserted.placed;
      assert.equal(shape.objectId, inserted.objectId);
      assert.equal(shape.anchorKind, placement);
      assert.equal(shape.from, "B2");
      if (placement === "twoCell") {
        assert.equal(shape.to, "E6");
      } else {
        assert.equal(shape.width.pt, 120);
        assert.equal(shape.height.pt, 40);
      }
      assert.equal(shape.geometry, geometry === "omit" ? "rect" : geometry);
      assert.equal(shape.kind, geometry === "line" ? "connector" : "shape");
      assert.equal(shape.fill, fill === "omit" ? undefined : fill === "hex" ? "#336699" : "none");
      assert.equal(shape.line, line === "omit" ? undefined : line === "hex" ? "#AA0000" : "none");
      assert.equal(
        shape.rotation,
        rotation === "omit" ? undefined : rotation === "positive" ? 45 : 270,
      );
      assert.equal(shape.text, text === "omit" ? undefined : "Note\nSecond line");
      assert.equal(getSession("shape-insert").dirty, true);
    });
  },
);

test("updates, moves and removes a shape, and keeps it across a save", async () => {
  await withTempDir(async (dir) => {
    const file = path.join(dir, "shapes.xlsx");
    await withSession("shape-edit", async () => {
      const box = insertSessionShape("shape-edit", 0, {
        anchor: "B2",
        width: { size: 100 },
        height: { size: 30 },
        text: "Draft",
        fill: "#FFEE00",
        textBox: true,
      });
      const connector = insertSessionShape("shape-edit", 0, {
        anchor: "A1",
        to: "C3",
        geometry: "straightConnector1",
      });

      const updated = updateSessionShape("shape-edit", 0, box.objectId, {
        text: "Final",
        anchor: "D4",
      });
      assert.equal(updated.placed.text, "Final");
      assert.equal(updated.placed.from, "D4");
      // Unstated fields are kept: size, fill and the text-box flag.
      assert.equal(updated.placed.width.pt, 100);
      assert.equal(updated.placed.fill, "#FFEE00");
      assert.equal(updated.placed.textBox, true);

      const moved = updateSessionShape("shape-edit", 0, connector.objectId, { to: "F8" });
      assert.equal(moved.placed.from, "A1");
      assert.equal(moved.placed.to, "F8");
      assert.throws(
        () => updateSessionShape("shape-edit", 0, connector.objectId, { width: { size: 10 } }),
        /two-cell shape is sized by its to cell/,
      );
      assert.throws(
        () => updateSessionShape("shape-edit", 0, box.objectId, { to: "H9" }),
        /keeps its anchor kind/,
      );
      assert.throws(
        () => insertSessionShape("shape-edit", 0, { anchor: "A1", fill: "#80FF0000", to: "B2" }),
        /opaque/,
      );
      assert.throws(
        () => insertSessionShape("shape-edit", 0, { anchor: "A1" }),
        /width and height/,
      );

      await saveSession("shape-edit", file);
    });
    await openSession(file, "shape-reopen");
    try {
      const listed = listSessionShapes("shape-reopen", 0);
      assert.equal(listed.count, 2);
      const [box] = listed.shapes;
      assert.equal(box.text, "Final");
      assert.equal(box.from, "D4");
      removeSessionShape("shape-reopen", 0, box.objectId);
      assert.equal(listSessionShapes("shape-reopen", 0).count, 1);
    } finally {
      closeSession("shape-reopen");
    }
  });
});

test("adds, updates and removes sparkline groups without resetting unstated settings", async () => {
  await withTempDir(async (dir) => {
    const file = path.join(dir, "sparklines.xlsx");
    await withSession("sparklines", async () => {
      setSessionRange(
        "sparklines",
        "A1",
        [
          [1, 3, 2, 5],
          [4, -1, 2, 0],
        ],
        0,
        false,
      );
      const added = addSessionSparklines(
        "sparklines",
        0,
        [
          { data: "A1:D1", location: "E1" },
          { data: "A2:D2", location: "E2" },
        ],
        { type: "column", high: true, colors: { series: "#1F77B4" }, maxAxis: "group" },
      );
      assert.equal(added.group.index, 0);
      assert.equal(added.group.type, "column");
      // A bare range is qualified with the host sheet.
      assert.deepEqual(added.group.sparklines, [
        { data: "'Sheet1'!A1:D1", location: "E1" },
        { data: "'Sheet1'!A2:D2", location: "E2" },
      ]);
      assert.equal(added.group.high, true);
      assert.equal(added.group.colors.series, "#1F77B4");
      assert.equal(added.group.maxAxis, "group");
      assert.equal(added.group.rightToLeft, undefined);

      const updated = updateSessionSparklines("sparklines", 0, 0, undefined, {
        negative: true,
        minAxis: -2,
        rightToLeft: true,
      });
      assert.equal(updated.group.negative, true);
      assert.equal(updated.group.rightToLeft, true);
      assert.equal(updated.group.minAxis, -2);
      // The engine rewrites every modelled field on update; these must survive.
      assert.equal(updated.group.type, "column");
      assert.equal(updated.group.high, true);
      assert.equal(updated.group.colors.series, "#1F77B4");
      assert.equal(updated.group.maxAxis, "group");

      const replaced = updateSessionSparklines(
        "sparklines",
        0,
        0,
        [{ data: "Sheet1!A1:D2", location: "F1" }],
        {},
      );
      assert.deepEqual(replaced.group.sparklines, [{ data: "'Sheet1'!A1:D2", location: "F1" }]);
      assert.equal(replaced.group.negative, true);
      assert.throws(() => updateSessionSparklines("sparklines", 0, 0, undefined, {}), /requires/);
      assert.throws(() => addSessionSparklines("sparklines", 0, [], {}), /at least one/);
      assert.throws(
        () => updateSessionSparklines("sparklines", 0, 3, undefined, { high: true }),
        /sparkline group index out of range/,
      );
      await saveSession("sparklines", file);
    });
    await openSession(file, "sparklines-reopen");
    try {
      const listed = listSessionSparklines("sparklines-reopen", 0);
      assert.equal(listed.count, 1);
      assert.equal(listed.groups[0].type, "column");
      assert.equal(listed.groups[0].rightToLeft, true);
      removeSessionSparklines("sparklines-reopen", 0, 0);
      assert.equal(listSessionSparklines("sparklines-reopen", 0).count, 0);
    } finally {
      closeSession("sparklines-reopen");
    }
  });
});

test("models a one-variable What-If data table", async () => {
  await withSession("data-table", () => {
    // Input cell A1; formula above the result range; substitution values to its left.
    setSessionRange(
      "data-table",
      "A1",
      [
        [2, null, null],
        [null, null, { f: "=A1*10" }],
        [null, 1, null],
        [null, 2, null],
        [null, 3, null],
      ],
      0,
      true,
    );
    assert.throws(() => addSessionDataTable("data-table", 0, { range: "C3:C5" }), /rowInput/);
    const added = addSessionDataTable("data-table", 0, { range: "C3:C5", colInput: "A1" });
    assert.deepEqual(added.dataTable, {
      range: "C3:C5",
      twoDimensional: false,
      colInput: { cell: "A1" },
    });
    for (const [ref, expected] of [
      ["C3", 10],
      ["C4", 20],
      ["C5", 30],
    ]) {
      assert.equal(getSessionCellByA1("data-table", ref).value.value, expected);
    }
    assert.equal(getSessionCellByA1("data-table", "A1").value.value, 2);
    assert.match(getSessionCellByA1("data-table", "C4").formula, /TABLE\(,A1\)/);
    assert.throws(() => setSessionRange("data-table", "C4", [[99]], 0, true), /data table/);
    assert.equal(getSessionDataTable("data-table", 0, "C5").dataTable.range, "C3:C5");
    assert.equal(getSessionDataTable("data-table", 0, "B5").dataTable, null);
    assert.equal(recalcSessionDataTables("data-table").status.ok, true);

    const removed = removeSessionDataTable("data-table", 0, "C4");
    assert.equal(removed.removed.range, "C3:C5");
    assert.equal(getSessionCellByA1("data-table", "C4").value.kind, "blank");
    assert.throws(() => removeSessionDataTable("data-table", 0, "C4"), /no data table/);
  });
});

test("sets formula display, split panes, scroll origin and the active sheet", async () => {
  await withSession("sheet-view", () => {
    callWorkbookMethod("sheet-view", undefined, "addSheet", ["Second"]);
    const result = setSessionSheetView("sheet-view", "Second", {
      showFormulas: true,
      showOutlineSymbols: false,
      splitX: 72,
      splitY: 36,
      topLeft: "C10",
      active: true,
    });
    assert.equal(result.statuses.length, 5);
    const view = callWorkbookMethod(
      "sheet-view",
      { kind: "Sheet", sheet: 1 },
      "getView",
      [],
    ).result;
    assert.equal(view.showFormulas, true);
    assert.equal(view.showOutlineSymbols, false);
    assert.equal(view.splitXTwips, 1440);
    assert.equal(view.splitYTwips, 720);
    assert.deepEqual(view.topLeft, { row: 9, col: 2 });
    assert.equal(callWorkbookMethod("sheet-view", undefined, "activeSheet", []).result.index, 1);

    // A freeze replaces the split.
    setSessionSheetView("sheet-view", 1, { freezeRows: 1 });
    const frozen = callWorkbookMethod(
      "sheet-view",
      { kind: "Sheet", sheet: 1 },
      "getView",
      [],
    ).result;
    assert.equal(frozen.freezeRows, 1);
    assert.equal(frozen.splitXTwips, 0);
    assert.throws(
      () => setSessionSheetView("sheet-view", 1, { freezeRows: 1, splitX: 10 }),
      /either frozen or split/,
    );
    assert.throws(
      () => setSessionSheetView("sheet-view", 1, { topLeft: "Sheet1!A1" }),
      /without a sheet name/,
    );
    // Hiding the active sheet activates the next visible one.
    setSessionSheetView("sheet-view", 1, { visibility: "hidden" });
    assert.equal(callWorkbookMethod("sheet-view", undefined, "activeSheet", []).result.index, 0);
    assert.throws(
      () => setSessionSheetView("sheet-view", 1, { active: true }),
      /setActiveSheet sheet is hidden/,
    );
  });
});

test("marks outline groups collapsed on rows and columns", async () => {
  await withSession("collapsed", () => {
    sessionDimension("collapsed", "row", "outline", { row: 1, level: 1 });
    sessionDimension("collapsed", "row", "collapsed", { row: 2, collapsed: true });
    sessionDimension("collapsed", "column", "collapsed", { first: 3, last: 4, collapsed: true });
    const rows = sessionDimension("collapsed", "row", "list", {}).value.rows;
    assert.equal(rows.find((row) => row.row === 2).collapsed, true);
    assert.equal(rows.find((row) => row.row === 1).hidden, false);
    const columns = sessionDimension("collapsed", "column", "list", {}).value.columns;
    assert.equal(columns.find((column) => column.first === 3).collapsed, true);
    assert.throws(
      () => sessionDimension("collapsed", "row", "collapsed", { row: 2 }),
      /collapsed is required/,
    );
  });
});

test("defined names carry a comment and a hidden flag", async () => {
  await withSession("name-attrs", () => {
    setSessionDefinedName("name-attrs", "Rate", "=0.1", undefined, {
      comment: "Tax rate",
      hidden: true,
    });
    let names = inspectSession("name-attrs", false, 0).workbook.definedNames;
    assert.deepEqual(
      names.find((entry) => entry.name === "Rate"),
      {
        name: "Rate",
        formula: "=0.1",
        localSheetId: -1,
        comment: "Tax rate",
        hidden: true,
      },
    );
    // Replacing the formula keeps both attributes.
    setSessionDefinedName("name-attrs", "Rate", "0.2");
    names = inspectSession("name-attrs", false, 0).workbook.definedNames;
    assert.equal(names.find((entry) => entry.name === "Rate").comment, "Tax rate");
    setSessionDefinedName("name-attrs", "Rate", "0.2", undefined, { comment: "", hidden: false });
    names = inspectSession("name-attrs", false, 0).workbook.definedNames;
    assert.deepEqual(
      names.find((entry) => entry.name === "Rate"),
      {
        name: "Rate",
        formula: "=0.2",
        localSheetId: -1,
      },
    );
    assert.throws(
      () => setSessionDefinedName("name-attrs", "Rate", "", undefined, { hidden: true }),
      /removing/,
    );
  });
});

test("tables take style toggles, totals-row cells and a rename that rewrites references", async () => {
  await withSession("table-totals", () => {
    setSessionRange(
      "table-totals",
      "A1",
      [
        ["Item", "Qty"],
        ["Pen", 2],
        ["Ink", 3],
        [null, null],
      ],
      0,
      true,
    );
    setSessionRange("table-totals", "E1", [[{ f: "=SUM(Items[Qty])" }]], 0, true);
    const created = tableOperation("table-totals", "create", {
      range: "A1:B3",
      name: "Items",
      style: "TableStyleMedium2",
      showFirstColumn: true,
      showRowStripes: false,
    });
    // The style toggles are no longer part of the table properties record; read them from the handle.
    const styleOptions = callWorkbookMethod(
      "table-totals",
      { kind: "Table", table: "Items" },
      "getStyleOptions",
      [],
    ).result;
    assert.equal(styleOptions.styleName, "TableStyleMedium2");
    assert.equal(styleOptions.showFirstColumn, true);
    assert.equal(styleOptions.showRowStripes, false);
    assert.equal(created.table.name, "Items");
    // Settings given while the totals row is off are stored for when it is shown.
    const stored = tableOperation("table-totals", "update", {
      index: 0,
      columns: [{}, { totalsRowFunction: "max" }],
    });
    assert.equal(stored.table.ref, "A1:B3");
    assert.equal(stored.table.raw.columns[1].totalsRowFunction, "max");
    const updated = tableOperation("table-totals", "update", {
      index: 0,
      range: "A1:B4",
      totalsRow: true,
      name: "Goods",
      columns: [{ totalsRowLabel: "Total" }, { totalsRowFunction: "sum" }],
    });
    assert.equal(updated.table.name, "Goods");
    assert.deepEqual(
      updated.table.raw.columns.map((column) => [
        column.totalsRowFunction,
        column.totalsRowLabel,
        column.calculatedFormula,
      ]),
      [
        [null, "Total", null],
        ["sum", null, null],
      ],
    );
    assert.equal(getSessionCellByA1("table-totals", "E1", 0, true).formula, "=SUM(Goods[Qty])");
    assert.equal(getSessionCellByA1("table-totals", "B4", 0, true).value.value, 5);
    assert.equal(getSessionCellByA1("table-totals", "A4").value.value, "Total");
    const cleared = tableOperation("table-totals", "update", {
      index: 0,
      columns: [{ totalsRowLabel: null }, { totalsRowFunction: null }],
    });
    assert.deepEqual(
      cleared.table.raw.columns.map((column) => [column.totalsRowFunction, column.totalsRowLabel]),
      [
        [null, null],
        [null, null],
      ],
    );
    assert.throws(
      () => tableOperation("table-totals", "update", { index: 0, columns: [{ name: "Thing" }] }),
      /columns\[0\]\.name names a column on create only/,
    );
    assert.throws(
      () =>
        tableOperation("table-totals", "update", {
          index: 0,
          columns: [{ totalsRowFunction: "none" }],
        }),
      /unknown totalsRowFunction "none"/,
    );
  });
});

test("table columns name the columns on create and carry the calculated-column mark", async () => {
  await withSession("table-calc", () => {
    setSessionRange(
      "table-calc",
      "A2",
      [
        [2, 0],
        [3, 0],
      ],
      0,
      true,
    );
    assert.throws(
      () =>
        tableOperation("table-calc", "create", {
          range: "A1:B3",
          name: "Calc",
          columns: [{ name: "Qty" }, {}],
        }),
      /give every column a name/,
    );
    const created = tableOperation("table-calc", "create", {
      range: "A1:B3",
      name: "Calc",
      columns: [{ name: "Qty" }, { name: "Ten", calculatedFormula: "=[@Qty]*10" }],
    });
    assert.deepEqual(
      created.table.raw.columns.map((column) => [column.name, column.calculatedFormula]),
      [
        ["Qty", null],
        ["Ten", "=[@Qty]*10"],
      ],
    );
    assert.equal(getSessionCellByA1("table-calc", "A1").value.value, "Qty");
    assert.equal(getSessionCellByA1("table-calc", "B2").value.value, 20);
    assert.equal(getSessionCellByA1("table-calc", "B3").value.value, 30);

    // The same formula refills nothing: a constant typed over one cell stays.
    setSessionRange("table-calc", "B3", [[7]], 0, true);
    tableOperation("table-calc", "update", {
      index: 0,
      columns: [{}, { calculatedFormula: "=[@Qty]*10" }],
    });
    assert.equal(getSessionCellByA1("table-calc", "B3").value.value, 7);

    // Removing the mark keeps the cells.
    const unmarked = tableOperation("table-calc", "update", {
      index: 0,
      columns: [{}, { calculatedFormula: null }],
    });
    assert.equal(unmarked.table.raw.columns[1].calculatedFormula, null);
    assert.equal(getSessionCellByA1("table-calc", "B2").value.value, 20);
  });
});

test("table create leaves column naming and totals defaults to the engine", async () => {
  await withSession("table-names", () => {
    setSessionRange(
      "table-names",
      "A1",
      [
        ["Qty", null, "qty", 5],
        [1, 2, 3, 4],
        [null, null, null, null],
      ],
      0,
      true,
    );
    const named = tableOperation("table-names", "create", {
      range: "A1:D3",
      name: "Numbered",
      totalsRow: true,
      columns: [{}, {}, { totalsRowFunction: "max" }],
    });
    assert.deepEqual(
      named.table.raw.columns.map((column) => [
        column.name,
        column.totalsRowFunction,
        column.totalsRowLabel,
      ]),
      [
        ["Qty", null, "Total"],
        ["Column1", null, null],
        ["qty2", "max", null],
        ["5", "sum", null],
      ],
    );
    assert.equal(getSessionCellByA1("table-names", "B1").value.value, "Column1");

    setSessionRange(
      "table-names",
      "F1",
      [
        [1, 2],
        [3, 4],
      ],
      0,
      true,
    );
    const headless = tableOperation("table-names", "create", {
      range: "F1:G2",
      name: "Headless",
      headerRow: false,
    });
    assert.deepEqual(
      headless.table.raw.columns.map((column) => column.name),
      ["Column1", "Column2"],
    );
  });
});

test("dates decode and encode under the 1904 date system", async () => {
  await withSession("date1904", () => {
    callWorkbookMethod("date1904", undefined, "setDate1904", [true]);
    // 1904 serial 43538 is the 1900 serial 45000: 2023-03-15.
    setSessionRange("date1904", "A1", [[43538]], 0, true);
    styleSessionRange("date1904", "A1", { numberFormat: "yyyy-mm-dd" }, 0, "existing");
    const cell = getSessionCellByA1("date1904", "A1");
    assert.equal(cell.formatted, "2023-03-15");
    assert.equal(cell.text, "2023-03-15");
  });
});

test("new tools are registered and reachable over stdio", async () => {
  await withClient(async (client) => {
    const names = (await client.listTools()).tools.map((tool) => tool.name);
    for (const name of [
      "formulon_shape_operation",
      "formulon_sparkline_operation",
      "formulon_data_table_operation",
    ]) {
      assert.equal(names.includes(name), true, name);
    }
    const sessionId = "sheet-objects-stdio";
    await callTool(client, "formulon_open_workbook", { sessionId });
    try {
      const inserted = textPayload(
        await callTool(client, "formulon_shape_operation", {
          sessionId,
          operation: "insert",
          anchor: "B2",
          width: { size: 2, unit: "mm" },
          height: { size: 30, unit: "px" },
          text: "Hi",
        }),
      );
      assert.equal(inserted.placed.text, "Hi");
      await callTool(client, "formulon_comment_operation", {
        sessionId,
        operation: "set",
        row: 0,
        col: 0,
        author: "me",
        text: "note",
      });
      textPayload(
        await callTool(client, "formulon_comment_operation", {
          sessionId,
          operation: "visibility",
          row: 0,
          col: 0,
          visible: true,
        }),
      );
      const visible = textPayload(
        await callTool(client, "formulon_workbook_call", {
          sessionId,
          target: { kind: "Sheet", sheet: 0 },
          method: "getCommentVisible",
          args: [0, 0],
        }),
      );
      assert.equal(visible.result, true);
      const sparkline = textPayload(
        await callTool(client, "formulon_sparkline_operation", {
          sessionId,
          operation: "add",
          sparklines: [{ data: "C1:F1", location: "G1" }],
        }),
      );
      assert.equal(sparkline.group.type, "line");
      const table = textPayload(
        await callTool(client, "formulon_data_table_operation", {
          sessionId,
          operation: "get",
          cell: "A1",
        }),
      );
      assert.equal(table.dataTable, null);
    } finally {
      await callTool(client, "formulon_close_workbook", { sessionId });
    }
  });
});
