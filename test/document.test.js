import assert from "node:assert/strict";
import { test } from "vitest";
import {
  buildSessionDocument,
  callWorkbookMethod,
  closeSession,
  getSessionRange,
  openSession,
  sessionPrintSettings,
} from "../dist/sessions.js";

/** Reads a session range as an `A1 -> {value, formula}` map. */
function cellMap(id, range) {
  const read = getSessionRange(id, range, {
    maxCells: 500,
    includeFormulas: true,
    recalc: true,
  });
  return Object.fromEntries(
    read.cells.map((cell) => [cell.a1, { value: cell.value.value, formula: cell.formula }]),
  );
}

const LINE_ITEMS = {
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
};

test("stacks blocks and binds cross-block references to the layout", async () => {
  await openSession(undefined, "doc-invoice");
  try {
    const built = buildSessionDocument(
      "doc-invoice",
      {
        start: "B2",
        blocks: [
          { type: "title", text: "Invoice" },
          { type: "spacer" },
          LINE_ITEMS,
          { type: "spacer" },
          {
            type: "summary",
            name: "total",
            items: [
              { label: "Subtotal", formula: "=SUM({table.Amount})", format: "number" },
              { label: "Tax", formula: "=ROUND({Subtotal}*0.1,0)", format: "number" },
              { label: "Total", formula: "={Subtotal}+{Tax}", format: "number", emphasis: true },
            ],
          },
        ],
      },
      undefined,
    );

    assert.equal(built.width, 4);
    assert.equal(built.range, "B2:E10");
    assert.equal(built.names["table.header"], "B4:E4");
    assert.equal(built.names["table.body"], "B5:E6");
    assert.equal(built.names["table.Amount"], "E5:E6");
    // The summary sits below the table, its label merged across two columns.
    assert.equal(built.names.total, "C8:E10");

    const cells = cellMap("doc-invoice", built.range);
    // A per-row formula binds to that row's own cells, not the anchor row's.
    assert.equal(cells.E5.formula, "=C5*D5");
    assert.equal(cells.E6.formula, "=C6*D6");
    assert.equal(cells.E5.value, 360000);
    // The summary's SUM covers exactly the body the table produced.
    assert.equal(cells[built.names.Subtotal].formula, "=SUM(E5:E6)");
    assert.equal(cells[built.names.Subtotal].value, 850000);
    // An item references an earlier item by label, bound to its actual cell.
    assert.equal(cells[built.names.Tax].formula, "=ROUND(E8*0.1,0)");
    assert.equal(cells[built.names.Total].formula, "=E8+E9");
    assert.equal(cells[built.names.Total].value, 935000);
  } finally {
    closeSession("doc-invoice");
  }
});

test("places sameRow blocks side by side", async () => {
  await openSession(undefined, "doc-samerow");
  try {
    const built = buildSessionDocument(
      "doc-samerow",
      {
        width: 4,
        blocks: [
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
          { type: "text", name: "after", text: "below" },
        ],
      },
      undefined,
    );

    assert.equal(built.names.to, "B2:C2");
    assert.equal(built.names.meta, "D2:E3");
    // The next block clears the tallest block in the group, not just the first.
    assert.equal(built.names.after, "B4:E4");

    const cells = cellMap("doc-samerow", built.range);
    // An ISO date under a date format is stored as a serial, not as text.
    assert.equal(typeof cells[built.names.Date].value, "number");
  } finally {
    closeSession("doc-samerow");
  }
});

test("refuses sameRow blocks that would overlap", async () => {
  await openSession(undefined, "doc-overlap");
  try {
    assert.throws(
      () =>
        buildSessionDocument(
          "doc-overlap",
          {
            width: 4,
            blocks: [
              { type: "text", text: "wide" },
              {
                type: "fields",
                align: "right",
                sameRow: true,
                items: [{ label: "No.", value: 1 }],
              },
            ],
          },
          undefined,
        ),
      /shares a row with block 0 and overlaps its columns/,
    );
  } finally {
    closeSession("doc-overlap");
  }
});

test("names an unknown reference instead of writing a broken formula", async () => {
  await openSession(undefined, "doc-unknown");
  try {
    assert.throws(
      () =>
        buildSessionDocument(
          "doc-unknown",
          {
            blocks: [
              LINE_ITEMS,
              { type: "summary", items: [{ label: "Total", formula: "=SUM({table.Total})" }] },
            ],
          },
          undefined,
        ),
      /unknown reference \{table\.Total\}.*known names/s,
    );
  } finally {
    closeSession("doc-unknown");
  }
});

test("leaves an Excel array constant alone", async () => {
  await openSession(undefined, "doc-array");
  try {
    const built = buildSessionDocument(
      "doc-array",
      {
        width: 2,
        blocks: [{ type: "fields", items: [{ label: "Sum", formula: "=SUM({1,2;3,4})" }] }],
      },
      undefined,
    );
    const cells = cellMap("doc-array", built.range);
    assert.equal(cells[built.names.Sum].formula, "=SUM({1,2;3,4})");
    assert.equal(cells[built.names.Sum].value, 10);
  } finally {
    closeSession("doc-array");
  }
});

test("lays out blank ruled rows for a form to fill in", async () => {
  await openSession(undefined, "doc-form");
  try {
    const built = buildSessionDocument(
      "doc-form",
      {
        blocks: [
          {
            type: "table",
            columns: [{ header: "Item" }, { header: "Amount", format: "number" }],
            rowCount: 5,
          },
        ],
      },
      undefined,
    );
    assert.equal(built.range, "B2:C7");
    // The empty rows are materialized as styled blanks, so the ruling renders.
    const { result } = callWorkbookMethod("doc-form", "getCellXfIndex", [0, 6, 1]);
    assert.ok(result.xfIndex > 0);
  } finally {
    closeSession("doc-form");
  }
});

test("sets the print area and repeats the table header", async () => {
  await openSession(undefined, "doc-print");
  try {
    const built = buildSessionDocument(
      "doc-print",
      { blocks: [{ type: "title", text: "Report" }, LINE_ITEMS], print: "a4-portrait-fit" },
      undefined,
    );
    assert.equal(built.pageCount, 1);

    const { settings } = sessionPrintSettings("doc-print", 0);
    assert.equal(settings.printArea, built.range);
    assert.equal(settings.pageSetup.orientation, "portrait");
    assert.equal(settings.pageSetup.fitToPage, true);
    // The header row is repeated so a table that runs long stays readable.
    assert.equal(settings.printTitles.repeatRows, "3:3");
  } finally {
    closeSession("doc-print");
  }
});

test("rejects an unknown print preset by name", async () => {
  await openSession(undefined, "doc-preset");
  try {
    assert.throws(
      () =>
        buildSessionDocument(
          "doc-preset",
          { blocks: [{ type: "title", text: "x" }], print: "a3-portrait" },
          undefined,
        ),
      /unknown print preset: a3-portrait/,
    );
  } finally {
    closeSession("doc-preset");
  }
});

test("refuses a table wider than the document", async () => {
  await openSession(undefined, "doc-narrow");
  try {
    assert.throws(
      () =>
        buildSessionDocument(
          "doc-narrow",
          {
            width: 2,
            blocks: [
              { type: "table", columns: [{ header: "a" }, { header: "b" }, { header: "c" }] },
            ],
          },
          undefined,
        ),
      /table needs 3 columns but the document is 2 wide/,
    );
  } finally {
    closeSession("doc-narrow");
  }
});

test("accepts positional table rows as well as keyed ones", async () => {
  await openSession(undefined, "doc-positional");
  try {
    const built = buildSessionDocument(
      "doc-positional",
      {
        blocks: [
          {
            type: "table",
            columns: [{ header: "Item" }, { header: "Amount", format: "number" }],
            rows: [
              ["Design", 360000],
              ["Build", 490000],
            ],
          },
        ],
      },
      undefined,
    );
    const cells = cellMap("doc-positional", built.range);
    assert.equal(cells.B3.value, "Design");
    assert.equal(cells.C4.value, 490000);
  } finally {
    closeSession("doc-positional");
  }
});
