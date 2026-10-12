import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { test } from "vitest";
import {
  applySessionMutations,
  callWorkbookMethod,
  closeSession,
  getSessionCell,
  getSessionCellByA1,
  getSessionRange,
  listSessions,
  openSession,
  saveSession,
  sessionDimension,
} from "../../dist/sessions.js";

/** Applies a custom number format to a cell via the low-level style APIs. */
function applyNumberFormat(id, sheet, row, col, formatCode) {
  const numFmtId = callWorkbookMethod(id, "addNumFmt", [formatCode]).result.numFmtId;
  const xfIndex = callWorkbookMethod(id, "addXf", [
    {
      fontIndex: 0,
      fillIndex: 0,
      borderIndex: 0,
      numFmtId,
      horizontalAlign: 0,
      verticalAlign: 0,
      wrapText: false,
    },
  ]).result.index;
  callWorkbookMethod(id, "setCellXfIndex", [sheet, row, col, xfIndex]);
}
test("supports calc mode, profile, partial recalc, and sheet protection", async () => {
  await openSession(undefined, "settings");
  try {
    assert.equal(callWorkbookMethod("settings", "calcMode", []).result.value, 0);
    assert.equal(callWorkbookMethod("settings", "setCalcMode", [1]).result.ok, true);
    assert.equal(callWorkbookMethod("settings", "calcMode", []).result.value, 1);

    assert.equal(
      callWorkbookMethod("settings", "excelProfileId", []).result.value,
      "win-365-en_US",
    );
    assert.equal(
      callWorkbookMethod("settings", "setExcelProfileId", ["mac-365-ja_JP"]).result.ok,
      true,
    );
    assert.equal(
      callWorkbookMethod("settings", "excelProfileId", []).result.value,
      "mac-365-ja_JP",
    );

    applySessionMutations(
      "settings",
      [
        { type: "number", a1: "Sheet1!A1", value: 20 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1+22" },
      ],
      false,
    );
    const partial = callWorkbookMethod("settings", "partialRecalc", [
      { sheet: 0, firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 1 },
    ]);
    assert.equal(partial.result.status.ok, true);
    assert.equal(partial.result.recomputed >= 1, true);
    assert.deepEqual(getSessionCellByA1("settings", "Sheet1!B1").value, {
      kind: "number",
      value: 42,
    });

    const protection = callWorkbookMethod("settings", "getSheetProtection", [0]).result.protection;
    const nextProtection = { ...protection, enabled: 1, sheet: 1, selectLockedCells: 1 };
    assert.equal(
      callWorkbookMethod("settings", "setSheetProtection", [0, nextProtection]).result.ok,
      true,
    );
    const protectedSheet = callWorkbookMethod("settings", "getSheetProtection", [0]);
    assert.equal(protectedSheet.result.status.ok, true);
    assert.equal(protectedSheet.result.protection.enabled, 1);
  } finally {
    closeSession("settings");
  }
});

test("decodes date, currency, and percent number formats on reads", async () => {
  await openSession(undefined, "numfmt");
  try {
    applySessionMutations(
      "numfmt",
      [
        { type: "number", row: 0, col: 0, value: 45000 },
        { type: "number", row: 1, col: 0, value: 1500 },
        { type: "number", row: 2, col: 0, value: 0.153 },
        { type: "number", row: 3, col: 0, value: 42 },
      ],
      false,
    );
    applyNumberFormat("numfmt", 0, 0, 0, "yyyy-mm-dd");
    applyNumberFormat("numfmt", 0, 1, 0, '"¥"#,##0');
    applyNumberFormat("numfmt", 0, 2, 0, "0.0%");

    const date = getSessionCell("numfmt", 0, 0, 0);
    assert.equal(date.formatKind, "date");
    assert.equal(date.formatted, "2023-03-15");
    assert.equal(date.value.value, 45000);

    const currency = getSessionCell("numfmt", 0, 1, 0);
    assert.equal(currency.formatKind, "currency");
    assert.equal(currency.formatted, undefined);

    const percent = getSessionCell("numfmt", 0, 2, 0);
    assert.equal(percent.formatKind, "percent");
    assert.equal(percent.formatted, "15.3%");

    // Unformatted number carries no format annotation.
    const plain = getSessionCell("numfmt", 0, 3, 0);
    assert.equal(plain.formatKind, undefined);
    assert.equal(plain.numberFormat, undefined);

    const range = getSessionRange("numfmt", "A1:A4", {
      maxCells: 100,
      recalc: false,
      includeFormulas: false,
    });
    const a1 = range.cells.find((cell) => cell.a1 === "A1");
    assert.equal(a1.formatted, "2023-03-15");
    const a4 = range.cells.find((cell) => cell.a1 === "A4");
    assert.equal(a4.formatKind, undefined);
  } finally {
    closeSession("numfmt");
  }
});

test("dimension operation lists and sets column width and row height", async () => {
  await openSession(undefined, "dims");
  try {
    const setWidth = sessionDimension("dims", "column", "size", {
      sheet: 0,
      first: 0,
      last: 2,
      size: 120,
    });
    assert.equal(setWidth.status.ok, true);
    const setHeight = sessionDimension("dims", "row", "size", { sheet: 0, row: 0, size: 30 });
    assert.equal(setHeight.status.ok, true);
    const cols = sessionDimension("dims", "column", "list", { sheet: 0 });
    assert.equal(cols.value.status.ok, true);

    sessionDimension("dims", "column", "hidden", { sheet: 0, first: 1, hidden: true });
    sessionDimension("dims", "column", "resetSize", { sheet: 0, first: 0, last: 2 });
    sessionDimension("dims", "row", "resetSize", { sheet: 0, row: 0 });
    const columns = callWorkbookMethod("dims", "getSheetColumns", [0]).result.columns;
    assert.equal(
      columns.every((span) => span.hasWidth === 0),
      true,
    );
    // Clearing the width keeps the column's other state.
    assert.equal(
      columns.some((span) => span.first <= 1 && span.last >= 1 && span.hidden === 1),
      true,
    );
    const rows = callWorkbookMethod("dims", "getSheetRowOverrides", [0]).result.rows;
    assert.equal(
      rows.every((row) => row.row !== 0 || row.hasHeight === 0),
      true,
    );
    assert.throws(
      () => sessionDimension("dims", "column", "size", { sheet: 0, first: 0 }),
      /size is required/,
    );
  } finally {
    closeSession("dims");
  }
});

// ---------------------------------------------------------------------------
// resetSize clears only the width or height of a column or row.
// ---------------------------------------------------------------------------

/**
 * Pairwise matrix (axis x hidden x outline x style x sized x span; the first
 * eight rows cover every value pair).
 */
const RESET_CASES = [
  ["column", "yes", "yes", "no", "yes", "single"],
  ["column", "no", "no", "yes", "no", "multi"],
  ["column", "yes", "no", "no", "no", "single"],
  ["column", "no", "yes", "no", "yes", "multi"],
  ["column", "yes", "no", "no", "yes", "multi"],
  ["row", "no", "yes", "yes", "yes", "single"],
  ["row", "yes", "no", "no", "no", "single"],
  ["column", "yes", "yes", "yes", "no", "multi"],
  // Beyond the pairwise rows: everything at once on each axis, and a bare size.
  ["column", "yes", "yes", "yes", "yes", "multi"],
  ["row", "yes", "yes", "yes", "yes", "single"],
  ["row", "no", "no", "no", "yes", "single"],
  ["column", "no", "no", "no", "yes", "single"],
  ["row", "yes", "no", "yes", "yes", "single"],
  ["row", "no", "yes", "no", "no", "single"],
];

const CONTROL_SIZE = { column: 40, row: 45 };

async function runResetCase([axis, hiddenKey, outlineKey, styleKey, sizedKey, spanKey], index) {
  const id = `reset-size-${index}`;
  const hidden = hiddenKey === "yes";
  const outline = outlineKey === "yes";
  const styled = styleKey === "yes";
  const sized = sizedKey === "yes";
  const dir = await mkdtemp(path.join(tmpdir(), "formulon-mcp-reset-"));
  await openSession(undefined, id);
  try {
    const call = (method, args) => callWorkbookMethod(id, method, args).result;
    const numFmtId = call("addNumFmt", ["0.00%"]).numFmtId;
    const xf = call("addXf", [{ ...call("getCellXf", [0]), status: undefined, numFmtId }]).index;
    const target = axis === "column" ? [3, spanKey === "multi" ? 5 : 3] : [4, 4];
    const [first, last] = target;
    const before = first - 1;
    const after = last + 1;

    const setup =
      axis === "column"
        ? {
            size: (a, b, size) => call("setColumnWidth", [0, a, b, size]),
            hidden: (a, b) => call("setColumnHidden", [0, a, b, true]),
            outline: (a, b) => call("setColumnOutline", [0, a, b, 2]),
            style: (a, b) => call("setColumnStyle", [0, a, b, xf]),
          }
        : {
            size: (a, b, size) => {
              for (let row = a; row <= b; row += 1) {
                call("setRowHeight", [0, row, size]);
              }
            },
            hidden: (a, b) => {
              for (let row = a; row <= b; row += 1) {
                call("setRowHidden", [0, row, true]);
              }
            },
            outline: (a, b) => {
              for (let row = a; row <= b; row += 1) {
                call("setRowOutline", [0, row, 2]);
              }
            },
            style: (a, b) => call("setRowStyle", [0, a, b, xf]),
          };
    // Neighbours carry their own size so a reset that spills over shows up.
    for (const neighbour of [before, after]) {
      setup.size(neighbour, neighbour, CONTROL_SIZE[axis]);
    }
    if (sized) {
      setup.size(first, last, axis === "column" ? 25 : 30);
    }
    if (hidden) {
      setup.hidden(first, last);
    }
    if (outline) {
      setup.outline(first, last);
    }
    if (styled) {
      setup.style(first, last);
    }
    await saveSession(id, path.join(dir, "setup.xlsx"));
    assert.equal(listSessions().find((session) => session.id === id).dirty, false);

    const stateOf = (position) => {
      if (axis === "column") {
        const entry = call("getSheetColumns", [0]).columns.find(
          (item) => item.first <= position && position <= item.last,
        );
        return entry ?? { hasWidth: 0, hidden: 0, outlineLevel: 0, hasStyle: 0, styleXf: 0 };
      }
      const entry = call("getSheetRowOverrides", [0]).rows.find((item) => item.row === position);
      return (
        entry ?? {
          hasHeight: 0,
          customHeight: 0,
          hidden: 0,
          outlineLevel: 0,
          hasStyle: 0,
          styleXf: 0,
        }
      );
    };
    const sizeKey = axis === "column" ? "hasWidth" : "hasHeight";
    const beforeNeighbours = [stateOf(before), stateOf(after)];
    for (let position = first; position <= last; position += 1) {
      assert.equal(stateOf(position)[sizeKey], sized ? 1 : 0, `setup of ${axis} ${position}`);
    }

    const reset =
      axis === "column"
        ? sessionDimension(id, "column", "resetSize", { sheet: 0, first, last })
        : sessionDimension(id, "row", "resetSize", { sheet: 0, row: first });
    assert.equal(reset.status.ok, true);
    assert.equal(reset.session.dirty, true);

    for (let position = first; position <= last; position += 1) {
      const state = stateOf(position);
      assert.equal(state[sizeKey], 0, `${axis} ${position} size cleared`);
      if (axis === "row") {
        assert.equal(state.customHeight, 0);
      }
      assert.equal(state.hidden, hidden ? 1 : 0, `${axis} ${position} hidden`);
      assert.equal(state.outlineLevel, outline ? 2 : 0, `${axis} ${position} outline`);
      assert.equal(state.hasStyle, styled ? 1 : 0, `${axis} ${position} style flag`);
      assert.equal(state.styleXf, styled ? xf : 0, `${axis} ${position} style xf`);
    }
    // The neighbours keep their own size and everything else.
    assert.deepEqual([stateOf(before), stateOf(after)], beforeNeighbours);
    assert.equal(stateOf(before)[sizeKey], 1);

    // The geometry falls back to the sheet default, the same as an untouched line.
    const untouched = 30;
    const measure = (position) =>
      axis === "column"
        ? call("getColumnWidthPt", [0, position, 0]).value
        : call("getRowHeightPt", [0, position]).value;
    for (let position = first; position <= last; position += 1) {
      if (!hidden) {
        assert.equal(measure(position), measure(untouched), `${axis} ${position} default size`);
      }
    }
  } finally {
    closeSession(id);
    await rm(dir, { recursive: true, force: true });
  }
}

for (const [index, row] of RESET_CASES.entries()) {
  test(`resetSize clears only the size: ${row.join(" / ")}`, async () => {
    await runResetCase(row, index);
  });
}

test("resetSize on an unset column or row succeeds without changing anything", async () => {
  await openSession(undefined, "reset-noop");
  try {
    const columns = () => callWorkbookMethod("reset-noop", "getSheetColumns", [0]).result.columns;
    const rows = () => callWorkbookMethod("reset-noop", "getSheetRowOverrides", [0]).result.rows;
    sessionDimension("reset-noop", "column", "size", { sheet: 0, first: 0, size: 20 });
    sessionDimension("reset-noop", "row", "size", { sheet: 0, row: 0, size: 30 });
    const columnsBefore = columns();
    const rowsBefore = rows();

    const column = sessionDimension("reset-noop", "column", "resetSize", {
      sheet: 0,
      first: 10,
      last: 14,
    });
    const row = sessionDimension("reset-noop", "row", "resetSize", { sheet: 0, row: 99 });
    for (const result of [column, row]) {
      assert.equal(result.status.ok, true);
      assert.equal(result.session.dirty, true);
    }
    assert.deepEqual(columns(), columnsBefore);
    assert.deepEqual(rows(), rowsBefore);

    // Resetting twice is the same as resetting once.
    sessionDimension("reset-noop", "column", "resetSize", { sheet: 0, first: 0 });
    sessionDimension("reset-noop", "column", "resetSize", { sheet: 0, first: 0 });
    assert.deepEqual(columns(), []);
    sessionDimension("reset-noop", "row", "resetSize", { sheet: 0, row: 0 });
    sessionDimension("reset-noop", "row", "resetSize", { sheet: 0, row: 0 });
    assert.deepEqual(rows(), []);

    // The edge columns and rows of the grid are valid targets.
    assert.equal(
      sessionDimension("reset-noop", "column", "resetSize", { sheet: 0, first: 16383 }).status.ok,
      true,
    );
    assert.equal(
      sessionDimension("reset-noop", "row", "resetSize", { sheet: 0, row: 1048575 }).status.ok,
      true,
    );
  } finally {
    closeSession("reset-noop");
  }
});

test("resetSize leaves the sheet's other columns and rows alone and keeps values", async () => {
  await openSession(undefined, "reset-scope");
  try {
    applySessionMutations("reset-scope", [{ type: "number", a1: "B2", value: 42 }], true);
    sessionDimension("reset-scope", "column", "size", { sheet: 0, first: 0, last: 4, size: 22 });
    sessionDimension("reset-scope", "column", "resetSize", { sheet: 0, first: 1, last: 2 });
    const columns = callWorkbookMethod("reset-scope", "getSheetColumns", [0]).result.columns;
    const sizedColumns = columns.flatMap((span) =>
      span.hasWidth === 1
        ? Array.from({ length: span.last - span.first + 1 }, (_, offset) => span.first + offset)
        : [],
    );
    assert.deepEqual(sizedColumns, [0, 3, 4]);
    assert.equal(getSessionCellByA1("reset-scope", "B2").value.value, 42);
  } finally {
    closeSession("reset-scope");
  }
});

test("resetSize rejects missing or out-of-grid targets before touching the session", async () => {
  await openSession(undefined, "reset-invalid");
  try {
    const dirty = () => listSessions().find((session) => session.id === "reset-invalid").dirty;
    assert.throws(
      () => sessionDimension("reset-invalid", "column", "resetSize", { sheet: 0 }),
      /first is required/,
    );
    assert.throws(
      () =>
        sessionDimension("reset-invalid", "column", "resetSize", { sheet: 0, first: 3, last: 1 }),
      /last \(1\) must be >= first \(3\)/,
    );
    assert.throws(
      () => sessionDimension("reset-invalid", "column", "resetSize", { sheet: 0, first: 16384 }),
      /exceeds Excel sheet limits/,
    );
    assert.throws(
      () => sessionDimension("reset-invalid", "row", "resetSize", { sheet: 0 }),
      /row is required/,
    );
    assert.throws(
      () => sessionDimension("reset-invalid", "row", "resetSize", { sheet: 0, row: 1048576 }),
      /exceeds Excel sheet limits/,
    );
    assert.throws(
      () => sessionDimension("reset-invalid", "row", "resetSize", { sheet: 5, row: 0 }),
      /sheet index out of bounds/,
    );
    assert.equal(dirty(), false);
  } finally {
    closeSession("reset-invalid");
  }
});
