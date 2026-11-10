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

/** Calls a method of the workbook, a `Sheet`, or a sheet's `ooxml` (`kind`), by 0.19 name. */
function call(id, kind, method, ...args) {
  const target = kind === "Workbook" ? undefined : { kind, sheet: 0 };
  return callWorkbookMethod(id, target, method, args).result;
}

/** Applies a custom number format to a cell via the low-level style APIs. */
function applyNumberFormat(id, sheet, row, col, formatCode) {
  const numFmtId = callWorkbookMethod(id, { kind: "WorkbookOoxml" }, "addNumberFormat", [
    formatCode,
  ]).result;
  const xfIndex = callWorkbookMethod(id, { kind: "WorkbookOoxml" }, "addCellXf", [
    {
      fontIndex: 0,
      fillIndex: 0,
      borderIndex: 0,
      numFmtId,
      horizontal: 0,
      vertical: 0,
      wrapText: false,
    },
  ]).result;
  callWorkbookMethod(id, { kind: "SheetOoxml", sheet }, "setCellXfIndex", [row, col, xfIndex]);
}
test("supports calc mode, profile, partial recalc, and sheet protection", async () => {
  await openSession(undefined, "settings");
  try {
    assert.equal(call("settings", "Workbook", "getCalcMode"), 0);
    // A setter returns nothing; a failure would have thrown.
    assert.equal(call("settings", "Workbook", "setCalcMode", 1), undefined);
    assert.equal(call("settings", "Workbook", "getCalcMode"), 1);

    assert.equal(call("settings", "Workbook", "getProfileId"), "win-365-en_US");
    assert.equal(call("settings", "Workbook", "setProfileId", "mac-365-ja_JP"), undefined);
    assert.equal(call("settings", "Workbook", "getProfileId"), "mac-365-ja_JP");

    applySessionMutations(
      "settings",
      [
        { type: "number", a1: "Sheet1!A1", value: 20 },
        { type: "formula", a1: "Sheet1!B1", formula: "=A1+22" },
      ],
      false,
    );
    const recomputed = call("settings", "Sheet", "recalcViewport", {
      firstRow: 0,
      lastRow: 0,
      firstCol: 0,
      lastCol: 1,
    });
    assert.equal(recomputed >= 1, true);
    assert.deepEqual(getSessionCellByA1("settings", "Sheet1!B1").value, {
      kind: "number",
      value: 42,
    });

    // An unprotected sheet reads null; the record's presence is the protection.
    assert.equal(call("settings", "Sheet", "getSheetProtection"), null);
    assert.equal(
      call("settings", "Sheet", "setSheetProtection", {
        allowSelectLockedCells: true,
        allowFormatCells: true,
      }),
      undefined,
    );
    const protectedSheet = call("settings", "Sheet", "getSheetProtection");
    assert.notEqual(protectedSheet, null);
    assert.equal(protectedSheet.allowSelectLockedCells, true);
    assert.equal(protectedSheet.allowFormatCells, true);
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
    const columns = call("dims", "Sheet", "listColumnOverrides");
    assert.equal(
      columns.every((span) => span.widthChars === null),
      true,
    );
    // Clearing the width keeps the column's other state.
    assert.equal(
      columns.some((span) => span.first <= 1 && span.last >= 1 && span.hidden === true),
      true,
    );
    const rows = call("dims", "Sheet", "listRowOverrides");
    assert.equal(
      rows.every((row) => row.row !== 0 || row.heightPt === null),
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
    const sheet = (method, ...args) => call(id, "Sheet", method, ...args);
    const styles = (method, ...args) => call(id, "SheetOoxml", method, ...args);
    const tables = (method, ...args) =>
      callWorkbookMethod(id, { kind: "WorkbookOoxml" }, method, args).result;
    const numFmtId = tables("addNumberFormat", "0.00%");
    const xf = tables("addCellXf", { ...tables("getCellXf", 0), numFmtId });
    const target = axis === "column" ? [3, spanKey === "multi" ? 5 : 3] : [4, 4];
    const [first, last] = target;
    const before = first - 1;
    const after = last + 1;

    const setup =
      axis === "column"
        ? {
            size: (a, b, size) => sheet("setColumnWidth", a, b, size),
            hidden: (a, b) => sheet("setColumnHidden", a, b, true),
            outline: (a, b) => sheet("setColumnOutline", a, b, 2),
            style: (a, b) => styles("setColumnXfIndex", a, b, xf),
          }
        : {
            size: (a, b, size) => sheet("setRowHeight", a, b, size),
            hidden: (a, b) => sheet("setRowHidden", a, b, true),
            outline: (a, b) => sheet("setRowOutline", a, b, 2),
            style: (a, b) => styles("setRowXfIndex", a, b, xf),
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

    // One line's state: whether it states a size, its flags, and the xf of its own (null when none).
    const stateOf = (position) => {
      const unset = { sized: false, customHeight: false, hidden: false, outlineLevel: 0 };
      if (axis === "column") {
        const entry = sheet("listColumnOverrides").find(
          (item) => item.first <= position && position <= item.last,
        );
        return {
          ...unset,
          ...(entry && {
            sized: entry.widthChars !== null,
            hidden: entry.hidden,
            outlineLevel: entry.outlineLevel,
          }),
          styleXf: styles("getColumnXfIndex", position),
        };
      }
      const entry = sheet("listRowOverrides").find((item) => item.row === position);
      return {
        ...unset,
        ...(entry && {
          sized: entry.heightPt !== null,
          customHeight: entry.customHeight,
          hidden: entry.hidden,
          outlineLevel: entry.outlineLevel,
        }),
        styleXf: styles("getRowXfIndex", position),
      };
    };
    const beforeNeighbours = [stateOf(before), stateOf(after)];
    for (let position = first; position <= last; position += 1) {
      assert.equal(stateOf(position).sized, sized, `setup of ${axis} ${position}`);
    }

    const reset =
      axis === "column"
        ? sessionDimension(id, "column", "resetSize", { sheet: 0, first, last })
        : sessionDimension(id, "row", "resetSize", { sheet: 0, row: first });
    assert.equal(reset.status.ok, true);
    assert.equal(reset.session.dirty, true);

    for (let position = first; position <= last; position += 1) {
      const state = stateOf(position);
      assert.equal(state.sized, false, `${axis} ${position} size cleared`);
      if (axis === "row") {
        assert.equal(state.customHeight, false);
      }
      assert.equal(state.hidden, hidden, `${axis} ${position} hidden`);
      assert.equal(state.outlineLevel, outline ? 2 : 0, `${axis} ${position} outline`);
      assert.equal(state.styleXf, styled ? xf : null, `${axis} ${position} style xf`);
    }
    // The neighbours keep their own size and everything else.
    assert.deepEqual([stateOf(before), stateOf(after)], beforeNeighbours);
    assert.equal(stateOf(before).sized, true);

    // The geometry falls back to the sheet default, the same as an untouched line.
    const untouched = 30;
    const measure = (position) =>
      axis === "column"
        ? sheet("getColumnWidthPt", position, 0)
        : sheet("getRowHeightPt", position);
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
    const columns = () => call("reset-noop", "Sheet", "listColumnOverrides");
    const rows = () => call("reset-noop", "Sheet", "listRowOverrides");
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
    const columns = call("reset-scope", "Sheet", "listColumnOverrides");
    const sizedColumns = columns.flatMap((span) =>
      span.widthChars !== null
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
